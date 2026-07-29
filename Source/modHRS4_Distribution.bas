Attribute VB_Name = "modHRS4_Distribution"
Option Explicit

'============================================================
' 発注まるめシステム Ver4.0 Part3-1
' 配分エンジン
'
' 目的:
' ・使用日プレビュー作業の使用数量上位2行へ発注数を配分
' ・取消行を通常は対象外にする
' ・削除項目など、取消行も対象にする運用に対応
' ・通常表示と納品日集約表示の双方に対応
' ・V4通常表示キャッシュ / V4集約表示キャッシュへ結果を同期
'
' このモジュール単体でコンパイルできます。
' modHRS_Finalとの接続はPart3-4で行います。
'============================================================

Private Const HRS4D_VERSION As String = "Ver4.0.0 Part3-1"

Private Const SH_INPUT As String = "発注入力"
Private Const SH_PREVIEW As String = "使用日プレビュー作業"
Private Const SH_V4_NORMAL As String = "V4通常表示キャッシュ"
Private Const SH_V4_AGGREGATE As String = "V4集約表示キャッシュ"
Private Const SH_V4_WRITEBACK As String = "V4書戻しキャッシュ"
Private Const SH_PERFORMANCE As String = "V4速度ログ"

Private Const ITEM_TOP As Long = 8
Private Const ITEM_BOTTOM As Long = 20

' 使用日プレビュー作業
Private Const PC_CANCEL As Long = 1
Private Const PC_USE_DATE As Long = 2
Private Const PC_MEAL As Long = 3
Private Const PC_USAGE As Long = 4
Private Const PC_DELIVERY As Long = 5
Private Const PC_DISTRIBUTION As Long = 6
Private Const PC_UNIT As Long = 7
Private Const PC_SOURCE_BOOK As Long = 9
Private Const PC_SOURCE_SHEET As Long = 10
Private Const PC_SOURCE_CELL As Long = 11
Private Const PC_CHANGED As Long = 12
Private Const PC_RAW_ROW As Long = 13
Private Const PC_AGG_TARGET_RAW_ROW As Long = 23
Private Const PC_AGG_RAW_ROW_LIST As Long = 24

' V4通常表示キャッシュ
Private Const NC_KEY As Long = 1
Private Const NC_VENDOR_NAME As Long = 3
Private Const NC_PRODUCT_CODE As Long = 4
Private Const NC_PRODUCT_NAME As Long = 5
Private Const NC_DELIVERY As Long = 7
Private Const NC_USE_DATE As Long = 8
Private Const NC_MEAL As Long = 9
Private Const NC_USAGE As Long = 10
Private Const NC_CANCEL As Long = 11
Private Const NC_DISTRIBUTION As Long = 12
Private Const NC_SOURCE_SHEET As Long = 14
Private Const NC_SOURCE_CELL As Long = 17
Private Const NC_RAW_ROW As Long = 20

' V4集約表示キャッシュ
Private Const AC_VENDOR_NAME As Long = 3
Private Const AC_PRODUCT_CODE As Long = 4
Private Const AC_PRODUCT_NAME As Long = 5
Private Const AC_DELIVERY As Long = 7
Private Const AC_USAGE_TOTAL As Long = 11
Private Const AC_CANCEL As Long = 12
Private Const AC_DISTRIBUTION As Long = 13
Private Const AC_TARGET_RAW_ROW As Long = 14
Private Const AC_RAW_ROW_LIST As Long = 15

Private mOldCalculation As XlCalculation
Private mOldScreenUpdating As Boolean
Private mOldEnableEvents As Boolean
Private mFastDepth As Long

'------------------------------------------------------------
' 公開入口
'------------------------------------------------------------

Public Sub HRS4D_ShowVersion()
    MsgBox "配分エンジン " & HRS4D_VERSION, vbInformation, "発注まるめシステム"
End Sub

Public Sub HRS4D_DistributeSelectedRow(ByVal selectedRow As Long, _
                                      Optional ByVal includeCancelled As Boolean = False)

    Dim wsInput As Worksheet
    Dim productName As String
    Dim orderValue As Variant
    Dim orderQty As Double

    If selectedRow < ITEM_TOP Or selectedRow > ITEM_BOTTOM Then Exit Sub
    If Not HRS4D_SheetExists(SH_INPUT) Then Exit Sub

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    productName = Trim$(CStr(wsInput.Cells(selectedRow, "B").value))
    If productName = "" Then Exit Sub

    orderValue = wsInput.Cells(selectedRow, "F").value

    If Trim$(CStr(orderValue)) = "" Then
        orderQty = 0
    ElseIf IsNumeric(orderValue) Then
        orderQty = CDbl(orderValue)
    Else
        MsgBox "発注数には数値を入力してください。", vbExclamation, "発注まるめシステム"
        Exit Sub
    End If

    HRS4D_DistributePreview orderQty, includeCancelled, True

End Sub

Public Sub HRS4D_DistributePreview(ByVal orderQty As Double, _
                                   Optional ByVal includeCancelled As Boolean = False, _
                                   Optional ByVal synchronizeV4Cache As Boolean = True)

    Dim startTime As Double
    Dim wsPreview As Worksheet
    Dim lastRow As Long
    Dim topRow1 As Long
    Dim topRow2 As Long
    Dim eligibleCount As Long
    Dim qty1 As Double
    Dim qty2 As Double

    On Error GoTo ErrHandler

    startTime = Timer
    HRS4D_BeginFast

    If Not HRS4D_SheetExists(SH_PREVIEW) Then
        Err.Raise vbObjectError + 4101, , "使用日プレビュー作業がありません。"
    End If

    Set wsPreview = ThisWorkbook.Worksheets(SH_PREVIEW)
    lastRow = HRS4D_LastRow(wsPreview, PC_USE_DATE)

    HRS4D_ClearPreviewDistribution wsPreview, lastRow

    If orderQty <= 0 Or lastRow < 2 Then GoTo SafeExit

    HRS4D_FindTop2Rows wsPreview, lastRow, includeCancelled, _
                       topRow1, topRow2, eligibleCount

    If eligibleCount = 0 Or topRow1 = 0 Then GoTo SafeExit

    HRS4D_SplitQuantity orderQty, topRow2 > 0, qty1, qty2

    HRS4D_SetPreviewDistribution wsPreview, topRow1, qty1

    If topRow2 > 0 Then
        HRS4D_SetPreviewDistribution wsPreview, topRow2, qty2
    End If

SafeExit:
    If synchronizeV4Cache Then
        HRS4D_SynchronizeV4Caches wsPreview
    End If

    HRS4D_WritePerformance "V4配分", HRS4D_Elapsed(startTime), eligibleCount, _
                           "発注数=" & CStr(orderQty)
    HRS4D_EndFast
    Exit Sub

ErrHandler:
    HRS4D_EndFast
    MsgBox "配分処理でエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, _
           vbExclamation, "発注まるめシステム"
End Sub

Public Sub HRS4D_ClearCurrentDistribution(Optional ByVal synchronizeV4Cache As Boolean = True)

    Dim wsPreview As Worksheet
    Dim lastRow As Long

    On Error GoTo ErrHandler
    HRS4D_BeginFast

    If Not HRS4D_SheetExists(SH_PREVIEW) Then GoTo SafeExit

    Set wsPreview = ThisWorkbook.Worksheets(SH_PREVIEW)
    lastRow = HRS4D_LastRow(wsPreview, PC_USE_DATE)
    HRS4D_ClearPreviewDistribution wsPreview, lastRow

    If synchronizeV4Cache Then
        HRS4D_SynchronizeV4Caches wsPreview
    End If

SafeExit:
    HRS4D_EndFast
    Exit Sub

ErrHandler:
    HRS4D_EndFast
    MsgBox "配分クリアでエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, _
           vbExclamation, "発注まるめシステム"
End Sub

Public Function HRS4D_GetDistributionTotal() As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    If Not HRS4D_SheetExists(SH_PREVIEW) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SH_PREVIEW)
    lastRow = HRS4D_LastRow(ws, PC_USE_DATE)

    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, PC_DISTRIBUTION).value) Then
            HRS4D_GetDistributionTotal = HRS4D_GetDistributionTotal + _
                                        CDbl(ws.Cells(r, PC_DISTRIBUTION).value)
        End If
    Next r

End Function

Public Function HRS4D_ValidateDistribution(ByVal expectedQty As Double, _
                                           Optional ByVal showMessage As Boolean = True) As Boolean

    Dim actualQty As Double
    Dim difference As Double

    actualQty = HRS4D_GetDistributionTotal()
    difference = Abs(actualQty - expectedQty)
    HRS4D_ValidateDistribution = (difference < 0.0000001)

    If showMessage Then
        If HRS4D_ValidateDistribution Then
            MsgBox "配分合計は発注数と一致しています。" & vbCrLf & _
                   "配分合計: " & HRS4D_NumberText(actualQty), _
                   vbInformation, "発注まるめシステム"
        Else
            MsgBox "配分合計が発注数と一致していません。" & vbCrLf & _
                   "発注数: " & HRS4D_NumberText(expectedQty) & vbCrLf & _
                   "配分合計: " & HRS4D_NumberText(actualQty), _
                   vbExclamation, "発注まるめシステム"
        End If
    End If

End Function

'------------------------------------------------------------
' 配分計算
'------------------------------------------------------------

Private Sub HRS4D_FindTop2Rows(ByVal ws As Worksheet, _
                              ByVal lastRow As Long, _
                              ByVal includeCancelled As Boolean, _
                              ByRef topRow1 As Long, _
                              ByRef topRow2 As Long, _
                              ByRef eligibleCount As Long)

    Dim r As Long
    Dim usageQty As Double
    Dim maxUsage1 As Double
    Dim maxUsage2 As Double
    Dim isEligible As Boolean

    topRow1 = 0
    topRow2 = 0
    eligibleCount = 0
    maxUsage1 = -1E+308
    maxUsage2 = -1E+308

    For r = 2 To lastRow

        isEligible = HRS4D_IsEligibleRow(ws, r, includeCancelled)

        If isEligible Then
            eligibleCount = eligibleCount + 1
            usageQty = HRS4D_NumberOrZero(ws.Cells(r, PC_USAGE).value)

            If topRow1 = 0 Or usageQty > maxUsage1 Then
                maxUsage2 = maxUsage1
                topRow2 = topRow1
                maxUsage1 = usageQty
                topRow1 = r
            ElseIf topRow2 = 0 Or usageQty > maxUsage2 Then
                maxUsage2 = usageQty
                topRow2 = r
            End If
        End If
    Next r

End Sub

Private Function HRS4D_IsEligibleRow(ByVal ws As Worksheet, _
                                     ByVal rowNumber As Long, _
                                     ByVal includeCancelled As Boolean) As Boolean

    If Trim$(CStr(ws.Cells(rowNumber, PC_USE_DATE).value)) = "" Then Exit Function

    If includeCancelled Then
        HRS4D_IsEligibleRow = True
    Else
        HRS4D_IsEligibleRow = Not HRS4D_IsCancelled(ws.Cells(rowNumber, PC_CANCEL).value)
    End If

End Function

Private Sub HRS4D_SplitQuantity(ByVal orderQty As Double, _
                               ByVal hasSecondRow As Boolean, _
                               ByRef firstRowQty As Double, _
                               ByRef secondRowQty As Double)

    If Not hasSecondRow Then
        firstRowQty = orderQty
        secondRowQty = 0
        Exit Sub
    End If

    If HRS4D_IsWholeNumber(orderQty) Then
        secondRowQty = Fix(orderQty / 2)
        firstRowQty = orderQty - secondRowQty
    Else
        firstRowQty = orderQty / 2
        secondRowQty = orderQty - firstRowQty
    End If

End Sub

Private Sub HRS4D_ClearPreviewDistribution(ByVal ws As Worksheet, _
                                           ByVal lastRow As Long)

    Dim clearLastRow As Long

    clearLastRow = lastRow
    If clearLastRow < 2 Then clearLastRow = 2

    ws.Range(ws.Cells(2, PC_DISTRIBUTION), _
             ws.Cells(clearLastRow, PC_DISTRIBUTION)).ClearContents

    ws.Range(ws.Cells(2, PC_CHANGED), _
             ws.Cells(clearLastRow, PC_CHANGED)).value = "FALSE"

End Sub

Private Sub HRS4D_SetPreviewDistribution(ByVal ws As Worksheet, _
                                         ByVal rowNumber As Long, _
                                         ByVal distributionQty As Double)

    If HRS4D_IsWholeNumber(distributionQty) Then
        ws.Cells(rowNumber, PC_DISTRIBUTION).value = CLng(distributionQty)
    Else
        ws.Cells(rowNumber, PC_DISTRIBUTION).value = distributionQty
    End If

    ws.Cells(rowNumber, PC_CHANGED).value = "TRUE"

End Sub

'------------------------------------------------------------
' V4キャッシュ同期
'------------------------------------------------------------

Private Sub HRS4D_SynchronizeV4Caches(ByVal wsPreview As Worksheet)

    Dim vendorName As String
    Dim productCode As String
    Dim productName As String

    vendorName = HRS4D_CurrentVendorName()
    productCode = Trim$(CStr(wsPreview.Range("N3").value))
    productName = Trim$(CStr(wsPreview.Range("N4").value))

    If productName = "" Then Exit Sub

    HRS4D_UpdateNormalCache wsPreview, vendorName, productCode, productName
    HRS4D_UpdateAggregateCache wsPreview, vendorName, productCode, productName
    HRS4D_UpdateWriteBackCache wsPreview, vendorName, productCode, productName

End Sub

Private Sub HRS4D_UpdateNormalCache(ByVal wsPreview As Worksheet, _
                                    ByVal vendorName As String, _
                                    ByVal productCode As String, _
                                    ByVal productName As String)

    Dim wsCache As Worksheet
    Dim previewLastRow As Long
    Dim cacheLastRow As Long
    Dim p As Long
    Dim cacheRow As Long
    Dim rawRow As Long
    Dim distributionValue As Variant
    Dim cancelValue As Boolean

    If Not HRS4D_SheetExists(SH_V4_NORMAL) Then Exit Sub

    Set wsCache = ThisWorkbook.Worksheets(SH_V4_NORMAL)
    previewLastRow = HRS4D_LastRow(wsPreview, PC_USE_DATE)
    cacheLastRow = HRS4D_LastRow(wsCache, NC_KEY)

    For p = 2 To previewLastRow
        rawRow = CLng(HRS4D_NumberOrZero(wsPreview.Cells(p, PC_RAW_ROW).value))
        distributionValue = wsPreview.Cells(p, PC_DISTRIBUTION).value
        cancelValue = HRS4D_IsCancelled(wsPreview.Cells(p, PC_CANCEL).value)

        cacheRow = HRS4D_FindNormalCacheRow(wsCache, cacheLastRow, _
                                            vendorName, productCode, productName, _
                                            rawRow, wsPreview.Cells(p, PC_SOURCE_SHEET).value, _
                                            wsPreview.Cells(p, PC_SOURCE_CELL).value)

        If cacheRow > 0 Then
            wsCache.Cells(cacheRow, NC_CANCEL).value = cancelValue
            wsCache.Cells(cacheRow, NC_DISTRIBUTION).value = distributionValue
        End If
    Next p

End Sub

Private Function HRS4D_FindNormalCacheRow(ByVal ws As Worksheet, _
                                          ByVal lastRow As Long, _
                                          ByVal vendorName As String, _
                                          ByVal productCode As String, _
                                          ByVal productName As String, _
                                          ByVal rawRow As Long, _
                                          ByVal sourceSheet As Variant, _
                                          ByVal sourceCell As Variant) As Long

    Dim r As Long

    For r = 2 To lastRow
        If HRS4D_ProductMatches(ws, r, NC_VENDOR_NAME, NC_PRODUCT_CODE, _
                                NC_PRODUCT_NAME, vendorName, productCode, productName) Then

            If rawRow > 0 Then
                If CLng(HRS4D_NumberOrZero(ws.Cells(r, NC_RAW_ROW).value)) = rawRow Then
                    HRS4D_FindNormalCacheRow = r
                    Exit Function
                End If
            ElseIf StrComp(Trim$(CStr(ws.Cells(r, NC_SOURCE_SHEET).value)), _
                           Trim$(CStr(sourceSheet)), vbTextCompare) = 0 And _
                   StrComp(Trim$(CStr(ws.Cells(r, NC_SOURCE_CELL).value)), _
                           Trim$(CStr(sourceCell)), vbTextCompare) = 0 Then
                HRS4D_FindNormalCacheRow = r
                Exit Function
            End If
        End If
    Next r

End Function

Private Sub HRS4D_UpdateAggregateCache(ByVal wsPreview As Worksheet, _
                                       ByVal vendorName As String, _
                                       ByVal productCode As String, _
                                       ByVal productName As String)

    Dim wsCache As Worksheet
    Dim previewLastRow As Long
    Dim cacheLastRow As Long
    Dim p As Long
    Dim r As Long
    Dim targetRawRow As Long
    Dim rowList As String

    If Not HRS4D_SheetExists(SH_V4_AGGREGATE) Then Exit Sub

    Set wsCache = ThisWorkbook.Worksheets(SH_V4_AGGREGATE)
    previewLastRow = HRS4D_LastRow(wsPreview, PC_USE_DATE)
    cacheLastRow = HRS4D_LastRow(wsCache, 1)

    For p = 2 To previewLastRow
        targetRawRow = CLng(HRS4D_NumberOrZero( _
                       wsPreview.Cells(p, PC_AGG_TARGET_RAW_ROW).value))
        rowList = Trim$(CStr(wsPreview.Cells(p, PC_AGG_RAW_ROW_LIST).value))

        If targetRawRow > 0 Or rowList <> "" Then
            For r = 2 To cacheLastRow
                If HRS4D_ProductMatches(wsCache, r, AC_VENDOR_NAME, AC_PRODUCT_CODE, _
                                        AC_PRODUCT_NAME, vendorName, productCode, productName) Then
                    If HRS4D_AggregateRowMatches(wsCache, r, targetRawRow, rowList) Then
                        wsCache.Cells(r, AC_CANCEL).value = _
                            HRS4D_IsCancelled(wsPreview.Cells(p, PC_CANCEL).value)
                        wsCache.Cells(r, AC_DISTRIBUTION).value = _
                            wsPreview.Cells(p, PC_DISTRIBUTION).value
                        Exit For
                    End If
                End If
            Next r
        End If
    Next p

End Sub

Private Function HRS4D_AggregateRowMatches(ByVal ws As Worksheet, _
                                           ByVal rowNumber As Long, _
                                           ByVal targetRawRow As Long, _
                                           ByVal rowList As String) As Boolean

    If targetRawRow > 0 Then
        If CLng(HRS4D_NumberOrZero(ws.Cells(rowNumber, AC_TARGET_RAW_ROW).value)) = _
           targetRawRow Then
            HRS4D_AggregateRowMatches = True
            Exit Function
        End If
    End If

    If rowList <> "" Then
        HRS4D_AggregateRowMatches = _
            (StrComp(HRS4D_NormalizeRowList(ws.Cells(rowNumber, AC_RAW_ROW_LIST).value), _
                     HRS4D_NormalizeRowList(rowList), vbTextCompare) = 0)
    End If

End Function

Private Sub HRS4D_UpdateWriteBackCache(ByVal wsPreview As Worksheet, _
                                       ByVal vendorName As String, _
                                       ByVal productCode As String, _
                                       ByVal productName As String)

    Dim wsWrite As Worksheet
    Dim lastWriteRow As Long
    Dim previewLastRow As Long
    Dim p As Long
    Dim r As Long
    Dim rawRow As Long
    Dim writeRawRow As Long
    Dim newValue As Variant

    If Not HRS4D_SheetExists(SH_V4_WRITEBACK) Then Exit Sub

    Set wsWrite = ThisWorkbook.Worksheets(SH_V4_WRITEBACK)
    lastWriteRow = HRS4D_LastRow(wsWrite, 1)
    previewLastRow = HRS4D_LastRow(wsPreview, PC_USE_DATE)

    For p = 2 To previewLastRow
        rawRow = CLng(HRS4D_NumberOrZero(wsPreview.Cells(p, PC_RAW_ROW).value))
        If rawRow <= 0 Then
            rawRow = CLng(HRS4D_NumberOrZero( _
                     wsPreview.Cells(p, PC_AGG_TARGET_RAW_ROW).value))
        End If

        If rawRow > 0 Then
            newValue = wsPreview.Cells(p, PC_DISTRIBUTION).value

            For r = 2 To lastWriteRow
                If StrComp(Trim$(CStr(wsWrite.Cells(r, 2).value)), _
                           vendorName, vbTextCompare) = 0 And _
                   HRS4D_SameProduct(wsWrite.Cells(r, 3).value, _
                                     wsWrite.Cells(r, 4).value, _
                                     productCode, productName) Then

                    writeRawRow = CLng(HRS4D_NumberOrZero(wsWrite.Cells(r, 8).value))
                    If writeRawRow = rawRow Then
                        wsWrite.Cells(r, 14).value = newValue
                        wsWrite.Cells(r, 15).value = True
                        Exit For
                    End If
                End If
            Next r
        End If
    Next p

End Sub

'------------------------------------------------------------
' 共通補助
'------------------------------------------------------------

Private Function HRS4D_CurrentVendorName() As String
    If HRS4D_SheetExists(SH_INPUT) Then
        HRS4D_CurrentVendorName = Trim$(CStr( _
            ThisWorkbook.Worksheets(SH_INPUT).Range("B3").value))
    End If
End Function

Private Function HRS4D_ProductMatches(ByVal ws As Worksheet, _
                                      ByVal rowNumber As Long, _
                                      ByVal vendorColumn As Long, _
                                      ByVal codeColumn As Long, _
                                      ByVal nameColumn As Long, _
                                      ByVal vendorName As String, _
                                      ByVal productCode As String, _
                                      ByVal productName As String) As Boolean

    If StrComp(Trim$(CStr(ws.Cells(rowNumber, vendorColumn).value)), _
               Trim$(vendorName), vbTextCompare) <> 0 Then Exit Function

    HRS4D_ProductMatches = HRS4D_SameProduct( _
        ws.Cells(rowNumber, codeColumn).value, _
        ws.Cells(rowNumber, nameColumn).value, _
        productCode, productName)

End Function

Private Function HRS4D_SameProduct(ByVal rowCode As Variant, _
                                   ByVal rowName As Variant, _
                                   ByVal productCode As String, _
                                   ByVal productName As String) As Boolean

    Dim normalizedRowCode As String
    Dim normalizedCode As String

    normalizedRowCode = HRS4D_NormalizeCode(rowCode)
    normalizedCode = HRS4D_NormalizeCode(productCode)

    If normalizedCode <> "" And normalizedRowCode <> "" Then
        HRS4D_SameProduct = _
            (StrComp(normalizedRowCode, normalizedCode, vbTextCompare) = 0)
    Else
        HRS4D_SameProduct = _
            (StrComp(Trim$(CStr(rowName)), Trim$(productName), vbTextCompare) = 0)
    End If

End Function

Private Function HRS4D_NormalizeCode(ByVal value As Variant) As String

    Dim textValue As String

    textValue = Trim$(CStr(value))
    textValue = Replace$(textValue, " ", "")
    textValue = Replace$(textValue, "　", "")

    If textValue <> "" And IsNumeric(textValue) Then
        On Error Resume Next
        textValue = Format$(CDbl(textValue), "0")
        On Error GoTo 0
    End If

    HRS4D_NormalizeCode = textValue

End Function

Private Function HRS4D_NormalizeRowList(ByVal value As Variant) As String

    Dim sourceText As String
    Dim parts() As String
    Dim numbers() As Long
    Dim i As Long
    Dim j As Long
    Dim count As Long
    Dim temp As Long
    Dim resultText As String

    sourceText = Trim$(CStr(value))
    sourceText = Replace$(sourceText, " ", "")
    sourceText = Replace$(sourceText, "、", ",")
    sourceText = Replace$(sourceText, ";", ",")

    If sourceText = "" Then Exit Function

    parts = Split(sourceText, ",")
    ReDim numbers(0 To UBound(parts))

    For i = LBound(parts) To UBound(parts)
        If IsNumeric(parts(i)) Then
            numbers(count) = CLng(Val(parts(i)))
            count = count + 1
        End If
    Next i

    If count = 0 Then Exit Function

    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If numbers(j) < numbers(i) Then
                temp = numbers(i)
                numbers(i) = numbers(j)
                numbers(j) = temp
            End If
        Next j
    Next i

    For i = 0 To count - 1
        If resultText <> "" Then resultText = resultText & ","
        resultText = resultText & CStr(numbers(i))
    Next i

    HRS4D_NormalizeRowList = resultText

End Function

Private Function HRS4D_IsCancelled(ByVal value As Variant) As Boolean

    Select Case UCase$(Trim$(CStr(value)))
        Case "■", "TRUE", "1", "YES", "取消"
            HRS4D_IsCancelled = True
    End Select

End Function

Private Function HRS4D_NumberOrZero(ByVal value As Variant) As Double
    If IsNumeric(value) Then HRS4D_NumberOrZero = CDbl(value)
End Function

Private Function HRS4D_IsWholeNumber(ByVal value As Double) As Boolean
    HRS4D_IsWholeNumber = (Abs(value - Fix(value)) < 0.0000001)
End Function

Private Function HRS4D_NumberText(ByVal value As Double) As String
    If HRS4D_IsWholeNumber(value) Then
        HRS4D_NumberText = CStr(CLng(value))
    Else
        HRS4D_NumberText = CStr(value)
    End If
End Function

Private Function HRS4D_LastRow(ByVal ws As Worksheet, _
                               ByVal columnNumber As Long) As Long
    HRS4D_LastRow = ws.Cells(ws.Rows.count, columnNumber).End(xlUp).Row
End Function

Private Function HRS4D_SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    HRS4D_SheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0

End Function

Private Sub HRS4D_BeginFast()

    If mFastDepth = 0 Then
        mOldCalculation = Application.Calculation
        mOldScreenUpdating = Application.ScreenUpdating
        mOldEnableEvents = Application.EnableEvents

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    End If

    mFastDepth = mFastDepth + 1

End Sub

Private Sub HRS4D_EndFast()

    If mFastDepth <= 0 Then Exit Sub

    mFastDepth = mFastDepth - 1

    If mFastDepth = 0 Then
        Application.Calculation = mOldCalculation
        Application.EnableEvents = mOldEnableEvents
        Application.ScreenUpdating = mOldScreenUpdating
    End If

End Sub

Private Function HRS4D_Elapsed(ByVal startTimer As Double) As Double
    If Timer >= startTimer Then
        HRS4D_Elapsed = Timer - startTimer
    Else
        HRS4D_Elapsed = (86400# - startTimer) + Timer
    End If
End Function

Private Sub HRS4D_WritePerformance(ByVal processName As String, _
                                   ByVal elapsedSeconds As Double, _
                                   ByVal itemCount As Long, _
                                   ByVal noteText As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    If Not HRS4D_SheetExists(SH_PERFORMANCE) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SH_PERFORMANCE)
    nextRow = HRS4D_LastRow(ws, 1) + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = processName
    ws.Cells(nextRow, 3).value = elapsedSeconds
    ws.Cells(nextRow, 4).value = itemCount
    ws.Cells(nextRow, 5).value = noteText

End Sub
