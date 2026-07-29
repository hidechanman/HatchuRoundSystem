Attribute VB_Name = "modHRS4_Session"
Option Explicit

'============================================================
' 発注まるめシステム Ver4.0 Part3-2
' セッション管理エンジン
'
' 目的:
' ・商品切替後も取消、使用数量、配分値を保持する
' ・V4セッションキャッシュをDictionaryへ読み込む
' ・Dictionaryの内容をシートへ一括保存する
' ・使用日プレビュー作業との保存・復元を行う
'
' 参照設定不要（Scripting.Dictionaryは遅延バインディング）
'============================================================

Private Const HRS4S_VERSION As String = "Ver4.0.0 Part3-2"

Private Const SH_INPUT As String = "発注入力"
Private Const SH_PREVIEW As String = "使用日プレビュー作業"
Private Const SH_SESSION As String = "V4セッションキャッシュ"
Private Const SH_PERFORMANCE As String = "V4速度ログ"

Private Const SESSION_COLS As Long = 18
Private Const PREVIEW_FIRST_ROW As Long = 2

' V4セッションキャッシュ列
Private Const SC_KEY As Long = 1
Private Const SC_VENDOR As Long = 2
Private Const SC_PRODUCT_CODE As Long = 3
Private Const SC_PRODUCT_NAME As Long = 4
Private Const SC_DELIVERY As Long = 5
Private Const SC_USE_DATE As Long = 6
Private Const SC_MEAL As Long = 7
Private Const SC_USAGE As Long = 8
Private Const SC_CANCEL As Long = 9
Private Const SC_DISTRIBUTION As Long = 10
Private Const SC_UNIT As Long = 11
Private Const SC_SOURCE_BOOK As Long = 12
Private Const SC_SOURCE_SHEET As Long = 13
Private Const SC_SOURCE_CELL As Long = 14
Private Const SC_RAW_ROW As Long = 15
Private Const SC_CHANGED As Long = 16
Private Const SC_UPDATED_AT As Long = 17
Private Const SC_DISPLAY_MODE As Long = 18

' 使用日プレビュー作業列
Private Const PC_CANCEL As Long = 1
Private Const PC_USE_DATE As Long = 2
Private Const PC_MEAL As Long = 3
Private Const PC_USAGE As Long = 4
Private Const PC_DELIVERY As Long = 5
Private Const PC_DISTRIBUTION As Long = 6
Private Const PC_UNIT As Long = 7
Private Const PC_PRODUCT_NAME As Long = 8
Private Const PC_SOURCE_SHEET As Long = 9
Private Const PC_SOURCE_CELL As Long = 10
Private Const PC_SOURCE_BOOK As Long = 11
Private Const PC_CHANGED As Long = 12
Private Const PC_RAW_ROW As Long = 13

Private mSession As Object
Private mLoaded As Boolean
Private mDirty As Boolean

Private mFastDepth As Long
Private mOldCalculation As XlCalculation
Private mOldScreenUpdating As Boolean
Private mOldEnableEvents As Boolean

'------------------------------------------------------------
' 公開入口
'------------------------------------------------------------

Public Sub HRS4_InitSession()
    HRS4S_Initialize True
End Sub

Public Sub HRS4_SaveSession( _
    Optional ByVal orderQty As Variant, _
    Optional ByVal confirmed As Boolean = False)

    HRS4S_SaveCurrentPreview orderQty, confirmed, True
End Sub

Public Sub HRS4_LoadSessionCache()
    HRS4S_LoadCache
End Sub

Public Sub HRS4_SaveSessionCache()
    HRS4S_FlushCache
End Sub

Public Function HRS4_RestoreProductSession() As Long
    HRS4_RestoreProductSession = HRS4S_RestoreCurrentPreview(True)
End Function

Public Function HRS4_GetSessionCount() As Long
    HRS4_GetSessionCount = HRS4S_GetCount()
End Function

Public Sub HRS4_ClearSession()
    HRS4S_ClearAll True
End Sub

Public Sub HRS4S_ShowVersion()
    MsgBox "セッション管理エンジン " & HRS4S_VERSION, _
           vbInformation, "発注まるめシステム"
End Sub

Public Sub HRS4S_Initialize(Optional ByVal loadFromSheet As Boolean = True)

    Set mSession = CreateObject("Scripting.Dictionary")
    mSession.CompareMode = vbTextCompare
    mLoaded = True
    mDirty = False

    HRS4S_EnsureSessionSheet

    If loadFromSheet Then
        HRS4S_LoadCache
    End If

End Sub

Public Function HRS4S_GetCount() As Long
    HRS4S_EnsureInitialized
    HRS4S_GetCount = mSession.count
End Function

Public Function HRS4S_IsDirty() As Boolean
    HRS4S_IsDirty = mDirty
End Function

Public Sub HRS4S_ClearAll(Optional ByVal flushToSheet As Boolean = True)

    HRS4S_EnsureInitialized
    mSession.RemoveAll
    mDirty = True

    If flushToSheet Then
        HRS4S_FlushCache
    End If

End Sub

Public Sub HRS4S_SaveRecord( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal deliveryDate As Variant, _
    ByVal useDate As Variant, _
    ByVal mealText As String, _
    ByVal usageQty As Variant, _
    ByVal isCancelled As Boolean, _
    ByVal distributionQty As Variant, _
    ByVal unitName As String, _
    ByVal sourceBook As String, _
    ByVal sourceSheet As String, _
    ByVal sourceCell As String, _
    ByVal rawRow As Long, _
    Optional ByVal isChanged As Boolean = True, _
    Optional ByVal displayMode As String = "通常")

    Dim keyText As String
    Dim recordData As Variant

    HRS4S_EnsureInitialized

    keyText = HRS4S_MakeKey(vendorName, productCode, productName, _
                            deliveryDate, useDate, mealText, _
                            sourceSheet, sourceCell, rawRow)

    recordData = Array( _
        keyText, Trim$(vendorName), HRS4S_NormalizeCode(productCode), _
        Trim$(productName), HRS4S_NormalizeDateValue(deliveryDate), _
        HRS4S_NormalizeDateValue(useDate), Trim$(mealText), _
        HRS4S_ValueOrBlank(usageQty), isCancelled, _
        HRS4S_ValueOrBlank(distributionQty), Trim$(unitName), _
        Trim$(sourceBook), Trim$(sourceSheet), Trim$(sourceCell), _
        rawRow, isChanged, Now, Trim$(displayMode))

    mSession(keyText) = recordData
    mDirty = True

End Sub

Public Function HRS4S_TryGetRecord( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal deliveryDate As Variant, _
    ByVal useDate As Variant, _
    ByVal mealText As String, _
    ByVal sourceSheet As String, _
    ByVal sourceCell As String, _
    ByVal rawRow As Long, _
    ByRef recordData As Variant) As Boolean

    Dim keyText As String

    HRS4S_EnsureInitialized

    keyText = HRS4S_MakeKey(vendorName, productCode, productName, _
                            deliveryDate, useDate, mealText, _
                            sourceSheet, sourceCell, rawRow)

    If mSession.Exists(keyText) Then
        recordData = mSession(keyText)
        HRS4S_TryGetRecord = True
    End If

End Function

Public Function HRS4S_DeleteRecord( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal deliveryDate As Variant, _
    ByVal useDate As Variant, _
    ByVal mealText As String, _
    ByVal sourceSheet As String, _
    ByVal sourceCell As String, _
    ByVal rawRow As Long) As Boolean

    Dim keyText As String

    HRS4S_EnsureInitialized

    keyText = HRS4S_MakeKey(vendorName, productCode, productName, _
                            deliveryDate, useDate, mealText, _
                            sourceSheet, sourceCell, rawRow)

    If mSession.Exists(keyText) Then
        mSession.Remove keyText
        mDirty = True
        HRS4S_DeleteRecord = True
    End If

End Function

Public Function HRS4S_DeleteProduct( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String, _
    Optional ByVal flushToSheet As Boolean = False) As Long

    Dim keys As Variant
    Dim keyItem As Variant
    Dim recordData As Variant

    HRS4S_EnsureInitialized
    If mSession.count = 0 Then Exit Function

    keys = mSession.keys

    For Each keyItem In keys
        recordData = mSession(CStr(keyItem))

        If HRS4S_ProductMatches(recordData, vendorName, _
                                productCode, productName) Then
            mSession.Remove CStr(keyItem)
            HRS4S_DeleteProduct = HRS4S_DeleteProduct + 1
        End If
    Next keyItem

    If HRS4S_DeleteProduct > 0 Then mDirty = True
    If flushToSheet And mDirty Then HRS4S_FlushCache

End Function

'------------------------------------------------------------
' キャッシュシート入出力
'------------------------------------------------------------

Public Sub HRS4S_LoadCache()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim data As Variant
    Dim r As Long
    Dim keyText As String
    Dim recordData As Variant

    On Error GoTo ErrHandler

    HRS4S_EnsureInitialized
    Set ws = HRS4S_EnsureSessionSheet()
    lastRow = HRS4S_LastRow(ws)

    mSession.RemoveAll

    If lastRow < 2 Then
        mDirty = False
        Exit Sub
    End If

    data = ws.Range(ws.Cells(2, 1), _
                    ws.Cells(lastRow, SESSION_COLS)).Value2

    For r = 1 To UBound(data, 1)
        keyText = Trim$(CStr(data(r, SC_KEY)))

        If keyText = "" Then
            keyText = HRS4S_MakeKey( _
                CStr(data(r, SC_VENDOR)), _
                CStr(data(r, SC_PRODUCT_CODE)), _
                CStr(data(r, SC_PRODUCT_NAME)), _
                data(r, SC_DELIVERY), data(r, SC_USE_DATE), _
                CStr(data(r, SC_MEAL)), _
                CStr(data(r, SC_SOURCE_SHEET)), _
                CStr(data(r, SC_SOURCE_CELL)), _
                CLng(Val(data(r, SC_RAW_ROW))))
        End If

        recordData = HRS4S_RowToRecord(data, r, keyText)
        mSession(keyText) = recordData
    Next r

    mDirty = False
    Exit Sub

ErrHandler:
    MsgBox "セッションキャッシュの読込でエラーが発生しました。" & _
           vbCrLf & Err.Number & " : " & Err.Description, _
           vbExclamation, "発注まるめシステム"
End Sub

Public Sub HRS4S_FlushCache()

    Dim startedAt As Double
    Dim ws As Worksheet
    Dim outputData() As Variant
    Dim keys As Variant
    Dim keyItem As Variant
    Dim recordData As Variant
    Dim r As Long
    Dim c As Long

    On Error GoTo ErrHandler

    startedAt = Timer
    HRS4S_BeginFast
    HRS4S_EnsureInitialized
    Set ws = HRS4S_EnsureSessionSheet()

    If ws.Rows.count > 1 Then
        ws.Range(ws.Cells(2, 1), _
                 ws.Cells(ws.Rows.count, SESSION_COLS)).ClearContents
    End If

    If mSession.count > 0 Then
        ReDim outputData(1 To mSession.count, 1 To SESSION_COLS)
        keys = mSession.keys

        For Each keyItem In keys
            r = r + 1
            recordData = mSession(CStr(keyItem))

            For c = 1 To SESSION_COLS
                outputData(r, c) = recordData(c - 1)
            Next c
        Next keyItem

        ws.Cells(2, 1).Resize(mSession.count, SESSION_COLS).Value2 = outputData
    End If

    ws.Columns(SC_PRODUCT_CODE).NumberFormat = "@"
    ws.Columns(SC_UPDATED_AT).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Visible = xlSheetVeryHidden

    mDirty = False
    HRS4S_WritePerformance "V4セッション保存", _
                           HRS4S_Elapsed(startedAt), _
                           mSession.count, "Dictionary→Sheet"
    HRS4S_EndFast
    Exit Sub

ErrHandler:
    HRS4S_EndFast
    MsgBox "セッションキャッシュの保存でエラーが発生しました。" & _
           vbCrLf & Err.Number & " : " & Err.Description, _
           vbExclamation, "発注まるめシステム"
End Sub

'------------------------------------------------------------
' 使用日プレビュー作業との同期
'------------------------------------------------------------

Public Sub HRS4S_SaveCurrentPreview( _
    Optional ByVal orderQty As Variant, _
    Optional ByVal confirmed As Boolean = False, _
    Optional ByVal flushToSheet As Boolean = True)

    Dim startedAt As Double
    Dim wsInput As Worksheet
    Dim wsPreview As Worksheet
    Dim vendorName As String
    Dim productCode As String
    Dim productName As String
    Dim displayMode As String
    Dim lastRow As Long
    Dim r As Long
    Dim usageValue As Variant
    Dim distributionValue As Variant
    Dim isCancelled As Boolean

    On Error GoTo ErrHandler

    startedAt = Timer
    HRS4S_BeginFast
    HRS4S_EnsureInitialized

    If Not HRS4S_SheetExists(SH_PREVIEW) Then GoTo SafeExit
    If Not HRS4S_SheetExists(SH_INPUT) Then GoTo SafeExit

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsPreview = ThisWorkbook.Worksheets(SH_PREVIEW)

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    productCode = HRS4S_NormalizeCode(wsPreview.Range("N3").value)
    productName = Trim$(CStr(wsPreview.Range("N4").value))
    displayMode = IIf(HRS4S_ToBoolean(wsPreview.Range("N7").value), _
                      "集約", "通常")

    If productName = "" Then GoTo SafeExit

    HRS4S_DeleteProduct vendorName, productCode, productName, False
    lastRow = HRS4S_LastPreviewRow(wsPreview)

    For r = PREVIEW_FIRST_ROW To lastRow
        If Trim$(CStr(wsPreview.Cells(r, PC_USE_DATE).value)) <> "" Then
            usageValue = wsPreview.Cells(r, PC_USAGE).value
            distributionValue = wsPreview.Cells(r, PC_DISTRIBUTION).value
            isCancelled = HRS4S_IsCancelledMark( _
                wsPreview.Cells(r, PC_CANCEL).value)

            HRS4S_SaveRecord vendorName, productCode, productName, _
                wsPreview.Cells(r, PC_DELIVERY).value, _
                wsPreview.Cells(r, PC_USE_DATE).value, _
                CStr(wsPreview.Cells(r, PC_MEAL).value), _
                usageValue, isCancelled, distributionValue, _
                CStr(wsPreview.Cells(r, PC_UNIT).value), _
                CStr(wsPreview.Cells(r, PC_SOURCE_BOOK).value), _
                CStr(wsPreview.Cells(r, PC_SOURCE_SHEET).value), _
                CStr(wsPreview.Cells(r, PC_SOURCE_CELL).value), _
                CLng(Val(wsPreview.Cells(r, PC_RAW_ROW).value)), _
                HRS4S_ToBoolean(wsPreview.Cells(r, PC_CHANGED).value), _
                displayMode
        End If
    Next r

    If flushToSheet Then HRS4S_FlushCache

    HRS4S_WritePerformance "V4プレビュー保存", _
                           HRS4S_Elapsed(startedAt), _
                           HRS4S_CountProduct(vendorName, _
                           productCode, productName), _
                           "確認=" & CStr(confirmed) & _
                           ", 発注数=" & HRS4S_ValueText(orderQty)

SafeExit:
    HRS4S_EndFast
    Exit Sub

ErrHandler:
    HRS4S_EndFast
    MsgBox "現在商品のセッション保存でエラーが発生しました。" & _
           vbCrLf & Err.Number & " : " & Err.Description, _
           vbExclamation, "発注まるめシステム"
End Sub

Public Function HRS4S_RestoreCurrentPreview( _
    Optional ByVal recalculateAggregate As Boolean = True) As Long

    Dim startedAt As Double
    Dim wsInput As Worksheet
    Dim wsPreview As Worksheet
    Dim vendorName As String
    Dim productCode As String
    Dim productName As String
    Dim lastRow As Long
    Dim r As Long
    Dim recordData As Variant

    On Error GoTo ErrHandler

    startedAt = Timer
    HRS4S_BeginFast
    HRS4S_EnsureInitialized

    If Not HRS4S_SheetExists(SH_PREVIEW) Then GoTo SafeExit
    If Not HRS4S_SheetExists(SH_INPUT) Then GoTo SafeExit

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsPreview = ThisWorkbook.Worksheets(SH_PREVIEW)

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    productCode = HRS4S_NormalizeCode(wsPreview.Range("N3").value)
    productName = Trim$(CStr(wsPreview.Range("N4").value))
    lastRow = HRS4S_LastPreviewRow(wsPreview)

    For r = PREVIEW_FIRST_ROW To lastRow
        If HRS4S_TryGetRecord(vendorName, productCode, productName, _
             wsPreview.Cells(r, PC_DELIVERY).value, _
             wsPreview.Cells(r, PC_USE_DATE).value, _
             CStr(wsPreview.Cells(r, PC_MEAL).value), _
             CStr(wsPreview.Cells(r, PC_SOURCE_SHEET).value), _
             CStr(wsPreview.Cells(r, PC_SOURCE_CELL).value), _
             CLng(Val(wsPreview.Cells(r, PC_RAW_ROW).value)), _
             recordData) Then

            wsPreview.Cells(r, PC_CANCEL).value = _
                IIf(CBool(recordData(SC_CANCEL - 1)), "■", "□")
            wsPreview.Cells(r, PC_USAGE).value = _
                recordData(SC_USAGE - 1)
            wsPreview.Cells(r, PC_DISTRIBUTION).value = _
                recordData(SC_DISTRIBUTION - 1)
            wsPreview.Cells(r, PC_UNIT).value = _
                recordData(SC_UNIT - 1)
            wsPreview.Cells(r, PC_CHANGED).value = "TRUE"

            HRS4S_RestoreCurrentPreview = _
                HRS4S_RestoreCurrentPreview + 1
        End If
    Next r

    If recalculateAggregate Then
        HRS4S_RecalculateAggregate wsPreview
    End If

    HRS4S_WritePerformance "V4プレビュー復元", _
                           HRS4S_Elapsed(startedAt), _
                           HRS4S_RestoreCurrentPreview, _
                           vendorName & " / " & productName

SafeExit:
    HRS4S_EndFast
    Exit Function

ErrHandler:
    HRS4S_EndFast
    MsgBox "現在商品のセッション復元でエラーが発生しました。" & _
           vbCrLf & Err.Number & " : " & Err.Description, _
           vbExclamation, "発注まるめシステム"
End Function

Public Function HRS4S_CountProduct( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String) As Long

    Dim keyItem As Variant
    Dim recordData As Variant

    HRS4S_EnsureInitialized

    For Each keyItem In mSession.keys
        recordData = mSession(CStr(keyItem))

        If HRS4S_ProductMatches(recordData, vendorName, _
                                productCode, productName) Then
            HRS4S_CountProduct = HRS4S_CountProduct + 1
        End If
    Next keyItem

End Function

'------------------------------------------------------------
' 内部処理
'------------------------------------------------------------

Private Sub HRS4S_EnsureInitialized()
    If Not mLoaded Or mSession Is Nothing Then
        HRS4S_Initialize True
    End If
End Sub

Private Function HRS4S_EnsureSessionSheet() As Worksheet

    Dim ws As Worksheet
    Dim headers As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SESSION)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
            ThisWorkbook.Worksheets.count))
        ws.Name = SH_SESSION
    End If

    headers = Array( _
        "キャッシュキー", "業者名", "商品番号", "商品名", _
        "納品日", "使用日", "区分", "使用数量", _
        "取消状態", "配分後", "単位", "発注書", _
        "シート名", "セル番地", "原票DB行", _
        "変更済", "更新日時", "表示モード")

    ws.Range("A1:R1").value = headers
    ws.Rows(1).Font.Bold = True
    ws.Visible = xlSheetVeryHidden

    Set HRS4S_EnsureSessionSheet = ws

End Function

Private Function HRS4S_RowToRecord( _
    ByVal data As Variant, _
    ByVal rowNumber As Long, _
    ByVal keyText As String) As Variant

    HRS4S_RowToRecord = Array( _
        keyText, CStr(data(rowNumber, SC_VENDOR)), _
        HRS4S_NormalizeCode(data(rowNumber, SC_PRODUCT_CODE)), _
        CStr(data(rowNumber, SC_PRODUCT_NAME)), _
        data(rowNumber, SC_DELIVERY), data(rowNumber, SC_USE_DATE), _
        CStr(data(rowNumber, SC_MEAL)), _
        data(rowNumber, SC_USAGE), _
        HRS4S_ToBoolean(data(rowNumber, SC_CANCEL)), _
        data(rowNumber, SC_DISTRIBUTION), _
        CStr(data(rowNumber, SC_UNIT)), _
        CStr(data(rowNumber, SC_SOURCE_BOOK)), _
        CStr(data(rowNumber, SC_SOURCE_SHEET)), _
        CStr(data(rowNumber, SC_SOURCE_CELL)), _
        CLng(Val(data(rowNumber, SC_RAW_ROW))), _
        HRS4S_ToBoolean(data(rowNumber, SC_CHANGED)), _
        data(rowNumber, SC_UPDATED_AT), _
        CStr(data(rowNumber, SC_DISPLAY_MODE)))

End Function

Private Function HRS4S_MakeKey( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal deliveryDate As Variant, _
    ByVal useDate As Variant, _
    ByVal mealText As String, _
    ByVal sourceSheet As String, _
    ByVal sourceCell As String, _
    ByVal rawRow As Long) As String

    Dim separatorText As String
    Dim productIdentity As String

    separatorText = ChrW$(&H1F)
    productIdentity = HRS4S_ProductIdentity(productCode, productName)

    HRS4S_MakeKey = _
        UCase$(Trim$(vendorName)) & separatorText & _
        productIdentity & separatorText & _
        HRS4S_DateKey(deliveryDate) & separatorText & _
        HRS4S_DateKey(useDate) & separatorText & _
        UCase$(Trim$(mealText)) & separatorText & _
        UCase$(Trim$(sourceSheet)) & separatorText & _
        UCase$(Trim$(sourceCell)) & separatorText & _
        CStr(rawRow)

End Function

Private Function HRS4S_ProductIdentity( _
    ByVal productCode As Variant, _
    ByVal productName As Variant) As String

    Dim codeText As String
    Dim nameText As String

    codeText = HRS4S_NormalizeCode(productCode)
    nameText = UCase$(Trim$(CStr(productName)))

    If codeText <> "" Then
        HRS4S_ProductIdentity = "C|" & codeText
    Else
        HRS4S_ProductIdentity = "N|" & nameText
    End If

End Function

Private Function HRS4S_ProductMatches( _
    ByVal recordData As Variant, _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String) As Boolean

    If StrComp(CStr(recordData(SC_VENDOR - 1)), _
               Trim$(vendorName), vbTextCompare) <> 0 Then Exit Function

    HRS4S_ProductMatches = _
        (HRS4S_ProductIdentity(recordData(SC_PRODUCT_CODE - 1), _
         recordData(SC_PRODUCT_NAME - 1)) = _
         HRS4S_ProductIdentity(productCode, productName))

End Function

Private Function HRS4S_NormalizeCode(ByVal value As Variant) As String

    Dim textValue As String

    If IsError(value) Or IsEmpty(value) Then Exit Function

    textValue = Trim$(CStr(value))
    If textValue = "" Then Exit Function

    If IsNumeric(textValue) Then
        HRS4S_NormalizeCode = Format$(CDbl(textValue), "0")
    Else
        HRS4S_NormalizeCode = textValue
    End If

End Function

Private Function HRS4S_DateKey(ByVal value As Variant) As String

    If IsError(value) Or IsEmpty(value) Then Exit Function
    If Trim$(CStr(value)) = "" Then Exit Function

    If IsDate(value) Then
        HRS4S_DateKey = Format$(CDate(value), "yyyy/mm/dd")
    Else
        HRS4S_DateKey = Trim$(CStr(value))
    End If

End Function

Private Function HRS4S_NormalizeDateValue( _
    ByVal value As Variant) As Variant

    If IsError(value) Or IsEmpty(value) Then
        HRS4S_NormalizeDateValue = ""
    ElseIf Trim$(CStr(value)) = "" Then
        HRS4S_NormalizeDateValue = ""
    ElseIf IsDate(value) Then
        HRS4S_NormalizeDateValue = CDate(value)
    Else
        HRS4S_NormalizeDateValue = Trim$(CStr(value))
    End If

End Function

Private Function HRS4S_ValueOrBlank(ByVal value As Variant) As Variant

    If IsError(value) Or IsEmpty(value) Then
        HRS4S_ValueOrBlank = ""
    ElseIf Trim$(CStr(value)) = "" Then
        HRS4S_ValueOrBlank = ""
    ElseIf IsNumeric(value) Then
        HRS4S_ValueOrBlank = CDbl(value)
    Else
        HRS4S_ValueOrBlank = value
    End If

End Function

Private Function HRS4S_ToBoolean(ByVal value As Variant) As Boolean

    Dim textValue As String

    If VarType(value) = vbBoolean Then
        HRS4S_ToBoolean = CBool(value)
        Exit Function
    End If

    If IsNumeric(value) Then
        HRS4S_ToBoolean = (CDbl(value) <> 0)
        Exit Function
    End If

    textValue = UCase$(Trim$(CStr(value)))
    HRS4S_ToBoolean = _
        (textValue = "TRUE") Or (textValue = "YES") Or _
        (textValue = "1") Or (textValue = "■") Or _
        (textValue = "取消") Or (textValue = "済")

End Function

Private Function HRS4S_IsCancelledMark(ByVal value As Variant) As Boolean
    HRS4S_IsCancelledMark = HRS4S_ToBoolean(value)
End Function

Private Function HRS4S_LastRow(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    Set lastCell = ws.Cells.Find( _
        What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False)

    If lastCell Is Nothing Then
        HRS4S_LastRow = 1
    Else
        HRS4S_LastRow = lastCell.Row
    End If

End Function

Private Function HRS4S_LastPreviewRow(ByVal ws As Worksheet) As Long

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.count, PC_USE_DATE).End(xlUp).Row
    If lastRow < PREVIEW_FIRST_ROW Then lastRow = PREVIEW_FIRST_ROW - 1
    HRS4S_LastPreviewRow = lastRow

End Function

Private Function HRS4S_SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    HRS4S_SheetExists = Not ws Is Nothing

End Function

Private Sub HRS4S_RecalculateAggregate(ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim r As Long
    Dim detailRows As String
    Dim parts As Variant
    Dim item As Variant
    Dim detailRow As Long
    Dim totalValue As Double
    Dim hasValue As Boolean
    Dim allCancelled As Boolean

    lastRow = ws.Cells(ws.Rows.count, 16).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For r = 2 To lastRow
        detailRows = Trim$(CStr(ws.Cells(r, 24).value))
        totalValue = 0
        hasValue = False
        allCancelled = True

        If detailRows <> "" Then
            parts = Split(detailRows, ",")

            For Each item In parts
                detailRow = CLng(Val(item))

                If detailRow >= PREVIEW_FIRST_ROW Then
                    If Not HRS4S_IsCancelledMark( _
                       ws.Cells(detailRow, PC_CANCEL).value) Then
                        allCancelled = False
                    End If

                    If IsNumeric(ws.Cells(detailRow, _
                       PC_DISTRIBUTION).value) Then
                        totalValue = totalValue + CDbl( _
                            ws.Cells(detailRow, _
                            PC_DISTRIBUTION).value)
                        hasValue = True
                    End If
                End If
            Next item
        End If

        ws.Cells(r, 15).value = IIf(allCancelled, "■", "□")

        If hasValue Then
            ws.Cells(r, 20).value = totalValue
        Else
            ws.Cells(r, 20).ClearContents
        End If
    Next r

End Sub

Private Sub HRS4S_BeginFast()

    If mFastDepth = 0 Then
        mOldCalculation = Application.Calculation
        mOldScreenUpdating = Application.ScreenUpdating
        mOldEnableEvents = Application.EnableEvents

        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    End If

    mFastDepth = mFastDepth + 1

End Sub

Private Sub HRS4S_EndFast()

    If mFastDepth <= 0 Then Exit Sub

    mFastDepth = mFastDepth - 1

    If mFastDepth = 0 Then
        Application.Calculation = mOldCalculation
        Application.ScreenUpdating = mOldScreenUpdating
        Application.EnableEvents = mOldEnableEvents
    End If

End Sub

Private Function HRS4S_Elapsed(ByVal startedAt As Double) As Double

    Dim currentTime As Double

    currentTime = Timer
    If currentTime < startedAt Then currentTime = currentTime + 86400#
    HRS4S_Elapsed = currentTime - startedAt

End Function

Private Sub HRS4S_WritePerformance( _
    ByVal processName As String, _
    ByVal elapsedSeconds As Double, _
    ByVal recordCount As Long, _
    ByVal noteText As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    On Error Resume Next

    If Not HRS4S_SheetExists(SH_PERFORMANCE) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SH_PERFORMANCE)
    nextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = processName
    ws.Cells(nextRow, 3).value = elapsedSeconds
    ws.Cells(nextRow, 4).value = recordCount
    ws.Cells(nextRow, 5).value = noteText

    On Error GoTo 0

End Sub

Private Function HRS4S_ValueText(ByVal value As Variant) As String

    If IsError(value) Or IsEmpty(value) Then
        HRS4S_ValueText = ""
    Else
        HRS4S_ValueText = CStr(value)
    End If

End Function
