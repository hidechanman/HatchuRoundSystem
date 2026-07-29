Attribute VB_Name = "modHRS4_Display"
Option Explicit

'=========================================================
' 発注まるめシステム Ver4.0 Part2
' 完成キャッシュ表示ブリッジ
'
' 目的:
' ・商品選択時に発注原票DBを検索しない
' ・商品選択時に集約処理を行わない
' ・V4通常表示キャッシュとV4集約表示キャッシュを
'   既存の使用日プレビュー作業へ配列で一括展開する
' ・既存の配分、取消、書戻し処理を壊さず接続する
'=========================================================

Private Const SH_INPUT As String = "発注入力"
Private Const SH_LEGACY_PREVIEW As String = "使用日プレビュー作業"
Private Const SH_SESSION As String = "配分セッションDB"

Private Const NORMAL_COLS As Long = 20
Private Const AGGREGATE_COLS As Long = 20

Public Sub HRS4_LoadSelectedProductToLegacyPreview( _
    ByVal productCode As String, _
    ByVal productName As String)

    Dim startedAt As Double
    Dim wsInput As Worksheet
    Dim wsPreview As Worksheet
    Dim normalRows As Variant
    Dim aggregateRows As Variant
    Dim normalOutput As Variant
    Dim aggregateOutput As Variant
    Dim rawToDetailRow As Object
    Dim vendorName As String
    Dim aggregateMode As Boolean
    Dim normalCount As Long
    Dim aggregateCount As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ErrHandler

    startedAt = Timer
    HRS4_BeginFast "Ver4.0完成キャッシュから商品を表示しています..."

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsPreview = ThisWorkbook.Worksheets(SH_LEGACY_PREVIEW)

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    aggregateMode = HRS4_ReadAggregateMode(wsPreview)

    normalRows = HRS4_GetProductRows( _
        vendorName, productCode, productName, False)

    aggregateRows = HRS4_GetProductRows( _
        vendorName, productCode, productName, True)

    Set rawToDetailRow = CreateObject("Scripting.Dictionary")
    rawToDetailRow.CompareMode = vbTextCompare

    normalOutput = HRS4_ConvertNormalRows( _
        normalRows, rawToDetailRow, normalCount)

    aggregateOutput = HRS4_ConvertAggregateRows( _
        aggregateRows, rawToDetailRow, normalOutput, _
        normalCount, aggregateCount)

    HRS4_ClearLegacyPreview wsPreview
    HRS4_WriteLegacyHeaders wsPreview, aggregateMode, _
        productCode, productName

    If normalCount > 0 Then
        wsPreview.Range("A2").Resize(normalCount, 12).Value2 = _
            normalOutput
    End If

    If aggregateCount > 0 Then
        wsPreview.Range("O2").Resize(aggregateCount, 12).Value2 = _
            aggregateOutput
    End If

    HRS4_OverlaySessionData wsPreview, vendorName, _
        productCode, productName, normalCount

    wsPreview.Range("N8").value = "FALSE"

    HRS4_WritePerformanceLog _
        "V4商品選択表示準備", _
        HRS4_ElapsedSeconds(startedAt), _
        normalCount + aggregateCount, _
        vendorName & " / " & productName

ExitHandler:
    HRS4_EndFast

    If errorNumber <> 0 Then
        MsgBox "Ver4.0の商品表示準備中にエラーが発生しました。" & _
               vbCrLf & _
               "番号: " & CStr(errorNumber) & vbCrLf & _
               "内容: " & errorDescription, _
               vbCritical, "発注まるめシステム"
    End If

    Exit Sub

ErrHandler:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Resume ExitHandler

End Sub

Public Sub HRS4_ToggleAggregateDisplay()

    Dim wsPreview As Worksheet
    Dim nextMode As Boolean

    Set wsPreview = ThisWorkbook.Worksheets(SH_LEGACY_PREVIEW)

    nextMode = Not HRS4_ReadAggregateMode(wsPreview)
    wsPreview.Range("N7").value = IIf(nextMode, "TRUE", "FALSE")
    wsPreview.Range("N2").value = 1
    wsPreview.Range("N5").value = 1

    HRS_RenderPreviewPage
    HRS_UpdateCurrentWriteBackFromV4
    HRS4_UpdateAggregateCaption nextMode

End Sub

Private Function HRS4_ConvertNormalRows( _
    ByVal sourceRows As Variant, _
    ByRef rawToDetailRow As Object, _
    ByRef outputCount As Long) As Variant

    Dim resultData() As Variant
    Dim r As Long
    Dim detailRow As Long
    Dim rawRow As Long

    If IsEmpty(sourceRows) Then
        ReDim resultData(1 To 1, 1 To 12)
        HRS4_ConvertNormalRows = resultData
        Exit Function
    End If

    outputCount = UBound(sourceRows, 1)
    ReDim resultData(1 To outputCount, 1 To 12)

    For r = 1 To outputCount
        detailRow = r + 1
        rawRow = CLng(Val(sourceRows(r, 20)))

        resultData(r, 1) = HRS4_NormalCancelMark(sourceRows(r, 11))
        resultData(r, 2) = sourceRows(r, 8)
        resultData(r, 3) = sourceRows(r, 9)
        resultData(r, 4) = sourceRows(r, 10)
        resultData(r, 5) = sourceRows(r, 7)
        resultData(r, 6) = sourceRows(r, 12)
        resultData(r, 7) = sourceRows(r, 6)
        resultData(r, 8) = sourceRows(r, 5)
        resultData(r, 9) = sourceRows(r, 14)
        resultData(r, 10) = sourceRows(r, 17)
        resultData(r, 11) = sourceRows(r, 13)
        resultData(r, 12) = "FALSE"

        If rawRow > 0 Then
            rawToDetailRow(CStr(rawRow)) = detailRow
        End If
    Next r

    HRS4_ConvertNormalRows = resultData

End Function

Private Function HRS4_ConvertAggregateRows( _
    ByVal sourceRows As Variant, _
    ByVal rawToDetailRow As Object, _
    ByVal normalOutput As Variant, _
    ByVal normalCount As Long, _
    ByRef outputCount As Long) As Variant

    Dim resultData() As Variant
    Dim r As Long
    Dim targetDetailRow As Long
    Dim detailRowList As String
    Dim allCancelled As Boolean

    If IsEmpty(sourceRows) Then
        ReDim resultData(1 To 1, 1 To 12)
        HRS4_ConvertAggregateRows = resultData
        Exit Function
    End If

    outputCount = UBound(sourceRows, 1)
    ReDim resultData(1 To outputCount, 1 To 12)

    For r = 1 To outputCount
        targetDetailRow = HRS4_MapRawRow( _
            rawToDetailRow, CLng(Val(sourceRows(r, 14))))

        detailRowList = HRS4_MapRawRowList( _
            rawToDetailRow, CStr(sourceRows(r, 15)))

        allCancelled = HRS4_AllMappedRowsCancelled( _
            normalOutput, normalCount, detailRowList)

        resultData(r, 1) = IIf(allCancelled, "■", "□")
        resultData(r, 2) = HRS4_UseDateRangeText( _
            sourceRows(r, 8), sourceRows(r, 9))
        resultData(r, 3) = "朝へ集約"
        resultData(r, 4) = sourceRows(r, 11)
        resultData(r, 5) = sourceRows(r, 7)
        resultData(r, 6) = sourceRows(r, 13)
        resultData(r, 7) = sourceRows(r, 5)
        resultData(r, 8) = sourceRows(r, 17)
        resultData(r, 9) = targetDetailRow
        resultData(r, 10) = detailRowList
        resultData(r, 11) = sourceRows(r, 8)
        resultData(r, 12) = sourceRows(r, 9)
    Next r

    HRS4_ConvertAggregateRows = resultData

End Function

Private Sub HRS4_ClearLegacyPreview(ByVal wsPreview As Worksheet)

    wsPreview.Range("A1:Z5000").ClearContents

End Sub

Private Sub HRS4_WriteLegacyHeaders( _
    ByVal wsPreview As Worksheet, _
    ByVal aggregateMode As Boolean, _
    ByVal productCode As String, _
    ByVal productName As String)

    wsPreview.Range("A1:L1").value = Array( _
        "取消", "使用日", "区分", "使用数量", _
        "納品日", "配分後", "単位/注意点", _
        "商品", "発注書", "セル番地", _
        "取込元ブック", "変更済")

    wsPreview.Range("O1:Z1").value = Array( _
        "取消", "使用日表示", "区分", "使用数量", _
        "納品日", "配分後", "商品", "発注書", _
        "書戻対象行", "明細行一覧", _
        "最初使用日", "最後使用日")

    wsPreview.Range("N1").value = "表示位置"
    wsPreview.Range("N2").value = 1
    wsPreview.Range("N3").NumberFormat = "@"
    wsPreview.Range("N3").value = productCode
    wsPreview.Range("N4").value = productName
    wsPreview.Range("N5").value = 1
    wsPreview.Range("N6").value = 1
    wsPreview.Range("N7").value = _
        IIf(aggregateMode, "TRUE", "FALSE")
    wsPreview.Range("N8").value = "FALSE"

End Sub

Private Sub HRS4_OverlaySessionData( _
    ByVal wsPreview As Worksheet, _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal normalCount As Long)

    Dim wsSession As Worksheet
    Dim lastRow As Long
    Dim sessionData As Variant
    Dim previewData As Variant
    Dim sessionIndex As Object
    Dim r As Long
    Dim keyText As String
    Dim itemData As Variant

    If normalCount <= 0 Then Exit Sub
    If Not HRS4_SheetExistsLocal(SH_SESSION) Then Exit Sub

    Set wsSession = ThisWorkbook.Worksheets(SH_SESSION)
    lastRow = HRS4_LastRow(wsSession, 1)

    If lastRow < 2 Then Exit Sub

    sessionData = wsSession.Range("A2:P" & lastRow).Value2
    Set sessionIndex = CreateObject("Scripting.Dictionary")
    sessionIndex.CompareMode = vbTextCompare

    For r = 1 To UBound(sessionData, 1)
        If StrComp(Trim$(CStr(sessionData(r, 1))), _
                   vendorName, vbTextCompare) = 0 Then
            If HRS4_ProductMatches( _
                sessionData(r, 2), sessionData(r, 3), _
                productCode, productName) Then

                keyText = HRS4_SessionKey( _
                    sessionData(r, 8), _
                    sessionData(r, 9), _
                    sessionData(r, 11), _
                    sessionData(r, 14), _
                    sessionData(r, 15))

                itemData = Array( _
                    sessionData(r, 7), _
                    sessionData(r, 10), _
                    sessionData(r, 12), _
                    sessionData(r, 13), _
                    sessionData(r, 4), _
                    sessionData(r, 14), _
                    sessionData(r, 15), _
                    sessionData(r, 16))
                sessionIndex(keyText) = itemData
            End If
        End If
    Next r

    If sessionIndex.count = 0 Then Exit Sub

    previewData = wsPreview.Range("A2:L" & normalCount + 1).Value2

    For r = 1 To UBound(previewData, 1)
        keyText = HRS4_SessionKey( _
            previewData(r, 2), _
            previewData(r, 3), _
            previewData(r, 5), _
            previewData(r, 9), _
            previewData(r, 10))

        If sessionIndex.Exists(keyText) Then
            itemData = sessionIndex(keyText)

            previewData(r, 1) = _
                IIf(CBool(itemData(0)), "■", "□")
            previewData(r, 4) = itemData(1)
            previewData(r, 6) = itemData(2)
            previewData(r, 7) = itemData(3)
            previewData(r, 8) = itemData(4)
            previewData(r, 9) = itemData(5)
            previewData(r, 10) = itemData(6)
            previewData(r, 11) = itemData(7)
            previewData(r, 12) = "TRUE"
        End If
    Next r

    wsPreview.Range("A2").Resize(normalCount, 12).Value2 = previewData

    HRS4_RecalculateAggregateFromDetails wsPreview

End Sub

Private Sub HRS4_RecalculateAggregateFromDetails( _
    ByVal wsPreview As Worksheet)

    Dim aggregateLastRow As Long
    Dim r As Long
    Dim detailRows As String
    Dim totalDistribution As Double
    Dim hasDistribution As Boolean
    Dim allCancelled As Boolean
    Dim parts As Variant
    Dim item As Variant
    Dim detailRow As Long

    aggregateLastRow = HRS4_LastRow(wsPreview, 16)
    If aggregateLastRow < 2 Then Exit Sub

    For r = 2 To aggregateLastRow
        detailRows = CStr(wsPreview.Cells(r, 24).value)
        totalDistribution = 0
        hasDistribution = False
        allCancelled = True

        If Trim$(detailRows) <> "" Then
            parts = Split(detailRows, ",")

            For Each item In parts
                detailRow = CLng(Val(item))

                If detailRow > 1 Then
                    If CStr(wsPreview.Cells(detailRow, 1).value) <> "■" Then
                        allCancelled = False
                    End If

                    If IsNumeric(wsPreview.Cells(detailRow, 6).value) Then
                        If Trim$(CStr( _
                            wsPreview.Cells(detailRow, 6).value)) <> "" Then

                            totalDistribution = totalDistribution + _
                                CDbl(wsPreview.Cells(detailRow, 6).value)
                            hasDistribution = True
                        End If
                    End If
                End If
            Next item
        Else
            allCancelled = False
        End If

        wsPreview.Cells(r, 15).value = _
            IIf(allCancelled, "■", "□")

        If hasDistribution And totalDistribution <> 0 Then
            wsPreview.Cells(r, 20).value = totalDistribution
        Else
            wsPreview.Cells(r, 20).ClearContents
        End If
    Next r

End Sub

Private Function HRS4_ReadAggregateMode( _
    ByVal wsPreview As Worksheet) As Boolean

    HRS4_ReadAggregateMode = _
        (UCase$(Trim$(CStr(wsPreview.Range("N7").value))) = "TRUE")

End Function

Private Sub HRS4_UpdateAggregateCaption( _
    ByVal aggregateMode As Boolean)

    Dim wsInput As Worksheet
    Dim captionText As String

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)

    captionText = IIf(aggregateMode, "集約 ON", "集約 OFF")

    On Error Resume Next
    wsInput.Shapes("HRS_AggregateMode"). _
        TextFrame.Characters.Text = captionText
    On Error GoTo 0

End Sub

Private Sub HRS_UpdateCurrentWriteBackFromV4()

    '既存版の今回書戻しプレビューは、使用日プレビュー作業を
    '参照して作られるため、既存公開処理へ戻す。
    '既存モジュール側の切替処理から呼ばれる場合は何もしない。
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & _
        "'!HRS_BuildCurrentWriteBackPreview"
    On Error GoTo 0

End Sub

Private Function HRS4_NormalCancelMark( _
    ByVal sourceMark As Variant) As String

    Select Case CStr(sourceMark)
        Case "■", "□", "●", "○"
            HRS4_NormalCancelMark = CStr(sourceMark)
        Case Else
            HRS4_NormalCancelMark = "○"
    End Select

End Function

Private Function HRS4_MapRawRow( _
    ByVal rawToDetailRow As Object, _
    ByVal rawRow As Long) As Long

    If rawRow <= 0 Then Exit Function

    If rawToDetailRow.Exists(CStr(rawRow)) Then
        HRS4_MapRawRow = CLng(rawToDetailRow(CStr(rawRow)))
    End If

End Function

Private Function HRS4_MapRawRowList( _
    ByVal rawToDetailRow As Object, _
    ByVal rawRowList As String) As String

    Dim parts As Variant
    Dim item As Variant
    Dim rawRow As Long
    Dim detailRow As Long
    Dim resultText As String

    If Trim$(rawRowList) = "" Then Exit Function

    parts = Split(rawRowList, ",")

    For Each item In parts
        rawRow = CLng(Val(item))
        detailRow = HRS4_MapRawRow(rawToDetailRow, rawRow)

        If detailRow > 0 Then
            If resultText <> "" Then resultText = resultText & ","
            resultText = resultText & CStr(detailRow)
        End If
    Next item

    HRS4_MapRawRowList = resultText

End Function

Private Function HRS4_AllMappedRowsCancelled( _
    ByVal normalOutput As Variant, _
    ByVal normalCount As Long, _
    ByVal detailRowList As String) As Boolean

    Dim parts As Variant
    Dim item As Variant
    Dim detailRow As Long
    Dim arrayRow As Long
    Dim hasRows As Boolean

    If normalCount <= 0 Then Exit Function
    If Trim$(detailRowList) = "" Then Exit Function

    HRS4_AllMappedRowsCancelled = True
    parts = Split(detailRowList, ",")

    For Each item In parts
        detailRow = CLng(Val(item))
        arrayRow = detailRow - 1

        If arrayRow >= 1 And arrayRow <= normalCount Then
            hasRows = True

            If CStr(normalOutput(arrayRow, 1)) <> "■" Then
                HRS4_AllMappedRowsCancelled = False
                Exit Function
            End If
        End If
    Next item

    If Not hasRows Then
        HRS4_AllMappedRowsCancelled = False
    End If

End Function

Private Function HRS4_UseDateRangeText( _
    ByVal firstDate As Variant, _
    ByVal lastDate As Variant) As String

    If Trim$(CStr(firstDate)) = "" Then Exit Function

    If IsDate(firstDate) And IsDate(lastDate) Then
        If CLng(CDate(firstDate)) = CLng(CDate(lastDate)) Then
            HRS4_UseDateRangeText = Format$(CDate(firstDate), "m/d")
        Else
            HRS4_UseDateRangeText = _
                Format$(CDate(firstDate), "m/d") & "～" & _
                Format$(CDate(lastDate), "m/d")
        End If
    ElseIf CStr(firstDate) = CStr(lastDate) Then
        HRS4_UseDateRangeText = CStr(firstDate)
    Else
        HRS4_UseDateRangeText = _
            CStr(firstDate) & "～" & CStr(lastDate)
    End If

End Function

Private Function HRS4_SessionKey( _
    ByVal useDate As Variant, _
    ByVal categoryText As Variant, _
    ByVal deliveryDate As Variant, _
    ByVal sheetName As Variant, _
    ByVal cellAddress As Variant) As String

    HRS4_SessionKey = _
        HRS4_DateKey(useDate) & "|" & _
        UCase$(Trim$(CStr(categoryText))) & "|" & _
        HRS4_DateKey(deliveryDate) & "|" & _
        UCase$(Trim$(CStr(sheetName))) & "|" & _
        UCase$(Trim$(CStr(cellAddress)))

End Function

Private Function HRS4_ProductMatches( _
    ByVal sessionCode As Variant, _
    ByVal sessionName As Variant, _
    ByVal productCode As String, _
    ByVal productName As String) As Boolean

    Dim sessionIdentity As String
    Dim selectedIdentity As String

    sessionIdentity = HRS4_ProductIdentity( _
        sessionCode, sessionName)
    selectedIdentity = HRS4_ProductIdentity( _
        productCode, productName)

    HRS4_ProductMatches = _
        (StrComp(sessionIdentity, selectedIdentity, _
                 vbTextCompare) = 0)

End Function

Private Function HRS4_SheetExistsLocal( _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    HRS4_SheetExistsLocal = Not ws Is Nothing

End Function
