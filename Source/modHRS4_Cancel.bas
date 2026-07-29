Attribute VB_Name = "modHRS4_Cancel"
Option Explicit

'============================================================
' 発注まるめシステム Ver4.0
' Part3-3 Cancel Engine
'============================================================

Private Const HRS4C_VERSION As String = "Ver4.0.0 Part3-3"
Private Const APP_TITLE As String = "発注まるめシステム"

Private Const SH_INPUT As String = "発注入力"
Private Const SH_PREVIEW As String = "使用日プレビュー作業"
Private Const SH_V4_NORMAL As String = "V4通常表示キャッシュ"
Private Const SH_V4_AGGREGATE As String = "V4集約表示キャッシュ"
Private Const SH_V4_WRITEBACK As String = "V4書戻しキャッシュ"
Private Const SH_PERFORMANCE As String = "V4速度ログ"

Private Const PREVIEW_TOP As Long = 8
Private Const PREVIEW_BOTTOM As Long = 20
Private Const PREVIEW_FIRST_ROW As Long = 2

' 使用日プレビュー作業
Private Const PC_CANCEL As Long = 1
Private Const PC_USE_DATE As Long = 2
Private Const PC_USAGE As Long = 4
Private Const PC_DISTRIBUTION As Long = 6
Private Const PC_CHANGED As Long = 12
Private Const PC_AGG_CANCEL As Long = 15
Private Const PC_AGG_RAW_ROW_LIST As Long = 24

' V4通常表示キャッシュ
Private Const NC_PRODUCT_CODE As Long = 4
Private Const NC_PRODUCT_NAME As Long = 5
Private Const NC_USAGE As Long = 10
Private Const NC_CANCEL As Long = 11
Private Const NC_DISTRIBUTION As Long = 12
Private Const NC_RAW_ROW As Long = 20

' V4集約表示キャッシュ
Private Const AC_PRODUCT_CODE As Long = 4
Private Const AC_PRODUCT_NAME As Long = 5
Private Const AC_USAGE_TOTAL As Long = 11
Private Const AC_CANCEL As Long = 12
Private Const AC_DISTRIBUTION As Long = 13
Private Const AC_RAW_ROW_LIST As Long = 15

' V4書戻しキャッシュ
' Part1/Part3-1の列構成差を吸収するため、見出し名で列を取得する。

'------------------------------------------------------------
' バージョン確認
'------------------------------------------------------------
Public Sub HRS4C_ShowVersion()
    MsgBox "Cancel Engine " & HRS4C_VERSION, vbInformation, APP_TITLE
End Sub

'------------------------------------------------------------
' 画面上の1行を取消／取消解除
' displayRow は発注入力シート上の行番号（8～20）
'------------------------------------------------------------
Public Sub HRS4C_TogglePreviewCancel(ByVal displayRow As Long)

    Dim startedAt As Double
    Dim wsPreview As Worksheet
    Dim sourceRow As Long
    Dim aggregateMode As Boolean
    Dim turnOn As Boolean
    Dim rowList As String

    On Error GoTo ErrHandler

    startedAt = Timer
    HRS4C_BeginFast

    If displayRow < PREVIEW_TOP Or displayRow > PREVIEW_BOTTOM Then
        Err.Raise vbObjectError + 4101, "HRS4C_TogglePreviewCancel", _
                  "表示行は8～20を指定してください。"
    End If

    Set wsPreview = HRS4C_GetSheet(SH_PREVIEW)
    sourceRow = HRS4C_DisplayToSourceRow(wsPreview, displayRow)

    If sourceRow < PREVIEW_FIRST_ROW Then GoTo SafeExit
    If Trim$(CStr(wsPreview.Cells(sourceRow, PC_USE_DATE).value)) = "" Then GoTo SafeExit

    aggregateMode = HRS4C_IsAggregateMode(wsPreview)

    If aggregateMode Then
        rowList = Trim$(CStr(wsPreview.Cells(sourceRow, PC_AGG_RAW_ROW_LIST).value))
        turnOn = (CStr(wsPreview.Cells(sourceRow, PC_AGG_CANCEL).value) <> "■")
        HRS4C_SetAggregateGroupCancel wsPreview, rowList, turnOn
        wsPreview.Cells(sourceRow, PC_AGG_CANCEL).value = IIf(turnOn, "■", "□")
    Else
        turnOn = (CStr(wsPreview.Cells(sourceRow, PC_CANCEL).value) <> "■")
        HRS4C_SetPreviewRowCancel wsPreview, sourceRow, turnOn
    End If

    HRS4C_AfterCancelChange True
    HRS4C_WritePerformance "V4取消切替", HRS4C_Elapsed(startedAt), 1, _
                           "表示行=" & CStr(displayRow)

SafeExit:
    HRS4C_EndFast
    Exit Sub

ErrHandler:
    HRS4C_EndFast
    MsgBox "取消切替でエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, vbExclamation, APP_TITLE
End Sub

' 旧名称から接続しやすい別名
Public Sub HRS4_CancelCurrentProduct(Optional ByVal turnOn As Boolean = True)
    HRS4C_SetCurrentProductCancel turnOn
End Sub

'------------------------------------------------------------
' 現在商品の全行を取消／取消解除
'------------------------------------------------------------
Public Sub HRS4C_ToggleAllCancel()

    Dim wsPreview As Worksheet
    Dim turnOn As Boolean

    On Error GoTo ErrHandler

    Set wsPreview = HRS4C_GetSheet(SH_PREVIEW)
    turnOn = Not HRS4C_AllCancelled(wsPreview)
    HRS4C_SetCurrentProductCancel turnOn
    Exit Sub

ErrHandler:
    MsgBox "全取消の切替でエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub HRS4C_SetCurrentProductCancel(ByVal turnOn As Boolean)

    Dim startedAt As Double
    Dim wsPreview As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim changedCount As Long

    On Error GoTo ErrHandler

    startedAt = Timer
    HRS4C_BeginFast
    Set wsPreview = HRS4C_GetSheet(SH_PREVIEW)
    lastRow = HRS4C_LastRow(wsPreview, PC_USE_DATE)

    For r = PREVIEW_FIRST_ROW To lastRow
        If Trim$(CStr(wsPreview.Cells(r, PC_USE_DATE).value)) <> "" Then
            HRS4C_SetPreviewRowCancel wsPreview, r, turnOn
            changedCount = changedCount + 1
        End If
    Next r

    HRS4C_RecalculateAggregateCancel wsPreview
    HRS4C_AfterCancelChange True
    HRS4C_WritePerformance "V4商品全取消", HRS4C_Elapsed(startedAt), _
                           changedCount, "状態=" & CStr(turnOn)

SafeExit:
    HRS4C_EndFast
    Exit Sub

ErrHandler:
    HRS4C_EndFast
    MsgBox "現在商品の取消処理でエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub HRS4_CancelAll()
    HRS4C_SetAllProductsCancel True
End Sub

Public Sub HRS4_ClearCancel()
    HRS4C_SetCurrentProductCancel False
End Sub

'------------------------------------------------------------
' V4通常表示キャッシュに存在する全商品を取消／解除
' 表示中商品は即時更新し、その他はキャッシュに保存する。
'------------------------------------------------------------
Public Sub HRS4C_SetAllProductsCancel(ByVal turnOn As Boolean)

    Dim startedAt As Double
    Dim changedCount As Long

    On Error GoTo ErrHandler

    startedAt = Timer
    HRS4C_BeginFast

    changedCount = HRS4C_SetCacheCancelAll(SH_V4_NORMAL, NC_CANCEL, turnOn)
    HRS4C_SetCacheCancelAll SH_V4_AGGREGATE, AC_CANCEL, turnOn
    HRS4C_SetWriteBackCancelAll turnOn
    HRS4C_SetPreviewAllWithoutFlush turnOn
    HRS4C_AfterCancelChange True

    HRS4C_WritePerformance "V4全商品取消", HRS4C_Elapsed(startedAt), _
                           changedCount, "状態=" & CStr(turnOn)

SafeExit:
    HRS4C_EndFast
    Exit Sub

ErrHandler:
    HRS4C_EndFast
    MsgBox "全商品の取消処理でエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, vbExclamation, APP_TITLE
End Sub

'------------------------------------------------------------
' 使用数量が0.1の行を一括取消
' onlyCurrentProduct=True: 表示中商品だけ
' False: V4通常表示キャッシュ全体
'------------------------------------------------------------
Public Sub HRS4C_CancelPointOne(Optional ByVal onlyCurrentProduct As Boolean = True)

    Dim startedAt As Double
    Dim changedCount As Long

    On Error GoTo ErrHandler

    startedAt = Timer
    HRS4C_BeginFast

    If onlyCurrentProduct Then
        changedCount = HRS4C_CancelPointOnePreview()
        HRS4C_RecalculateAggregateCancel HRS4C_GetSheet(SH_PREVIEW)
    Else
        changedCount = HRS4C_CancelPointOneNormalCache()
        HRS4C_RefreshPreviewFromNormalCache
    End If

    HRS4C_AfterCancelChange True
    HRS4C_WritePerformance "V4 0.1一括取消", HRS4C_Elapsed(startedAt), _
                           changedCount, "現在商品のみ=" & CStr(onlyCurrentProduct)

SafeExit:
    HRS4C_EndFast
    Exit Sub

ErrHandler:
    HRS4C_EndFast
    MsgBox "0.1一括取消でエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub HRS4_CancelPointOne()
    HRS4C_CancelPointOne True
End Sub

'------------------------------------------------------------
' 指定行の取消状態を返す
'------------------------------------------------------------
Public Function HRS4C_IsCancelled(ByVal previewRow As Long) As Boolean

    Dim wsPreview As Worksheet

    If Not HRS4C_SheetExists(SH_PREVIEW) Then Exit Function
    Set wsPreview = ThisWorkbook.Worksheets(SH_PREVIEW)

    If previewRow < PREVIEW_FIRST_ROW Then Exit Function
    HRS4C_IsCancelled = (CStr(wsPreview.Cells(previewRow, PC_CANCEL).value) = "■")
End Function

Public Function HRS4_IsCanceled(ByVal previewRow As Long) As Boolean
    HRS4_IsCanceled = HRS4C_IsCancelled(previewRow)
End Function

Public Function HRS4C_AllCancelled(ByVal wsPreview As Worksheet) As Boolean

    Dim lastRow As Long
    Dim r As Long
    Dim hasData As Boolean

    lastRow = HRS4C_LastRow(wsPreview, PC_USE_DATE)
    HRS4C_AllCancelled = True

    For r = PREVIEW_FIRST_ROW To lastRow
        If Trim$(CStr(wsPreview.Cells(r, PC_USE_DATE).value)) <> "" Then
            hasData = True
            If CStr(wsPreview.Cells(r, PC_CANCEL).value) <> "■" Then
                HRS4C_AllCancelled = False
                Exit Function
            End If
        End If
    Next r

    If Not hasData Then HRS4C_AllCancelled = False
End Function

'------------------------------------------------------------
' 画面表示のグレー状態だけを更新
' 取消行は灰色、通常行は標準色へ戻す。
'------------------------------------------------------------
Public Sub HRS4C_UpdateCancelDisplay()

    Dim wsInput As Worksheet
    Dim wsPreview As Worksheet
    Dim displayRow As Long
    Dim sourceRow As Long
    Dim isCancelled As Boolean
    Dim targetRange As Range

    On Error GoTo SafeExit

    If Not HRS4C_SheetExists(SH_INPUT) Then Exit Sub
    If Not HRS4C_SheetExists(SH_PREVIEW) Then Exit Sub

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsPreview = ThisWorkbook.Worksheets(SH_PREVIEW)

    For displayRow = PREVIEW_TOP To PREVIEW_BOTTOM
        sourceRow = HRS4C_DisplayToSourceRow(wsPreview, displayRow)
        Set targetRange = wsInput.Range("F" & displayRow & ":L" & displayRow)

        If sourceRow >= PREVIEW_FIRST_ROW Then
            If Trim$(CStr(wsPreview.Cells(sourceRow, PC_USE_DATE).value)) <> "" Then
                isCancelled = HRS4C_EffectiveCancel(wsPreview, sourceRow)
                HRS4C_ApplyGray targetRange, isCancelled
            Else
                HRS4C_ApplyGray targetRange, False
            End If
        End If
    Next displayRow

    If HRS4C_AllCancelled(wsPreview) Then
        wsInput.Range("K5").value = "■ 全取消"
    Else
        wsInput.Range("K5").value = "□ 全取消"
    End If

SafeExit:
End Sub

'============================================================
' 内部処理
'============================================================

Private Sub HRS4C_SetPreviewRowCancel(ByVal wsPreview As Worksheet, _
                                      ByVal rowNo As Long, _
                                      ByVal turnOn As Boolean)
    wsPreview.Cells(rowNo, PC_CANCEL).value = IIf(turnOn, "■", "□")
    wsPreview.Cells(rowNo, PC_CHANGED).value = "TRUE"
End Sub

Private Sub HRS4C_SetAggregateGroupCancel(ByVal wsPreview As Worksheet, _
                                           ByVal rowList As String, _
                                           ByVal turnOn As Boolean)

    Dim parts As Variant
    Dim item As Variant
    Dim rowNo As Long

    If Trim$(rowList) = "" Then Exit Sub
    parts = Split(rowList, ",")

    For Each item In parts
        rowNo = CLng(Val(item))
        If rowNo >= PREVIEW_FIRST_ROW Then
            HRS4C_SetPreviewRowCancel wsPreview, rowNo, turnOn
        End If
    Next item
End Sub

Private Sub HRS4C_RecalculateAggregateCancel(ByVal wsPreview As Worksheet)

    Dim lastRow As Long
    Dim r As Long
    Dim rowList As String

    lastRow = HRS4C_LastRow(wsPreview, PC_USE_DATE)

    For r = PREVIEW_FIRST_ROW To lastRow
        rowList = Trim$(CStr(wsPreview.Cells(r, PC_AGG_RAW_ROW_LIST).value))
        If rowList <> "" Then
            wsPreview.Cells(r, PC_AGG_CANCEL).value = _
                IIf(HRS4C_GroupAllCancelled(wsPreview, rowList), "■", "□")
        End If
    Next r
End Sub

Private Function HRS4C_GroupAllCancelled(ByVal wsPreview As Worksheet, _
                                          ByVal rowList As String) As Boolean

    Dim parts As Variant
    Dim item As Variant
    Dim rowNo As Long
    Dim hasRow As Boolean

    If Trim$(rowList) = "" Then Exit Function
    HRS4C_GroupAllCancelled = True
    parts = Split(rowList, ",")

    For Each item In parts
        rowNo = CLng(Val(item))
        If rowNo >= PREVIEW_FIRST_ROW Then
            hasRow = True
            If CStr(wsPreview.Cells(rowNo, PC_CANCEL).value) <> "■" Then
                HRS4C_GroupAllCancelled = False
                Exit Function
            End If
        End If
    Next item

    If Not hasRow Then HRS4C_GroupAllCancelled = False
End Function

Private Function HRS4C_EffectiveCancel(ByVal wsPreview As Worksheet, _
                                        ByVal rowNo As Long) As Boolean
    If HRS4C_IsAggregateMode(wsPreview) Then
        HRS4C_EffectiveCancel = _
            (CStr(wsPreview.Cells(rowNo, PC_AGG_CANCEL).value) = "■")
    Else
        HRS4C_EffectiveCancel = _
            (CStr(wsPreview.Cells(rowNo, PC_CANCEL).value) = "■")
    End If
End Function

Private Function HRS4C_CancelPointOnePreview() As Long

    Dim wsPreview As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set wsPreview = HRS4C_GetSheet(SH_PREVIEW)
    lastRow = HRS4C_LastRow(wsPreview, PC_USE_DATE)

    For r = PREVIEW_FIRST_ROW To lastRow
        If HRS4C_IsPointOne(wsPreview.Cells(r, PC_USAGE).value) Then
            If CStr(wsPreview.Cells(r, PC_CANCEL).value) <> "■" Then
                HRS4C_SetPreviewRowCancel wsPreview, r, True
                HRS4C_CancelPointOnePreview = HRS4C_CancelPointOnePreview + 1
            End If
        End If
    Next r
End Function

Private Function HRS4C_CancelPointOneNormalCache() As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    If Not HRS4C_SheetExists(SH_V4_NORMAL) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_V4_NORMAL)
    lastRow = HRS4C_LastRow(ws, NC_PRODUCT_NAME)

    For r = 2 To lastRow
        If HRS4C_IsPointOne(ws.Cells(r, NC_USAGE).value) Then
            If Not HRS4C_ToBoolean(ws.Cells(r, NC_CANCEL).value) Then
                ws.Cells(r, NC_CANCEL).value = True
                HRS4C_CancelPointOneNormalCache = _
                    HRS4C_CancelPointOneNormalCache + 1
            End If
        End If
    Next r

    HRS4C_RebuildAggregateCancelFromNormal
    HRS4C_SetWriteBackPointOne
End Function

Private Sub HRS4C_SetPreviewAllWithoutFlush(ByVal turnOn As Boolean)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    If Not HRS4C_SheetExists(SH_PREVIEW) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_PREVIEW)
    lastRow = HRS4C_LastRow(ws, PC_USE_DATE)

    For r = PREVIEW_FIRST_ROW To lastRow
        If Trim$(CStr(ws.Cells(r, PC_USE_DATE).value)) <> "" Then
            HRS4C_SetPreviewRowCancel ws, r, turnOn
        End If
    Next r

    HRS4C_RecalculateAggregateCancel ws
End Sub

Private Function HRS4C_SetCacheCancelAll(ByVal sheetName As String, _
                                          ByVal cancelColumn As Long, _
                                          ByVal turnOn As Boolean) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    If Not HRS4C_SheetExists(sheetName) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = HRS4C_LastUsedRow(ws)

    For r = 2 To lastRow
        If HRS4C_RowHasData(ws, r) Then
            ws.Cells(r, cancelColumn).value = turnOn
            HRS4C_SetCacheCancelAll = HRS4C_SetCacheCancelAll + 1
        End If
    Next r
End Function

Private Sub HRS4C_RebuildAggregateCancelFromNormal()

    Dim wsNormal As Worksheet
    Dim wsAggregate As Worksheet
    Dim lastAgg As Long
    Dim r As Long
    Dim rowList As String

    If Not HRS4C_SheetExists(SH_V4_NORMAL) Then Exit Sub
    If Not HRS4C_SheetExists(SH_V4_AGGREGATE) Then Exit Sub

    Set wsNormal = ThisWorkbook.Worksheets(SH_V4_NORMAL)
    Set wsAggregate = ThisWorkbook.Worksheets(SH_V4_AGGREGATE)
    lastAgg = HRS4C_LastRow(wsAggregate, AC_PRODUCT_NAME)

    For r = 2 To lastAgg
        rowList = Trim$(CStr(wsAggregate.Cells(r, AC_RAW_ROW_LIST).value))
        If rowList <> "" Then
            wsAggregate.Cells(r, AC_CANCEL).value = _
                HRS4C_NormalRowsAllCancelled(wsNormal, rowList)
        End If
    Next r
End Sub

Private Function HRS4C_NormalRowsAllCancelled(ByVal wsNormal As Worksheet, _
                                               ByVal rowList As String) As Boolean

    Dim parts As Variant
    Dim item As Variant
    Dim rawRow As Long
    Dim foundRow As Long
    Dim hasRow As Boolean

    HRS4C_NormalRowsAllCancelled = True
    parts = Split(rowList, ",")

    For Each item In parts
        rawRow = CLng(Val(item))
        If rawRow > 0 Then
            foundRow = HRS4C_FindRawRow(wsNormal, rawRow)
            If foundRow > 0 Then
                hasRow = True
                If Not HRS4C_ToBoolean(wsNormal.Cells(foundRow, NC_CANCEL).value) Then
                    HRS4C_NormalRowsAllCancelled = False
                    Exit Function
                End If
            End If
        End If
    Next item

    If Not hasRow Then HRS4C_NormalRowsAllCancelled = False
End Function

Private Function HRS4C_FindRawRow(ByVal ws As Worksheet, _
                                   ByVal rawRow As Long) As Long

    Dim hit As Range

    Set hit = ws.Columns(NC_RAW_ROW).Find(What:=rawRow, _
              After:=ws.Cells(1, NC_RAW_ROW), LookIn:=xlValues, _
              LookAt:=xlWhole, SearchOrder:=xlByRows, _
              SearchDirection:=xlNext, MatchCase:=False)

    If Not hit Is Nothing Then HRS4C_FindRawRow = hit.Row
End Function

Private Sub HRS4C_SetWriteBackCancelAll(ByVal turnOn As Boolean)

    Dim ws As Worksheet
    Dim cancelCol As Long
    Dim changedCol As Long
    Dim lastRow As Long
    Dim r As Long

    If Not HRS4C_SheetExists(SH_V4_WRITEBACK) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_V4_WRITEBACK)

    cancelCol = HRS4C_FindHeaderColumn(ws, Array("取消", "取消状態", "取消フラグ"))
    changedCol = HRS4C_FindHeaderColumn(ws, Array("変更", "変更状態", "変更フラグ"))
    If cancelCol = 0 Then Exit Sub

    lastRow = HRS4C_LastUsedRow(ws)
    For r = 2 To lastRow
        If HRS4C_RowHasData(ws, r) Then
            ws.Cells(r, cancelCol).value = turnOn
            If changedCol > 0 Then ws.Cells(r, changedCol).value = True
        End If
    Next r
End Sub

Private Sub HRS4C_SetWriteBackPointOne()

    Dim ws As Worksheet
    Dim usageCol As Long
    Dim cancelCol As Long
    Dim changedCol As Long
    Dim lastRow As Long
    Dim r As Long

    If Not HRS4C_SheetExists(SH_V4_WRITEBACK) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_V4_WRITEBACK)

    usageCol = HRS4C_FindHeaderColumn(ws, Array("使用数量", "使用量", "数量"))
    cancelCol = HRS4C_FindHeaderColumn(ws, Array("取消", "取消状態", "取消フラグ"))
    changedCol = HRS4C_FindHeaderColumn(ws, Array("変更", "変更状態", "変更フラグ"))
    If usageCol = 0 Or cancelCol = 0 Then Exit Sub

    lastRow = HRS4C_LastUsedRow(ws)
    For r = 2 To lastRow
        If HRS4C_IsPointOne(ws.Cells(r, usageCol).value) Then
            ws.Cells(r, cancelCol).value = True
            If changedCol > 0 Then ws.Cells(r, changedCol).value = True
        End If
    Next r
End Sub

Private Sub HRS4C_RefreshPreviewFromNormalCache()
    ' 表示中商品はセッション復元または旧描画処理側で更新される。
    ' 現段階ではプレビュー上の0.1も同じ状態に揃える。
    HRS4C_CancelPointOnePreview
End Sub

Private Sub HRS4C_AfterCancelChange(ByVal flushSession As Boolean)

    On Error Resume Next

    ' 現在のプレビュー状態をV4セッションへ保存
    Application.Run "HRS4S_SaveCurrentPreview", Empty, True, flushSession

    ' 既存描画・書戻しプレビューへ接続（存在する場合のみ）
    Application.Run "HRS_RenderPreviewPage"
    Application.Run "HRS_BuildCurrentWriteBackPreview"

    HRS4C_UpdateCancelDisplay
    On Error GoTo 0
End Sub

Private Sub HRS4C_ApplyGray(ByVal targetRange As Range, _
                             ByVal isCancelled As Boolean)
    With targetRange
        If isCancelled Then
            .Font.Color = RGB(128, 128, 128)
            .Interior.Color = RGB(230, 230, 230)
        Else
            .Font.ColorIndex = xlAutomatic
            .Interior.Pattern = xlNone
        End If
    End With
End Sub

Private Function HRS4C_DisplayToSourceRow(ByVal wsPreview As Worksheet, _
                                           ByVal displayRow As Long) As Long
    Dim pos As Long
    pos = CLng(Val(wsPreview.Range("N2").value))
    HRS4C_DisplayToSourceRow = pos + (displayRow - PREVIEW_TOP) + 1
End Function

Private Function HRS4C_IsAggregateMode(ByVal wsPreview As Worksheet) As Boolean
    HRS4C_IsAggregateMode = _
        (UCase$(Trim$(CStr(wsPreview.Range("N7").value))) = "TRUE")
End Function

Private Function HRS4C_IsPointOne(ByVal valueData As Variant) As Boolean
    If IsError(valueData) Or IsEmpty(valueData) Then Exit Function
    If Not IsNumeric(valueData) Then Exit Function
    HRS4C_IsPointOne = (Abs(CDbl(valueData) - 0.1) < 0.0000001)
End Function

Private Function HRS4C_FindHeaderColumn(ByVal ws As Worksheet, _
                                         ByVal headerCandidates As Variant) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim candidate As Variant
    Dim headerText As String

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For c = 1 To lastCol
        headerText = Trim$(CStr(ws.Cells(1, c).value))
        For Each candidate In headerCandidates
            If StrComp(headerText, CStr(candidate), vbTextCompare) = 0 Then
                HRS4C_FindHeaderColumn = c
                Exit Function
            End If
        Next candidate
    Next c
End Function

Private Function HRS4C_RowHasData(ByVal ws As Worksheet, _
                                   ByVal rowNo As Long) As Boolean
    HRS4C_RowHasData = (Application.CountA(ws.Rows(rowNo)) > 0)
End Function

Private Function HRS4C_LastRow(ByVal ws As Worksheet, _
                                ByVal columnNo As Long) As Long
    HRS4C_LastRow = ws.Cells(ws.Rows.count, columnNo).End(xlUp).Row
End Function

Private Function HRS4C_LastUsedRow(ByVal ws As Worksheet) As Long

    Dim hit As Range

    Set hit = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
              LookIn:=xlFormulas, LookAt:=xlPart, _
              SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
              MatchCase:=False)

    If hit Is Nothing Then
        HRS4C_LastUsedRow = 1
    Else
        HRS4C_LastUsedRow = hit.Row
    End If
End Function

Private Function HRS4C_GetSheet(ByVal sheetName As String) As Worksheet
    If Not HRS4C_SheetExists(sheetName) Then
        Err.Raise vbObjectError + 4102, "HRS4C_GetSheet", _
                  "シート「" & sheetName & "」が見つかりません。"
    End If
    Set HRS4C_GetSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Private Function HRS4C_SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    HRS4C_SheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Private Function HRS4C_ToBoolean(ByVal valueData As Variant) As Boolean

    Dim textValue As String

    If VarType(valueData) = vbBoolean Then
        HRS4C_ToBoolean = CBool(valueData)
        Exit Function
    End If

    If IsNumeric(valueData) Then
        HRS4C_ToBoolean = (CDbl(valueData) <> 0)
        Exit Function
    End If

    textValue = UCase$(Trim$(CStr(valueData)))
    HRS4C_ToBoolean = _
        (textValue = "TRUE" Or textValue = "■" Or textValue = "取消")
End Function

Private Sub HRS4C_BeginFast()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub

Private Sub HRS4C_EndFast()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function HRS4C_Elapsed(ByVal startedAt As Double) As Double
    Dim currentValue As Double
    currentValue = Timer
    If currentValue < startedAt Then currentValue = currentValue + 86400#
    HRS4C_Elapsed = currentValue - startedAt
End Function

Private Sub HRS4C_WritePerformance(ByVal processName As String, _
                                    ByVal elapsedSeconds As Double, _
                                    ByVal recordCount As Long, _
                                    ByVal detailText As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    On Error Resume Next
    If Not HRS4C_SheetExists(SH_PERFORMANCE) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SH_PERFORMANCE)
    nextRow = HRS4C_LastUsedRow(ws) + 1

    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = processName
    ws.Cells(nextRow, 3).value = elapsedSeconds
    ws.Cells(nextRow, 4).value = recordCount
    ws.Cells(nextRow, 5).value = detailText
    ws.Cells(nextRow, 6).value = HRS4C_VERSION
    On Error GoTo 0
End Sub
