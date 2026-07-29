Attribute VB_Name = "modHRS5_WriteBack"
Option Explicit

'============================================================
' 発注まるめシステム Ver5.0.2
' Module  : modHRS5_WriteBack
' Purpose : 配分セッションDBの内容を元の発注書へ書き戻す
'
' 方針:
' ・旧 modHRS_Final と併存できるよう、公開名は HRS5WB_ で統一
' ・外部参照設定を追加しない
' ・書込前検証、確認、取消線反映、履歴保存、ログ記録を一元化
'============================================================

Private Const APP_NAME As String = "発注まるめシステム"
Private Const MODULE_VERSION As String = "Ver5.0.2"

Private Const SH_SESSION As String = "配分セッションDB"
Private Const SH_HISTORY As String = "配分履歴DB"
Private Const SH_LOG As String = "ログ"

' 配分セッションDB列
Private Const SC_UPDATED_AT As Long = 1
Private Const SC_VENDOR_NAME As Long = 2
Private Const SC_PRODUCT_CODE As Long = 3
Private Const SC_PRODUCT_NAME As Long = 4
Private Const SC_ORDER_QTY As Long = 5
Private Const SC_CONFIRMED As Long = 6
Private Const SC_CANCELLED As Long = 7
Private Const SC_DELIVERY_DATE As Long = 8
Private Const SC_USE_DATE As Long = 9
Private Const SC_MEAL As Long = 10
Private Const SC_ORIGINAL_QTY As Long = 11
Private Const SC_DISTRIBUTION_QTY As Long = 12
Private Const SC_SOURCE_ROW As Long = 13
Private Const SC_SOURCE_SHEET As Long = 14
Private Const SC_SOURCE_CELL As Long = 15
Private Const SC_SOURCE_BOOK As Long = 16
Private Const SESSION_COLS As Long = 16

Private mBusy As Boolean
Private mLastWritten As Long
Private mLastNotFound As Long
Private mLastSkipped As Long
Private mLastError As String

'------------------------------------------------------------
' 公開API
'------------------------------------------------------------
Public Sub HRS5WB_ShowVersion()
    MsgBox "modHRS5_WriteBack " & MODULE_VERSION, vbInformation, APP_NAME
End Sub

Public Function HRS5WB_IsBusy() As Boolean
    HRS5WB_IsBusy = mBusy
End Function

Public Function HRS5WB_GetLastWrittenCount() As Long
    HRS5WB_GetLastWrittenCount = mLastWritten
End Function

Public Function HRS5WB_GetLastNotFoundCount() As Long
    HRS5WB_GetLastNotFoundCount = mLastNotFound
End Function

Public Function HRS5WB_GetLastSkippedCount() As Long
    HRS5WB_GetLastSkippedCount = mLastSkipped
End Function

Public Function HRS5WB_GetLastError() As String
    HRS5WB_GetLastError = mLastError
End Function

' 画面ボタンおよびControllerから呼び出す入口
Public Sub HRS5WB_WriteBack()

    Dim wsSession As Worksheet
    Dim validCount As Long
    Dim answer As VbMsgBoxResult

    If mBusy Then
        MsgBox "書き戻し処理はすでに実行中です。", vbExclamation, APP_NAME
        Exit Sub
    End If

    mBusy = True
    mLastWritten = 0
    mLastNotFound = 0
    mLastSkipped = 0
    mLastError = vbNullString

    On Error GoTo ErrHandler

    If Not HRS5WB_SheetExists(SH_SESSION, ThisWorkbook) Then
        MsgBox "「" & SH_SESSION & "」シートがありません。" & vbCrLf & _
               "先に読込・配分処理を実行してください。", _
               vbExclamation, APP_NAME
        GoTo ExitHandler
    End If

    Set wsSession = ThisWorkbook.Worksheets(SH_SESSION)
    validCount = HRS5WB_CountWritableRows(wsSession)

    If validCount = 0 Then
        MsgBox "書き戻せる配分データがありません。" & vbCrLf & _
               "商品を選択し、発注数を配分してから実行してください。", _
               vbInformation, APP_NAME
        GoTo ExitHandler
    End If

    answer = MsgBox( _
        "元の発注書へ配分結果を書き戻します。" & vbCrLf & vbCrLf & _
        "対象データ：" & validCount & "件" & vbCrLf & _
        "・配分値を対象セルの1行下へ入力" & vbCrLf & _
        "・取消状態を元の数量セルへ反映" & vbCrLf & _
        "・書き戻し後に配分履歴DBへ保存" & vbCrLf & vbCrLf & _
        "元の発注書ブックが開いていることを確認してください。" & vbCrLf & _
        "このまま実行しますか？", _
        vbYesNo + vbQuestion + vbDefaultButton2, APP_NAME)

    If answer <> vbYes Then
        HRS5WB_WriteLog "WRITEBACK_CANCEL", "利用者が書き戻しを中止"
        GoTo ExitHandler
    End If

    HRS5WB_BeginFast "発注書へ書き戻しています..."
    HRS5WB_WriteRows wsSession

    If mLastWritten > 0 Then
        HRS5WB_ArchiveSession wsSession
    End If

    HRS5WB_WriteLog "WRITEBACK", _
        "書込=" & mLastWritten & _
        "、未確認=" & mLastNotFound & _
        "、スキップ=" & mLastSkipped

ExitHandler:
    HRS5WB_EndFast
    mBusy = False

    If Len(mLastError) = 0 And validCount > 0 And answer = vbYes Then
        MsgBox "書き戻しが完了しました。" & vbCrLf & vbCrLf & _
               "書込：" & mLastWritten & "件" & vbCrLf & _
               "対象ブック・シート未確認：" & mLastNotFound & "件" & vbCrLf & _
               "入力対象外：" & mLastSkipped & "件", _
               vbInformation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    mLastError = CStr(Err.Number) & " : " & Err.Description
    HRS5WB_WriteLog "ERROR:HRS5WB_WriteBack", mLastError
    MsgBox "書き戻し処理でエラーが発生しました。" & vbCrLf & _
           mLastError, vbCritical, APP_NAME
    Resume ExitHandler

End Sub

' Controllerから確認なしで呼ぶ内部向け入口
' 戻り値 True: 1件以上書込 / False: 書込なしまたはエラー
Public Function HRS5WB_WriteBackSilent() As Boolean

    Dim wsSession As Worksheet

    If mBusy Then Exit Function

    mBusy = True
    mLastWritten = 0
    mLastNotFound = 0
    mLastSkipped = 0
    mLastError = vbNullString

    On Error GoTo ErrHandler

    If Not HRS5WB_SheetExists(SH_SESSION, ThisWorkbook) Then GoTo ExitHandler

    Set wsSession = ThisWorkbook.Worksheets(SH_SESSION)
    If HRS5WB_CountWritableRows(wsSession) = 0 Then GoTo ExitHandler

    HRS5WB_BeginFast "発注書へ書き戻しています..."
    HRS5WB_WriteRows wsSession

    If mLastWritten > 0 Then
        HRS5WB_ArchiveSession wsSession
        HRS5WB_WriteBackSilent = True
    End If

    HRS5WB_WriteLog "WRITEBACK_SILENT", _
        "書込=" & mLastWritten & _
        "、未確認=" & mLastNotFound & _
        "、スキップ=" & mLastSkipped

ExitHandler:
    HRS5WB_EndFast
    mBusy = False
    Exit Function

ErrHandler:
    mLastError = CStr(Err.Number) & " : " & Err.Description
    HRS5WB_WriteLog "ERROR:HRS5WB_WriteBackSilent", mLastError
    Resume ExitHandler

End Function

'------------------------------------------------------------
' 書き戻し本体
'------------------------------------------------------------
Private Sub HRS5WB_WriteRows(ByVal wsSession As Worksheet)

    Dim lastRow As Long
    Dim r As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim writeCell As Range
    Dim sourceBookName As String
    Dim sourceSheetName As String
    Dim cellAddress As String
    Dim distributionValue As Variant
    Dim isCancelled As Boolean

    lastRow = HRS5WB_LastRow(wsSession, SC_UPDATED_AT)

    For r = 2 To lastRow

        If Not HRS5WB_IsWritableRow(wsSession, r) Then
            mLastSkipped = mLastSkipped + 1
            GoTo ContinueRow
        End If

        sourceBookName = Trim$(CStr(wsSession.Cells(r, SC_SOURCE_BOOK).Value2))
        sourceSheetName = Trim$(CStr(wsSession.Cells(r, SC_SOURCE_SHEET).Value2))
        cellAddress = Trim$(CStr(wsSession.Cells(r, SC_SOURCE_CELL).Value2))
        distributionValue = wsSession.Cells(r, SC_DISTRIBUTION_QTY).Value
        isCancelled = HRS5WB_ToBoolean(wsSession.Cells(r, SC_CANCELLED).Value)

        Set wb = HRS5WB_FindOpenWorkbook(sourceBookName)

        If wb Is Nothing Then
            mLastNotFound = mLastNotFound + 1
            HRS5WB_WriteLog "WRITEBACK_NOT_FOUND", _
                "ブック未確認：" & sourceBookName & _
                " / " & sourceSheetName & "!" & cellAddress
            GoTo ContinueRow
        End If

        If Not HRS5WB_SheetExists(sourceSheetName, wb) Then
            mLastNotFound = mLastNotFound + 1
            HRS5WB_WriteLog "WRITEBACK_NOT_FOUND", _
                "シート未確認：" & sourceBookName & _
                " / " & sourceSheetName
            GoTo ContinueRow
        End If

        Set ws = wb.Worksheets(sourceSheetName)
        Set targetCell = HRS5WB_GetRangeSafely(ws, cellAddress)

        If targetCell Is Nothing Then
            mLastNotFound = mLastNotFound + 1
            HRS5WB_WriteLog "WRITEBACK_NOT_FOUND", _
                "セル未確認：" & sourceBookName & _
                " / " & sourceSheetName & "!" & cellAddress
            GoTo ContinueRow
        End If

        targetCell.Font.Strikethrough = isCancelled

        Set writeCell = targetCell.Offset(1, 0)
        writeCell.Value = distributionValue
        writeCell.HorizontalAlignment = xlCenter
        writeCell.VerticalAlignment = xlCenter
        writeCell.Font.Bold = True
        writeCell.NumberFormat = "General"

        mLastWritten = mLastWritten + 1

ContinueRow:
        Set writeCell = Nothing
        Set targetCell = Nothing
        Set ws = Nothing
        Set wb = Nothing
    Next r

End Sub

Private Function HRS5WB_CountWritableRows(ByVal wsSession As Worksheet) As Long

    Dim lastRow As Long
    Dim r As Long

    lastRow = HRS5WB_LastRow(wsSession, SC_UPDATED_AT)

    For r = 2 To lastRow
        If HRS5WB_IsWritableRow(wsSession, r) Then
            HRS5WB_CountWritableRows = HRS5WB_CountWritableRows + 1
        End If
    Next r

End Function

Private Function HRS5WB_IsWritableRow( _
    ByVal wsSession As Worksheet, _
    ByVal rowNumber As Long) As Boolean

    Dim sourceBookName As String
    Dim sourceSheetName As String
    Dim cellAddress As String
    Dim distributionValue As Variant
    Dim isCancelled As Boolean

    sourceBookName = Trim$(CStr(wsSession.Cells(rowNumber, SC_SOURCE_BOOK).Value2))
    sourceSheetName = Trim$(CStr(wsSession.Cells(rowNumber, SC_SOURCE_SHEET).Value2))
    cellAddress = Trim$(CStr(wsSession.Cells(rowNumber, SC_SOURCE_CELL).Value2))
    distributionValue = wsSession.Cells(rowNumber, SC_DISTRIBUTION_QTY).Value
    isCancelled = HRS5WB_ToBoolean(wsSession.Cells(rowNumber, SC_CANCELLED).Value)

    If Len(sourceBookName) = 0 Then Exit Function
    If Len(sourceSheetName) = 0 Then Exit Function
    If Len(cellAddress) = 0 Then Exit Function

    ' 配分値が空でも取消線だけを書き戻す必要がある場合は対象とする。
    If Len(Trim$(CStr(distributionValue))) > 0 Or isCancelled Then
        HRS5WB_IsWritableRow = True
    End If

End Function

'------------------------------------------------------------
' 履歴保存
'------------------------------------------------------------
Private Sub HRS5WB_ArchiveSession(ByVal wsSession As Worksheet)

    Dim wsHistory As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim sourceData As Variant
    Dim outputData() As Variant
    Dim r As Long
    Dim c As Long
    Dim rowCount As Long

    lastRow = HRS5WB_LastRow(wsSession, SC_UPDATED_AT)
    If lastRow < 2 Then Exit Sub

    Set wsHistory = HRS5WB_EnsureHistorySheet()

    sourceData = wsSession.Range( _
        wsSession.Cells(2, 1), _
        wsSession.Cells(lastRow, SESSION_COLS)).Value

    For r = 1 To UBound(sourceData, 1)
        If HRS5WB_RowArrayHasData(sourceData, r) Then
            rowCount = rowCount + 1
        End If
    Next r

    If rowCount = 0 Then Exit Sub

    ReDim outputData(1 To rowCount, 1 To SESSION_COLS)
    rowCount = 0

    For r = 1 To UBound(sourceData, 1)
        If HRS5WB_RowArrayHasData(sourceData, r) Then
            rowCount = rowCount + 1
            For c = 1 To SESSION_COLS
                outputData(rowCount, c) = sourceData(r, c)
            Next c
        End If
    Next r

    targetRow = HRS5WB_LastRow(wsHistory, 1) + 1
    If targetRow < 2 Then targetRow = 2

    wsHistory.Cells(targetRow, 1).Resize(rowCount, SESSION_COLS).Value = outputData
    wsHistory.Columns(SC_PRODUCT_CODE).NumberFormat = "@"
    wsHistory.Columns(SC_UPDATED_AT).NumberFormat = "yyyy/mm/dd hh:mm:ss"

End Sub

Private Function HRS5WB_EnsureHistorySheet() As Worksheet

    Dim ws As Worksheet

    If HRS5WB_SheetExists(SH_HISTORY, ThisWorkbook) Then
        Set HRS5WB_EnsureHistorySheet = ThisWorkbook.Worksheets(SH_HISTORY)
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SH_HISTORY

    ws.Cells(1, 1).Resize(1, SESSION_COLS).Value = Array( _
        "更新日時", "業者名", "商品番号", "商品名", _
        "発注数", "確認済", "取消", "納品日", _
        "使用日", "食事区分", "元数量", "配分数量", _
        "元行", "元シート", "元セル", "元ブック")

    ws.Rows(1).Font.Bold = True
    ws.Rows(1).AutoFilter
    ws.Columns(SC_PRODUCT_CODE).NumberFormat = "@"

    Set HRS5WB_EnsureHistorySheet = ws

End Function

Private Function HRS5WB_RowArrayHasData( _
    ByRef sourceData As Variant, _
    ByVal rowNumber As Long) As Boolean

    HRS5WB_RowArrayHasData = _
        Len(Trim$(CStr(sourceData(rowNumber, SC_VENDOR_NAME)))) > 0 Or _
        Len(Trim$(CStr(sourceData(rowNumber, SC_PRODUCT_CODE)))) > 0 Or _
        Len(Trim$(CStr(sourceData(rowNumber, SC_PRODUCT_NAME)))) > 0

End Function

'------------------------------------------------------------
' 共通処理
'------------------------------------------------------------
Private Function HRS5WB_FindOpenWorkbook( _
    ByVal bookName As String) As Workbook

    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, bookName, vbTextCompare) = 0 Then
            Set HRS5WB_FindOpenWorkbook = wb
            Exit Function
        End If
    Next wb

End Function

Private Function HRS5WB_GetRangeSafely( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String) As Range

    On Error Resume Next
    Set HRS5WB_GetRangeSafely = ws.Range(cellAddress)
    On Error GoTo 0

End Function

Private Function HRS5WB_SheetExists( _
    ByVal sheetName As String, _
    ByVal wb As Workbook) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    HRS5WB_SheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0

End Function

Private Function HRS5WB_LastRow( _
    ByVal ws As Worksheet, _
    ByVal columnNumber As Long) As Long

    Dim result As Long

    result = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row
    If result < 1 Then result = 1
    HRS5WB_LastRow = result

End Function

Private Function HRS5WB_ToBoolean(ByVal sourceValue As Variant) As Boolean

    Dim textValue As String

    If VarType(sourceValue) = vbBoolean Then
        HRS5WB_ToBoolean = CBool(sourceValue)
        Exit Function
    End If

    If IsNumeric(sourceValue) Then
        HRS5WB_ToBoolean = (CDbl(sourceValue) <> 0)
        Exit Function
    End If

    textValue = UCase$(Trim$(CStr(sourceValue)))

    Select Case textValue
        Case "TRUE", "YES", "Y", "ON", "1", "■", "取消", "取消済"
            HRS5WB_ToBoolean = True
    End Select

End Function

Private Sub HRS5WB_BeginFast(ByVal statusText As String)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = statusText

End Sub

Private Sub HRS5WB_EndFast()

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Private Sub HRS5WB_WriteLog( _
    ByVal processName As String, _
    ByVal detailText As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    On Error GoTo SafeExit

    If Not HRS5WB_SheetExists(SH_LOG, ThisWorkbook) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    nextRow = HRS5WB_LastRow(ws, 1) + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = processName
    ws.Cells(nextRow, 3).Value = detailText
    ws.Cells(nextRow, 4).Value = Environ$("Username")

SafeExit:
End Sub
