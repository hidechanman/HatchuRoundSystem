Attribute VB_Name = "modHRS4_Core"
Option Explicit

'=========================================================
' 発注まるめシステム Ver4.0 Part1
' 高速データエンジン 共通処理
'=========================================================

Public Const HRS4_APP_VERSION As String = "Ver4.0 Part1"

Public Const HRS4_SH_RAW As String = "発注原票DB"
Public Const HRS4_SH_NORMAL As String = "V4通常表示キャッシュ"
Public Const HRS4_SH_AGGREGATE As String = "V4集約表示キャッシュ"
Public Const HRS4_SH_WRITEBACK As String = "V4書戻しキャッシュ"
Public Const HRS4_SH_SESSION As String = "V4セッションキャッシュ"
Public Const HRS4_SH_INDEX As String = "V4商品索引"
Public Const HRS4_SH_PERFORMANCE As String = "V4速度ログ"

Private mFastDepth As Long
Private mOldCalculation As XlCalculation
Private mOldScreenUpdating As Boolean
Private mOldEnableEvents As Boolean
Private mOldDisplayAlerts As Boolean
Private mOldStatusBar As Variant

Public Sub HRS4_BeginFast(ByVal statusText As String)

    If mFastDepth = 0 Then
        mOldCalculation = Application.Calculation
        mOldScreenUpdating = Application.ScreenUpdating
        mOldEnableEvents = Application.EnableEvents
        mOldDisplayAlerts = Application.DisplayAlerts
        mOldStatusBar = Application.StatusBar

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    End If

    mFastDepth = mFastDepth + 1
    Application.StatusBar = statusText

End Sub

Public Sub HRS4_EndFast()

    If mFastDepth <= 0 Then Exit Sub

    mFastDepth = mFastDepth - 1

    If mFastDepth = 0 Then
        Application.Calculation = mOldCalculation
        Application.ScreenUpdating = mOldScreenUpdating
        Application.EnableEvents = mOldEnableEvents
        Application.DisplayAlerts = mOldDisplayAlerts
        Application.StatusBar = mOldStatusBar
    End If

End Sub

Public Function HRS4_GetOrCreateSheet( _
    ByVal sheetName As String, _
    Optional ByVal veryHidden As Boolean = True) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
    End If

    If veryHidden Then
        ws.Visible = xlSheetVeryHidden
    Else
        ws.Visible = xlSheetVisible
    End If

    Set HRS4_GetOrCreateSheet = ws

End Function

Public Function HRS4_LastRow( _
    ByVal ws As Worksheet, _
    ByVal columnNumber As Long) As Long

    Dim lastCell As Range

    Set lastCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If lastCell Is Nothing Then
        HRS4_LastRow = 1
    Else
        HRS4_LastRow = lastCell.Row
    End If

End Function

Public Function HRS4_NormalizeCode(ByVal value As Variant) As String

    Dim textValue As String

    If IsError(value) Or IsEmpty(value) Then Exit Function

    textValue = Trim$(CStr(value))

    If textValue = "" Then Exit Function

    If IsNumeric(textValue) Then
        HRS4_NormalizeCode = Format$(CDbl(textValue), "0")
    Else
        HRS4_NormalizeCode = textValue
    End If

End Function

Public Function HRS4_ProductIdentity( _
    ByVal productCode As Variant, _
    ByVal productName As Variant) As String

    Dim codeText As String
    Dim nameText As String

    codeText = HRS4_NormalizeCode(productCode)
    nameText = UCase$(Trim$(CStr(productName)))

    If codeText <> "" Then
        HRS4_ProductIdentity = "C|" & codeText
    Else
        HRS4_ProductIdentity = "N|" & nameText
    End If

End Function

Public Function HRS4_DateKey(ByVal value As Variant) As String

    If IsError(value) Or IsEmpty(value) Then Exit Function
    If Trim$(CStr(value)) = "" Then Exit Function

    If IsDate(value) Then
        HRS4_DateKey = Format$(CDate(value), "yyyy/mm/dd")
    Else
        HRS4_DateKey = Trim$(CStr(value))
    End If

End Function

Public Function HRS4_IsMorning(ByVal mealText As Variant) As Boolean

    Dim textValue As String

    textValue = UCase$(Trim$(CStr(mealText)))

    HRS4_IsMorning = _
        (InStr(textValue, "朝") > 0) Or _
        (InStr(textValue, "MORNING") > 0)

End Function

Public Function HRS4_ElapsedSeconds(ByVal startTimer As Double) As Double

    Dim currentTimer As Double

    currentTimer = Timer

    If currentTimer < startTimer Then
        currentTimer = currentTimer + 86400#
    End If

    HRS4_ElapsedSeconds = currentTimer - startTimer

End Function

Public Sub HRS4_WritePerformanceLog( _
    ByVal processName As String, _
    ByVal elapsedSeconds As Double, _
    ByVal recordCount As Long, _
    Optional ByVal noteText As String = "")

    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = HRS4_GetOrCreateSheet(HRS4_SH_PERFORMANCE, False)

    If Trim$(CStr(ws.Cells(1, 1).value)) = "" Then
        ws.Range("A1:E1").value = Array( _
            "日時", "処理", "秒数", "件数", "備考")
        ws.Rows(1).Font.Bold = True
    End If

    nextRow = HRS4_LastRow(ws, 1) + 1

    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = processName
    ws.Cells(nextRow, 3).value = Round(elapsedSeconds, 3)
    ws.Cells(nextRow, 4).value = recordCount
    ws.Cells(nextRow, 5).value = noteText

    ws.Columns("A:E").AutoFit

End Sub

Public Sub HRS4_ResetExcelState()

    mFastDepth = 0

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False

    MsgBox "Excelの動作状態を復旧しました。", _
           vbInformation, "発注まるめシステム"

End Sub
