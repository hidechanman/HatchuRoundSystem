Attribute VB_Name = "modProductNameSettings"
Option Explicit

'============================================================
' 発注まるめシステム
' Module  : modProductNameSettings
' Version : Ver1.3.4.9
' Purpose : 商品名をキーにした表記ゆれ管理
'============================================================

Private Const SH_NAME_SETTINGS As String = "商品名設定"

Public Sub HRS_SetupProductNameSettings()

    Dim ws As Worksheet

    On Error GoTo ErrHandler

    If Not HRS_NameSheetExists(SH_NAME_SETTINGS) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = SH_NAME_SETTINGS
    Else
        Set ws = ThisWorkbook.Worksheets(SH_NAME_SETTINGS)
    End If

    ws.Range("A1").value = "発注書の商品名"
    ws.Range("B1").value = "管理名"
    ws.Range("C1").value = "使用"
    ws.Range("D1").value = "備考"


    With ws.Range("A1:D1")
        .Font.Bold = True
        .Font.Color = vbBlack
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With

    HRS_AddDefaultNameRow ws, "牛乳200ml", "牛乳200ml", "特殊発注商品"
    HRS_AddDefaultNameRow ws, "牛乳 200ml", "牛乳200ml", "表記ゆれ"
    HRS_AddDefaultNameRow ws, "牛乳1L", "牛乳1L", "特殊発注商品"
    HRS_AddDefaultNameRow ws, "牛乳 1L", "牛乳1L", "表記ゆれ"
    HRS_AddDefaultNameRow ws, "牛乳（1L）", "牛乳1L", "表記ゆれ"
    HRS_AddDefaultNameRow ws, "牛乳1000ml", "牛乳1L", "表記ゆれ"
    HRS_AddDefaultNameRow ws, "ヨーグルト", "ヨーグルト", "特殊発注商品"

    ws.Columns("A:B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 9
    ws.Columns("D").ColumnWidth = 24
    ws.Columns("A:D").Font.Color = vbBlack
    ws.Columns("A:D").VerticalAlignment = xlCenter

    Exit Sub

ErrHandler:
    MsgBox "商品名設定シートの作成中にエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, vbCritical, "発注まるめシステム"
End Sub

Public Function HRS_GetManagedProductName(ByVal sourceName As String) As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim sourceKey As String
    Dim rowKey As String
    Dim enabledValue As String

    sourceName = Trim$(sourceName)
    If sourceName = "" Then Exit Function

    If Not HRS_NameSheetExists(SH_NAME_SETTINGS) Then
        HRS_GetManagedProductName = sourceName
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(SH_NAME_SETTINGS)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    sourceKey = HRS_NormalizeNameKey(sourceName)

    For r = 2 To lastRow
        rowKey = HRS_NormalizeNameKey(CStr(ws.Cells(r, "A").value))
        enabledValue = UCase$(Trim$(CStr(ws.Cells(r, "C").value)))

        If rowKey = sourceKey And enabledValue <> "OFF" And enabledValue <> "FALSE" And enabledValue <> "0" Then
            If Trim$(CStr(ws.Cells(r, "B").value)) <> "" Then
                HRS_GetManagedProductName = Trim$(CStr(ws.Cells(r, "B").value))
            Else
                HRS_GetManagedProductName = sourceName
            End If
            Exit Function
        End If
    Next r

    HRS_GetManagedProductName = sourceName

End Function

Private Sub HRS_AddDefaultNameRow(ByVal ws As Worksheet, _
                                  ByVal sourceName As String, _
                                  ByVal managedName As String, _
                                  ByVal noteText As String)

    Dim lastRow As Long
    Dim r As Long
    Dim targetRow As Long
    Dim sourceKey As String

    sourceKey = HRS_NormalizeNameKey(sourceName)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    For r = 2 To lastRow
        If HRS_NormalizeNameKey(CStr(ws.Cells(r, "A").value)) = sourceKey Then
            Exit Sub
        End If
    Next r

    targetRow = lastRow + 1
    If targetRow < 2 Then targetRow = 2

    ws.Cells(targetRow, "A").value = sourceName
    ws.Cells(targetRow, "B").value = managedName
    ws.Cells(targetRow, "C").value = "ON"
    ws.Cells(targetRow, "D").value = noteText
    ws.Range("A" & targetRow & ":D" & targetRow).Borders.LineStyle = xlContinuous
    ws.Range("A" & targetRow & ":D" & targetRow).Font.Color = vbBlack

End Sub

Private Function HRS_NormalizeNameKey(ByVal valueText As String) As String

    Dim s As String

    s = Trim$(valueText)
    s = Replace(s, "　", "")
    s = Replace(s, " ", "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, "Ｌ", "L")
    s = Replace(s, "ｌ", "L")
    s = Replace(s, "（", "(")
    s = Replace(s, "）", ")")

    HRS_NormalizeNameKey = LCase$(s)

End Function

Private Function HRS_NameSheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    HRS_NameSheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0

End Function
