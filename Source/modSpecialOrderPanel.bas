Attribute VB_Name = "modSpecialOrderPanel"
Option Explicit

'============================================================
' 発注まるめシステム
' Module  : modSpecialOrderPanel
' Version : Ver1.3.5.6
' Purpose : 既存レイアウトを変更せず、選択商品の表示欄を使用して
'           特殊商品の定数を商品名で保存する。
'           牛乳1L選択時のみ料理用詳細を右下へ展開する。
'============================================================

Private Const SH_INPUT As String = "発注入力"
Private Const SH_SPECIAL As String = "特殊商品"
Private Const SH_PREVIEW_CACHE As String = "使用日プレビュー作業"

Private Const EDIT_PRODUCT_CELL As String = "A24"
Private Const EDIT_LABEL_CELL As String = "C24"
Private Const EDIT_VALUE_CELL As String = "D24"
Private Const EDIT_STATUS_CELL As String = "F24"
Private Const EDIT_AREA As String = "A24:H24"

Private Const OLD_BTN_TOGGLE As String = "btnSpecialOrderList"
Private Const OLD_BTN_SAVE As String = "btnSpecialOrderSave"

'------------------------------------------------------------
' 初回セットアップ
' ・以前の特殊商品パネル用ボタンを削除
' ・特殊商品シートと商品名設定を整備
' ・既存の「発注時（選択商品）」欄を初期表示へ戻す
' ※セル結合、行挿入、列幅変更は行わない
'------------------------------------------------------------
Public Sub HRS_SetupSpecialOrderPanel()

    Dim wsInput As Worksheet

    On Error GoTo ErrHandler

    If Not HRS_SpecialSheetExists(SH_INPUT) Then
        Err.Raise vbObjectError + 3501, , "「" & SH_INPUT & "」シートがありません。"
    End If

    HRS_EnsureSpecialProductRows
    HRS_SetupProductNameSettings

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)

    On Error Resume Next
    wsInput.Shapes(OLD_BTN_TOGGLE).Delete
    wsInput.Shapes(OLD_BTN_SAVE).Delete
    On Error GoTo ErrHandler

    '旧パネルは表示しない。既存画面の列幅・結合状態は変更しない。
    wsInput.Columns("S:V").Hidden = True

    HRS_ShowSelectedSpecialOrder ""
    HRS_ClearMilkCookingDetail wsInput

    MsgBox "特殊商品の表示方法を更新しました。" & vbCrLf & _
           "商品一覧で牛乳200ml・牛乳1L・ヨーグルトを選択すると、" & vbCrLf & _
           "左下の「発注時（選択商品）」欄に定数入力が表示されます。", _
           vbInformation, "発注まるめシステム"
    Exit Sub

ErrHandler:
    MsgBox "特殊商品表示の初期設定中にエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, _
           vbCritical, "発注まるめシステム"
End Sub

'互換性維持：旧ボタンから呼ばれても選択商品欄を更新するだけ
Public Sub HRS_ToggleSpecialOrderPanel()
    Dim productName As String
    productName = HRS_GetCurrentSelectedProductName()
    HRS_ShowSelectedSpecialOrder productName
End Sub

'------------------------------------------------------------
' 選択商品欄へ特殊商品の定数入力を表示
' 普通の商品では従来の案内文へ戻す
'------------------------------------------------------------
Public Sub HRS_ShowSelectedSpecialOrder(ByVal productName As String)

    Dim ws As Worksheet
    Dim canonicalName As String
    Dim constantValue As Double

    On Error GoTo ExitHandler

    Set ws = ThisWorkbook.Worksheets(SH_INPUT)

    '値と入力用の書式だけを戻す。結合・行挿入・列幅変更は行わない。
    With ws.Range(EDIT_AREA)
        .ClearContents
        .Font.Color = vbBlack
        .Font.Bold = False
        .Interior.Color = RGB(242, 242, 242)
        .NumberFormat = "General"
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With
    ws.Rows("24").Hidden = False
    ws.Rows("24").RowHeight = 34

    If Not HRS_IsSpecialOrderProduct(productName) Then
        Exit Sub
    End If

    canonicalName = HRS_GetCanonicalSpecialName(productName)
    constantValue = HRS_GetSpecialConstant(canonicalName)

    ws.Range(EDIT_PRODUCT_CELL).value = "【特殊】" & canonicalName
    With ws.Range(EDIT_PRODUCT_CELL)
        .Font.Bold = True
        .Font.Color = RGB(192, 0, 0)
        .Font.Size = 12
        .Interior.Color = RGB(226, 239, 218)
    End With

    ws.Range(EDIT_LABEL_CELL).value = "定数入力 →"
    With ws.Range(EDIT_LABEL_CELL)
        .Font.Bold = True
        .Font.Color = vbBlack
        .HorizontalAlignment = xlRight
    End With

    With ws.Range(EDIT_VALUE_CELL)
        .NumberFormat = "General"
        .value = constantValue
        .Interior.Color = RGB(255, 242, 204)
        .Font.Color = vbBlack
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = RGB(112, 173, 71)
    End With

    ws.Range(EDIT_STATUS_CELL).value = "定数入力・自動保存"
    With ws.Range(EDIT_STATUS_CELL)
        .Font.Color = RGB(0, 128, 0)
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
    End With

    If HRS_IsMilk1L(canonicalName) Then
        ws.Range(EDIT_STATUS_CELL).value = "料理用は自動計算"
    Else
        ws.Range(EDIT_STATUS_CELL).value = "定数入力・自動保存"
    End If

ExitHandler:
End Sub

'------------------------------------------------------------
' D23へ入力した定数を商品名で自動保存
' Worksheet_Changeから呼び出す
'------------------------------------------------------------
Public Sub HRS_SaveSelectedSpecialConstant(Optional ByVal showMessage As Boolean = False)

    Dim wsInput As Worksheet
    Dim wsSpecial As Worksheet
    Dim productName As String
    Dim inputValue As Variant
    Dim targetRow As Long

    On Error GoTo ErrHandler

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsSpecial = ThisWorkbook.Worksheets(SH_SPECIAL)

    productName = HRS_GetCurrentSelectedProductName()
    If Not HRS_IsSpecialOrderProduct(productName) Then Exit Sub

    productName = HRS_GetCanonicalSpecialName(productName)
    inputValue = wsInput.Range(EDIT_VALUE_CELL).value

    If Len(Trim$(CStr(inputValue))) = 0 Then inputValue = 0

    If Not IsNumeric(inputValue) Then
        MsgBox productName & "の定数は、0以上の数値で入力してください。", _
               vbExclamation, "発注まるめシステム"
        Exit Sub
    End If

    If CDbl(inputValue) < 0 Then
        MsgBox productName & "の定数は、0以上で入力してください。", _
               vbExclamation, "発注まるめシステム"
        Exit Sub
    End If

    targetRow = HRS_FindSpecialProductRow(wsSpecial, productName)
    If targetRow = 0 Then
        targetRow = wsSpecial.Cells(wsSpecial.Rows.count, "B").End(xlUp).Row + 1
        If targetRow < 2 Then targetRow = 2
        wsSpecial.Cells(targetRow, "B").value = productName
    End If

    wsSpecial.Cells(targetRow, "C").value = "定数"
    wsSpecial.Cells(targetRow, "D").NumberFormat = "General"
    wsSpecial.Cells(targetRow, "D").value = CDbl(inputValue)

    If HRS_IsMilk1L(productName) Then
        wsSpecial.Cells(targetRow, "E").value = "ON"
        wsSpecial.Cells(targetRow, "F").value = _
            "定数のみ入力。料理用は自動計算し、牛乳1L選択時だけ表示。"
    Else
        wsSpecial.Cells(targetRow, "E").ClearContents
    End If

    wsInput.Range(EDIT_STATUS_CELL).value = "保存済"
    wsInput.Range(EDIT_STATUS_CELL).Font.Color = RGB(0, 128, 0)
    wsInput.Range(EDIT_STATUS_CELL).Font.Bold = True

    If showMessage Then
        MsgBox productName & "の定数を保存しました。", _
               vbInformation, "発注まるめシステム"
    End If
    Exit Sub

ErrHandler:
    MsgBox "特殊商品の定数保存中にエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, _
           vbCritical, "発注まるめシステム"
End Sub

'旧マクロ名との互換性
Public Sub HRS_SaveSpecialOrderConstants()
    HRS_SaveSelectedSpecialConstant True
End Sub

'------------------------------------------------------------
' 商品名から保存済み定数を取得
'------------------------------------------------------------
Public Function HRS_GetSpecialConstant(ByVal productName As String) As Double

    Dim ws As Worksheet
    Dim targetRow As Long
    Dim rawValue As Variant

    If Not HRS_SpecialSheetExists(SH_SPECIAL) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SH_SPECIAL)
    targetRow = HRS_FindSpecialProductRow(ws, productName)
    If targetRow = 0 Then Exit Function

    rawValue = ws.Cells(targetRow, "D").value
    HRS_GetSpecialConstant = HRS_ParseNumericValue(rawValue)
End Function

Public Function HRS_IsSpecialOrderProduct(ByVal productName As String) As Boolean
    HRS_IsSpecialOrderProduct = _
        HRS_IsMilk200(productName) Or _
        HRS_IsMilk1L(productName) Or _
        HRS_IsYogurt(productName)
End Function

Public Function HRS_IsMilk200(ByVal productName As String) As Boolean
    Dim s As String
    s = HRS_NormalizeSpecialName(HRS_GetManagedProductName(productName))
    HRS_IsMilk200 = (InStr(s, "牛乳200ml") > 0) Or (InStr(s, "牛乳200") > 0)
End Function

Public Function HRS_IsMilk1L(ByVal productName As String) As Boolean
    Dim s As String
    s = HRS_NormalizeSpecialName(HRS_GetManagedProductName(productName))
    HRS_IsMilk1L = _
        (InStr(s, "牛乳1l") > 0) Or _
        (InStr(s, "牛乳1000ml") > 0) Or _
        (InStr(s, "牛乳1000") > 0)
End Function

Public Function HRS_IsYogurt(ByVal productName As String) As Boolean
    HRS_IsYogurt = _
        (InStr(HRS_NormalizeSpecialName(HRS_GetManagedProductName(productName)), "ヨーグルト") > 0)
End Function

Private Function HRS_GetCanonicalSpecialName(ByVal productName As String) As String
    If HRS_IsMilk200(productName) Then
        HRS_GetCanonicalSpecialName = "牛乳200ml"
    ElseIf HRS_IsMilk1L(productName) Then
        HRS_GetCanonicalSpecialName = "牛乳1L"
    ElseIf HRS_IsYogurt(productName) Then
        HRS_GetCanonicalSpecialName = "ヨーグルト"
    Else
        HRS_GetCanonicalSpecialName = Trim$(productName)
    End If
End Function

Private Function HRS_GetCurrentSelectedProductName() As String
    Dim wsCache As Worksheet

    On Error Resume Next
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)
    On Error GoTo 0

    If Not wsCache Is Nothing Then
        HRS_GetCurrentSelectedProductName = Trim$(CStr(wsCache.Range("N4").value))
    End If
End Function

'============================================================
' 特殊商品シート整備
'============================================================
Private Sub HRS_EnsureSpecialProductRows()

    Dim ws As Worksheet
    Dim productNames As Variant
    Dim i As Long
    Dim r As Long

    If Not HRS_SpecialSheetExists(SH_SPECIAL) Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = SH_SPECIAL
    Else
        Set ws = ThisWorkbook.Worksheets(SH_SPECIAL)
    End If

    If Len(Trim$(CStr(ws.Cells(1, "A").value))) = 0 Then
        ws.Range("A1:F1").value = Array( _
            "商品番号", "商品名", "特殊区分", "定数", "料理用", "備考")
        ws.Range("A1:F1").Font.Bold = True
    End If

    productNames = Array("牛乳200ml", "牛乳1L", "ヨーグルト")

    For i = LBound(productNames) To UBound(productNames)
        r = HRS_FindSpecialProductRow(ws, CStr(productNames(i)))
        If r = 0 Then
            r = ws.Cells(ws.Rows.count, "B").End(xlUp).Row + 1
            If r < 2 Then r = 2
            ws.Cells(r, "B").value = CStr(productNames(i))
            ws.Cells(r, "C").value = "定数"
            ws.Cells(r, "D").value = 0
        End If

        If HRS_IsMilk1L(CStr(productNames(i))) Then
            ws.Cells(r, "E").value = "ON"
            ws.Cells(r, "F").value = _
                "定数のみ入力。料理用は自動計算し、牛乳1L選択時だけ表示。"
        End If
    Next i
End Sub

Private Function HRS_FindSpecialProductRow(ByVal ws As Worksheet, ByVal productName As String) As Long
    Dim lastRow As Long
    Dim r As Long

    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    For r = 2 To lastRow
        If HRS_SameSpecialProduct(CStr(ws.Cells(r, "B").value), productName) Then
            HRS_FindSpecialProductRow = r
            Exit Function
        End If
    Next r
End Function

Private Function HRS_SameSpecialProduct(ByVal name1 As String, ByVal name2 As String) As Boolean
    If HRS_IsMilk200(name1) And HRS_IsMilk200(name2) Then
        HRS_SameSpecialProduct = True
    ElseIf HRS_IsMilk1L(name1) And HRS_IsMilk1L(name2) Then
        HRS_SameSpecialProduct = True
    ElseIf HRS_IsYogurt(name1) And HRS_IsYogurt(name2) Then
        HRS_SameSpecialProduct = True
    End If
End Function

Private Function HRS_NormalizeSpecialName(ByVal value As String) As String
    Dim s As String
    s = LCase$(Trim$(value))
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    s = Replace(s, "ｍｌ", "ml")
    s = Replace(s, "ＭＬ", "ml")
    s = Replace(s, "ｌ", "l")
    s = Replace(s, "Ｌ", "l")
    HRS_NormalizeSpecialName = s
End Function

Private Function HRS_ParseNumericValue(ByVal value As Variant) As Double
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    If IsNumeric(value) Then
        HRS_ParseNumericValue = CDbl(value)
        Exit Function
    End If

    s = StrConv(CStr(value), vbNarrow)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "-" Then
            result = result & ch
        ElseIf Len(result) > 0 Then
            Exit For
        End If
    Next i

    If IsNumeric(result) Then HRS_ParseNumericValue = CDbl(result)
End Function

Private Function HRS_SpecialSheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    HRS_SpecialSheetExists = Not ws Is Nothing
End Function

'------------------------------------------------------------
' 牛乳1Lの料理用配分
' 納品日単位で次を計算する。
'   納品日の合計数 - ((7 + 1.4 + 1.4) × 使用日の日数)
' 結果は同じ納品日の最初の行だけ、配分後(F列)へ設定する。
' 画面の納品日・使用日の並びは変更しない。
'------------------------------------------------------------
Public Sub HRS_UpdateMilkCookingDetail(ByVal productName As String)

    HRS_ShowSelectedSpecialOrder productName
    HRS_ApplyMilk1LDistribution productName

End Sub

Public Sub HRS_ApplyMilk1LDistribution(ByVal productName As String)

    Dim wsCache As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim deliveryKey As String
    Dim useDateKey As String
    Dim qty As Double
    Dim resultQty As Double
    Dim totals As Object
    Dim firstRows As Object
    Dim useDates As Object
    Dim dateSet As Object
    Dim key As Variant

    On Error GoTo ExitHandler

    If Not HRS_IsMilk1L(productName) Then Exit Sub
    If Not HRS_SpecialSheetExists(SH_PREVIEW_CACHE) Then Exit Sub

    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)
    Set totals = CreateObject("Scripting.Dictionary")
    Set firstRows = CreateObject("Scripting.Dictionary")
    Set useDates = CreateObject("Scripting.Dictionary")

    lastRow = wsCache.Cells(wsCache.Rows.count, "B").End(xlUp).Row

    '以前の自動計算値をクリアしてから再計算する。
    For r = 2 To lastRow
        If Trim$(CStr(wsCache.Cells(r, "B").value)) <> "" Then
            wsCache.Cells(r, "M").ClearContents
        End If
    Next r

    For r = 2 To lastRow
        If Trim$(CStr(wsCache.Cells(r, "B").value)) <> "" Then
            deliveryKey = HRS_FormatDeliveryKey(wsCache.Cells(r, "E").value)
            If deliveryKey <> "" Then
                If Not totals.Exists(deliveryKey) Then
                    totals.Add deliveryKey, 0#
                    firstRows.Add deliveryKey, r
                    Set dateSet = CreateObject("Scripting.Dictionary")
                    useDates.Add deliveryKey, dateSet
                End If

                qty = HRS_ParseNumericValue(wsCache.Cells(r, "D").value)
                totals(deliveryKey) = CDbl(totals(deliveryKey)) + qty

                useDateKey = HRS_FormatUseDateKey(wsCache.Cells(r, "B").value)
                If useDateKey <> "" Then
                    Set dateSet = useDates(deliveryKey)
                    If Not dateSet.Exists(useDateKey) Then dateSet.Add useDateKey, True
                End If
            End If
        End If
    Next r

    For Each key In totals.keys
        Set dateSet = useDates(CStr(key))
        resultQty = CDbl(totals(CStr(key))) - (9.8 * CDbl(dateSet.count))
        resultQty = Fix(resultQty)
        If resultQty < 0 Then resultQty = 0

        r = CLng(firstRows(CStr(key)))
        wsCache.Cells(r, "M").NumberFormat = "General"
        wsCache.Cells(r, "M").value = resultQty
        wsCache.Cells(r, "L").value = "自動"
    Next key

ExitHandler:
End Sub

Private Function HRS_FormatUseDateKey(ByVal sourceValue As Variant) As String
    If IsDate(sourceValue) Then
        HRS_FormatUseDateKey = Format$(CDate(sourceValue), "yyyy/m/d")
    Else
        HRS_FormatUseDateKey = Trim$(CStr(sourceValue))
    End If
End Function

Private Function HRS_FormatDeliveryKey(ByVal sourceValue As Variant) As String
    If IsDate(sourceValue) Then
        HRS_FormatDeliveryKey = Format$(CDate(sourceValue), "yyyy/m/d")
    Else
        HRS_FormatDeliveryKey = Trim$(CStr(sourceValue))
    End If
End Function

'------------------------------------------------------------
' 牛乳1Lの料理用表示領域を初期化する。
' Ver1.3.5.6: Ver1.3.5.1で欠落していた補助処理を追加。
'------------------------------------------------------------
Private Sub HRS_ClearMilkCookingDetail(ByVal wsInput As Worksheet)
    On Error Resume Next
    wsInput.Range("J27:Q28").NumberFormat = "General"
    On Error GoTo 0
End Sub
