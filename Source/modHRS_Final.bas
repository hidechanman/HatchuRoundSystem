Attribute VB_Name = "modHRS_Final"
Option Explicit

' Office列挙定数を数値で保持し、参照設定差によるコンパイルエラーを防止
Private Const HRS_MSO_SHAPE_ROUNDED_RECTANGLE As Long = 5
Private Const HRS_MSO_SHAPE_DOWN_TRIANGLE As Long = 52
Private Const HRS_MSO_BRING_TO_FRONT As Long = 0

'=========================================================
' 発注まるめシステム Ver2.5.5 Final
' 1標準モジュール統合版
'
' 主な機能:
' ・必要シートと画面の自動作成
' ・発注書から発注原票DBへ一括読込
' ・再読込前に現在の発注原票DBを履歴シートへ保存
' ・発注原票DBが同一のファイルは重複取込を防止
' ・印刷ヘッダーから業者コード・業者名を抽出
' ・業者名のみの重複なしプルダウン
' ・商品一覧を7～19行の固定13件表示
' ・商品選択で使用日プレビューを自動表示
' ・発注数を使用量上位2行へ自動配分
' ・発注数・取消・配分修正をセッションDBへ即時保存
' ・全取消の状態を保存・復元
' ・商品一覧A列に確認状態を表示
' ・確認済み商品行をグレー表示
' ・通常プレビューは緑、変更・取消はグレー
' ・書き戻しイメージはI21:Q27の1か所のみ
' ・実発注書へ配分値と取消線を書き戻し
'=========================================================

Private Const APP_NAME As String = "発注まるめシステム"
Private Const APP_VERSION As String = "Ver4.0.0 Part2"

Private Const SH_INPUT As String = "発注入力"
Private Const SH_RAW As String = "発注原票DB"
Private Const SH_RAW_HISTORY As String = "発注原票履歴DB"
Private Const SH_IMPORT_SETTING As String = "読込設定"
Private Const SH_VENDOR As String = "業者マスタ"
Private Const SH_PRODUCT As String = "商品マスタ"
Private Const SH_SESSION As String = "配分セッションDB"
Private Const SH_HISTORY As String = "配分履歴DB"
Private Const SH_UNIT As String = "発注単位設定"
Private Const SH_RULE As String = "旬間発注時"
Private Const SH_DELETE As String = "削除項目"
Private Const SH_STOCK As String = "在庫チェック"
Private Const SH_LOG As String = "ログ"
Private Const SH_PRODUCT_CACHE As String = "商品一覧作業"
Private Const SH_PREVIEW_CACHE As String = "使用日プレビュー作業"
Private Const SH_ALL_AGGREGATE_CACHE As String = "納品日集約作業"
Private Const SH_ZERO_USAGE As String = "数量なし表示設定"
Private mZeroUsageSettings As Object
Private mZeroUsageSettingsLoaded As Boolean


Private Const ITEM_TOP As Long = 8
Private Const ITEM_BOTTOM As Long = 20
Private Const ITEM_PAGE_SIZE As Long = 13

Private Const PREVIEW_TOP As Long = 8
Private Const PREVIEW_BOTTOM As Long = 20
Private Const PREVIEW_PAGE_SIZE As Long = 13

Private Const HEADER_DELIVERY_ROW As Long = 4
Private Const HEADER_USE_ROW As Long = 5
Private Const HEADER_MEAL_ROW As Long = 6
Private Const FIRST_USAGE_COL As Long = 4

Private Const COLOR_TITLE As Long = 7951928
Private Const COLOR_BLUE As Long = 12611584
Private Const COLOR_GREEN As Long = 13434828
Private Const COLOR_GRAY As Long = 14277081
Private Const COLOR_INPUT As Long = 13434879
Private Const COLOR_ORANGE As Long = 49407

Private mFastDepth As Long
Private mOldCalc As XlCalculation
Private mOldEvents As Boolean
Private mOldScreen As Boolean
Private mOldAlerts As Boolean

'=========================================================
' 初期導入
'=========================================================
Public Sub HRS_InstallAll()

    Dim ws As Worksheet
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim currentStep As String

    On Error GoTo ErrHandler

    currentStep = "高速化設定"
    HRS_BeginFast "システムを初期化しています..."

    currentStep = "DBシート作成"
    HRS_CreateDatabaseSheets

    currentStep = "発注入力画面作成"
    HRS_CreateInputLayout

    currentStep = "納品日別集約キャッシュ作成"
    HRS_BuildAllDeliveryAggregateCache
    HRS4_AfterImport

    currentStep = "発注入力シート表示"
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)

    'VBEから実行した場合にActiveWindowがExcel画面を
    '参照していないことがあるため、対象ウィンドウを明示する。
    If ThisWorkbook.Windows.count > 0 Then
        ThisWorkbook.Activate
        ws.Activate
        ThisWorkbook.Windows(1).DisplayGridlines = False
    End If

    currentStep = "ログ記録"
    HRS_WriteLog "INSTALL", "一括初期化完了"

ExitHandler:
    HRS_EndFast

    If errorNumber = 0 Then
        MsgBox APP_NAME & " " & APP_VERSION & vbCrLf & _
               "一括初期化が完了しました。" & vbCrLf & vbCrLf & _
               "次に、発注書ブックを前面にして" & vbCrLf & _
               "HRS_ImportActiveOrderBook を実行してください。", _
               vbInformation, APP_NAME
    Else
        HRS_ShowError "HRS_InstallAll（" & currentStep & "）", _
                      errorNumber, errorDescription
    End If
    Exit Sub

ErrHandler:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Resume ExitHandler

End Sub

Private Sub HRS_CreateDatabaseSheets()

    Dim currentStep As String

    On Error GoTo ErrHandler

    currentStep = "発注原票DB"
    HRS_EnsureTable SH_RAW, Array( _
        "登録日時", "取込元ブック", "シート名", "業者コード", "業者名", _
        "商品番号", "商品名", "単位", "納品日", "使用日", "区分", _
        "使用数量", "元行", "元列", "セル番地", "特殊商品", "備考")

    currentStep = "発注原票DB履歴"
    HRS_EnsureTable SH_RAW_HISTORY, Array( _
        "履歴ID", "保存日時", "新取込元ブック", _
        "登録日時", "取込元ブック", "シート名", "業者コード", "業者名", _
        "商品番号", "商品名", "単位", "納品日", "使用日", "区分", _
        "使用数量", "元行", "元列", "セル番地", "特殊商品", "備考")

    currentStep = "取込設定"
    HRS_SetupImportSettings

    currentStep = "業者マスタ"
    HRS_EnsureTable SH_VENDOR, Array( _
        "業者コード", "業者名", "登録日", "更新日", "表示順", "プルダウン表示")

    currentStep = "業者マスタ設定"
    HRS_SetupVendorMasterSettings

    currentStep = "商品マスタ"
    HRS_EnsureTable SH_PRODUCT, Array( _
        "商品番号", "商品名", "規格", "単位", "業者コード", "業者名", _
        "特殊商品", "登録日", "更新日")

    currentStep = "数量なし表示設定"
    HRS_EnsureTable SH_ZERO_USAGE, Array( _
        "数量なし表示", "商品番号", "商品名", "業者コード", "業者名", _
        "登録日", "更新日", "備考")
    currentStep = "数量なし表示プルダウン"
    HRS_SetupZeroUsageDisplaySheet

    currentStep = "配分セッションDB"
    HRS_EnsureTable SH_SESSION, Array( _
        "更新日時", "業者名", "商品番号", "商品名", "発注数", "確認済", _
        "取消", "使用日", "区分", "使用数量", "納品日", "配分後", _
        "単位・注意点", "発注書", "セル番地", "取込元ブック")

    currentStep = "配分履歴DB"
    HRS_EnsureTable SH_HISTORY, Array( _
        "保存日時", "業者名", "商品番号", "商品名", "発注数", "確認済", _
        "取消", "使用日", "区分", "使用数量", "納品日", "配分後", _
        "単位・注意点", "発注書", "セル番地", "取込元ブック")

    currentStep = "発注単位設定"
    HRS_EnsureTable SH_UNIT, Array( _
        "商品番号", "商品名", "発注単位", "換算数", "表示内容", "備考")

    currentStep = "旬間発注時"
    HRS_EnsureTable SH_RULE, Array( _
        "番号", "商品名", "項目3", "項目4", "項目5", "内容")

    currentStep = "削除項目"
    HRS_EnsureDeleteItemsSheet
    currentStep = "在庫確認"
    HRS_EnsureStockCheckSheet

    currentStep = "ログ"
    HRS_EnsureTable SH_LOG, Array( _
        "日時", "処理", "内容", "ユーザー")

    currentStep = "商品一覧作業"
    HRS_GetOrCreateSheet SH_PRODUCT_CACHE, True
    currentStep = "使用日プレビュー作業"
    HRS_GetOrCreateSheet SH_PREVIEW_CACHE, True
    currentStep = "納品日集約作業"
    HRS_EnsureTable SH_ALL_AGGREGATE_CACHE, Array( _
        "業者名", "商品番号", "商品名", "納品日", "使用数量合計", _
        "最初使用日", "最後使用日", "書戻対象元行", "朝あり", _
        "発注書", "明細元行一覧", "集約キー")
    ThisWorkbook.Worksheets(SH_ALL_AGGREGATE_CACHE).Visible = xlSheetVeryHidden


    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "HRS_CreateDatabaseSheets（" & currentStep & "）", Err.Description

End Sub

Private Sub HRS_EnsureDeleteItemsSheet()

    Dim ws As Worksheet

    Set ws = HRS_GetOrCreateSheet(SH_DELETE, False)

    If Trim$(CStr(ws.Range("B1").value)) = "" Then
        ws.Range("B1").value = "初期取消候補商品"
    End If

    If Trim$(CStr(ws.Range("D1").value)) = "" Then
        ws.Range("D1").value = "確認済商品"
    End If

    With ws.Range("B1")
        .Font.Bold = True
        .Interior.Color = RGB(255, 199, 206)
        .Borders.LineStyle = xlContinuous
    End With

    With ws.Range("D1")
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
        .Borders.LineStyle = xlContinuous
    End With

    ws.Columns("B").ColumnWidth = 34
    ws.Columns("D").ColumnWidth = 34

End Sub


Private Sub HRS_EnsureStockCheckSheet()

    Dim ws As Worksheet

    Set ws = HRS_GetOrCreateSheet(SH_STOCK, False)

    If Trim$(CStr(ws.Range("B1").value)) = "" Then
        ws.Range("B1").value = "商品名"
    End If

    If Trim$(CStr(ws.Range("C1").value)) = "" Then
        ws.Range("C1").value = "在庫数"
    End If

    With ws.Range("B1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns("B").ColumnWidth = 36
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("C").NumberFormat = "0.##"

End Sub

Private Sub HRS_EnsureTable(ByVal sheetName As String, ByVal headers As Variant)

    Dim ws As Worksheet
    Dim i As Long
    Dim headerLastCol As Long

    Set ws = HRS_GetOrCreateSheet(sheetName, False)
    headerLastCol = UBound(headers) + 1

    '過去レイアウトの結合セルが残っていても見出しを書けるようにする。
    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(1, headerLastCol)).UnMerge
    ws.AutoFilterMode = False
    On Error GoTo 0

    For i = LBound(headers) To UBound(headers)
        If Trim$(CStr(ws.Cells(1, i + 1).value)) = "" Then
            ws.Cells(1, i + 1).value = headers(i)
        End If
    Next i

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, headerLastCol))
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With

    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(1, headerLastCol)).AutoFilter
    On Error GoTo 0

End Sub

'=========================================================
' 画面作成
'=========================================================
Public Sub HRS_CreateInputLayout()

    Dim ws As Worksheet

    On Error GoTo ErrHandler

    Set ws = HRS_GetOrCreateSheet(SH_INPUT, False)

    HRS_BeginFast "発注入力画面を作成しています..."

    HRS_RemoveSystemShapes ws

    On Error Resume Next
    ws.Range("A1:S52").UnMerge
    ws.Range("A1:S52").ClearContents
    ws.Range("A1:S52").ClearFormats
    ws.Range("A1:S52").Validation.Delete
    On Error GoTo ErrHandler

    With ws
        .Cells.Font.Name = "Meiryo UI"
        .Cells.Font.Size = 10

        .Columns("A").ColumnWidth = 13
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 9
        .Columns("E").ColumnWidth = 10
        .Columns("F").ColumnWidth = 10
        .Columns("G").ColumnWidth = 10
        .Columns("H").ColumnWidth = 10
        .Columns("I").ColumnWidth = 9
        .Columns("J").ColumnWidth = 3
        .Columns("K").ColumnWidth = 10
        .Columns("L").ColumnWidth = 12
        .Columns("M").ColumnWidth = 8
        .Columns("N").ColumnWidth = 12
        .Columns("O").ColumnWidth = 11
        .Columns("Q").ColumnWidth = 12
        .Columns("P").ColumnWidth = 12
        .Columns("R").ColumnWidth = 22
        .Columns("S").ColumnWidth = 16

        .Range("A2:S2").Interior.Color = RGB(31, 78, 121)
        .Range("A2").value = APP_NAME & " " & APP_VERSION
        .Range("A2").Font.Color = vbWhite
        .Range("A2").Font.Bold = True
        .Range("A2").Font.Size = 18

        .Range("A3").value = "業者選択"
        .Range("A3").Font.Bold = True
        .Range("B3").Interior.Color = RGB(255, 242, 204)
        .Range("B3").Borders.LineStyle = xlContinuous

        .Range("H3").value = "商品検索"
        .Range("H3").Font.Bold = True
        .Range("I3:K3").Merge
        .Range("I3").Interior.Color = RGB(255, 242, 204)
        .Range("I3:K3").Borders.LineStyle = xlContinuous

        .Range("A4:S4").Clear
        .Rows("4").RowHeight = 12

        .Range("K5:L5").Merge
        .Range("K5").value = "□ 全取消"
        .Range("K5:L5").Font.Size = 14
        .Range("K5:L5").Font.Bold = True
        .Range("K5:L5").HorizontalAlignment = xlCenter
        .Range("K5:L5").VerticalAlignment = xlCenter
        .Range("K5:L5").Borders.LineStyle = xlContinuous
        .Range("K5:L5").Interior.Color = RGB(255, 242, 204)

        .Range("A6:I6").Interior.Color = RGB(0, 112, 192)
        .Range("A6").value = "商品一覧・発注数入力"
        .Range("A6:I6").Font.Color = vbWhite
        .Range("A6:I6").Font.Bold = True
        .Range("A6:I6").Borders.LineStyle = xlContinuous

        .Range("A7").value = "商品番号"
        .Range("B7").value = "商品名"
        .Range("C7").value = "合計"
        .Range("D7").value = "単位"
        .Range("E7").value = "在庫数"
        .Range("F7").value = "発注数"
        .Range("G7").value = "前回合計"
        .Range("H7").value = "前回在庫"
        .Range("I7").value = "確認"

        .Range("A7:I20").Borders.LineStyle = xlContinuous
        .Range("A7:I20").Font.Size = 11
        .Range("A7:I7").Interior.Color = RGB(217, 225, 242)
        .Range("A7:I7").Font.Bold = True
        .Range("F8:F20").Interior.Color = RGB(255, 242, 204)
        .Range("F8:F20").Locked = False
        .Range("F8:F20").NumberFormat = "General"
        .Range("I8:I20").HorizontalAlignment = xlCenter
        .Range("I8:I20").VerticalAlignment = xlCenter
        .Range("I8:I20").Font.Size = 11
        .Range("I8:I20").Font.Bold = True

        .Range("K6:S6").Interior.Color = RGB(0, 176, 80)
        .Range("K6").value = "選択商品の使用日プレビュー"
        .Range("K6:S6").Font.Color = vbWhite
        .Range("K6:S6").Font.Bold = True
        .Range("K6:S6").Borders.LineStyle = xlContinuous

        .Range("K7").value = "取消"
        .Range("L7").value = "納品日"
        .Range("M7").value = "曜日"
        .Range("N7").value = "使用日"
        .Range("O7").value = "区分"
        .Range("P7").value = "使用数量"
        .Range("Q7").value = "配分後"
        .Range("R7").value = "商品名"
        .Range("S7").value = "発注書"

        .Range("K7:S20").Borders.LineStyle = xlContinuous
        .Range("K7:S20").Font.Size = 11
        .Range("K7:S7").Interior.Color = RGB(226, 239, 218)
        .Range("K7:S7").Font.Bold = True
        .Range("K8:S20").Interior.Color = RGB(226, 239, 218)
        .Range("L8:L20").Font.Size = 11
        .Range("L8:L20").Font.Bold = True
        .Range("L8:L20").HorizontalAlignment = xlCenter
        .Range("Q8:Q20").Interior.Color = RGB(255, 242, 204)
        .Range("Q8:Q20").NumberFormat = "General"

        '旬間発注時：タイトル22行、内容23行
        .Range("A22:I22").Merge
        .Range("A22:I22").Interior.Color = RGB(112, 173, 71)
        .Range("A22").value = "旬間発注時（選択商品）"
        .Range("A22:I22").Font.Color = vbWhite
        .Range("A22:I22").Font.Bold = True
        .Range("A22:I22").Borders.LineStyle = xlContinuous
        .Range("A22:I22").HorizontalAlignment = xlCenter
        .Range("A22:I22").VerticalAlignment = xlCenter

        .Range("A23:I23").Merge
        .Range("A23").value = "商品を選択すると内容を表示します。"
        .Range("A23:I23").Interior.Color = RGB(242, 242, 242)
        .Range("A23:I23").Borders.LineStyle = xlContinuous
        .Range("A23:I23").WrapText = True
        .Range("A23").HorizontalAlignment = xlCenter
        .Range("A23").VerticalAlignment = xlCenter
        .Range("A23:I23").Font.Size = 14
        .Range("A23:I23").Font.Bold = True
        .Range("A23:I23").ShrinkToFit = True

        '前回書き戻し：タイトル25行、表26～31行（必要時に少し下へスクロールして確認）
        .Range("A25:I25").Interior.Color = RGB(255, 192, 0)
        .Range("A25").value = "前回の書き戻しイメージ"
        .Range("A25:I25").Font.Bold = True
        .Range("A25:I25").Borders.LineStyle = xlContinuous
        .Range("A26:I31").Interior.Color = RGB(242, 242, 242)
        .Range("A26:I31").Borders.LineStyle = xlContinuous

        '今回書き戻し：タイトル22行、表23～28行
        .Range("K22:S22").Interior.Color = RGB(255, 192, 0)
        .Range("K22").value = "今回の書き戻しイメージ"
        .Range("K22:S22").Font.Bold = True
        .Range("K22:S22").Borders.LineStyle = xlContinuous
        .Range("K23:S28").Interior.Color = RGB(242, 242, 242)
        .Range("K23:S28").Borders.LineStyle = xlContinuous
        .Range("K29:S52").Clear

        .Rows("2").RowHeight = 30
        .Rows("3").RowHeight = 28
        .Rows("5").RowHeight = 28
        .Rows("6").RowHeight = 28
        .Rows("7").RowHeight = 24
        .Rows("8:20").RowHeight = 34
        .Rows("22:23").RowHeight = 34
        .Rows("25").RowHeight = 28
        .Rows("26:31").RowHeight = 30
        .Rows("22:31").Hidden = False
        .Range("A2:S30").VerticalAlignment = xlCenter
    End With

    '3行目の主要ボタンを色分け
    '記号はVBAで文字化けしにくい半角記号を使用
    HRS_AddButton ws, "Import", "[IN] 発注書再読込", _
        "HRS_ImportActiveOrderBook", ws.Range("D3:E3"), _
        RGB(0, 112, 192)

    HRS_AddButton ws, "VendorRefresh", "UP 一覧更新", _
        "HRS_RefreshAll", ws.Range("F3:G3"), _
        RGB(112, 173, 71)

    HRS_AddButton ws, "ProductSearch", "検索", _
        "HRS_SearchProduct", ws.Range("L3"), _
        RGB(112, 48, 160)

    HRS_AddButton ws, "WriteBack", "[OK] 実際に書き戻す", _
        "HRS_WriteBackToOrderBook", ws.Range("Q3"), _
        RGB(237, 125, 49)

    HRS_AddButton ws, "ProductPrev", "<< 前", _
        "HRS_ProductPrev", ws.Range("H6")
    HRS_AddButton ws, "ProductNext", "次 >>", _
        "HRS_ProductNext", ws.Range("I6")

    HRS_AddButton ws, "AggregateMode", "集約 OFF", _
        "HRS_ToggleDeliveryAggregate", ws.Range("M5:O5"), _
        RGB(112, 48, 160)

    HRS_AddButton ws, "PreviewPrev", "<< 前", _
        "HRS_PreviewPrev", ws.Range("R6")
    HRS_AddButton ws, "PreviewNext", "次 >>", _
        "HRS_PreviewNext", ws.Range("S6")

    HRS_AddButton ws, "PrevWBPrev", "<< 前", _
        "HRS_PreviousWriteBackPrev", ws.Range("H25")
    HRS_AddButton ws, "PrevWBNext", "次 >>", _
        "HRS_PreviousWriteBackNext", ws.Range("I25")

    HRS_AddButton ws, "CurrWBPrev", "<< 前", _
        "HRS_CurrentWriteBackPrev", ws.Range("R22")
    HRS_AddButton ws, "CurrWBNext", "次 >>", _
        "HRS_CurrentWriteBackNext", ws.Range("S22")

    HRS_UpdateAggregateButtonCaption

ExitHandler:
    HRS_EndFast
    Exit Sub

ErrHandler:
    HRS_ShowError "HRS_CreateInputLayout", Err.Number, Err.Description
    Resume ExitHandler

End Sub

Private Sub HRS_AddButton( _
    ByVal ws As Worksheet, _
    ByVal suffixName As String, _
    ByVal caption As String, _
    ByVal macroName As String, _
    ByVal targetCell As Range, _
    Optional ByVal fillColor As Long = -1)

    Dim shp As Shape
    Dim buttonColor As Long

    If fillColor = -1 Then
        buttonColor = RGB(91, 155, 213)
    Else
        buttonColor = fillColor
    End If

    Set shp = ws.Shapes.AddShape( _
        HRS_MSO_SHAPE_ROUNDED_RECTANGLE, _
        targetCell.Left + 2, _
        targetCell.Top + 2, _
        targetCell.Width - 4, _
        targetCell.Height - 4)

    With shp
        .Name = "HRS_" & suffixName
        .TextFrame.Characters.Text = caption
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.Characters.Font.Name = "Meiryo UI"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = vbWhite
        .Fill.ForeColor.RGB = buttonColor
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 1
        .OnAction = macroName
        .Placement = xlMoveAndSize
    End With

End Sub

Private Sub HRS_RemoveSystemShapes(ByVal ws As Worksheet)

    Dim i As Long

    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "HRS_" Then
            ws.Shapes(i).Delete
        End If
    Next i

End Sub

'=========================================================
' 発注書読込
'=========================================================
Public Sub HRS_ImportActiveOrderBook()

    Dim selectedPath As Variant
    Dim sourceBook As Workbook
    Dim alreadyOpen As Boolean
    Dim wb As Workbook
    Dim answer As VbMsgBoxResult
    Dim historyAnswer As VbMsgBoxResult
    Dim saveCurrentHistory As Boolean
    Dim currentDataExists As Boolean
    Dim wsCurrentRaw As Worksheet

    selectedPath = Application.GetOpenFilename( _
        FileFilter:="Excelファイル (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="取り込む発注書ファイルを選択してください", _
        MultiSelect:=False)

    If VarType(selectedPath) = vbBoolean Then
        If selectedPath = False Then Exit Sub
    End If

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, CStr(selectedPath), vbTextCompare) = 0 Then
            Set sourceBook = wb
            alreadyOpen = True
            Exit For
        End If
    Next wb

    If sourceBook Is Nothing Then
        On Error GoTo OpenError
        Set sourceBook = Workbooks.Open( _
            Filename:=CStr(selectedPath), _
            ReadOnly:=True, _
            UpdateLinks:=False)
        On Error GoTo 0
    End If

    If sourceBook Is ThisWorkbook Then
        MsgBox "発注まるめシステム自身は取り込めません。", vbExclamation, APP_NAME
        GoTo ExitHandler
    End If

    HRS_CreateDatabaseSheets

    If HRS_GetImportSetting("重複取込防止", True) Then
        If HRS_IsSameAsCurrentRawDB(sourceBook) Then
            MsgBox "同じ発注書です。" & vbCrLf & vbCrLf & _
                   "現在の発注原票DBと内容が同じため、取り込みませんでした。" & vbCrLf & _
                   "選択ファイル：" & sourceBook.Name, _
                   vbExclamation, APP_NAME
            HRS_WriteLog "IMPORT_SKIP", "同一発注書のため取込中止：" & sourceBook.Name
            GoTo ExitHandler
        End If
    End If

    Set wsCurrentRaw = ThisWorkbook.Worksheets(SH_RAW)
    currentDataExists = (HRS_LastRow(wsCurrentRaw, 1) >= 2)

    If currentDataExists Then

        historyAnswer = MsgBox( _
            "現在の発注原票DBを履歴DBへ保存しますか？" & vbCrLf & vbCrLf & _
            "【はい】" & vbCrLf & _
            "　現在のデータを発注原票履歴DBへ保存してから、新しい発注書を取り込みます。" & vbCrLf & vbCrLf & _
            "【いいえ】" & vbCrLf & _
            "　現在のデータを保存せず、新しい発注書を取り込みます。" & vbCrLf & vbCrLf & _
            "【キャンセル】" & vbCrLf & _
            "　読込処理を中止します。", _
            vbYesNoCancel + vbQuestion, APP_NAME)

        If historyAnswer = vbCancel Then
            HRS_WriteLog "IMPORT_CANCEL", _
                "履歴保存確認でキャンセル：" & sourceBook.Name
            GoTo ExitHandler
        End If

        saveCurrentHistory = (historyAnswer = vbYes)

        answer = MsgBox( _
            "現在の発注原票DBは削除され、新しい発注書へ置き換えられます。" & vbCrLf & vbCrLf & _
            "取込元：" & sourceBook.Name & vbCrLf & _
            "履歴保存：" & IIf(saveCurrentHistory, "保存する", "保存しない") & vbCrLf & vbCrLf & _
            "このまま続行しますか？", _
            vbYesNo + vbExclamation + vbDefaultButton2, APP_NAME)

    Else

        saveCurrentHistory = False

        answer = MsgBox( _
            "選択した発注書を取り込みます。" & vbCrLf & vbCrLf & _
            "取込元：" & sourceBook.Name & vbCrLf & vbCrLf & _
            "このまま続行しますか？", _
            vbYesNo + vbQuestion + vbDefaultButton2, APP_NAME)

    End If

    If answer <> vbYes Then
        HRS_WriteLog "IMPORT_CANCEL", _
            "最終確認で中止：" & sourceBook.Name
        GoTo ExitHandler
    End If

    HRS_ImportCore sourceBook, saveCurrentHistory

ExitHandler:
    If Not sourceBook Is Nothing Then
        If Not alreadyOpen Then sourceBook.Close SaveChanges:=False
    End If
    Exit Sub

OpenError:
    MsgBox "選択した発注書を開けませんでした。" & vbCrLf & _
           Err.Number & " : " & Err.Description, vbCritical, APP_NAME
    Resume ExitHandler

End Sub


Private Sub HRS_SetupImportSettings()

    Dim ws As Worksheet

    Set ws = HRS_GetOrCreateSheet(SH_IMPORT_SETTING, False)

    If Trim$(CStr(ws.Range("A1").value)) = "" Then
        ws.Range("A1").value = "設定項目"
        ws.Range("B1").value = "設定値"
        ws.Range("C1").value = "説明"
    End If

    HRS_EnsureImportSettingRow ws, "履歴保存確認", "ON", _
        "読込時に、現在の発注原票DBを履歴へ保存するか毎回確認します。"

    HRS_EnsureImportSettingRow ws, "重複取込防止", "ON", _
        "選択した発注書の内容が現在の発注原票DBと同じ場合は取り込みません。"

    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
    End With

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 70
    ws.Columns("C").WrapText = True

End Sub

Private Sub HRS_EnsureImportSettingRow(ByVal ws As Worksheet, _
                                       ByVal settingName As String, _
                                       ByVal defaultValue As String, _
                                       ByVal descriptionText As String)

    Dim foundCell As Range
    Dim targetRow As Long
    Dim listText As String

    Set foundCell = ws.Columns("A").Find( _
        What:=settingName, _
        After:=ws.Cells(1, "A"), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    If foundCell Is Nothing Then
        targetRow = HRS_LastRow(ws, 1) + 1
        If targetRow < 2 Then targetRow = 2
        ws.Cells(targetRow, "A").value = settingName
        ws.Cells(targetRow, "B").value = defaultValue
    Else
        targetRow = foundCell.Row
        If Trim$(CStr(ws.Cells(targetRow, "B").value)) = "" Then
            ws.Cells(targetRow, "B").value = defaultValue
        End If
    End If

    ws.Cells(targetRow, "C").value = descriptionText

    listText = "ON" & Application.International(xlListSeparator) & "OFF"

    On Error Resume Next
    ws.Cells(targetRow, "B").Validation.Delete
    Err.Clear
    ws.Cells(targetRow, "B").Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:=listText

    If Err.Number <> 0 Then
        Err.Clear
        ws.Cells(targetRow, "B").Validation.Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="ON,OFF"
    End If

    If Err.Number = 0 Then
        ws.Cells(targetRow, "B").Validation.IgnoreBlank = True
        ws.Cells(targetRow, "B").Validation.InCellDropdown = True
        ws.Cells(targetRow, "B").Validation.ShowError = True
    End If
    Err.Clear
    On Error GoTo 0

End Sub

Private Function HRS_GetImportSetting(ByVal settingName As String, _
                                      ByVal defaultValue As Boolean) As Boolean

    Dim ws As Worksheet
    Dim foundCell As Range
    Dim settingText As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_IMPORT_SETTING)
    On Error GoTo 0

    If ws Is Nothing Then
        HRS_GetImportSetting = defaultValue
        Exit Function
    End If

    Set foundCell = ws.Columns("A").Find( _
        What:=settingName, _
        After:=ws.Cells(1, "A"), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    If foundCell Is Nothing Then
        HRS_GetImportSetting = defaultValue
        Exit Function
    End If

    settingText = UCase$(Trim$(CStr(ws.Cells(foundCell.Row, "B").value)))

    If settingText = "OFF" Or settingText = "FALSE" Or settingText = "0" Then
        HRS_GetImportSetting = False
    ElseIf settingText = "ON" Or settingText = "TRUE" Or settingText = "1" Then
        HRS_GetImportSetting = True
    Else
        HRS_GetImportSetting = defaultValue
    End If

End Function

Private Sub HRS_ArchiveCurrentRawDB(ByVal newSourceBookName As String)

    Dim wsRaw As Worksheet
    Dim wsHistory As Worksheet
    Dim lastRawRow As Long
    Dim nextHistoryRow As Long
    Dim rowCount As Long
    Dim historyID As String
    Dim sourceData As Variant
    Dim outputData() As Variant
    Dim r As Long
    Dim c As Long

    Set wsRaw = ThisWorkbook.Worksheets(SH_RAW)
    Set wsHistory = ThisWorkbook.Worksheets(SH_RAW_HISTORY)

    lastRawRow = HRS_LastRow(wsRaw, 1)
    If lastRawRow < 2 Then Exit Sub

    rowCount = lastRawRow - 1
    sourceData = wsRaw.Range("A2:Q" & lastRawRow).Value2
    ReDim outputData(1 To rowCount, 1 To 20)

    historyID = Format$(Now, "yyyymmdd_hhnnss") & "_" & _
                Format$(HRS_LastRow(wsHistory, 1), "000000")

    For r = 1 To rowCount
        outputData(r, 1) = historyID
        outputData(r, 2) = Now
        outputData(r, 3) = newSourceBookName

        For c = 1 To 17
            outputData(r, c + 3) = sourceData(r, c)
        Next c
    Next r

    nextHistoryRow = HRS_LastRow(wsHistory, 1) + 1
    If nextHistoryRow < 2 Then nextHistoryRow = 2

    wsHistory.Cells(nextHistoryRow, 1).Resize(rowCount, 20).value = outputData
    wsHistory.Columns("A:T").AutoFit

    HRS_WriteLog "RAW_ARCHIVE", _
        "履歴ID：" & historyID & "、保存件数：" & rowCount & _
        "、次回取込元：" & newSourceBookName

End Sub

Private Function HRS_IsSameAsCurrentRawDB(ByVal sourceBook As Workbook) As Boolean

    Dim wsRaw As Worksheet
    Dim currentMap As Object
    Dim sourceMap As Object
    Dim currentCount As Long
    Dim sourceCount As Long
    Dim keyValue As Variant

    Set wsRaw = ThisWorkbook.Worksheets(SH_RAW)

    If HRS_LastRow(wsRaw, 1) < 2 Then Exit Function

    Set currentMap = HRS_BuildCurrentRawMap(wsRaw, currentCount)
    Set sourceMap = HRS_BuildSourceRawMap(sourceBook, sourceCount)

    If currentCount <> sourceCount Then Exit Function
    If currentMap.count <> sourceMap.count Then Exit Function

    For Each keyValue In currentMap.keys
        If Not sourceMap.Exists(CStr(keyValue)) Then Exit Function
        If CLng(currentMap(keyValue)) <> CLng(sourceMap(keyValue)) Then Exit Function
    Next keyValue

    HRS_IsSameAsCurrentRawDB = True

End Function

Private Function HRS_BuildCurrentRawMap(ByVal wsRaw As Worksheet, _
                                        ByRef recordCount As Long) As Object

    Dim mapValue As Object
    Dim lastRow As Long
    Dim dataValue As Variant
    Dim r As Long
    Dim recordKey As String

    Set mapValue = CreateObject("Scripting.Dictionary")
    mapValue.CompareMode = vbBinaryCompare

    lastRow = HRS_LastRow(wsRaw, 1)

    If lastRow >= 2 Then
        dataValue = wsRaw.Range("A2:Q" & lastRow).Value2

        For r = 1 To UBound(dataValue, 1)
            recordKey = HRS_MakeRawRecordKey( _
                CStr(dataValue(r, 3)), CStr(dataValue(r, 4)), CStr(dataValue(r, 5)), _
                CStr(dataValue(r, 6)), CStr(dataValue(r, 7)), CStr(dataValue(r, 8)), _
                dataValue(r, 9), dataValue(r, 10), CStr(dataValue(r, 11)), _
                dataValue(r, 12), CLng(Val(CStr(dataValue(r, 13)))), _
                CLng(Val(CStr(dataValue(r, 14)))), CStr(dataValue(r, 15)), _
                CStr(dataValue(r, 16)), CStr(dataValue(r, 17)))

            HRS_AddRecordKey mapValue, recordKey
            recordCount = recordCount + 1
        Next r
    End If

    Set HRS_BuildCurrentRawMap = mapValue

End Function

Private Function HRS_BuildSourceRawMap(ByVal sourceBook As Workbook, _
                                       ByRef recordCount As Long) As Object

    Dim mapValue As Object
    Dim ws As Worksheet
    Dim vendorCode As String
    Dim vendorName As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim subtotalCol As Long
    Dim productRow As Long
    Dim dataCol As Long
    Dim productCode As String
    Dim baseName As String
    Dim specText As String
    Dim productName As String
    Dim unitText As String
    Dim qty As Variant
    Dim recordKey As String
    Dim deliveryValue As Variant
    Dim useValue As Variant
    Dim mealText As String

    Set mapValue = CreateObject("Scripting.Dictionary")
    mapValue.CompareMode = vbBinaryCompare

    For Each ws In sourceBook.Worksheets

        vendorCode = ""
        vendorName = HRS_ExtractVendorFromHeader(ws, vendorCode)
        If vendorName = "" Then vendorName = "未登録業者"

        If HRS_CountProductRows(ws) > 0 Then

            lastRow = HRS_LastRow(ws, 0)
            lastCol = HRS_LastColumn(ws)
            subtotalCol = HRS_FindSubtotalColumn(ws)
            If subtotalCol = 0 Then subtotalCol = lastCol + 1

            For productRow = 7 To lastRow

                productCode = HRS_ProductCode(ws.Cells(productRow, 1))

                If HRS_LooksLikeProductCode(productCode) Then

                    baseName = HRS_CleanText(HRS_CellText(ws.Cells(productRow + 1, 1)))
                    specText = HRS_CleanText(HRS_CellText(ws.Cells(productRow + 2, 1)))
                    unitText = HRS_CleanText(HRS_CellText(ws.Cells(productRow, 2)))

                    If baseName <> "" Then
                        productName = baseName

                        For dataCol = FIRST_USAGE_COL To subtotalCol - 1

                            If HRS_IsUsageColumn(ws, dataCol) Then

                                qty = ws.Cells(productRow, dataCol).Value2

                                If HRS_IsValidQty(qty) Then

                                    deliveryValue = HRS_HeaderLeftFill( _
                                        ws, HEADER_DELIVERY_ROW, dataCol)
                                    useValue = HRS_HeaderLeftFill( _
                                        ws, HEADER_USE_ROW, dataCol)
                                    mealText = HRS_CleanText( _
                                        HRS_CellText(ws.Cells(HEADER_MEAL_ROW, dataCol)))

                                    recordKey = HRS_MakeRawRecordKey( _
                                        ws.Name, vendorCode, vendorName, _
                                        productCode, productName, unitText, _
                                        deliveryValue, useValue, mealText, _
                                        CDbl(qty), productRow, dataCol, _
                                        ws.Cells(productRow, dataCol).Address(False, False), _
                                        IIf(HRS_IsSpecial(productName), "TRUE", "FALSE"), _
                                        "規格：" & specText)

                                    HRS_AddRecordKey mapValue, recordKey
                                    recordCount = recordCount + 1

                                End If
                            End If
                        Next dataCol
                    End If
                End If
            Next productRow
        End If
    Next ws

    Set HRS_BuildSourceRawMap = mapValue

End Function

Private Sub HRS_AddRecordKey(ByVal mapValue As Object, _
                             ByVal recordKey As String)

    If mapValue.Exists(recordKey) Then
        mapValue(recordKey) = CLng(mapValue(recordKey)) + 1
    Else
        mapValue.Add recordKey, 1
    End If

End Sub

Private Function HRS_MakeRawRecordKey( _
    ByVal sheetName As String, _
    ByVal vendorCode As String, _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal unitText As String, _
    ByVal deliveryValue As Variant, _
    ByVal useValue As Variant, _
    ByVal mealText As String, _
    ByVal qtyValue As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceCol As Long, _
    ByVal cellAddress As String, _
    ByVal specialText As String, _
    ByVal noteText As String) As String

    Dim SEP As String

    SEP = ChrW$(30)

    HRS_MakeRawRecordKey = _
        HRS_KeyText(sheetName) & SEP & _
        HRS_KeyText(vendorCode) & SEP & _
        HRS_KeyText(vendorName) & SEP & _
        HRS_KeyText(productCode) & SEP & _
        HRS_KeyText(productName) & SEP & _
        HRS_KeyText(unitText) & SEP & _
        HRS_KeyValue(deliveryValue) & SEP & _
        HRS_KeyValue(useValue) & SEP & _
        HRS_KeyText(mealText) & SEP & _
        HRS_KeyNumber(qtyValue) & SEP & _
        CStr(sourceRow) & SEP & CStr(sourceCol) & SEP & _
        HRS_KeyText(cellAddress) & SEP & _
        HRS_KeyText(specialText) & SEP & _
        HRS_KeyText(noteText)

End Function

Private Function HRS_KeyText(ByVal valueText As String) As String

    valueText = Replace(valueText, vbCr, " ")
    valueText = Replace(valueText, vbLf, " ")
    valueText = Replace(valueText, ChrW(30), " ")
    HRS_KeyText = Trim$(valueText)

End Function

Private Function HRS_KeyValue(ByVal valueData As Variant) As String

    If IsError(valueData) Or IsEmpty(valueData) Then
        HRS_KeyValue = ""
    ElseIf IsDate(valueData) Then
        HRS_KeyValue = Format$(CDate(valueData), "yyyy/mm/dd hh:nn:ss")
    ElseIf IsNumeric(valueData) Then
        HRS_KeyValue = HRS_KeyNumber(valueData)
    Else
        HRS_KeyValue = HRS_KeyText(CStr(valueData))
    End If

End Function

Private Function HRS_KeyNumber(ByVal valueData As Variant) As String

    If IsNumeric(valueData) Then
        HRS_KeyNumber = Format$(CDbl(valueData), "0.###############")
    Else
        HRS_KeyNumber = HRS_KeyText(CStr(valueData))
    End If

End Function


Private Function HRS_IsPointOne(ByVal valueData As Variant) As Boolean

    If IsError(valueData) Or IsEmpty(valueData) Then Exit Function
    If Not IsNumeric(valueData) Then Exit Function

    HRS_IsPointOne = (Abs(CDbl(valueData) - 0.1) < 0.0000001)

End Function

Private Sub HRS_ApplyPointOneStrikeToSource(ByVal sourceBook As Workbook)

    Dim wsRaw As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim sourceSheetName As String
    Dim cellAddress As String
    Dim targetSheet As Worksheet
    Dim targetCell As Range
    Dim appliedCount As Long

    On Error GoTo ErrHandler

    Set wsRaw = ThisWorkbook.Worksheets(SH_RAW)
    lastRow = HRS_LastRow(wsRaw, 1)

    For r = 2 To lastRow

        If HRS_IsPointOne(wsRaw.Cells(r, 12).value) Then

            sourceSheetName = CStr(wsRaw.Cells(r, 3).value)
            cellAddress = CStr(wsRaw.Cells(r, 15).value)

            If HRS_SheetExists(sourceSheetName, sourceBook) And _
               Len(cellAddress) > 0 Then

                Set targetSheet = sourceBook.Worksheets(sourceSheetName)
                Set targetCell = targetSheet.Range(cellAddress)
                targetCell.Font.Strikethrough = True
                appliedCount = appliedCount + 1

            End If
        End If

        Set targetSheet = Nothing
        Set targetCell = Nothing
    Next r

    HRS_WriteLog "POINT_ONE_STRIKE", _
        "読込時0.1取消線：" & appliedCount & "件、対象：" & sourceBook.Name
    Exit Sub

ErrHandler:
    HRS_WriteLog "POINT_ONE_STRIKE_ERROR", _
        CStr(Err.Number) & "：" & Err.Description
    Err.Clear

End Sub

Public Sub HRS_RefreshPointOneStrike()

    Dim wsInput As Worksheet
    Dim productCode As String
    Dim productName As String

    On Error GoTo ErrHandler

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)

    productCode = HRS_NormalizeCode(wsInput.Range("B5").value)
    productName = Trim$(CStr(wsInput.Range("B6").value))

    If productCode = "" And productName = "" Then Exit Sub

    HRS_BuildPreviewCache productCode, productName
    HRS_RenderPreviewPage
    HRS_BuildCurrentWriteBackPreview

    MsgBox "0.1の取消線状態を更新しました。", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    HRS_ShowError "HRS_RefreshPointOneStrike", _
        Err.Number, Err.Description

End Sub

Private Sub HRS_ImportCore(ByVal sourceBook As Workbook, _
                           ByVal saveCurrentHistory As Boolean)

    Dim wsRaw As Worksheet
    Dim ws As Worksheet
    Dim buffer() As Variant
    Dim bufferCount As Long
    Dim nextRow As Long
    Dim vendorCode As String
    Dim vendorName As String
    Dim totalRecords As Long
    Dim targetSheets As Long

    On Error GoTo ErrHandler

    HRS_BeginFast "発注書を読み込んでいます..."
    HRS_InvalidateZeroUsageSettingsCache
    HRS_LoadZeroUsageSettingsCache

    Set wsRaw = ThisWorkbook.Worksheets(SH_RAW)

    If saveCurrentHistory Then
        HRS_ArchiveCurrentRawDB sourceBook.Name
    Else
        HRS_WriteLog "RAW_ARCHIVE_SKIP", _
            "利用者が履歴保存を選択しなかったため保存なし：" & sourceBook.Name
    End If

    HRS_ClearBelowHeader wsRaw, "A", "Q"

    ReDim buffer(1 To 5000, 1 To 17)
    nextRow = 2

    For Each ws In sourceBook.Worksheets

        vendorCode = ""
        vendorName = HRS_ExtractVendorFromHeader(ws, vendorCode)

        If vendorName = "" Then
            vendorName = "未登録業者"
        End If

        If HRS_CountProductRows(ws) > 0 Then
            targetSheets = targetSheets + 1

            totalRecords = totalRecords + HRS_ImportOneSheet( _
                ws, sourceBook.Name, vendorCode, vendorName, _
                wsRaw, buffer, bufferCount, nextRow)

            HRS_RegisterVendor vendorCode, vendorName
        End If

    Next ws

    HRS_FlushBuffer wsRaw, buffer, bufferCount, nextRow

    wsRaw.Columns("D").NumberFormat = "@"
    wsRaw.Columns("F").NumberFormat = "@"
    wsRaw.Columns("A:Q").AutoFit

    HRS_ApplyPointOneStrikeToSource sourceBook

    Application.StatusBar = "納品日別の集約キャッシュを作成しています..."
    HRS_BuildAllDeliveryAggregateCache
    HRS4_AfterImport

    HRS_RefreshVendorList

ExitHandler:
    HRS_EndFast

    If Err.Number = 0 Then
        ThisWorkbook.Worksheets(SH_INPUT).Activate
        MsgBox "発注書の読込が完了しました。" & vbCrLf & _
               "対象シート：" & targetSheets & "枚" & vbCrLf & _
               "登録明細：" & totalRecords & "件" & vbCrLf & _
               "数量なし表示を使う商品は「" & SH_ZERO_USAGE & _
               "」シートのA列をONにして再読込してください。", _
               vbInformation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    HRS_ShowError "HRS_ImportCore", Err.Number, Err.Description
    Resume ExitHandler

End Sub

Private Function HRS_ImportOneSheet( _
    ByVal sourceSheet As Worksheet, _
    ByVal sourceBookName As String, _
    ByVal vendorCode As String, _
    ByVal vendorName As String, _
    ByVal wsRaw As Worksheet, _
    ByRef buffer() As Variant, _
    ByRef bufferCount As Long, _
    ByRef nextRow As Long) As Long

    Dim lastRow As Long
    Dim lastCol As Long
    Dim subtotalCol As Long
    Dim productRow As Long
    Dim dataCol As Long
    Dim productCode As String
    Dim baseName As String
    Dim specText As String
    Dim productName As String
    Dim unitText As String
    Dim qty As Variant
    Dim countValue As Long
    Dim includeWithoutQty As Boolean

    lastRow = HRS_LastRow(sourceSheet, 0)
    lastCol = HRS_LastColumn(sourceSheet)
    subtotalCol = HRS_FindSubtotalColumn(sourceSheet)

    If subtotalCol = 0 Then subtotalCol = lastCol + 1

    For productRow = 7 To lastRow

        productCode = HRS_ProductCode(sourceSheet.Cells(productRow, 1))

        If HRS_LooksLikeProductCode(productCode) Then

            baseName = HRS_CleanText( _
                HRS_CellText(sourceSheet.Cells(productRow + 1, 1)))

            specText = HRS_CleanText( _
                HRS_CellText(sourceSheet.Cells(productRow + 2, 1)))

            unitText = HRS_CleanText( _
                HRS_CellText(sourceSheet.Cells(productRow, 2)))

            If baseName <> "" Then

                '商品名は4行ブロックの2行目だけを使用する。
                '規格行は商品名へ連結しない。
                productName = baseName

                HRS_RegisterProduct productCode, productName, specText, _
                                    unitText, vendorCode, vendorName

                '発注書を読むたび、未登録商品だけ設定シート末尾へ追加する。
                HRS_RegisterZeroUsageProduct productCode, productName, _
                    vendorCode, vendorName
                includeWithoutQty = HRS_IsZeroUsageDisplayEnabled( _
                    productCode, productName, vendorCode, vendorName)

                For dataCol = FIRST_USAGE_COL To subtotalCol - 1

                    If HRS_IsUsageColumn(sourceSheet, dataCol) Then

                        qty = sourceSheet.Cells(productRow, dataCol).Value2

                        If HRS_IsValidQty(qty) Or includeWithoutQty Then

                            bufferCount = bufferCount + 1

                            buffer(bufferCount, 1) = Now
                            buffer(bufferCount, 2) = sourceBookName
                            buffer(bufferCount, 3) = sourceSheet.Name
                            buffer(bufferCount, 4) = vendorCode
                            buffer(bufferCount, 5) = vendorName
                            buffer(bufferCount, 6) = productCode
                            buffer(bufferCount, 7) = productName
                            buffer(bufferCount, 8) = unitText
                            buffer(bufferCount, 9) = HRS_HeaderLeftFill( _
                                sourceSheet, HEADER_DELIVERY_ROW, dataCol)
                            buffer(bufferCount, 10) = HRS_HeaderLeftFill( _
                                sourceSheet, HEADER_USE_ROW, dataCol)
                            buffer(bufferCount, 11) = HRS_CleanText( _
                                HRS_CellText(sourceSheet.Cells(HEADER_MEAL_ROW, dataCol)))
                            If HRS_IsValidQty(qty) Then
                                buffer(bufferCount, 12) = CDbl(qty)
                            Else
                                '数量なし表示対象は、使用数量を空欄のまま保持する。
                                buffer(bufferCount, 12) = ""
                            End If
                            buffer(bufferCount, 13) = productRow
                            buffer(bufferCount, 14) = dataCol
                            buffer(bufferCount, 15) = _
                                sourceSheet.Cells(productRow, dataCol).Address(False, False)
                            buffer(bufferCount, 16) = _
                                IIf(HRS_IsSpecial(productName), "TRUE", "FALSE")
                            buffer(bufferCount, 17) = "規格：" & specText

                            countValue = countValue + 1

                            If bufferCount = 5000 Then
                                HRS_FlushBuffer wsRaw, buffer, bufferCount, nextRow
                            End If

                        End If
                    End If
                Next dataCol
            End If
        End If
    Next productRow

    HRS_ImportOneSheet = countValue

End Function

Private Sub HRS_FlushBuffer(ByVal wsRaw As Worksheet, _
                            ByRef buffer() As Variant, _
                            ByRef bufferCount As Long, _
                            ByRef nextRow As Long)

    Dim outputData() As Variant
    Dim r As Long
    Dim c As Long

    If bufferCount = 0 Then Exit Sub

    ReDim outputData(1 To bufferCount, 1 To 17)

    For r = 1 To bufferCount
        For c = 1 To 17
            outputData(r, c) = buffer(r, c)
        Next c
    Next r

    wsRaw.Cells(nextRow, 1).Resize(bufferCount, 17).value = outputData
    nextRow = nextRow + bufferCount
    bufferCount = 0

    Erase buffer
    ReDim buffer(1 To 5000, 1 To 17)

End Sub

'=========================================================
' 業者一覧・商品一覧
'=========================================================
Public Sub HRS_RefreshAll()

    HRS_RefreshVendorList
    HRS_EnsureVendorDropdown
    HRS_OnVendorChanged

End Sub


Public Sub HRS_SetupVendorMasterSettings()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim nextOrder As Long

    Set ws = ThisWorkbook.Worksheets(SH_VENDOR)

    ws.Cells(1, 5).value = "表示順"
    ws.Cells(1, 6).value = "プルダウン表示"

    lastRow = Application.Max(HRS_LastRow(ws, 1), HRS_LastRow(ws, 2))
    nextOrder = 1

    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 2).value)) <> "" Then
            If IsNumeric(ws.Cells(r, 5).value) Then
                If CLng(Val(ws.Cells(r, 5).value)) >= nextOrder Then
                    nextOrder = CLng(Val(ws.Cells(r, 5).value)) + 1
                End If
            End If
        End If
    Next r

    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 2).value)) <> "" Then
            If Not IsNumeric(ws.Cells(r, 5).value) Or Val(ws.Cells(r, 5).value) <= 0 Then
                ws.Cells(r, 5).value = nextOrder
                nextOrder = nextOrder + 1
            End If
            If Trim$(CStr(ws.Cells(r, 6).value)) = "" Then
                ws.Cells(r, 6).value = "ON"
            End If
        End If
    Next r

    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C:D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("E").NumberFormat = "0"

    If lastRow < 2 Then lastRow = 2
    HRS_SetVendorVisibilityValidation ws.Range("F2:F" & CStr(Application.Max(lastRow, 500)))

End Sub

Private Sub HRS_SetVendorVisibilityValidation(ByVal targetRange As Range)

    Dim oneCell As Range
    Dim listText As String
    Dim separatorText As String

    If targetRange Is Nothing Then Exit Sub

    separatorText = Application.International(xlListSeparator)
    If Len(separatorText) = 0 Then separatorText = ","
    listText = "ON" & separatorText & "OFF"

    For Each oneCell In targetRange.Cells
        HRS_AddOnOffValidation oneCell, listText
    Next oneCell

End Sub

Private Sub HRS_AddOnOffValidation(ByVal targetCell As Range, ByVal listText As String)

    If targetCell Is Nothing Then Exit Sub
    If targetCell.MergeCells Then Exit Sub

    On Error Resume Next
    targetCell.Validation.Delete
    Err.Clear

    targetCell.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, listText

    If Err.Number <> 0 Then
        Err.Clear
        targetCell.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "ON,OFF"
    End If

    If Err.Number = 0 Then
        targetCell.Validation.IgnoreBlank = True
        targetCell.Validation.InCellDropdown = True
        targetCell.Validation.ShowInput = False
        targetCell.Validation.ShowError = True
        targetCell.Validation.ErrorTitle = "入力内容を確認してください"
        targetCell.Validation.ErrorMessage = "ON または OFF を選択してください。"
    End If

    Err.Clear
    On Error GoTo 0

End Sub

Private Function HRS_FindVendorMasterRowByName(ByVal vendorName As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim targetText As String

    Set ws = ThisWorkbook.Worksheets(SH_VENDOR)
    targetText = HRS_NormalizeText(vendorName)
    lastRow = HRS_LastRow(ws, 2)

    For r = 2 To lastRow
        If HRS_NormalizeText(CStr(ws.Cells(r, 2).value)) = targetText Then
            HRS_FindVendorMasterRowByName = r
            Exit Function
        End If
    Next r

End Function

Private Function HRS_IsVendorDropdownVisible(ByVal sourceValue As Variant) As Boolean

    Dim textValue As String

    textValue = UCase$(Trim$(CStr(sourceValue)))

    Select Case textValue
        Case "", "ON", "TRUE", "1", "表示"
            HRS_IsVendorDropdownVisible = True
        Case Else
            HRS_IsVendorDropdownVisible = False
    End Select

End Function

Private Function HRS_GetVendorDisplayOrder(ByVal ws As Worksheet, _
                                           ByVal masterRow As Long) As Double

    If IsNumeric(ws.Cells(masterRow, 5).value) Then
        If CDbl(ws.Cells(masterRow, 5).value) > 0 Then
            HRS_GetVendorDisplayOrder = CDbl(ws.Cells(masterRow, 5).value)
            Exit Function
        End If
    End If

    HRS_GetVendorDisplayOrder = 1000000# + CDbl(masterRow)

End Function

Private Function HRS_ArrayContainsVendor(ByRef vendorNames() As String, _
                                         ByVal itemCount As Long, _
                                         ByVal vendorName As String) As Boolean

    Dim i As Long

    If itemCount <= 0 Then Exit Function

    For i = 1 To itemCount
        If StrComp(vendorNames(i), vendorName, vbTextCompare) = 0 Then
            HRS_ArrayContainsVendor = True
            Exit Function
        End If
    Next i

End Function

Public Sub HRS_RefreshVendorList()

    Dim wsInput As Worksheet
    Dim wsRaw As Worksheet
    Dim wsVendor As Worksheet
    Dim dict As Object
    Dim lastRow As Long
    Dim r As Long
    Dim vendorName As String
    Dim outputRow As Long
    Dim key As Variant
    Dim currentVendor As String
    Dim masterRow As Long
    Dim itemCount As Long
    Dim vendorNames() As String
    Dim vendorOrders() As Double
    Dim i As Long
    Dim j As Long
    Dim tempName As String
    Dim tempOrder As Double

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsRaw = ThisWorkbook.Worksheets(SH_RAW)
    Set wsVendor = ThisWorkbook.Worksheets(SH_VENDOR)
    Set dict = CreateObject("Scripting.Dictionary")

    HRS_SetupVendorMasterSettings

    wsInput.Columns("AA:AC").Hidden = False
    wsInput.Range("AA1:AC5000").ClearContents

    lastRow = HRS_LastRow(wsRaw, 1)

    For r = 2 To lastRow
        vendorName = Trim$(CStr(wsRaw.Cells(r, 5).value))
        If vendorName <> "" Then
            If Not dict.Exists(vendorName) Then dict.Add vendorName, vendorName
        End If
    Next r

    '今回読み込んだ発注書にある業者をマスタへ補完する。
    For Each key In dict.keys
        masterRow = HRS_FindVendorMasterRowByName(CStr(key))
        If masterRow = 0 Then
            HRS_RegisterVendor "", CStr(key)
        End If
    Next key

    '表示対象だけを配列へ取り込み、表示順→業者名で並べる。
    For Each key In dict.keys
        masterRow = HRS_FindVendorMasterRowByName(CStr(key))
        If masterRow > 0 Then
            If HRS_IsVendorDropdownVisible(wsVendor.Cells(masterRow, 6).value) Then
                itemCount = itemCount + 1
                If itemCount = 1 Then
                    ReDim vendorNames(1 To 1)
                    ReDim vendorOrders(1 To 1)
                Else
                    ReDim Preserve vendorNames(1 To itemCount)
                    ReDim Preserve vendorOrders(1 To itemCount)
                End If
                vendorNames(itemCount) = CStr(key)
                vendorOrders(itemCount) = HRS_GetVendorDisplayOrder(wsVendor, masterRow)
            End If
        End If
    Next key

    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount
            If vendorOrders(i) > vendorOrders(j) Or _
               (vendorOrders(i) = vendorOrders(j) And _
                StrComp(vendorNames(i), vendorNames(j), vbTextCompare) > 0) Then
                tempOrder = vendorOrders(i)
                vendorOrders(i) = vendorOrders(j)
                vendorOrders(j) = tempOrder

                tempName = vendorNames(i)
                vendorNames(i) = vendorNames(j)
                vendorNames(j) = tempName
            End If
        Next j
    Next i

    outputRow = 1
    For i = 1 To itemCount
        wsInput.Cells(outputRow, "AA").value = vendorNames(i)
        outputRow = outputRow + 1
    Next i

    On Error Resume Next
    wsInput.Range("B3").MergeArea.Validation.Delete
    On Error GoTo 0

    If itemCount > 0 Then
        currentVendor = Trim$(CStr(wsInput.Range("B3").value))
        If Not HRS_ArrayContainsVendor(vendorNames, itemCount, currentVendor) Then
            currentVendor = CStr(wsInput.Range("AA1").value)
            wsInput.Range("B3").value = currentVendor
        End If
    Else
        wsInput.Range("B3").ClearContents
    End If

    HRS_EnsureVendorDropdown
    wsInput.Columns("AA:AC").Hidden = True

End Sub

Public Sub HRS_OnVendorChanged()

    Dim wsInput As Worksheet
    Dim wsRaw As Worksheet
    Dim wsCache As Worksheet
    Dim dict As Object
    Dim vendorName As String
    Dim rowVendor As String
    Dim productCode As String
    Dim productName As String
    Dim keyText As String
    Dim lastRow As Long
    Dim r As Long
    Dim outputRow As Long
    Dim item As Variant
    Dim key As Variant
    Dim deleteStatus As Long
    Dim sessionConfirmed As Boolean

    On Error GoTo ErrHandler

    HRS_BeginFast "商品一覧を更新しています..."

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsRaw = ThisWorkbook.Worksheets(SH_RAW)
    Set wsCache = ThisWorkbook.Worksheets(SH_PRODUCT_CACHE)
    Set dict = CreateObject("Scripting.Dictionary")

    vendorName = Trim$(CStr(wsInput.Range("B3").value))

    wsCache.Range("A1:L5000").ClearContents
    wsCache.Range("A1:K1").value = Array( _
        "商品番号", "商品名", "合計", "在庫数", "発注数", _
        "前回合計", "前回在庫", "確認", "削除状態", "業者名", "単位")

    wsCache.Range("L1").value = 1
    wsCache.Range("L2").value = vendorName

    lastRow = HRS_LastRow(wsRaw, 1)

    For r = 2 To lastRow

        rowVendor = Trim$(CStr(wsRaw.Cells(r, 5).value))

        If rowVendor = vendorName Then

            productCode = HRS_NormalizeCode(wsRaw.Cells(r, 6).value)
            productName = Trim$(CStr(wsRaw.Cells(r, 7).value))

            If productName <> "" Then

                keyText = productCode & "|" & productName

                If Not dict.Exists(keyText) Then
                    dict.Add keyText, Array( _
                        productCode, productName, _
                        CDbl(Val(wsRaw.Cells(r, 12).value)), _
                        Trim$(CStr(wsRaw.Cells(r, 8).value)))
                Else
                    item = dict(keyText)
                    item(2) = CDbl(item(2)) + _
                              CDbl(Val(wsRaw.Cells(r, 12).value))
                    dict(keyText) = item
                End If

            End If
        End If
    Next r

    outputRow = 2

    For Each key In dict.keys

        item = dict(key)
        deleteStatus = HRS_GetDeleteItemStatus(CStr(item(1)))
        sessionConfirmed = HRS_IsSessionConfirmed( _
            vendorName, CStr(item(0)), CStr(item(1)))

        wsCache.Cells(outputRow, 1).NumberFormat = "@"
        wsCache.Cells(outputRow, 1).value = item(0)
        wsCache.Cells(outputRow, 2).value = item(1)
        wsCache.Cells(outputRow, 3).value = item(2)
        wsCache.Cells(outputRow, 4).value = _
            HRS_GetStockQty(CStr(item(1)))
        wsCache.Cells(outputRow, 5).value = _
            HRS_GetSavedOrderQty( _
                vendorName, CStr(item(0)), CStr(item(1)))
        wsCache.Cells(outputRow, 6).value = _
            HRS_GetPreviousTotalQty( _
                vendorName, CStr(item(0)), CStr(item(1)))
        wsCache.Cells(outputRow, 7).value = _
            HRS_GetPreviousStockQty(CStr(item(1)))

        'B列の赤文字商品は、初期状態では確認を外す。
        '発注数入力または手動確認後だけ■にする。
        If deleteStatus = 1 Then
            If sessionConfirmed Then
                wsCache.Cells(outputRow, 8).value = "■"
            Else
                wsCache.Cells(outputRow, 8).value = "□"
            End If

        ElseIf deleteStatus = 2 Then
            wsCache.Cells(outputRow, 8).value = "■"

        ElseIf sessionConfirmed Then
            wsCache.Cells(outputRow, 8).value = "■"
        Else
            wsCache.Cells(outputRow, 8).value = "□"
        End If

        wsCache.Cells(outputRow, 9).value = deleteStatus
        wsCache.Cells(outputRow, 10).value = vendorName
        wsCache.Cells(outputRow, 11).value = item(3)
        outputRow = outputRow + 1

    Next key

    HRS_SyncStockProducts
    HRS_RenderProductPage
    HRS_ClearSelectedPanels

ExitHandler:
    HRS_EndFast
    Exit Sub

ErrHandler:
    HRS_ShowError "HRS_OnVendorChanged", Err.Number, Err.Description
    Resume ExitHandler

End Sub

Public Sub HRS_RenderProductPage()

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim total As Long
    Dim pos As Long
    Dim i As Long
    Dim sourceRow As Long
    Dim displayRow As Long
    Dim lastShow As Long
    Dim deleteStatus As Long
    Dim isConfirmed As Boolean

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PRODUCT_CACHE)

    wsInput.Range("A8:I20").ClearContents
    wsInput.Range("A8:I20").Interior.ColorIndex = xlNone
    wsInput.Range("A8:I20").Font.Color = vbBlack
    wsInput.Range("A8:I20").Font.Strikethrough = False
    wsInput.Range("F8:F20").Interior.Color = RGB(255, 242, 204)
    wsInput.Range("I8:I20").Font.Size = 11

    total = HRS_LastRow(wsCache, 1) - 1
    If total < 0 Then total = 0

    pos = CLng(Val(wsCache.Range("L1").value))
    If pos < 1 Then pos = 1

    If total > 0 And pos > total Then
        pos = total - ITEM_PAGE_SIZE + 1
        If pos < 1 Then pos = 1
        wsCache.Range("L1").value = pos
    End If

    For i = 0 To ITEM_PAGE_SIZE - 1

        sourceRow = pos + i + 1
        displayRow = ITEM_TOP + i

        If sourceRow > total + 1 Then Exit For

        deleteStatus = CLng(Val(wsCache.Cells(sourceRow, 9).value))
        isConfirmed = _
            (CStr(wsCache.Cells(sourceRow, 8).value) = "■")

        wsInput.Cells(displayRow, "A").NumberFormat = "@"
        wsInput.Cells(displayRow, "A").value = _
            wsCache.Cells(sourceRow, 1).value
        wsInput.Cells(displayRow, "B").value = _
            wsCache.Cells(sourceRow, 2).value
        wsInput.Cells(displayRow, "C").value = _
            wsCache.Cells(sourceRow, 3).value
        wsInput.Cells(displayRow, "D").value = _
            wsCache.Cells(sourceRow, 11).value
        wsInput.Cells(displayRow, "E").value = _
            wsCache.Cells(sourceRow, 4).value
        If Trim$(CStr(wsCache.Cells(sourceRow, 5).value)) = "" Then
            wsInput.Cells(displayRow, "F").ClearContents
        ElseIf IsNumeric(wsCache.Cells(sourceRow, 5).value) Then
            If CDbl(wsCache.Cells(sourceRow, 5).value) = _
               Fix(CDbl(wsCache.Cells(sourceRow, 5).value)) Then

                wsInput.Cells(displayRow, "F").value = _
                    CLng(wsCache.Cells(sourceRow, 5).value)
            Else
                wsInput.Cells(displayRow, "F").value = _
                    CDbl(wsCache.Cells(sourceRow, 5).value)
            End If
        Else
            wsInput.Cells(displayRow, "F").value = _
                wsCache.Cells(sourceRow, 5).value
        End If

        wsInput.Cells(displayRow, "F").NumberFormat = "General"

        If CDbl(Val(wsCache.Cells(sourceRow, 6).value)) = 0 Then
            wsInput.Cells(displayRow, "G").value = ""
        Else
            wsInput.Cells(displayRow, "G").value = _
                wsCache.Cells(sourceRow, 6).value
        End If

        If CDbl(Val(wsCache.Cells(sourceRow, 7).value)) = 0 Then
            wsInput.Cells(displayRow, "H").value = ""
        Else
            wsInput.Cells(displayRow, "H").value = _
                wsCache.Cells(sourceRow, 7).value
        End If

        wsInput.Cells(displayRow, "I").value = _
            wsCache.Cells(sourceRow, 8).value

        If deleteStatus = 1 Then
            '赤文字は常に維持する。
            wsInput.Cells(displayRow, "B").Font.Color = RGB(255, 0, 0)
            wsInput.Cells(displayRow, "B").Font.Bold = True

            If isConfirmed Then
                '発注数入力または確認後は背景をグレーにする。
                wsInput.Range( _
                    "A" & displayRow & ":I" & displayRow).Interior.Color = _
                    RGB(217, 217, 217)
            Else
                '初期状態は背景をグレーにしない。
                wsInput.Cells(displayRow, "F").Interior.Color = _
                    RGB(255, 242, 204)
            End If

        ElseIf deleteStatus = 2 Then
            wsInput.Range( _
                "A" & displayRow & ":I" & displayRow).Interior.Color = _
                RGB(217, 217, 217)

        ElseIf isConfirmed Then
            wsInput.Range( _
                "A" & displayRow & ":I" & displayRow).Interior.Color = _
                RGB(217, 217, 217)
        Else
            wsInput.Cells(displayRow, "F").Interior.Color = _
                RGB(255, 242, 204)
        End If

    Next i

    If total = 0 Then
        wsInput.Range("E6").value = "0件"
    Else
        lastShow = WorksheetFunction.Min( _
            pos + ITEM_PAGE_SIZE - 1, total)

        wsInput.Range("E6").value = _
            pos & "～" & lastShow & "/" & total & "件"
    End If

    With wsInput.Range("E6:F6")
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With

End Sub

Public Sub HRS_SearchProduct()

    Dim wsInput As Worksheet
    Dim wsRaw As Worksheet
    Dim wsCache As Worksheet
    Dim searchText As String
    Dim rawData As Variant
    Dim lastRow As Long
    Dim dict As Object
    Dim r As Long
    Dim keyText As String
    Dim vendorName As String
    Dim productCode As String
    Dim productName As String
    Dim item As Variant
    Dim key As Variant
    Dim outputRow As Long
    Dim deleteStatus As Long
    Dim sessionConfirmed As Boolean

    On Error GoTo ErrHandler

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsRaw = ThisWorkbook.Worksheets(SH_RAW)
    Set wsCache = ThisWorkbook.Worksheets(SH_PRODUCT_CACHE)

    searchText = HRS_NormalizeText(CStr(wsInput.Range("I3").value))

    If searchText = "" Then
        HRS_OnVendorChanged
        Exit Sub
    End If

    HRS_BeginFast "全業者の商品を検索しています..."

    lastRow = HRS_LastRow(wsRaw, 1)

    If lastRow < 2 Then
        MsgBox "検索対象の商品データがありません。", _
               vbInformation, APP_NAME
        GoTo ExitHandler
    End If

    'セルを1行ずつ読まず、DBを配列へ一括取得する。
    rawData = wsRaw.Range("A2:Q" & lastRow).Value2
    Set dict = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(rawData, 1)

        productName = Trim$(CStr(rawData(r, 7)))

        If productName <> "" Then

            '完全一致ではなく部分一致。
            If InStr(1, HRS_NormalizeText(productName), _
                     searchText, vbTextCompare) > 0 Then

                vendorName = Trim$(CStr(rawData(r, 5)))
                productCode = HRS_NormalizeCode(rawData(r, 6))

                '同一商品でも業者が違えば別結果として表示する。
                keyText = vendorName & "|" & _
                          productCode & "|" & productName

                If Not dict.Exists(keyText) Then
                    dict.Add keyText, Array( _
                        vendorName, productCode, productName, _
                        CDbl(Val(rawData(r, 12))), _
                        Trim$(CStr(rawData(r, 8))))
                Else
                    item = dict(keyText)
                    item(3) = CDbl(item(3)) + _
                              CDbl(Val(rawData(r, 12)))
                    dict(keyText) = item
                End If
            End If
        End If
    Next r

    wsCache.Range("A1:L5000").ClearContents
    wsCache.Range("A1:K1").value = Array( _
        "商品番号", "商品名", "合計", "在庫数", "発注数", _
        "前回合計", "前回在庫", "確認", "削除状態", "業者名", "単位")

    wsCache.Range("L1").value = 1
    wsCache.Range("L2").value = "検索結果"

    outputRow = 2

    For Each key In dict.keys

        item = dict(key)
        vendorName = CStr(item(0))
        productCode = CStr(item(1))
        productName = CStr(item(2))

        deleteStatus = HRS_GetDeleteItemStatus(productName)
        sessionConfirmed = HRS_IsSessionConfirmed( _
            vendorName, productCode, productName)

        wsCache.Cells(outputRow, 1).NumberFormat = "@"
        wsCache.Cells(outputRow, 1).value = productCode
        wsCache.Cells(outputRow, 2).value = productName
        wsCache.Cells(outputRow, 3).value = item(3)
        wsCache.Cells(outputRow, 4).value = _
            HRS_GetStockQty(productName)
        wsCache.Cells(outputRow, 5).value = _
            HRS_GetSavedOrderQty( _
                vendorName, productCode, productName)
        wsCache.Cells(outputRow, 6).value = _
            HRS_GetPreviousTotalQty( _
                vendorName, productCode, productName)
        wsCache.Cells(outputRow, 7).value = _
            HRS_GetPreviousStockQty(productName)

        If deleteStatus = 1 Then
            If sessionConfirmed Then
                wsCache.Cells(outputRow, 8).value = "■"
            Else
                wsCache.Cells(outputRow, 8).value = "□"
            End If
        ElseIf deleteStatus = 2 Or sessionConfirmed Then
            wsCache.Cells(outputRow, 8).value = "■"
        Else
            wsCache.Cells(outputRow, 8).value = "□"
        End If

        wsCache.Cells(outputRow, 9).value = deleteStatus
        wsCache.Cells(outputRow, 10).value = vendorName
        wsCache.Cells(outputRow, 11).value = item(4)

        outputRow = outputRow + 1
    Next key

    HRS_RenderProductPage
    HRS_ClearSelectedPanels

    If dict.count = 0 Then
        MsgBox "該当する商品が見つかりませんでした。", _
               vbInformation, APP_NAME
    Else
        wsInput.Activate
        wsInput.Range("A8").Select
    End If

ExitHandler:
    HRS_EndFast
    Exit Sub

ErrHandler:
    HRS_ShowError "HRS_SearchProduct", _
        Err.Number, Err.Description
    Resume ExitHandler

End Sub

Public Sub HRS_ProductPrev()

    Dim wsCache As Worksheet
    Dim pos As Long

    Set wsCache = ThisWorkbook.Worksheets(SH_PRODUCT_CACHE)

    pos = CLng(Val(wsCache.Range("L1").value))
    pos = pos - ITEM_PAGE_SIZE
    If pos < 1 Then pos = 1

    wsCache.Range("L1").value = pos
    HRS_RenderProductPage
    HRS_ClearSelectedPanels

End Sub

Public Sub HRS_ProductNext()

    Dim wsCache As Worksheet
    Dim total As Long
    Dim pos As Long

    Set wsCache = ThisWorkbook.Worksheets(SH_PRODUCT_CACHE)

    total = HRS_LastRow(wsCache, 1) - 1
    pos = CLng(Val(wsCache.Range("L1").value))
    pos = pos + ITEM_PAGE_SIZE

    If pos > total Then
        pos = total - ITEM_PAGE_SIZE + 1
    End If

    If pos < 1 Then pos = 1

    wsCache.Range("L1").value = pos
    HRS_RenderProductPage
    HRS_ClearSelectedPanels

End Sub

'=========================================================
' 商品選択・プレビュー
'=========================================================
Public Sub HRS_OnProductSelected(ByVal selectedRow As Long)

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim productCode As String
    Dim productName As String
    Dim sourceRow As Long
    Dim pos As Long
    Dim resultVendor As String

    If selectedRow < ITEM_TOP Or _
       selectedRow > ITEM_BOTTOM Then Exit Sub

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PRODUCT_CACHE)

    productCode = Trim$(CStr( _
        wsInput.Cells(selectedRow, "A").value))

    productName = Trim$(CStr( _
        wsInput.Cells(selectedRow, "B").value))

    If productName = "" Then Exit Sub

    '検索結果の場合、キャッシュJ列に元の業者名が入っている。
    pos = CLng(Val(wsCache.Range("L1").value))
    If pos < 1 Then pos = 1

    sourceRow = pos + _
                (selectedRow - ITEM_TOP) + 1

    resultVendor = Trim$(CStr( _
        wsCache.Cells(sourceRow, 10).value))

    If resultVendor <> "" Then
        'イベント停止中なので一覧を再構築せず、業者だけ切り替える。
        wsInput.Range("B3").value = resultVendor
    End If

    HRS4_LoadSelectedProductToLegacyPreview productCode, productName
    HRS_ApplyDeleteItemToPreview productCode, productName
    HRS_RenderPreviewPage
    HRS_ShowRule productName
    HRS_BuildCurrentWriteBackPreview
    HRS_BuildPreviousWriteBackPreview productCode, productName
    HRS_UpdateMilkCookingDetail productName

End Sub


'=========================================================
' 読込時納品日集約キャッシュ
'=========================================================
Private Sub HRS_BuildAllDeliveryAggregateCache()

    Dim wsRaw As Worksheet
    Dim wsAgg As Worksheet
    Dim rawData As Variant
    Dim resultData() As Variant
    Dim dict As Object
    Dim lastRow As Long
    Dim rowCount As Long
    Dim r As Long
    Dim groupIndex As Long
    Dim groupCount As Long
    Dim keyText As String
    Dim productIdentity As String
    Dim deliveryKey As String
    Dim categoryText As String
    Dim currentDeliveryDate As Variant

    Dim vendorNames() As String
    Dim productCodes() As String
    Dim productNames() As String
    Dim deliveryDates() As Variant
    Dim usageTotals() As Double
    Dim firstUseDates() As Variant
    Dim lastUseDates() As Variant
    Dim targetRawRows() As Long
    Dim morningFound() As Boolean
    Dim orderSheets() As String
    Dim rawRowLists() As String
    Dim aggregateKeys() As String

    Set wsRaw = ThisWorkbook.Worksheets(SH_RAW)
    Set wsAgg = ThisWorkbook.Worksheets(SH_ALL_AGGREGATE_CACHE)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    HRS_ClearBelowHeader wsAgg, "A", "L"

    lastRow = HRS_LastRow(wsRaw, 1)
    If lastRow < 2 Then Exit Sub

    rawData = wsRaw.Range("A2:Q" & lastRow).Value2
    rowCount = UBound(rawData, 1)

    ReDim vendorNames(1 To rowCount)
    ReDim productCodes(1 To rowCount)
    ReDim productNames(1 To rowCount)
    ReDim deliveryDates(1 To rowCount)
    ReDim usageTotals(1 To rowCount)
    ReDim firstUseDates(1 To rowCount)
    ReDim lastUseDates(1 To rowCount)
    ReDim targetRawRows(1 To rowCount)
    ReDim morningFound(1 To rowCount)
    ReDim orderSheets(1 To rowCount)
    ReDim rawRowLists(1 To rowCount)
    ReDim aggregateKeys(1 To rowCount)

    For r = 1 To rowCount
        If Trim$(CStr(rawData(r, 7))) <> "" Then

            productIdentity = HRS_AggregateProductIdentity( _
                CStr(rawData(r, 6)), CStr(rawData(r, 7)))

            currentDeliveryDate = rawData(r, 9)
            deliveryKey = HRS_AggregateKey(currentDeliveryDate)

            If deliveryKey = "" Then
                deliveryKey = "BLANK_RAW_" & CStr(r + 1)
            End If

            keyText = UCase$(Trim$(CStr(rawData(r, 5)))) & "|" & _
                      productIdentity & "|" & deliveryKey

            If Not dict.Exists(keyText) Then
                groupCount = groupCount + 1
                groupIndex = groupCount
                dict.Add keyText, groupIndex

                vendorNames(groupIndex) = CStr(rawData(r, 5))
                productCodes(groupIndex) = HRS_NormalizeCode(rawData(r, 6))
                productNames(groupIndex) = CStr(rawData(r, 7))
                deliveryDates(groupIndex) = currentDeliveryDate
                firstUseDates(groupIndex) = rawData(r, 10)
                lastUseDates(groupIndex) = rawData(r, 10)
                targetRawRows(groupIndex) = r + 1
                orderSheets(groupIndex) = CStr(rawData(r, 3))
                rawRowLists(groupIndex) = CStr(r + 1)
                aggregateKeys(groupIndex) = deliveryKey

                categoryText = CStr(rawData(r, 11))
                morningFound(groupIndex) = _
                    HRS_IsMorningCategory(categoryText)
            Else
                groupIndex = CLng(dict(keyText))
                lastUseDates(groupIndex) = rawData(r, 10)
                rawRowLists(groupIndex) = _
                    rawRowLists(groupIndex) & "," & CStr(r + 1)

                categoryText = CStr(rawData(r, 11))

                If HRS_IsMorningCategory(categoryText) And _
                   Not morningFound(groupIndex) Then
                    targetRawRows(groupIndex) = r + 1
                    morningFound(groupIndex) = True
                    orderSheets(groupIndex) = CStr(rawData(r, 3))
                End If
            End If

            If Trim$(CStr(rawData(r, 12))) <> "" And _
               IsNumeric(rawData(r, 12)) Then
                usageTotals(groupIndex) = _
                    usageTotals(groupIndex) + CDbl(rawData(r, 12))
            End If
        End If
    Next r

    If groupCount = 0 Then Exit Sub

    ReDim resultData(1 To groupCount, 1 To 12)

    For groupIndex = 1 To groupCount
        resultData(groupIndex, 1) = vendorNames(groupIndex)
        resultData(groupIndex, 2) = productCodes(groupIndex)
        resultData(groupIndex, 3) = productNames(groupIndex)
        resultData(groupIndex, 4) = deliveryDates(groupIndex)
        resultData(groupIndex, 5) = usageTotals(groupIndex)
        resultData(groupIndex, 6) = firstUseDates(groupIndex)
        resultData(groupIndex, 7) = lastUseDates(groupIndex)
        resultData(groupIndex, 8) = targetRawRows(groupIndex)
        resultData(groupIndex, 9) = IIf(morningFound(groupIndex), "TRUE", "FALSE")
        resultData(groupIndex, 10) = orderSheets(groupIndex)
        resultData(groupIndex, 11) = rawRowLists(groupIndex)
        resultData(groupIndex, 12) = aggregateKeys(groupIndex)
    Next groupIndex

    wsAgg.Range("A2").Resize(groupCount, 12).Value2 = resultData
    wsAgg.Columns("B").NumberFormat = "@"
    wsAgg.Visible = xlSheetVeryHidden

End Sub

Private Function HRS_AggregateProductIdentity( _
    ByVal productCode As String, _
    ByVal productName As String) As String

    productCode = HRS_NormalizeCode(productCode)

    If productCode <> "" Then
        HRS_AggregateProductIdentity = "C|" & productCode
    Else
        HRS_AggregateProductIdentity = _
            "N|" & UCase$(Trim$(productName))
    End If

End Function

Private Sub HRS_LoadPrebuiltAggregateForCurrentProduct()

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim wsAgg As Worksheet
    Dim aggData As Variant
    Dim detailData As Variant
    Dim outputData() As Variant
    Dim aggLastRow As Long
    Dim detailLastRow As Long
    Dim aggIndex As Long
    Dim detailIndex As Long
    Dim outputCount As Long
    Dim vendorName As String
    Dim productCode As String
    Dim productName As String
    Dim identityText As String
    Dim aggIdentity As String
    Dim deliveryKey As String
    Dim detailKey As String
    Dim rowList As String
    Dim targetDetailRow As Long
    Dim firstDetailRow As Long
    Dim morningDetailRow As Long
    Dim allCancelled As Boolean
    Dim distributionTotal As Double
    Dim hasDistribution As Boolean
    Dim categoryText As String

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)
    Set wsAgg = ThisWorkbook.Worksheets(SH_ALL_AGGREGATE_CACHE)

    wsCache.Range("O1:Z5000").ClearContents
    wsCache.Range("O1:Z1").value = Array( _
        "取消", "使用日表示", "区分", "使用数量", "納品日", _
        "配分後", "商品", "発注書", "書戻対象行", "明細行一覧", _
        "最初使用日", "最後使用日")

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    productCode = HRS_NormalizeCode(wsCache.Range("N3").value)
    productName = Trim$(CStr(wsCache.Range("N4").value))
    identityText = HRS_AggregateProductIdentity(productCode, productName)

    aggLastRow = HRS_LastRow(wsAgg, 1)
    detailLastRow = HRS_LastRow(wsCache, 2)

    If aggLastRow < 2 Or detailLastRow < 2 Then
        wsCache.Range("N8").value = "FALSE"
        Exit Sub
    End If

    aggData = wsAgg.Range("A2:L" & aggLastRow).Value2
    detailData = wsCache.Range("A2:N" & detailLastRow).Value2
    ReDim outputData(1 To UBound(aggData, 1), 1 To 12)

    For aggIndex = 1 To UBound(aggData, 1)

        aggIdentity = HRS_AggregateProductIdentity( _
            CStr(aggData(aggIndex, 2)), _
            CStr(aggData(aggIndex, 3)))

        If StrComp(Trim$(CStr(aggData(aggIndex, 1))), _
                   vendorName, vbTextCompare) = 0 And _
           StrComp(aggIdentity, identityText, vbTextCompare) = 0 Then

            deliveryKey = CStr(aggData(aggIndex, 12))
            rowList = ""
            targetDetailRow = 0
            firstDetailRow = 0
            morningDetailRow = 0
            allCancelled = True
            distributionTotal = 0
            hasDistribution = False

            For detailIndex = 1 To UBound(detailData, 1)

                If Trim$(CStr(detailData(detailIndex, 2))) <> "" Then
                    detailKey = HRS_AggregateKey(detailData(detailIndex, 5))

                    If detailKey = "" Then
                        detailKey = "BLANK_DETAIL_" & CStr(detailIndex + 1)
                    End If

                    If detailKey = deliveryKey Or _
                       (Left$(deliveryKey, 10) = "BLANK_RAW_" And _
                        Left$(detailKey, 13) = "BLANK_DETAIL_") Then

                        If firstDetailRow = 0 Then
                            firstDetailRow = detailIndex + 1
                        End If

                        If rowList <> "" Then rowList = rowList & ","
                        rowList = rowList & CStr(detailIndex + 1)

                        categoryText = CStr(detailData(detailIndex, 3))

                        If morningDetailRow = 0 And _
                           HRS_IsMorningCategory(categoryText) Then
                            morningDetailRow = detailIndex + 1
                        End If

                        If CStr(detailData(detailIndex, 1)) <> "■" Then
                            allCancelled = False
                        End If

                        If Trim$(CStr(detailData(detailIndex, 6))) <> "" And _
                           IsNumeric(detailData(detailIndex, 6)) Then
                            distributionTotal = distributionTotal + _
                                CDbl(detailData(detailIndex, 6))
                            hasDistribution = True
                        End If
                    End If
                End If
            Next detailIndex

            If morningDetailRow > 0 Then
                targetDetailRow = morningDetailRow
            Else
                targetDetailRow = firstDetailRow
            End If

            outputCount = outputCount + 1
            outputData(outputCount, 1) = IIf(allCancelled, "■", "□")
            outputData(outputCount, 2) = HRS_AggregateUseDateText( _
                aggData(aggIndex, 6), aggData(aggIndex, 7))
            outputData(outputCount, 3) = "朝へ集約"
            outputData(outputCount, 4) = aggData(aggIndex, 5)
            outputData(outputCount, 5) = aggData(aggIndex, 4)

            If hasDistribution And distributionTotal <> 0 Then
                outputData(outputCount, 6) = distributionTotal
            Else
                outputData(outputCount, 6) = ""
            End If

            outputData(outputCount, 7) = aggData(aggIndex, 3)
            outputData(outputCount, 8) = aggData(aggIndex, 10)
            outputData(outputCount, 9) = targetDetailRow
            outputData(outputCount, 10) = rowList
            outputData(outputCount, 11) = aggData(aggIndex, 6)
            outputData(outputCount, 12) = aggData(aggIndex, 7)
        End If
    Next aggIndex

    If outputCount > 0 Then
        wsCache.Range("O2").Resize(outputCount, 12).Value2 = outputData
    End If

    wsCache.Range("N8").value = "FALSE"

End Sub

Private Sub HRS_BuildPreviewCache( _
    ByVal productCode As String, _
    ByVal productName As String)

    Dim wsInput As Worksheet
    Dim wsRaw As Worksheet
    Dim wsCache As Worksheet
    Dim vendorName As String
    Dim lastRow As Long
    Dim r As Long
    Dim outputRow As Long
    Dim sessionRow As Long

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsRaw = ThisWorkbook.Worksheets(SH_RAW)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    Dim aggregateMode As Boolean

    aggregateMode = HRS_IsDeliveryAggregateMode(wsCache)
    wsCache.Range("A1:Z5000").ClearContents

    wsCache.Range("A1:L1").value = Array( _
        "取消", "使用日", "区分", "使用数量", "納品日", "配分後", _
        "単位/注意点", "商品", "発注書", "セル番地", _
        "取込元ブック", "変更済")

    'N列は表示位置・選択商品専用。変更済みL列と分離する。
    wsCache.Range("N1").value = "表示位置"
    wsCache.Range("N2").value = 1
    wsCache.Range("N3").NumberFormat = "@"
    wsCache.Range("N3").value = productCode
    wsCache.Range("N4").value = productName
    wsCache.Range("N5").value = 1
    wsCache.Range("N6").value = 1
    wsCache.Range("N7").value = IIf(aggregateMode, "TRUE", "FALSE")
    wsCache.Range("N8").value = "TRUE"

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    lastRow = HRS_LastRow(wsRaw, 1)
    outputRow = 2

    For r = 2 To lastRow

        If Trim$(CStr(wsRaw.Cells(r, 5).value)) = vendorName Then

            If HRS_NormalizeCode(wsRaw.Cells(r, 6).value) = _
               productCode Or _
               HRS_ProductNameMatches( _
                   CStr(wsRaw.Cells(r, 7).value), productName) Then

                sessionRow = HRS_FindSessionRow( _
                    vendorName, productCode, productName, _
                    CStr(wsRaw.Cells(r, 10).value), _
                    CStr(wsRaw.Cells(r, 11).value), _
                    CStr(wsRaw.Cells(r, 9).value), _
                    CStr(wsRaw.Cells(r, 3).value), _
                    CStr(wsRaw.Cells(r, 15).value))

                If sessionRow > 0 Then
                    HRS_CopySessionToPreview _
                        sessionRow, wsCache, outputRow
                Else
                    If HRS_IsPointOne(wsRaw.Cells(r, 12).value) Then
                        wsCache.Cells(outputRow, 1).value = "●"
                    Else
                        wsCache.Cells(outputRow, 1).value = "○"
                    End If
                    wsCache.Cells(outputRow, 2).value = _
                        wsRaw.Cells(r, 10).value
                    wsCache.Cells(outputRow, 3).value = _
                        wsRaw.Cells(r, 11).value
                    wsCache.Cells(outputRow, 4).value = _
                        wsRaw.Cells(r, 12).value
                    wsCache.Cells(outputRow, 5).value = _
                        wsRaw.Cells(r, 9).value
                    wsCache.Cells(outputRow, 6).value = ""
                    wsCache.Cells(outputRow, 7).value = _
                        HRS_GetUnitNote(productCode, productName)
                    wsCache.Cells(outputRow, 8).value = _
                        wsRaw.Cells(r, 7).value
                    wsCache.Cells(outputRow, 9).value = _
                        wsRaw.Cells(r, 3).value
                    wsCache.Cells(outputRow, 10).value = _
                        wsRaw.Cells(r, 15).value
                    wsCache.Cells(outputRow, 11).value = _
                        wsRaw.Cells(r, 2).value
                    wsCache.Cells(outputRow, 12).value = "FALSE"
                End If

                outputRow = outputRow + 1
            End If
        End If
    Next r

    '商品選択時に読込済みの集約キャッシュを展開する。
    HRS_LoadPrebuiltAggregateForCurrentProduct

End Sub

Private Sub HRS_CopySessionToPreview(ByVal sessionRow As Long, _
                                     ByVal wsCache As Worksheet, _
                                     ByVal outputRow As Long)

    Dim wsSession As Worksheet

    Set wsSession = ThisWorkbook.Worksheets(SH_SESSION)

    wsCache.Cells(outputRow, 1).value = _
        IIf(CBool(wsSession.Cells(sessionRow, 7).value), "■", "□")
    wsCache.Cells(outputRow, 2).value = wsSession.Cells(sessionRow, 8).value
    wsCache.Cells(outputRow, 3).value = wsSession.Cells(sessionRow, 9).value
    wsCache.Cells(outputRow, 4).value = wsSession.Cells(sessionRow, 10).value
    wsCache.Cells(outputRow, 5).value = wsSession.Cells(sessionRow, 11).value
    wsCache.Cells(outputRow, 6).value = wsSession.Cells(sessionRow, 12).value
    wsCache.Cells(outputRow, 7).value = wsSession.Cells(sessionRow, 13).value
    wsCache.Cells(outputRow, 8).value = wsSession.Cells(sessionRow, 4).value
    wsCache.Cells(outputRow, 9).value = wsSession.Cells(sessionRow, 14).value
    wsCache.Cells(outputRow, 10).value = wsSession.Cells(sessionRow, 15).value
    wsCache.Cells(outputRow, 11).value = wsSession.Cells(sessionRow, 16).value
    wsCache.Cells(outputRow, 12).value = "TRUE"

End Sub

Private Function HRS_IsAggregateCacheDirty(ByVal wsCache As Worksheet) As Boolean

    HRS_IsAggregateCacheDirty = _
        (UCase$(Trim$(CStr(wsCache.Range("N8").value))) <> "FALSE")

End Function

Private Sub HRS_MarkAggregateCacheDirty()

    Dim wsCache As Worksheet

    If Not HRS_SheetExists(SH_PREVIEW_CACHE, ThisWorkbook) Then Exit Sub

    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)
    wsCache.Range("N8").value = "TRUE"

End Sub

Private Function HRS_IsDeliveryAggregateMode(ByVal wsCache As Worksheet) As Boolean

    HRS_IsDeliveryAggregateMode = _
        (UCase$(Trim$(CStr(wsCache.Range("N7").value))) = "TRUE")

End Function

Private Function HRS_AggregateKey(ByVal deliveryValue As Variant) As String

    If IsDate(deliveryValue) Then
        HRS_AggregateKey = Format$(CDate(deliveryValue), "yyyymmdd")
    Else
        HRS_AggregateKey = Trim$(CStr(deliveryValue))
    End If

End Function

Private Function HRS_IsMorningCategory(ByVal categoryText As String) As Boolean

    HRS_IsMorningCategory = _
        (InStr(1, Trim$(categoryText), "朝", vbTextCompare) > 0)

End Function

Private Function HRS_RowListContains(ByVal rowList As String, _
                                     ByVal rowNumber As Long) As Boolean

    Dim parts As Variant
    Dim item As Variant

    If Trim$(rowList) = "" Then Exit Function

    parts = Split(rowList, ",")

    For Each item In parts
        If CLng(Val(item)) = rowNumber Then
            HRS_RowListContains = True
            Exit Function
        End If
    Next item

End Function

Private Function HRS_GroupAllCancelled(ByVal wsCache As Worksheet, _
                                       ByVal rowList As String) As Boolean

    Dim parts As Variant
    Dim item As Variant
    Dim detailRow As Long
    Dim hasRow As Boolean

    If Trim$(rowList) = "" Then Exit Function

    HRS_GroupAllCancelled = True
    parts = Split(rowList, ",")

    For Each item In parts
        detailRow = CLng(Val(item))
        If detailRow > 0 Then
            hasRow = True
            If CStr(wsCache.Cells(detailRow, 1).value) <> "■" Then
                HRS_GroupAllCancelled = False
                Exit Function
            End If
        End If
    Next item

    If Not hasRow Then HRS_GroupAllCancelled = False

End Function

Private Function HRS_AggregateUseDateText(ByVal firstDate As Variant, _
                                          ByVal lastDate As Variant) As String

    If Trim$(CStr(firstDate)) = "" Then Exit Function

    If IsDate(firstDate) And IsDate(lastDate) Then
        If CLng(CDate(firstDate)) = CLng(CDate(lastDate)) Then
            HRS_AggregateUseDateText = Format$(CDate(firstDate), "m/d")
        Else
            HRS_AggregateUseDateText = _
                Format$(CDate(firstDate), "m/d") & "～" & _
                Format$(CDate(lastDate), "m/d")
        End If
    ElseIf CStr(firstDate) = CStr(lastDate) Then
        HRS_AggregateUseDateText = CStr(firstDate)
    Else
        HRS_AggregateUseDateText = _
            CStr(firstDate) & "～" & CStr(lastDate)
    End If

End Function

Private Sub HRS_BuildDeliveryAggregateCache()

    '集約構造は発注書読込時に作成済み。
    'ここでは現在商品の配分・取消状態だけを重ねて表示する。
    HRS_LoadPrebuiltAggregateForCurrentProduct

End Sub

Private Sub HRS_SetAggregateGroupCancel(ByVal wsCache As Worksheet, _
                                        ByVal rowList As String, _
                                        ByVal turnOn As Boolean)

    Dim parts As Variant
    Dim item As Variant
    Dim detailRow As Long

    If Trim$(rowList) = "" Then Exit Sub

    parts = Split(rowList, ",")

    For Each item In parts
        detailRow = CLng(Val(item))
        If detailRow > 0 Then
            wsCache.Cells(detailRow, 1).value = _
                IIf(turnOn, "■", "□")
            wsCache.Cells(detailRow, 12).value = "TRUE"
        End If
    Next item

End Sub

Private Sub HRS_SetAggregateDistribution(ByVal wsCache As Worksheet, _
                                         ByVal rowList As String, _
                                         ByVal targetRow As Long, _
                                         ByVal newValue As Variant)

    Dim parts As Variant
    Dim item As Variant
    Dim detailRow As Long
    Dim numericValue As Double

    If Trim$(rowList) = "" Then Exit Sub

    parts = Split(rowList, ",")

    For Each item In parts
        detailRow = CLng(Val(item))
        If detailRow > 0 Then
            wsCache.Cells(detailRow, 6).ClearContents
            wsCache.Cells(detailRow, 12).value = "TRUE"
        End If
    Next item

    If Trim$(CStr(newValue)) <> "" Then
        numericValue = CDbl(newValue)

        If numericValue = Fix(numericValue) Then
            wsCache.Cells(targetRow, 6).value = CLng(numericValue)
        Else
            wsCache.Cells(targetRow, 6).value = numericValue
        End If
    End If

End Sub

Private Sub HRS_UpdateAggregateButtonCaption()

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim captionText As String

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    If HRS_IsDeliveryAggregateMode(wsCache) Then
        captionText = "集約 ON"
    Else
        captionText = "集約 OFF"
    End If

    On Error Resume Next
    wsInput.Shapes("HRS_AggregateMode").TextFrame.Characters.Text = captionText
    On Error GoTo 0

End Sub

Public Sub HRS_ToggleDeliveryAggregate()

    Dim wsCache As Worksheet

    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    wsCache.Range("N7").value = _
        IIf(HRS_IsDeliveryAggregateMode(wsCache), "FALSE", "TRUE")
    wsCache.Range("N2").value = 1
    wsCache.Range("N5").value = 1

    HRS_RenderPreviewPage
    HRS_BuildCurrentWriteBackPreview
    HRS_UpdateAggregateButtonCaption

End Sub

Private Function HRS_GetJapaneseWeekday(ByVal dateValue As Variant) As String

    Dim dayNames As Variant
    Dim dateSerial As Date

    If IsEmpty(dateValue) Or Trim$(CStr(dateValue)) = "" Then Exit Function
    If Not IsDate(dateValue) Then Exit Function

    dateSerial = CDate(dateValue)
    dayNames = Array("日", "月", "火", "水", "木", "金", "土")
    HRS_GetJapaneseWeekday = CStr(dayNames(Weekday(dateSerial, vbSunday) - 1))

End Function

Public Sub HRS_RenderPreviewPage()

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim total As Long
    Dim pos As Long
    Dim i As Long
    Dim sourceRow As Long
    Dim displayRow As Long
    Dim lastShow As Long
    Dim distributionValue As Variant
    Dim deliveryDateValue As Variant
    Dim aggregateMode As Boolean
    Dim dataStartCol As Long

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    aggregateMode = HRS_IsDeliveryAggregateMode(wsCache)

    If aggregateMode Then
        dataStartCol = 15
        total = HRS_LastRow(wsCache, 16) - 1
    Else
        dataStartCol = 1
        total = HRS_LastRow(wsCache, 2) - 1
    End If

    wsInput.Range("K8:S20").ClearContents
    wsInput.Range("K8:S20").Font.Strikethrough = False
    wsInput.Range("K8:S20").Font.Size = 11
    wsInput.Range("K8:S20").Interior.Color = RGB(226, 239, 218)
    wsInput.Range("L8:L20").Font.Size = 11
    wsInput.Range("Q8:Q20").Interior.Color = RGB(255, 242, 204)
    wsInput.Range("Q8:Q20").NumberFormat = "General"
    wsInput.Range("P8:P20").NumberFormat = "General"
    wsInput.Range("L8:L20").NumberFormat = "yyyy/m/d"

    If aggregateMode Then
        wsInput.Range("N8:N20").NumberFormat = "@"
    Else
        wsInput.Range("N8:N20").NumberFormat = "yyyy/m/d"
    End If

    If total < 0 Then total = 0

    pos = CLng(Val(wsCache.Range("N2").value))
    If pos < 1 Then pos = 1

    For i = 0 To PREVIEW_PAGE_SIZE - 1

        sourceRow = pos + i + 1
        displayRow = PREVIEW_TOP + i

        If sourceRow > total + 1 Then Exit For

        If aggregateMode Then
            deliveryDateValue = wsCache.Cells(sourceRow, 19).value

            wsInput.Cells(displayRow, "K").value = _
                wsCache.Cells(sourceRow, 15).value
            wsInput.Cells(displayRow, "L").value = deliveryDateValue
            wsInput.Cells(displayRow, "M").value = _
                HRS_GetJapaneseWeekday(deliveryDateValue)
            wsInput.Cells(displayRow, "N").value = _
                wsCache.Cells(sourceRow, 16).value
            wsInput.Cells(displayRow, "O").value = _
                wsCache.Cells(sourceRow, 17).value
            wsInput.Cells(displayRow, "P").value = _
                wsCache.Cells(sourceRow, 18).value

            distributionValue = wsCache.Cells(sourceRow, 20).value

            wsInput.Cells(displayRow, "R").value = _
                wsCache.Cells(sourceRow, 21).value
            wsInput.Cells(displayRow, "S").value = _
                wsCache.Cells(sourceRow, 22).value
        Else
            deliveryDateValue = wsCache.Cells(sourceRow, 5).value

            wsInput.Cells(displayRow, "K").value = _
                wsCache.Cells(sourceRow, 1).value
            wsInput.Cells(displayRow, "L").value = deliveryDateValue
            wsInput.Cells(displayRow, "M").value = _
                HRS_GetJapaneseWeekday(deliveryDateValue)
            wsInput.Cells(displayRow, "N").value = _
                wsCache.Cells(sourceRow, 2).value
            wsInput.Cells(displayRow, "O").value = _
                wsCache.Cells(sourceRow, 3).value
            wsInput.Cells(displayRow, "P").value = _
                wsCache.Cells(sourceRow, 4).value

            distributionValue = wsCache.Cells(sourceRow, 6).value

            wsInput.Cells(displayRow, "R").value = _
                wsCache.Cells(sourceRow, 8).value
            wsInput.Cells(displayRow, "S").value = _
                wsCache.Cells(sourceRow, 9).value
        End If

        If Trim$(CStr(distributionValue)) = "" Then
            wsInput.Cells(displayRow, "Q").ClearContents
        ElseIf IsNumeric(distributionValue) Then
            If CDbl(distributionValue) = Fix(CDbl(distributionValue)) Then
                wsInput.Cells(displayRow, "Q").value = CLng(distributionValue)
            Else
                wsInput.Cells(displayRow, "Q").value = CDbl(distributionValue)
            End If
        Else
            wsInput.Cells(displayRow, "Q").value = distributionValue
        End If

        wsInput.Cells(displayRow, "Q").NumberFormat = "General"

        If CStr(wsInput.Cells(displayRow, "K").value) = "■" Then
            wsInput.Range("L" & displayRow & ":S" & displayRow).Font.Strikethrough = True
        End If

        If Trim$(CStr(distributionValue)) <> "" And _
           IsNumeric(distributionValue) And _
           CDbl(distributionValue) > 0 Then

            With wsInput.Range("K" & displayRow & ":S" & displayRow)
                .Interior.Color = RGB(217, 217, 217)
            End With
        Else
            wsInput.Range("K" & displayRow & ":S" & displayRow). _
                Interior.Color = RGB(226, 239, 218)
            wsInput.Cells(displayRow, "Q").Interior.Color = _
                RGB(255, 242, 204)
        End If

    Next i

    HRS_UpdateAllCancelCaption
    HRS_UpdateAggregateButtonCaption

    If total = 0 Then
        wsInput.Range("N6").value = "0件"
    Else
        lastShow = WorksheetFunction.Min(pos + PREVIEW_PAGE_SIZE - 1, total)
        wsInput.Range("N6").value = pos & "～" & lastShow & "/" & total & "件"
    End If

    With wsInput.Range("N6:N6")
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With

End Sub

Public Sub HRS_PreviewPrev()

    Dim wsCache As Worksheet
    Dim pos As Long

    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    pos = CLng(Val(wsCache.Range("N2").value))
    pos = pos - PREVIEW_PAGE_SIZE
    If pos < 1 Then pos = 1

    wsCache.Range("N2").value = pos

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    HRS_RenderPreviewPage
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Public Sub HRS_PreviewNext()

    Dim wsCache As Worksheet
    Dim total As Long
    Dim pos As Long

    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    total = HRS_LastRow(wsCache, 2) - 1
    pos = CLng(Val(wsCache.Range("N2").value))
    pos = pos + PREVIEW_PAGE_SIZE

    If pos > total Then
        pos = total - PREVIEW_PAGE_SIZE + 1
    End If

    If pos < 1 Then pos = 1

    wsCache.Range("N2").value = pos

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    HRS_RenderPreviewPage
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

'=========================================================
' 発注数・配分
'=========================================================

Public Sub HRS_ToggleProductConfirm(ByVal selectedRow As Long)

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim vendorName As String
    Dim productCode As String
    Dim productName As String
    Dim orderQty As Double
    Dim confirmed As Boolean
    Dim deleteStatus As Long
    Dim cacheRow As Long

    If selectedRow < ITEM_TOP Or selectedRow > ITEM_BOTTOM Then Exit Sub

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PRODUCT_CACHE)

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    productCode = Trim$(CStr(wsInput.Cells(selectedRow, "A").value))
    productName = Trim$(CStr(wsInput.Cells(selectedRow, "B").value))
    orderQty = CDbl(Val(wsInput.Cells(selectedRow, "F").value))

    If productName = "" Then Exit Sub

    deleteStatus = HRS_GetDeleteItemStatus(productName)

    If deleteStatus = 2 Then
        confirmed = True
    Else
        confirmed = (CStr(wsInput.Cells(selectedRow, "I").value) <> "■")
    End If

    HRS_BuildPreviewCache productCode, productName
    If deleteStatus > 0 Then HRS_ApplyDeleteItemToPreview productCode, productName

    HRS_SavePreviewToSession vendorName, productCode, productName, orderQty, confirmed

    cacheRow = HRS_FindProductCacheRow(productCode, productName)
    If cacheRow > 0 Then wsCache.Cells(cacheRow, 8).value = IIf(confirmed, "■", "□")

    wsInput.Cells(selectedRow, "I").value = IIf(confirmed, "■", "□")
    '確認チェックだけではグレーにしない。配分後入力がある商品だけグレー表示。
    HRS_UpdateDisplayedProductFill selectedRow

    If deleteStatus = 1 Then
        wsInput.Cells(selectedRow, "B").Font.Color = RGB(255, 0, 0)
        wsInput.Cells(selectedRow, "B").Font.Bold = True
    End If

    '選択位置と表示中の商品は変更しない。
    HRS_RenderPreviewPage
    HRS_BuildCurrentWriteBackPreview
    HRS_UpdateMilkCookingDetail productName

End Sub

Public Sub HRS_OnOrderQtyChanged(ByVal selectedRow As Long)

    Dim wsInput As Worksheet
    Dim vendorName As String
    Dim productCode As String
    Dim productName As String
    Dim orderQty As Double
    Dim deleteStatus As Long
    Dim inputValue As Variant

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    productCode = Trim$(CStr( _
        wsInput.Cells(selectedRow, "A").value))
    productName = Trim$(CStr( _
        wsInput.Cells(selectedRow, "B").value))
    inputValue = wsInput.Cells(selectedRow, "F").value

    If productName = "" Then Exit Sub

    If Trim$(CStr(inputValue)) = "" Then

        orderQty = 0
        wsInput.Cells(selectedRow, "F").ClearContents

    ElseIf IsNumeric(inputValue) Then

        orderQty = CDbl(inputValue)

        If orderQty = Fix(orderQty) Then
            wsInput.Cells(selectedRow, "F").value = CLng(orderQty)
        Else
            wsInput.Cells(selectedRow, "F").value = orderQty
        End If

        wsInput.Cells(selectedRow, "F").NumberFormat = "General"

    Else

        MsgBox "発注数には数値を入力してください。", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    deleteStatus = HRS_GetDeleteItemStatus(productName)

    HRS_BuildPreviewCache productCode, productName
    HRS_ApplyDeleteItemToPreview productCode, productName

    If deleteStatus = 1 Then
        '削除項目B列の赤文字商品：
        '全取消は維持しつつ使用数量上位2か所へ配分する。
        HRS_DistributeOrderQty orderQty, True
    Else
        HRS_DistributeOrderQty orderQty, False
    End If

    HRS_SavePreviewToSession _
        vendorName, productCode, productName, orderQty, True

    wsInput.Cells(selectedRow, "I").value = "■"

    '配分後入力の有無に合わせて商品一覧の塗りつぶしを更新する。
    HRS_UpdateDisplayedProductFill selectedRow

    If deleteStatus = 1 Then
        wsInput.Cells(selectedRow, "B").Font.Color = RGB(255, 0, 0)
        wsInput.Cells(selectedRow, "B").Font.Bold = True
    End If

    HRS_RenderPreviewPage
    HRS_BuildCurrentWriteBackPreview

End Sub

Private Sub HRS_DistributeOrderQty( _
    ByVal orderQty As Double, _
    Optional ByVal includeCancelled As Boolean = False)

    Dim wsCache As Worksheet
    Dim lastRow As Long
    Dim top1 As Long
    Dim top2 As Long
    Dim max1 As Double
    Dim max2 As Double
    Dim r As Long
    Dim usageQty As Double
    Dim firstQty As Double
    Dim secondQty As Double
    Dim eligibleCount As Long
    Dim isEligible As Boolean

    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)
    lastRow = HRS_LastRow(wsCache, 2)

    wsCache.Range( _
        "F2:F" & WorksheetFunction.Max(lastRow, 2)).ClearContents

    wsCache.Range( _
        "L2:L" & WorksheetFunction.Max(lastRow, 2)).value = "FALSE"

    If orderQty <= 0 Then Exit Sub

    For r = 2 To lastRow

        isEligible = False

        If Trim$(CStr(wsCache.Cells(r, 2).value)) <> "" Then

            If includeCancelled Then
                '赤文字商品は全取消状態でも配分対象にする。
                isEligible = True
            ElseIf CStr(wsCache.Cells(r, 1).value) <> "■" Then
                isEligible = True
            End If

        End If

        If isEligible Then

            eligibleCount = eligibleCount + 1
            usageQty = CDbl(Val(wsCache.Cells(r, 4).value))

            If top1 = 0 Or usageQty > max1 Then

                max2 = max1
                top2 = top1
                max1 = usageQty
                top1 = r

            ElseIf top2 = 0 Or usageQty > max2 Then

                max2 = usageQty
                top2 = r

            End If
        End If
    Next r

    If eligibleCount = 0 Then Exit Sub

    If orderQty = Fix(orderQty) Then

        firstQty = Fix(orderQty / 2)
        secondQty = orderQty - firstQty

        If top1 > 0 Then
            wsCache.Cells(top1, 6).value = CLng(secondQty)
            wsCache.Cells(top1, 12).value = "TRUE"
        End If

        If top2 > 0 Then
            wsCache.Cells(top2, 6).value = CLng(firstQty)
            wsCache.Cells(top2, 12).value = "TRUE"
        ElseIf top1 > 0 Then
            wsCache.Cells(top1, 6).value = CLng(orderQty)
        End If

    Else

        firstQty = orderQty / 2
        secondQty = orderQty - firstQty

        If top1 > 0 Then
            wsCache.Cells(top1, 6).value = secondQty
            wsCache.Cells(top1, 12).value = "TRUE"
        End If

        If top2 > 0 Then
            wsCache.Cells(top2, 6).value = firstQty
            wsCache.Cells(top2, 12).value = "TRUE"
        ElseIf top1 > 0 Then
            wsCache.Cells(top1, 6).value = orderQty
        End If

    End If

End Sub

Public Sub HRS_OnDistributionChanged(ByVal displayRow As Long)

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim sourceRow As Long
    Dim pos As Long
    Dim newValue As Variant
    Dim numericValue As Double
    Dim vendorName As String
    Dim productCode As String
    Dim productName As String
    Dim orderQty As Double
    Dim allCancelledBefore As Boolean
    Dim deleteStatus As Long
    Dim aggregateMode As Boolean
    Dim rowList As String
    Dim targetRow As Long

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    pos = CLng(Val(wsCache.Range("N2").value))
    If pos < 1 Then pos = 1

    sourceRow = pos + (displayRow - PREVIEW_TOP) + 1
    newValue = wsInput.Cells(displayRow, "Q").value
    aggregateMode = HRS_IsDeliveryAggregateMode(wsCache)

    allCancelledBefore = HRS_AllCancelled(wsCache)
    deleteStatus = HRS_GetDeleteItemStatus(CStr(wsCache.Range("N4").value))

    If aggregateMode Then

        rowList = CStr(wsCache.Cells(sourceRow, 24).value)
        targetRow = CLng(Val(wsCache.Cells(sourceRow, 23).value))

        If targetRow <= 0 Or Trim$(rowList) = "" Then Exit Sub

        If Trim$(CStr(newValue)) = "" Then
            HRS_SetAggregateDistribution wsCache, rowList, targetRow, ""

            wsInput.Range("K" & displayRow & ":S" & displayRow). _
                Interior.Color = RGB(226, 239, 218)
            wsInput.Cells(displayRow, "Q").Interior.Color = _
                RGB(255, 242, 204)

            If allCancelledBefore Or deleteStatus > 0 Then
                HRS_SetAggregateGroupCancel wsCache, rowList, True
            Else
                HRS_SetAggregateGroupCancel wsCache, rowList, False
            End If
        ElseIf IsNumeric(newValue) Then
            numericValue = CDbl(newValue)
            HRS_SetAggregateDistribution wsCache, rowList, targetRow, numericValue

            '集約ONでも、配分後へ正数を入力した時点で行全体をグレーにする。
            If numericValue > 0 Then
                wsInput.Range("K" & displayRow & ":S" & displayRow). _
                    Interior.Color = RGB(217, 217, 217)
            Else
                wsInput.Range("K" & displayRow & ":S" & displayRow). _
                    Interior.Color = RGB(226, 239, 218)
                wsInput.Cells(displayRow, "Q").Interior.Color = _
                    RGB(255, 242, 204)
            End If

            If allCancelledBefore Or deleteStatus > 0 Or numericValue > 0 Then
                HRS_SetAggregateGroupCancel wsCache, rowList, True
            Else
                HRS_SetAggregateGroupCancel wsCache, rowList, False
            End If
        Else
            MsgBox "配分後には数値を入力してください。", vbExclamation, APP_NAME
            HRS_RenderPreviewPage
            Exit Sub
        End If
    Else

        If Trim$(CStr(newValue)) = "" Then
            wsCache.Cells(sourceRow, 6).ClearContents

            If allCancelledBefore Or deleteStatus > 0 Then
                wsCache.Cells(sourceRow, 1).value = "■"
            Else
                wsCache.Cells(sourceRow, 1).value = "□"
            End If
        ElseIf IsNumeric(newValue) Then
            numericValue = CDbl(newValue)

            If numericValue = Fix(numericValue) Then
                wsCache.Cells(sourceRow, 6).value = CLng(numericValue)
            Else
                wsCache.Cells(sourceRow, 6).value = numericValue
            End If

            If allCancelledBefore Or deleteStatus > 0 Then
                wsCache.Cells(sourceRow, 1).value = "■"
            ElseIf numericValue > 0 Then
                wsCache.Cells(sourceRow, 1).value = "■"
            Else
                wsCache.Cells(sourceRow, 1).value = "□"
            End If
        Else
            MsgBox "配分後には数値を入力してください。", vbExclamation, APP_NAME
            HRS_RenderPreviewPage
            Exit Sub
        End If

        wsCache.Cells(sourceRow, 12).value = "TRUE"
    End If

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    productCode = CStr(wsCache.Range("N3").value)
    productName = CStr(wsCache.Range("N4").value)
    orderQty = HRS_GetSavedOrderQty(vendorName, productCode, productName)

    HRS_SavePreviewToSession vendorName, productCode, productName, orderQty, True
    HRS_LoadPrebuiltAggregateForCurrentProduct
    HRS_MarkDisplayedProductConfirmed productCode, productName
    HRS_RenderPreviewPage
    HRS_BuildCurrentWriteBackPreview
    HRS_UpdateMilkCookingDetail productName

End Sub

'=========================================================
' 取消
'=========================================================
Public Sub HRS_TogglePreviewCancel(ByVal displayRow As Long)

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim sourceRow As Long
    Dim pos As Long
    Dim vendorName As String
    Dim productCode As String
    Dim productName As String
    Dim orderQty As Double
    Dim aggregateMode As Boolean
    Dim rowList As String
    Dim turnOn As Boolean

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    pos = CLng(Val(wsCache.Range("N2").value))
    sourceRow = pos + (displayRow - PREVIEW_TOP) + 1
    aggregateMode = HRS_IsDeliveryAggregateMode(wsCache)

    If aggregateMode Then
        rowList = CStr(wsCache.Cells(sourceRow, 24).value)
        turnOn = (CStr(wsCache.Cells(sourceRow, 15).value) <> "■")
        HRS_SetAggregateGroupCancel wsCache, rowList, turnOn
    Else
        If CStr(wsCache.Cells(sourceRow, 1).value) = "■" Then
            wsCache.Cells(sourceRow, 1).value = "□"
        Else
            wsCache.Cells(sourceRow, 1).value = "■"
        End If
        wsCache.Cells(sourceRow, 12).value = "TRUE"
    End If

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    productCode = CStr(wsCache.Range("N3").value)
    productName = CStr(wsCache.Range("N4").value)
    orderQty = HRS_GetSavedOrderQty(vendorName, productCode, productName)

    HRS_SavePreviewToSession vendorName, productCode, productName, orderQty, True
    HRS_LoadPrebuiltAggregateForCurrentProduct
    HRS_RenderPreviewPage
    HRS_BuildCurrentWriteBackPreview

End Sub

Public Sub HRS_ToggleAllCancel()

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim turnOn As Boolean
    Dim vendorName As String
    Dim productCode As String
    Dim productName As String
    Dim orderQty As Double

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    lastRow = HRS_LastRow(wsCache, 2)
    turnOn = Not HRS_AllCancelled(wsCache)

    For r = 2 To lastRow
        If Trim$(CStr(wsCache.Cells(r, 2).value)) <> "" Then
            wsCache.Cells(r, 1).value = IIf(turnOn, "■", "□")
            wsCache.Cells(r, 12).value = "TRUE"
        End If
    Next r

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    productCode = CStr(wsCache.Range("N3").value)
    productName = CStr(wsCache.Range("N4").value)
    orderQty = HRS_GetSavedOrderQty(vendorName, productCode, productName)

    HRS_SavePreviewToSession vendorName, productCode, productName, orderQty, True
    HRS_RenderPreviewPage
    HRS_BuildCurrentWriteBackPreview

End Sub

Private Function HRS_AllCancelled(ByVal wsCache As Worksheet) As Boolean

    Dim lastRow As Long
    Dim r As Long
    Dim hasData As Boolean

    lastRow = HRS_LastRow(wsCache, 2)
    HRS_AllCancelled = True

    For r = 2 To lastRow
        If Trim$(CStr(wsCache.Cells(r, 2).value)) <> "" Then
            hasData = True
            If CStr(wsCache.Cells(r, 1).value) <> "■" Then
                HRS_AllCancelled = False
                Exit Function
            End If
        End If
    Next r

    If Not hasData Then HRS_AllCancelled = False

End Function

Private Sub HRS_UpdateAllCancelCaption()

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    If HRS_AllCancelled(wsCache) Then
        wsInput.Range("K5").value = "■ 全取消"
    Else
        wsInput.Range("K5").value = "□ 全取消"
    End If

    With wsInput.Range("K5:L5")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    wsInput.Range("J4").Font.Bold = True

End Sub

Private Function HRS_GetDeleteItemStatus( _
    ByVal productName As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    If Not HRS_SheetExists(SH_DELETE, ThisWorkbook) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SH_DELETE)

    'B列を最優先
    lastRow = HRS_LastRow(ws, 2)

    For r = 1 To lastRow
        If HRS_ProductNameMatches( _
            CStr(ws.Cells(r, 2).value), productName) Then

            HRS_GetDeleteItemStatus = 1
            Exit Function
        End If
    Next r

    lastRow = HRS_LastRow(ws, 4)

    For r = 1 To lastRow
        If HRS_ProductNameMatches( _
            CStr(ws.Cells(r, 4).value), productName) Then

            HRS_GetDeleteItemStatus = 2
            Exit Function
        End If
    Next r

End Function

Private Sub HRS_ApplyDeleteItemToPreview( _
    ByVal productCode As String, _
    ByVal productName As String)

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim deleteStatus As Long
    Dim vendorName As String
    Dim orderQty As Double
    Dim confirmed As Boolean

    deleteStatus = HRS_GetDeleteItemStatus(productName)

    If deleteStatus = 0 Then Exit Sub

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    lastRow = HRS_LastRow(wsCache, 2)

    For r = 2 To lastRow
        If Trim$(CStr(wsCache.Cells(r, 2).value)) <> "" Then
            wsCache.Cells(r, 1).value = "■"
            wsCache.Cells(r, 12).value = "TRUE"
        End If
    Next r

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    orderQty = HRS_GetSavedOrderQty( _
        vendorName, productCode, productName)

    If deleteStatus = 2 Then
        confirmed = True
    Else
        confirmed = HRS_IsSessionConfirmed( _
            vendorName, productCode, productName)
    End If

    '既にセッションがある場合、またはD列確認済み商品のみ保存。
    'B列の赤文字商品は、開いただけでは確認済みにしない。
    If HRS_HasAnySession( _
        vendorName, productCode, productName) Or _
        deleteStatus = 2 Then

        HRS_SavePreviewToSession _
            vendorName, productCode, productName, orderQty, confirmed
    End If

End Sub

'=========================================================
' セッション保存
'=========================================================
Private Sub HRS_SavePreviewToSession( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal orderQty As Double, _
    ByVal confirmed As Boolean)

    Dim wsSession As Worksheet
    Dim wsCache As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim targetRow As Long

    Set wsSession = ThisWorkbook.Worksheets(SH_SESSION)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    lastRow = HRS_LastRow(wsCache, 2)

    HRS_DeleteSessionForProduct vendorName, productCode, productName

    For r = 2 To lastRow

        If Trim$(CStr(wsCache.Cells(r, 2).value)) <> "" Then

            targetRow = HRS_LastRow(wsSession, 1) + 1
            If targetRow < 2 Then targetRow = 2

            wsSession.Cells(targetRow, 1).value = Now
            wsSession.Cells(targetRow, 2).value = vendorName
            wsSession.Cells(targetRow, 3).NumberFormat = "@"
            wsSession.Cells(targetRow, 3).value = productCode
            wsSession.Cells(targetRow, 4).value = productName
            wsSession.Cells(targetRow, 5).value = orderQty
            wsSession.Cells(targetRow, 6).value = confirmed
            wsSession.Cells(targetRow, 7).value = _
                (CStr(wsCache.Cells(r, 1).value) = "■")
            wsSession.Cells(targetRow, 8).value = wsCache.Cells(r, 2).value
            wsSession.Cells(targetRow, 9).value = wsCache.Cells(r, 3).value
            wsSession.Cells(targetRow, 10).value = wsCache.Cells(r, 4).value
            wsSession.Cells(targetRow, 11).value = wsCache.Cells(r, 5).value
            wsSession.Cells(targetRow, 12).value = wsCache.Cells(r, 6).value
            wsSession.Cells(targetRow, 13).value = wsCache.Cells(r, 7).value
            wsSession.Cells(targetRow, 14).value = wsCache.Cells(r, 9).value
            wsSession.Cells(targetRow, 15).value = wsCache.Cells(r, 10).value
            wsSession.Cells(targetRow, 16).value = wsCache.Cells(r, 11).value
        End If
    Next r

End Sub

Private Sub HRS_DeleteSessionForProduct(ByVal vendorName As String, _
                                        ByVal productCode As String, _
                                        ByVal productName As String)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SH_SESSION)
    lastRow = HRS_LastRow(ws, 1)

    For r = lastRow To 2 Step -1
        If CStr(ws.Cells(r, 2).value) = vendorName And _
           CStr(ws.Cells(r, 3).value) = productCode And _
           CStr(ws.Cells(r, 4).value) = productName Then
            ws.Rows(r).Delete
        End If
    Next r

End Sub

Private Function HRS_FindSessionRow( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal useDateText As String, _
    ByVal mealText As String, _
    ByVal deliveryText As String, _
    ByVal sheetName As String, _
    ByVal cellAddress As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SH_SESSION)
    lastRow = HRS_LastRow(ws, 1)

    For r = 2 To lastRow

        If CStr(ws.Cells(r, 2).value) = vendorName And _
           CStr(ws.Cells(r, 3).value) = productCode And _
           CStr(ws.Cells(r, 4).value) = productName And _
           CStr(ws.Cells(r, 8).value) = useDateText And _
           CStr(ws.Cells(r, 9).value) = mealText And _
           CStr(ws.Cells(r, 11).value) = deliveryText And _
           CStr(ws.Cells(r, 14).value) = sheetName And _
           CStr(ws.Cells(r, 15).value) = cellAddress Then

            HRS_FindSessionRow = r
            Exit Function
        End If
    Next r

End Function



Private Function HRS_HasAnySession( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String) As Boolean

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SH_SESSION)
    lastRow = HRS_LastRow(ws, 1)

    For r = 2 To lastRow
        If CStr(ws.Cells(r, 2).value) = vendorName And _
           CStr(ws.Cells(r, 3).value) = productCode And _
           CStr(ws.Cells(r, 4).value) = productName Then

            HRS_HasAnySession = True
            Exit Function
        End If
    Next r

End Function

Private Function HRS_IsSessionConfirmed( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String) As Boolean

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SH_SESSION)
    lastRow = HRS_LastRow(ws, 1)

    For r = lastRow To 2 Step -1

        If CStr(ws.Cells(r, 2).value) = vendorName And _
           CStr(ws.Cells(r, 3).value) = productCode And _
           CStr(ws.Cells(r, 4).value) = productName Then

            HRS_IsSessionConfirmed = _
                CBool(ws.Cells(r, 6).value)

            Exit Function
        End If
    Next r

End Function

Private Function HRS_HasSession(ByVal vendorName As String, _
                                ByVal productCode As String, _
                                ByVal productName As String) As Boolean

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SH_SESSION)
    lastRow = HRS_LastRow(ws, 1)

    For r = 2 To lastRow
        If CStr(ws.Cells(r, 2).value) = vendorName And _
           CStr(ws.Cells(r, 3).value) = productCode And _
           CStr(ws.Cells(r, 4).value) = productName Then
            HRS_HasSession = CBool(ws.Cells(r, 6).value)
            Exit Function
        End If
    Next r

End Function

Private Function HRS_GetSavedOrderQty(ByVal vendorName As String, _
                                      ByVal productCode As String, _
                                      ByVal productName As String) As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SH_SESSION)
    lastRow = HRS_LastRow(ws, 1)

    For r = 2 To lastRow
        If CStr(ws.Cells(r, 2).value) = vendorName And _
           CStr(ws.Cells(r, 3).value) = productCode And _
           CStr(ws.Cells(r, 4).value) = productName Then
            HRS_GetSavedOrderQty = CDbl(Val(ws.Cells(r, 5).value))
            Exit Function
        End If
    Next r

End Function


Public Sub HRS_CurrentWriteBackPrev()

    Dim wsCache As Worksheet
    Dim pos As Long

    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    pos = CLng(Val(wsCache.Range("N5").value))
    pos = pos - 8
    If pos < 1 Then pos = 1

    wsCache.Range("N5").value = pos

    Application.ScreenUpdating = False
    HRS_BuildCurrentWriteBackPreview
    Application.ScreenUpdating = True

End Sub

Public Sub HRS_CurrentWriteBackNext()

    Dim wsCache As Worksheet
    Dim total As Long
    Dim pos As Long

    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    total = HRS_LastRow(wsCache, 2) - 1
    pos = CLng(Val(wsCache.Range("N5").value))
    pos = pos + 8

    If pos > total Then pos = WorksheetFunction.Max(1, total - 6)

    wsCache.Range("N5").value = pos

    Application.ScreenUpdating = False
    HRS_BuildCurrentWriteBackPreview
    Application.ScreenUpdating = True

End Sub

Public Sub HRS_PreviousWriteBackPrev()

    Dim wsCache As Worksheet
    Dim pos As Long
    Dim productCode As String
    Dim productName As String

    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    pos = CLng(Val(wsCache.Range("N6").value))
    pos = pos - 8
    If pos < 1 Then pos = 1

    wsCache.Range("N6").value = pos
    productCode = CStr(wsCache.Range("N3").value)
    productName = CStr(wsCache.Range("N4").value)

    Application.ScreenUpdating = False
    HRS_BuildPreviousWriteBackPreviewPaged productCode, productName, pos
    Application.ScreenUpdating = True

End Sub

Public Sub HRS_PreviousWriteBackNext()

    Dim wsCache As Worksheet
    Dim pos As Long
    Dim productCode As String
    Dim productName As String

    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    pos = CLng(Val(wsCache.Range("N6").value))
    If pos < 1 Then pos = 1
    pos = pos + 8

    wsCache.Range("N6").value = pos
    productCode = CStr(wsCache.Range("N3").value)
    productName = CStr(wsCache.Range("N4").value)

    Application.ScreenUpdating = False
    HRS_BuildPreviousWriteBackPreviewPaged productCode, productName, pos
    Application.ScreenUpdating = True

End Sub

Private Sub HRS_BuildPreviousWriteBackPreviewPaged( _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal pageStart As Long)

    Dim wsInput As Worksheet
    Dim wsHistory As Worksheet
    Dim lastRow As Long, r As Long, outputCol As Long
    Dim vendorName As String, foundIndex As Long
    Dim totalMatches As Long, lastShow As Long
    Dim useDateText As String, originalQty As Variant, distributionQty As Variant
    Dim isMilk As Boolean, deliveryKey As String, useKey As String
    Dim totals As Object, useDates As Object, firstCols As Object, dateSet As Object
    Dim key As Variant, cookingQty As Double

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsHistory = ThisWorkbook.Worksheets(SH_HISTORY)
    isMilk = HRS_IsMilk1L(productName)

    wsInput.Range("A26:I31").ClearContents
    wsInput.Range("A26:I31").Font.Strikethrough = False
    wsInput.Range("A26:I31").Interior.Color = RGB(242, 242, 242)
    wsInput.Range("B28:I28").NumberFormat = "General"
    wsInput.Range("B29:I29").NumberFormat = "@"

    wsInput.Range("A26").value = "納品日"
    wsInput.Range("A27").value = "使用日"
    wsInput.Range("A28").value = productName
    wsInput.Range("A29").value = "配分後"
    If isMilk Then
        wsInput.Range("A30").value = "料理用"
        wsInput.Range("A31").value = "発注書"
    Else
        wsInput.Range("A30").value = "発注書"
    End If

    wsInput.Range("A26:A31").Font.Bold = True
    wsInput.Range("A26:A31").Interior.Color = RGB(217, 225, 242)

    If isMilk Then
        Set totals = CreateObject("Scripting.Dictionary")
        Set useDates = CreateObject("Scripting.Dictionary")
        Set firstCols = CreateObject("Scripting.Dictionary")
    End If

    vendorName = Trim$(CStr(wsInput.Range("B3").value))
    lastRow = HRS_LastRow(wsHistory, 1)
    outputCol = 2

    For r = lastRow To 2 Step -1
        If CStr(wsHistory.Cells(r, 2).value) = vendorName And _
           (HRS_NormalizeCode(wsHistory.Cells(r, 3).value) = productCode Or _
            HRS_ProductNameMatches(CStr(wsHistory.Cells(r, 4).value), productName)) Then
            totalMatches = totalMatches + 1
        End If
    Next r

    If totalMatches = 0 Then
        wsInput.Range("F25").value = "0件"
    Else
        lastShow = WorksheetFunction.Min(pageStart + 7, totalMatches)
        wsInput.Range("F25").value = pageStart & "～" & lastShow & "/" & totalMatches & "件"
    End If
    wsInput.Range("F25").HorizontalAlignment = xlRight
    wsInput.Range("F25").Font.Bold = True

    For r = lastRow To 2 Step -1
        If CStr(wsHistory.Cells(r, 2).value) = vendorName And _
           (HRS_NormalizeCode(wsHistory.Cells(r, 3).value) = productCode Or _
            HRS_ProductNameMatches(CStr(wsHistory.Cells(r, 4).value), productName)) Then

            foundIndex = foundIndex + 1
            If foundIndex >= pageStart Then
                useDateText = CStr(wsHistory.Cells(r, 8).value)
                If Trim$(CStr(wsHistory.Cells(r, 9).value)) <> "" Then _
                    useDateText = useDateText & " " & CStr(wsHistory.Cells(r, 9).value)

                originalQty = wsHistory.Cells(r, 10).value
                distributionQty = wsHistory.Cells(r, 12).value
                wsInput.Cells(26, outputCol).value = wsHistory.Cells(r, 11).value
                wsInput.Cells(27, outputCol).value = useDateText
                wsInput.Cells(28, outputCol).NumberFormat = "General"
                wsInput.Cells(28, outputCol).value = originalQty
                With wsInput.Cells(29, outputCol)
                    .NumberFormat = "@"
                    If Trim$(CStr(distributionQty)) <> "" Then
                        .value = HRS_DistributionDisplayText(distributionQty)
                    Else
                        .ClearContents
                    End If
                    .NumberFormat = "@"
                End With
                wsInput.Cells(IIf(isMilk, 31, 30), outputCol).value = wsHistory.Cells(r, 14).value

                If CBool(wsHistory.Cells(r, 7).value) Then
                    wsInput.Cells(28, outputCol).Font.Strikethrough = True
                    wsInput.Cells(28, outputCol).Interior.Color = RGB(217, 217, 217)
                End If

                If isMilk Then
                    deliveryKey = HRS_FormatDateKeyPublic(wsHistory.Cells(r, 11).value)
                    If deliveryKey <> "" Then
                        If Not totals.Exists(deliveryKey) Then
                            totals.Add deliveryKey, 0#
                            firstCols.Add deliveryKey, outputCol
                            Set dateSet = CreateObject("Scripting.Dictionary")
                            useDates.Add deliveryKey, dateSet
                        End If
                        totals(deliveryKey) = CDbl(totals(deliveryKey)) + CDbl(Val(originalQty))
                        useKey = HRS_FormatDateKeyPublic(wsHistory.Cells(r, 8).value)
                        If useKey <> "" Then
                            Set dateSet = useDates(deliveryKey)
                            If Not dateSet.Exists(useKey) Then dateSet.Add useKey, True
                        End If
                    End If
                End If

                outputCol = outputCol + 1
                If outputCol > 9 Then Exit For
            End If
        End If
    Next r

    If isMilk Then
        For Each key In totals.keys
            Set dateSet = useDates(CStr(key))
            cookingQty = Fix(CDbl(totals(CStr(key))) - (9.8 * CDbl(dateSet.count)))
            If cookingQty < 0 Then cookingQty = 0
            wsInput.Cells(30, CLng(firstCols(CStr(key)))).NumberFormat = "General"
            wsInput.Cells(30, CLng(firstCols(CStr(key)))).value = cookingQty
        Next key
    End If

    wsInput.Range("A26:I31").Borders.LineStyle = xlContinuous
    wsInput.Range("A26:I31").HorizontalAlignment = xlCenter
    wsInput.Range("A26:I31").VerticalAlignment = xlCenter
    wsInput.Range("A26:I31").WrapText = True
End Sub

'=========================================================
' 書き戻しプレビュー・実書戻し
'=========================================================
Private Function HRS_DistributionDisplayText(ByVal sourceValue As Variant) As String

    Dim numericValue As Double

    If IsEmpty(sourceValue) Or Trim$(CStr(sourceValue)) = "" Then Exit Function

    'セルやDBからDate型として戻った値は、日付文字列ではなく
    'Excelのシリアル値（1、1.7など）として表示する。
    If IsDate(sourceValue) And VarType(sourceValue) = vbDate Then
        numericValue = CDbl(CDate(sourceValue))
        If numericValue = Fix(numericValue) Then
            HRS_DistributionDisplayText = CStr(CLng(numericValue))
        Else
            HRS_DistributionDisplayText = CStr(numericValue)
        End If
    ElseIf IsNumeric(sourceValue) Then
        numericValue = CDbl(sourceValue)
        If numericValue = Fix(numericValue) Then
            HRS_DistributionDisplayText = CStr(CLng(numericValue))
        Else
            HRS_DistributionDisplayText = CStr(numericValue)
        End If
    Else
        HRS_DistributionDisplayText = CStr(sourceValue)
    End If

End Function

Private Sub HRS_BuildCurrentWriteBackPreview()

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim outputCol As Long
    Dim productName As String
    Dim displayUseDate As String
    Dim pageStart As Long
    Dim currentIndex As Long
    Dim totalCount As Long
    Dim lastShow As Long
    Dim originalQty As Variant
    Dim distributionQty As Variant
    Dim cookingQty As Variant
    Dim isMilk As Boolean
    Dim aggregateMode As Boolean
    Dim cancelMark As String
    Dim deliveryValue As Variant
    Dim orderSheetValue As Variant

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    productName = CStr(wsCache.Range("N4").value)
    isMilk = HRS_IsMilk1L(productName)
    aggregateMode = HRS_IsDeliveryAggregateMode(wsCache)

    HRS_ApplyMilk1LDistribution productName

    '集約データは商品選択時または配分変更時に更新済み。

    wsInput.Range("K23:S28").ClearContents
    wsInput.Range("K23:S28").Font.Strikethrough = False
    wsInput.Range("K23:S28").Interior.Color = RGB(242, 242, 242)
    wsInput.Range("L25:S25").NumberFormat = "General"
    wsInput.Range("L26:S26").NumberFormat = "@"
    wsInput.Range("L27:S27").NumberFormat = "General"

    pageStart = CLng(Val(wsCache.Range("N5").value))
    If pageStart < 1 Then pageStart = 1

    wsInput.Range("K23").value = "納品日"
    wsInput.Range("K24").value = "使用日"
    wsInput.Range("K25").value = productName
    wsInput.Range("K26").value = "配分後"

    If isMilk Then
        wsInput.Range("K27").value = "料理用"
        wsInput.Range("K28").value = "発注書"
    Else
        wsInput.Range("K27").value = "発注書"
    End If

    wsInput.Range("K23:K28").Font.Bold = True
    wsInput.Range("K23:K28").Interior.Color = RGB(217, 225, 242)

    outputCol = 12
    currentIndex = 0
    totalCount = 0

    If aggregateMode Then
        lastRow = HRS_LastRow(wsCache, 16)
        For r = 2 To lastRow
            If Trim$(CStr(wsCache.Cells(r, 16).value)) <> "" Then
                totalCount = totalCount + 1
            End If
        Next r
    Else
        lastRow = HRS_LastRow(wsCache, 2)
        For r = 2 To lastRow
            If Trim$(CStr(wsCache.Cells(r, 2).value)) <> "" Then
                totalCount = totalCount + 1
            End If
        Next r
    End If

    If totalCount = 0 Then
        wsInput.Range("P22").value = "0件"
    Else
        lastShow = WorksheetFunction.Min(pageStart + 7, totalCount)
        wsInput.Range("P22").value = _
            pageStart & "～" & lastShow & "/" & totalCount & "件"
    End If

    wsInput.Range("P22").HorizontalAlignment = xlRight
    wsInput.Range("P22").Font.Bold = True

    For r = 2 To lastRow

        If aggregateMode Then
            If Trim$(CStr(wsCache.Cells(r, 16).value)) = "" Then GoTo ContinueLoop

            displayUseDate = CStr(wsCache.Cells(r, 16).value) & " 朝へ集約"
            originalQty = wsCache.Cells(r, 18).value
            distributionQty = wsCache.Cells(r, 20).value
            deliveryValue = wsCache.Cells(r, 19).value
            orderSheetValue = wsCache.Cells(r, 22).value
            cancelMark = CStr(wsCache.Cells(r, 15).value)
            cookingQty = ""
        Else
            If Trim$(CStr(wsCache.Cells(r, 2).value)) = "" Then GoTo ContinueLoop

            displayUseDate = CStr(wsCache.Cells(r, 2).value)

            If Trim$(CStr(wsCache.Cells(r, 3).value)) <> "" Then
                displayUseDate = _
                    displayUseDate & " " & CStr(wsCache.Cells(r, 3).value)
            End If

            originalQty = wsCache.Cells(r, 4).value
            distributionQty = wsCache.Cells(r, 6).value
            cookingQty = wsCache.Cells(r, 13).value
            deliveryValue = wsCache.Cells(r, 5).value
            orderSheetValue = wsCache.Cells(r, 9).value
            cancelMark = CStr(wsCache.Cells(r, 1).value)
        End If

        currentIndex = currentIndex + 1

        If currentIndex >= pageStart Then

            wsInput.Cells(23, outputCol).value = deliveryValue
            wsInput.Cells(24, outputCol).value = displayUseDate
            wsInput.Cells(25, outputCol).NumberFormat = "General"
            wsInput.Cells(25, outputCol).value = originalQty

            With wsInput.Cells(26, outputCol)
                .NumberFormat = "@"
                If Trim$(CStr(distributionQty)) <> "" Then
                    .value = HRS_DistributionDisplayText(distributionQty)
                Else
                    .ClearContents
                End If
                .NumberFormat = "@"
            End With

            If isMilk And Trim$(CStr(cookingQty)) <> "" Then
                wsInput.Cells(27, outputCol).NumberFormat = "General"
                wsInput.Cells(27, outputCol).value = cookingQty
            End If

            wsInput.Cells(IIf(isMilk, 28, 27), outputCol).value = _
                orderSheetValue

            If cancelMark = "■" Then
                wsInput.Cells(25, outputCol).Font.Strikethrough = True
                wsInput.Cells(25, outputCol).Interior.Color = RGB(217, 217, 217)
            End If

            If Val(distributionQty) <> 0 Then
                wsInput.Cells(26, outputCol).Font.Bold = True
            End If

            outputCol = outputCol + 1
            If outputCol > 19 Then Exit For
        End If

ContinueLoop:
    Next r

    wsInput.Range("K23:S28").Borders.LineStyle = xlContinuous
    wsInput.Range("K23:S28").HorizontalAlignment = xlCenter
    wsInput.Range("K23:S28").VerticalAlignment = xlCenter
    wsInput.Range("K23:S28").WrapText = True

End Sub

Public Sub HRS_WriteBackToOrderBook()

    Dim wsSession As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sourceBookName As String
    Dim sourceSheetName As String
    Dim cellAddress As String
    Dim targetCell As Range
    Dim writeCell As Range
    Dim notFound As Long
    Dim written As Long

    On Error GoTo ErrHandler

    HRS_BeginFast "発注書へ書き戻しています..."

    Set wsSession = ThisWorkbook.Worksheets(SH_SESSION)
    lastRow = HRS_LastRow(wsSession, 1)

    For r = 2 To lastRow

        sourceBookName = CStr(wsSession.Cells(r, 16).value)
        sourceSheetName = CStr(wsSession.Cells(r, 14).value)
        cellAddress = CStr(wsSession.Cells(r, 15).value)

        Set wb = HRS_FindOpenWorkbook(sourceBookName)

        If wb Is Nothing Then
            notFound = notFound + 1
        ElseIf Not HRS_SheetExists(sourceSheetName, wb) Then
            notFound = notFound + 1
        ElseIf cellAddress = "" Then
            notFound = notFound + 1
        Else
            Set ws = wb.Worksheets(sourceSheetName)
            Set targetCell = ws.Range(cellAddress)

            targetCell.Font.Strikethrough = _
                CBool(wsSession.Cells(r, 7).value)

            Set writeCell = targetCell.Offset(1, 0)
            writeCell.value = wsSession.Cells(r, 12).value
            writeCell.HorizontalAlignment = xlCenter
            writeCell.Font.Bold = True

            written = written + 1
        End If

        Set wb = Nothing
        Set ws = Nothing
        Set targetCell = Nothing
        Set writeCell = Nothing
    Next r

    HRS_ArchiveSession

ExitHandler:
    HRS_EndFast

    If Err.Number = 0 Then
        MsgBox "書き戻しが完了しました。" & vbCrLf & _
               "書込：" & written & "件" & vbCrLf & _
               "対象ブック・シート未確認：" & notFound & "件", _
               vbInformation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    HRS_ShowError "HRS_WriteBackToOrderBook", Err.Number, Err.Description
    Resume ExitHandler

End Sub

Private Sub HRS_ArchiveSession()

    Dim wsSession As Worksheet
    Dim wsHistory As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long

    Set wsSession = ThisWorkbook.Worksheets(SH_SESSION)
    Set wsHistory = ThisWorkbook.Worksheets(SH_HISTORY)

    lastRow = HRS_LastRow(wsSession, 1)
    If lastRow < 2 Then Exit Sub

    targetRow = HRS_LastRow(wsHistory, 1) + 1
    If targetRow < 2 Then targetRow = 2

    wsSession.Range("A2:P" & lastRow).Copy
    wsHistory.Cells(targetRow, 1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False

End Sub



Private Function HRS_GetPreviousTotalQty( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String) As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim totalValue As Double
    Dim latestStamp As Variant
    Dim currentStamp As Variant

    If Not HRS_SheetExists(SH_HISTORY, ThisWorkbook) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SH_HISTORY)
    lastRow = HRS_LastRow(ws, 1)

    For r = lastRow To 2 Step -1

        If CStr(ws.Cells(r, 2).value) = vendorName And _
           (HRS_NormalizeCode(ws.Cells(r, 3).value) = productCode Or _
            HRS_ProductNameMatches(CStr(ws.Cells(r, 4).value), productName)) Then

            currentStamp = ws.Cells(r, 1).value

            If IsEmpty(latestStamp) Then latestStamp = currentStamp

            If CStr(currentStamp) = CStr(latestStamp) Then
                totalValue = totalValue + CDbl(Val(ws.Cells(r, 10).value))
            Else
                Exit For
            End If
        End If
    Next r

    HRS_GetPreviousTotalQty = totalValue

End Function

Private Function HRS_GetPreviousStockQty( _
    ByVal productName As String) As Double

    '在庫チェックの現在値を前回在庫として表示する。
    HRS_GetPreviousStockQty = HRS_GetStockQty(productName)

End Function

Private Function HRS_GetPreviousOrderQty( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String) As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    If Not HRS_SheetExists(SH_HISTORY, ThisWorkbook) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SH_HISTORY)
    lastRow = HRS_LastRow(ws, 1)

    '履歴の最後の行が最新。発注数は商品内で同じため最初の一致を返す。
    For r = lastRow To 2 Step -1

        If CStr(ws.Cells(r, 2).value) = vendorName And _
           (HRS_NormalizeCode(ws.Cells(r, 3).value) = productCode Or _
            HRS_ProductNameMatches( _
                CStr(ws.Cells(r, 4).value), productName)) Then

            HRS_GetPreviousOrderQty = _
                CDbl(Val(ws.Cells(r, 5).value))

            Exit Function
        End If
    Next r

End Function

Private Function HRS_GetStockQty( _
    ByVal productName As String) As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    If Not HRS_SheetExists(SH_STOCK, ThisWorkbook) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SH_STOCK)
    lastRow = HRS_LastRow(ws, 2)

    For r = 1 To lastRow

        If HRS_ProductNameMatches( _
            CStr(ws.Cells(r, 2).value), productName) Then

            HRS_GetStockQty = _
                CDbl(Val(ws.Cells(r, 3).value))

            Exit Function
        End If
    Next r

End Function

Private Sub HRS_SyncStockProducts()

    Dim wsProduct As Worksheet
    Dim wsStock As Worksheet
    Dim lastProductRow As Long
    Dim productRow As Long
    Dim stockRow As Long
    Dim productName As String
    Dim existsValue As Boolean
    Dim r As Long

    Set wsProduct = ThisWorkbook.Worksheets(SH_PRODUCT_CACHE)
    Set wsStock = ThisWorkbook.Worksheets(SH_STOCK)

    lastProductRow = HRS_LastRow(wsProduct, 2)

    For productRow = 2 To lastProductRow

        productName = Trim$(CStr( _
            wsProduct.Cells(productRow, 2).value))

        If productName <> "" Then

            existsValue = False

            For r = 2 To HRS_LastRow(wsStock, 2)
                If HRS_ProductNameMatches( _
                    CStr(wsStock.Cells(r, 2).value), productName) Then

                    existsValue = True
                    Exit For
                End If
            Next r

            If Not existsValue Then
                stockRow = HRS_LastRow(wsStock, 2) + 1
                If stockRow < 2 Then stockRow = 2
                wsStock.Cells(stockRow, 2).value = productName
            End If
        End If
    Next productRow

End Sub

Private Sub HRS_BuildPreviousWriteBackPreview( _
    ByVal productCode As String, _
    ByVal productName As String)

    Dim wsCache As Worksheet

    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)
    wsCache.Range("N6").value = 1

    HRS_BuildPreviousWriteBackPreviewPaged _
        productCode, productName, 1

End Sub

'=========================================================
' 旬間発注時・単位
'=========================================================
Private Sub HRS_ShowRule(ByVal productName As String)

    Dim wsInput As Worksheet
    Dim wsRule As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim resultText As String

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)

    If Not HRS_SheetExists(SH_RULE, ThisWorkbook) Then
        wsInput.Range("A23").value = "旬間発注時シートがありません。"
        HRS_FormatRuleDisplay wsInput
        Exit Sub
    End If

    Set wsRule = ThisWorkbook.Worksheets(SH_RULE)
    lastRow = HRS_LastRow(wsRule, 2)

    For r = 1 To lastRow
        If HRS_ProductNameMatches(CStr(wsRule.Cells(r, 2).value), productName) Then
            resultText = Trim$(CStr(wsRule.Cells(r, 6).value))
            Exit For
        End If
    Next r

    If resultText = "" Then resultText = "登録内容はありません。"

    wsInput.Range("A23").value = resultText
    HRS_FormatRuleDisplay wsInput

End Sub


Private Sub HRS_FormatRuleDisplay(ByVal wsInput As Worksheet)

    With wsInput.Range("A23:I23")
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .ShrinkToFit = True
    End With

End Sub

Private Function HRS_GetUnitNote(ByVal productCode As String, _
                                 ByVal productName As String) As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SH_UNIT)
    lastRow = HRS_LastRow(ws, 1)

    For r = 2 To lastRow
        If HRS_NormalizeCode(ws.Cells(r, 1).value) = productCode Or _
           Trim$(CStr(ws.Cells(r, 2).value)) = productName Then
            HRS_GetUnitNote = Trim$(CStr(ws.Cells(r, 5).value))
            Exit Function
        End If
    Next r

End Function

'=========================================================
' マスタ
'=========================================================
Private Sub HRS_RegisterVendor(ByVal vendorCode As String, _
                               ByVal vendorName As String)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim targetRow As Long

    If vendorName = "" Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SH_VENDOR)
    lastRow = HRS_LastRow(ws, 1)

    For r = 2 To lastRow
        If vendorCode <> "" Then
            If HRS_NormalizeCode(ws.Cells(r, 1).value) = _
               HRS_NormalizeCode(vendorCode) Then
                targetRow = r
                Exit For
            End If
        ElseIf HRS_NormalizeText(CStr(ws.Cells(r, 2).value)) = _
               HRS_NormalizeText(vendorName) Then
            targetRow = r
            Exit For
        End If
    Next r

    If targetRow = 0 Then
        targetRow = lastRow + 1
        If targetRow < 2 Then targetRow = 2
        ws.Cells(targetRow, 3).value = Now
    End If

    ws.Cells(targetRow, 1).NumberFormat = "@"
    ws.Cells(targetRow, 1).value = vendorCode
    ws.Cells(targetRow, 2).value = vendorName
    ws.Cells(targetRow, 4).value = Now

    If Not IsNumeric(ws.Cells(targetRow, 5).value) Or _
       Val(ws.Cells(targetRow, 5).value) <= 0 Then
        ws.Cells(targetRow, 5).value = HRS_NextVendorDisplayOrder(ws)
    End If

    If Trim$(CStr(ws.Cells(targetRow, 6).value)) = "" Then
        ws.Cells(targetRow, 6).value = "ON"
    End If

End Sub


Private Function HRS_NextVendorDisplayOrder(ByVal ws As Worksheet) As Long

    Dim lastRow As Long
    Dim r As Long
    Dim maxOrder As Long

    lastRow = Application.Max(HRS_LastRow(ws, 2), HRS_LastRow(ws, 5))

    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, 5).value) Then
            If CLng(Val(ws.Cells(r, 5).value)) > maxOrder Then
                maxOrder = CLng(Val(ws.Cells(r, 5).value))
            End If
        End If
    Next r

    HRS_NextVendorDisplayOrder = maxOrder + 1
    If HRS_NextVendorDisplayOrder < 1 Then HRS_NextVendorDisplayOrder = 1

End Function

Private Sub HRS_SetupZeroUsageDisplaySheet()

    Dim ws As Worksheet
    Dim validationRange As Range
    Dim listSeparator As String
    Dim listFormula As String
    Dim validationError As Long

    Set ws = ThisWorkbook.Worksheets(SH_ZERO_USAGE)

    ws.Columns("A").NumberFormat = "@"
    ws.Columns("B").NumberFormat = "@"
    ws.Columns("D").NumberFormat = "@"

    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 32
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("F:G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 30

    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(255, 192, 0)

    '全104万行へ入力規則を設定すると、環境によって1004になるため
    '実運用で十分な5000行までに限定する。
    Set validationRange = ws.Range("A2:A5000")

    On Error Resume Next
    validationRange.Validation.Delete
    Err.Clear

    listSeparator = Application.International(xlListSeparator)
    listFormula = "ON" & listSeparator & "OFF"

    validationRange.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:=listFormula

    validationError = Err.Number

    '環境依存で区切り文字が受け付けられない場合の予備処理
    If validationError <> 0 Then
        Err.Clear
        validationRange.Validation.Delete
        validationRange.Validation.Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="ON,OFF"
        validationError = Err.Number
    End If
    On Error GoTo 0

    '入力規則を作成できた場合だけ詳細設定を行う。
    If validationError = 0 Then
        With validationRange.Validation
            .IgnoreBlank = False
            .InCellDropdown = True
            .InputTitle = "数量なし表示"
            .InputMessage = "ON または OFF を選択してください。"
            .ErrorTitle = "入力エラー"
            .ErrorMessage = "ON または OFF をプルダウンから選択してください。"
            .ShowInput = True
            .ShowError = True
        End With
    End If

    If Trim$(CStr(ws.Range("J1").value)) = "" Then
        ws.Range("J1").value = "設定方法"
        ws.Range("J2").value = "数量が空欄でも使用日・区分を表示する商品は、A列をONにしてください。"
        ws.Range("J3").value = "発注書の再読込後から反映されます。既存の並び順と設定は維持されます。"
        ws.Range("J4").value = "新商品は読込時に末尾へOFFで追加されます。"
        ws.Range("J5").value = "A列はセル右側の▼からON/OFFを選択できます。"
        ws.Range("J1").Font.Bold = True
        ws.Columns("J").ColumnWidth = 75
        ws.Range("J1:J5").WrapText = True
    End If

End Sub

Private Function HRS_ZeroUsageLookupKey( _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal vendorCode As String, _
    ByVal vendorName As String) As String

    If HRS_NormalizeCode(productCode) <> "" Then
        HRS_ZeroUsageLookupKey = _
            "C|" & HRS_NormalizeCode(productCode) & "|" & _
            HRS_NormalizeCode(vendorCode)
    Else
        HRS_ZeroUsageLookupKey = _
            "N|" & UCase$(Trim$(productName)) & "|" & _
            UCase$(Trim$(vendorName))
    End If

End Function

Private Sub HRS_LoadZeroUsageSettingsCache()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim keyText As String
    Dim settingValue As String

    Set mZeroUsageSettings = CreateObject("Scripting.Dictionary")
    mZeroUsageSettings.CompareMode = vbTextCompare
    mZeroUsageSettingsLoaded = True

    If Not HRS_SheetExists(SH_ZERO_USAGE, ThisWorkbook) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SH_ZERO_USAGE)
    lastRow = HRS_LastRow(ws, 2)

    For r = 2 To lastRow
        keyText = HRS_ZeroUsageLookupKey( _
            CStr(ws.Cells(r, 2).value), _
            CStr(ws.Cells(r, 3).value), _
            CStr(ws.Cells(r, 4).value), _
            CStr(ws.Cells(r, 5).value))

        If keyText <> "" Then
            settingValue = CStr(ws.Cells(r, 1).value)

            If HRS_IsOnSettingValue(settingValue) Then
                mZeroUsageSettings(keyText) = True
            ElseIf Not mZeroUsageSettings.Exists(keyText) Then
                mZeroUsageSettings(keyText) = False
            End If
        End If
    Next r

End Sub

Private Sub HRS_InvalidateZeroUsageSettingsCache()

    mZeroUsageSettingsLoaded = False
    Set mZeroUsageSettings = Nothing

End Sub

Private Sub HRS_RegisterZeroUsageProduct( _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal vendorCode As String, _
    ByVal vendorName As String)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim targetRow As Long
    Dim firstMatchRow As Long
    Dim onMatchRow As Long
    Dim settingValue As String

    Set ws = ThisWorkbook.Worksheets(SH_ZERO_USAGE)
    lastRow = HRS_LastRow(ws, 2)

    '同じ商品の重複行がある場合、ONになっている行を最優先する。
    For r = 2 To lastRow
        If HRS_ZeroUsageProductMatches(ws, r, productCode, productName, _
                                       vendorCode, vendorName) Then

            If firstMatchRow = 0 Then firstMatchRow = r

            settingValue = UCase$(Trim$(CStr(ws.Cells(r, 1).value)))

            If HRS_IsOnSettingValue(settingValue) Then
                onMatchRow = r
                Exit For
            End If
        End If
    Next r

    If onMatchRow > 0 Then
        targetRow = onMatchRow
    ElseIf firstMatchRow > 0 Then
        targetRow = firstMatchRow
    Else
        targetRow = lastRow + 1
        If targetRow < 2 Then targetRow = 2

        '本当に新しい商品のみOFFで追加する。
        ws.Cells(targetRow, 1).NumberFormat = "@"
        ws.Cells(targetRow, 1).value = "OFF"
        ws.Cells(targetRow, 6).value = Now
        HRS_InvalidateZeroUsageSettingsCache
    End If

    '既存のA列設定は一切上書きしない。
    ws.Cells(targetRow, 2).NumberFormat = "@"
    ws.Cells(targetRow, 2).value = productCode
    ws.Cells(targetRow, 3).value = productName
    ws.Cells(targetRow, 4).NumberFormat = "@"
    ws.Cells(targetRow, 4).value = vendorCode
    ws.Cells(targetRow, 5).value = vendorName
    ws.Cells(targetRow, 7).value = Now

End Sub

Private Function HRS_IsOnSettingValue(ByVal settingValue As String) As Boolean

    settingValue = UCase$(Trim$(settingValue))

    HRS_IsOnSettingValue = _
        (settingValue = "ON" Or _
         settingValue = "TRUE" Or _
         settingValue = "1" Or _
         settingValue = "○" Or _
         settingValue = "〇" Or _
         settingValue = "有効")

End Function

Private Function HRS_IsZeroUsageDisplayEnabled( _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal vendorCode As String, _
    ByVal vendorName As String) As Boolean

    Dim keyText As String
    Dim fallbackKey As String

    If Not mZeroUsageSettingsLoaded Then
        HRS_LoadZeroUsageSettingsCache
    End If

    If mZeroUsageSettings Is Nothing Then Exit Function

    keyText = HRS_ZeroUsageLookupKey( _
        productCode, productName, vendorCode, vendorName)

    If mZeroUsageSettings.Exists(keyText) Then
        HRS_IsZeroUsageDisplayEnabled = CBool(mZeroUsageSettings(keyText))
        Exit Function
    End If

    If HRS_NormalizeCode(productCode) <> "" Then
        fallbackKey = "C|" & HRS_NormalizeCode(productCode) & "|"

        If mZeroUsageSettings.Exists(fallbackKey) Then
            HRS_IsZeroUsageDisplayEnabled = _
                CBool(mZeroUsageSettings(fallbackKey))
        End If
    End If

End Function

Private Function HRS_ZeroUsageProductMatches( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal vendorCode As String, _
    ByVal vendorName As String) As Boolean

    Dim savedProductCode As String
    Dim savedVendorCode As String
    Dim savedProductName As String
    Dim savedVendorName As String

    savedProductCode = HRS_NormalizeCode(ws.Cells(rowNumber, 2).value)
    savedVendorCode = HRS_NormalizeCode(ws.Cells(rowNumber, 4).value)
    savedProductName = Trim$(CStr(ws.Cells(rowNumber, 3).value))
    savedVendorName = Trim$(CStr(ws.Cells(rowNumber, 5).value))

    '商品番号がある場合は商品番号＋業者を優先する。
    If HRS_NormalizeCode(productCode) <> "" And savedProductCode <> "" Then
        If savedProductCode = HRS_NormalizeCode(productCode) Then
            If savedVendorCode <> "" And HRS_NormalizeCode(vendorCode) <> "" Then
                HRS_ZeroUsageProductMatches = _
                    (savedVendorCode = HRS_NormalizeCode(vendorCode))
            Else
                HRS_ZeroUsageProductMatches = _
                    HRS_ProductNameMatches(savedVendorName, vendorName)
            End If
        End If
        Exit Function
    End If

    '商品番号がない場合は商品名＋業者名で判定する。
    HRS_ZeroUsageProductMatches = _
        HRS_ProductNameMatches(savedProductName, productName) And _
        HRS_ProductNameMatches(savedVendorName, vendorName)

End Function

Private Sub HRS_RegisterProduct(ByVal productCode As String, _
                                ByVal productName As String, _
                                ByVal specText As String, _
                                ByVal unitText As String, _
                                ByVal vendorCode As String, _
                                ByVal vendorName As String)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim targetRow As Long

    Set ws = ThisWorkbook.Worksheets(SH_PRODUCT)
    lastRow = HRS_LastRow(ws, 1)

    For r = 2 To lastRow
        If HRS_NormalizeCode(ws.Cells(r, 1).value) = _
           HRS_NormalizeCode(productCode) Then
            targetRow = r
            Exit For
        End If
    Next r

    If targetRow = 0 Then
        targetRow = lastRow + 1
        If targetRow < 2 Then targetRow = 2
        ws.Cells(targetRow, 8).value = Now
    End If

    ws.Cells(targetRow, 1).NumberFormat = "@"
    ws.Cells(targetRow, 1).value = productCode
    ws.Cells(targetRow, 2).value = productName
    ws.Cells(targetRow, 3).value = specText
    ws.Cells(targetRow, 4).value = unitText
    ws.Cells(targetRow, 5).NumberFormat = "@"
    ws.Cells(targetRow, 5).value = vendorCode
    ws.Cells(targetRow, 6).value = vendorName
    ws.Cells(targetRow, 7).value = IIf(HRS_IsSpecial(productName), "TRUE", "FALSE")
    ws.Cells(targetRow, 9).value = Now

End Sub

'=========================================================
' ヘッダー・帳票解析
'=========================================================
Private Function HRS_ExtractVendorFromHeader(ByVal ws As Worksheet, _
                                             ByRef vendorCode As String) As String

    Dim headerText As String
    Dim regEx As Object
    Dim matches As Object
    Dim matchItem As Object
    Dim candidateName As String

    headerText = HRS_CleanHeader(ws.PageSetup.LeftHeader) & " | " & _
                 HRS_CleanHeader(ws.PageSetup.CenterHeader) & " | " & _
                 HRS_CleanHeader(ws.PageSetup.RightHeader)

    Set regEx = CreateObject("VBScript.RegExp")

    With regEx
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "([0-9]{2,10})[ 　]*[:：][ 　]*([^|""\r\n]+)"
    End With

    Set matches = regEx.Execute(headerText)

    For Each matchItem In matches
        vendorCode = HRS_OnlyDigits(CStr(matchItem.SubMatches(0)))
        candidateName = HRS_CleanVendorName(CStr(matchItem.SubMatches(1)))

        If HRS_IsValidVendor(candidateName) Then
            HRS_ExtractVendorFromHeader = candidateName
            Exit Function
        End If
    Next matchItem

End Function

Private Function HRS_CleanHeader(ByVal sourceText As String) As String

    Dim resultText As String
    Dim startPos As Long
    Dim endPos As Long
    Dim i As Long

    resultText = sourceText

    Do
        startPos = InStr(resultText, "&""")
        If startPos = 0 Then Exit Do

        endPos = InStr(startPos + 2, resultText, """")
        If endPos = 0 Then Exit Do

        resultText = Left$(resultText, startPos - 1) & _
                     Mid$(resultText, endPos + 1)
    Loop

    resultText = Replace(resultText, "&L", "")
    resultText = Replace(resultText, "&C", "")
    resultText = Replace(resultText, "&R", "")
    resultText = Replace(resultText, "&P", "")
    resultText = Replace(resultText, "&N", "")
    resultText = Replace(resultText, "&D", "")
    resultText = Replace(resultText, "&T", "")
    resultText = Replace(resultText, "&F", "")
    resultText = Replace(resultText, "&A", "")
    resultText = Replace(resultText, "&B", "")
    resultText = Replace(resultText, "&I", "")
    resultText = Replace(resultText, "&U", "")

    For i = 1 To 72
        resultText = Replace(resultText, "&" & CStr(i), "")
    Next i

    HRS_CleanHeader = Trim$(resultText)

End Function

Private Function HRS_CleanVendorName(ByVal sourceText As String) As String

    Dim resultText As String
    Dim cutPos As Long

    resultText = HRS_CleanText(sourceText)

    cutPos = InStr(resultText, "即日")
    If cutPos > 1 Then resultText = Trim$(Left$(resultText, cutPos - 1))

    cutPos = InStr(resultText, "発注書")
    If cutPos > 1 Then resultText = Trim$(Left$(resultText, cutPos - 1))

    HRS_CleanVendorName = resultText

End Function

Private Function HRS_IsValidVendor(ByVal sourceText As String) As Boolean

    Dim s As String

    s = HRS_NormalizeText(sourceText)

    If s = "" Then Exit Function
    If s = "食品名" Then Exit Function
    If s = "商品名" Then Exit Function
    If s = "業者名" Then Exit Function
    If InStr(s, "納品日") > 0 Then Exit Function
    If InStr(s, "使用日") > 0 Then Exit Function
    If InStr(s, "発注書") > 0 Then Exit Function

    HRS_IsValidVendor = True

End Function

Private Function HRS_CountProductRows(ByVal ws As Worksheet) As Long

    Dim lastRow As Long
    Dim r As Long

    lastRow = HRS_LastRow(ws, 0)

    For r = 7 To lastRow
        If HRS_LooksLikeProductCode(HRS_ProductCode(ws.Cells(r, 1))) Then
            HRS_CountProductRows = HRS_CountProductRows + 1
        End If
    Next r

End Function

Private Function HRS_FindSubtotalColumn(ByVal ws As Worksheet) As Long

    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim textValue As String

    lastCol = HRS_LastColumn(ws)

    For c = 1 To lastCol
        For r = 1 To 12
            textValue = HRS_NormalizeText(HRS_CellText(ws.Cells(r, c)))

            If InStr(textValue, "小計") > 0 Then
                HRS_FindSubtotalColumn = c
                Exit Function
            End If
        Next r
    Next c

End Function

Private Function HRS_IsUsageColumn(ByVal ws As Worksheet, _
                                   ByVal colNumber As Long) As Boolean

    Dim mealText As String
    Dim headerText As String
    Dim r As Long

    For r = 1 To 8
        headerText = headerText & _
            HRS_NormalizeText(HRS_CellText(ws.Cells(r, colNumber)))
    Next r

    If InStr(headerText, "小計") > 0 Then Exit Function
    If InStr(headerText, "合計") > 0 Then Exit Function
    If InStr(headerText, "購入") > 0 Then Exit Function
    If InStr(headerText, "単価") > 0 Then Exit Function
    If InStr(headerText, "金額") > 0 Then Exit Function

    mealText = HRS_NormalizeText( _
        HRS_CellText(ws.Cells(HEADER_MEAL_ROW, colNumber)))

    If mealText = "朝" Or mealText = "昼" Or mealText = "夕" Then
        HRS_IsUsageColumn = True
    End If

End Function

Private Function HRS_HeaderLeftFill(ByVal ws As Worksheet, _
                                    ByVal rowNumber As Long, _
                                    ByVal colNumber As Long) As String

    Dim c As Long
    Dim sourceValue As Variant
    Dim displayText As String

    For c = colNumber To FIRST_USAGE_COL Step -1

        sourceValue = ws.Cells(rowNumber, c).value
        displayText = HRS_CleanText(HRS_CellText(ws.Cells(rowNumber, c)))

        If IsDate(sourceValue) Then
            HRS_HeaderLeftFill = Format$(CDate(sourceValue), "yyyy/m/d")
            Exit Function
        End If

        If displayText <> "" Then
            HRS_HeaderLeftFill = displayText
            Exit Function
        End If

    Next c

End Function

Private Function HRS_ProductCode(ByVal targetCell As Range) As String

    Dim resultText As String

    resultText = HRS_CellText(targetCell)

    If resultText = "" Or InStr(resultText, "#") > 0 Then
        If Not IsError(targetCell.value) Then
            resultText = Trim$(CStr(targetCell.value))
        End If
    End If

    resultText = Replace(resultText, " ", "")
    resultText = Replace(resultText, "　", "")

    HRS_ProductCode = resultText

End Function

Private Function HRS_LooksLikeProductCode(ByVal sourceText As String) As Boolean

    Dim i As Long
    Dim ch As String

    If Len(sourceText) < 4 Or Len(sourceText) > 14 Then Exit Function

    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    HRS_LooksLikeProductCode = True

End Function

Private Function HRS_IsValidQty(ByVal sourceValue As Variant) As Boolean

    If IsError(sourceValue) Then Exit Function
    If IsEmpty(sourceValue) Then Exit Function
    If Not IsNumeric(sourceValue) Then Exit Function
    If CDbl(sourceValue) = 0 Then Exit Function

    HRS_IsValidQty = True

End Function

Private Function HRS_IsSpecial(ByVal productName As String) As Boolean

    Dim s As String

    s = HRS_NormalizeText(productName)

    HRS_IsSpecial = _
        (InStr(s, "牛乳200") > 0) Or _
        (InStr(s, "牛乳1l") > 0) Or _
        (InStr(s, "牛乳1000") > 0) Or _
        (InStr(s, "ヨーグルト") > 0)

End Function


Public Sub HRS_ApplyLayout_Ver1356()
    HRS_ApplyLayout_Ver1358
End Sub

Public Sub HRS_ApplyLayout_Ver1358()
    HRS_ApplyLayout_Ver1359
End Sub

Public Sub HRS_ApplyLayout_Ver1360()
    HRS_SetupVendorMasterSettings
    HRS_ApplyLayout_Ver1359
    HRS_RefreshVendorList
End Sub

Public Sub HRS_ApplyLayout_Ver1359()

    Dim ws As Worksheet
    Dim productCode As String
    Dim productName As String

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    Application.ScreenUpdating = False

    '上段の入力作業領域は動かさず、下段だけを確実に表示する。
    ws.Rows("22:31").Hidden = False
    '8・9行目を含む明細行の高さを均一にする。
    ws.Rows("8:20").RowHeight = 34
    ws.Rows("22:24").RowHeight = 34
    ws.Rows("25").RowHeight = 28
    ws.Rows("26:31").RowHeight = 30

    '今回の書き戻しイメージ（右側・上寄せ）
    With ws.Range("K22:S22")
        .Interior.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("K22").value = "今回の書き戻しイメージ"
    With ws.Range("K23:S28")
        .Interior.Color = RGB(242, 242, 242)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With

    '前回の書き戻しイメージ（左下・必要時に少しスクロールして確認）
    With ws.Range("A25:I25")
        .Interior.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("A25").value = "前回の書き戻しイメージ"
    With ws.Range("A26:I31")
        .Interior.Color = RGB(242, 242, 242)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With

    '数値欄が日付書式へ変わらないよう標準を固定する。
    ws.Range("K25:Q27").NumberFormat = "General"
    ws.Range("B28:H30").NumberFormat = "General"

    HRS_EnsureVendorDropdown

    productCode = CStr(ThisWorkbook.Worksheets(SH_PREVIEW_CACHE).Range("N3").value)
    productName = CStr(ThisWorkbook.Worksheets(SH_PREVIEW_CACHE).Range("N4").value)
    If Trim$(productName) <> "" Then
        HRS_BuildCurrentWriteBackPreview
        HRS_BuildPreviousWriteBackPreview productCode, productName
        HRS_ShowSelectedSpecialOrder productName
    End If

    Application.ScreenUpdating = True
    MsgBox "画面配置を更新しました。" & vbCrLf & _
           "使用日プレビューと今回の書き戻しは上側のまま、" & vbCrLf & _
           "前回の書き戻しは左下で必要時に確認できます。", _
           vbInformation, "発注まるめシステム"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "画面配置の更新中にエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, _
           vbCritical, "発注まるめシステム"
End Sub

'=========================================================
' 業者選択：常時表示のフォームドロップダウン
'=========================================================
Public Sub HRS_EnsureVendorDropdown()

    Const XL_FORM_DROPDOWN As Long = 2

    Dim ws As Worksheet
    Dim shp As Shape
    Dim anchor As Range
    Dim lastVendorRow As Long
    Dim currentVendor As String
    Dim i As Long
    Dim selectedIndex As Long
    Dim visibleLines As Long

    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    Set anchor = ws.Range("B3").MergeArea

    On Error Resume Next
    ws.Shapes("ddVendorAlways").Delete
    ws.Shapes("btnVendorAlways").Delete
    On Error GoTo 0

    lastVendorRow = ws.Cells(ws.Rows.count, "AA").End(xlUp).Row
    If lastVendorRow < 1 Or Trim$(CStr(ws.Range("AA1").value)) = "" Then Exit Sub

    On Error Resume Next
    anchor.Validation.Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddFormControl( _
        XL_FORM_DROPDOWN, _
        anchor.Left, _
        anchor.Top, _
        anchor.Width, _
        anchor.Height)

    With shp
        .Name = "ddVendorAlways"
        .OnAction = "HRS_VendorDropdownChanged"
        .Placement = xlMoveAndSize
        .Visible = True
        .Locked = True
    End With

    visibleLines = lastVendorRow
    If visibleLines > 15 Then visibleLines = 15
    If visibleLines < 2 Then visibleLines = 2

    With shp.ControlFormat
        .ListFillRange = "'" & ws.Name & "'!$AA$1:$AA$" & CStr(lastVendorRow)
        .LinkedCell = "'" & ws.Name & "'!$AC$1"
        .DropDownLines = visibleLines
    End With

    currentVendor = Trim$(CStr(ws.Range("B3").value))
    selectedIndex = 0

    For i = 1 To lastVendorRow
        If StrComp(Trim$(CStr(ws.Cells(i, "AA").value)), currentVendor, vbTextCompare) = 0 Then
            selectedIndex = i
            Exit For
        End If
    Next i

    If selectedIndex = 0 Then selectedIndex = 1

    ws.Range("AC1").value = selectedIndex
    shp.ControlFormat.ListIndex = selectedIndex
    ws.Range("B3").value = CStr(ws.Cells(selectedIndex, "AA").value)

End Sub

Public Sub HRS_VendorDropdownChanged()

    Dim ws As Worksheet
    Dim selectedIndex As Long
    Dim vendorName As String

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    selectedIndex = CLng(Val(ws.Range("AC1").value))
    If selectedIndex < 1 Then Exit Sub

    vendorName = Trim$(CStr(ws.Cells(selectedIndex, "AA").value))
    If vendorName = "" Then Exit Sub

    Application.EnableEvents = False
    ws.Range("B3").value = vendorName
    Application.EnableEvents = True

    HRS_OnVendorChanged
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    MsgBox "業者の切り替え中にエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, _
           vbExclamation, APP_NAME

End Sub

Public Sub HRS_OpenVendorDropdown()
    HRS_EnsureVendorDropdown
End Sub

Public Function HRS_FormatDateKeyPublic(ByVal sourceValue As Variant) As String
    If IsDate(sourceValue) Then
        HRS_FormatDateKeyPublic = Format$(CDate(sourceValue), "yyyy/m/d")
    Else
        HRS_FormatDateKeyPublic = Trim$(CStr(sourceValue))
    End If
End Function

'=========================================================
' 共通
'=========================================================
Private Sub HRS_BeginFast(Optional ByVal statusText As String = "")

    If mFastDepth = 0 Then
        mOldCalc = Application.Calculation
        mOldEvents = Application.EnableEvents
        mOldScreen = Application.ScreenUpdating
        mOldAlerts = Application.DisplayAlerts

        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        If statusText <> "" Then Application.StatusBar = statusText
    End If

    mFastDepth = mFastDepth + 1

End Sub

Private Sub HRS_EndFast()

    If mFastDepth > 0 Then mFastDepth = mFastDepth - 1

    If mFastDepth = 0 Then
        Application.Calculation = mOldCalc
        Application.EnableEvents = mOldEvents
        Application.ScreenUpdating = mOldScreen
        Application.DisplayAlerts = mOldAlerts
        Application.StatusBar = False
    End If

End Sub

Public Sub HRS_ResetExcelState()

    mFastDepth = 0

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.CutCopyMode = False

    MsgBox "Excelの状態を通常へ戻しました。", vbInformation, APP_NAME

End Sub

Private Function HRS_GetOrCreateSheet(ByVal sheetName As String, _
                                      ByVal veryHidden As Boolean) As Worksheet

    Dim ws As Worksheet

    If HRS_SheetExists(sheetName, ThisWorkbook) Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
    Else
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
    End If

    If veryHidden Then
        ws.Visible = xlSheetVeryHidden
    Else
        ws.Visible = xlSheetVisible
    End If

    Set HRS_GetOrCreateSheet = ws

End Function

Private Function HRS_SheetExists(ByVal sheetName As String, _
                                 ByVal wb As Workbook) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    HRS_SheetExists = Not ws Is Nothing
    On Error GoTo 0

End Function

Private Function HRS_LastRow(ByVal ws As Worksheet, _
                             ByVal colNumber As Long) As Long

    Dim lastCell As Range

    If colNumber > 0 Then
        HRS_LastRow = ws.Cells(ws.Rows.count, colNumber).End(xlUp).Row
        If HRS_LastRow < 1 Then HRS_LastRow = 1
        Exit Function
    End If

    Set lastCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious)

    If lastCell Is Nothing Then
        HRS_LastRow = 1
    Else
        HRS_LastRow = lastCell.Row
    End If

End Function

Private Function HRS_LastColumn(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    Set lastCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious)

    If lastCell Is Nothing Then
        HRS_LastColumn = 1
    Else
        HRS_LastColumn = lastCell.Column
    End If

End Function

Private Function HRS_CellText(ByVal targetCell As Range) As String

    Dim resultText As String

    On Error Resume Next
    resultText = Trim$(CStr(targetCell.Text))
    On Error GoTo 0

    If resultText = "" Or InStr(resultText, "#") > 0 Then
        If Not IsError(targetCell.value) Then
            resultText = Trim$(CStr(targetCell.value))
        End If
    End If

    HRS_CellText = resultText

End Function

Private Function HRS_CleanText(ByVal sourceText As String) As String

    Dim resultText As String

    resultText = Trim$(sourceText)
    resultText = Replace(resultText, vbCr, " ")
    resultText = Replace(resultText, vbLf, " ")
    resultText = Replace(resultText, "　", " ")

    Do While InStr(resultText, "  ") > 0
        resultText = Replace(resultText, "  ", " ")
    Loop

    HRS_CleanText = Trim$(resultText)

End Function


Private Function HRS_ProductNameMatches( _
    ByVal listName As String, _
    ByVal productName As String) As Boolean

    Dim leftName As String
    Dim rightName As String

    leftName = HRS_NormalizeText(HRS_GetManagedProductName(listName))
    rightName = HRS_NormalizeText(HRS_GetManagedProductName(productName))

    If leftName = "" Or rightName = "" Then Exit Function

    If leftName = rightName Then
        HRS_ProductNameMatches = True
        Exit Function
    End If

    '規格が商品名へ併記されている場合にも対応する。
    If Len(leftName) >= 3 And _
       InStr(rightName, leftName) > 0 Then

        HRS_ProductNameMatches = True
        Exit Function
    End If

    If Len(rightName) >= 3 And _
       InStr(leftName, rightName) > 0 Then

        HRS_ProductNameMatches = True
    End If

End Function


Private Function HRS_FindProductCacheRow( _
    ByVal productCode As String, _
    ByVal productName As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SH_PRODUCT_CACHE)
    lastRow = HRS_LastRow(ws, 2)

    For r = 2 To lastRow
        If HRS_NormalizeCode(ws.Cells(r, 1).value) = HRS_NormalizeCode(productCode) And _
           HRS_ProductNameMatches(CStr(ws.Cells(r, 2).value), productName) Then
            HRS_FindProductCacheRow = r
            Exit Function
        End If
    Next r

End Function

Private Sub HRS_ReapplyDistributionGray(ByVal displayRow As Long)

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim distributionValue As Variant
    Dim productName As String
    Dim deleteStatus As Long

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    distributionValue = wsInput.Cells(displayRow, "Q").value
    productName = CStr(wsCache.Range("N4").value)
    deleteStatus = HRS_GetDeleteItemStatus(productName)

    '削除項目は配分後を消してもグレーを維持する。
    If deleteStatus > 0 Then
        wsInput.Range("L" & displayRow & ":S" & displayRow).Interior.Color = RGB(217, 217, 217)
        Exit Sub
    End If

    If Trim$(CStr(distributionValue)) <> "" And IsNumeric(distributionValue) Then
        If CDbl(distributionValue) <> 0 Then
            wsInput.Range("L" & displayRow & ":S" & displayRow).Interior.Color = RGB(217, 217, 217)
        Else
            wsInput.Range("L" & displayRow & ":S" & displayRow).Interior.Color = RGB(226, 239, 218)
            wsInput.Cells(displayRow, "Q").Interior.Color = RGB(255, 242, 204)
        End If
    Else
        wsInput.Range("L" & displayRow & ":S" & displayRow).Interior.Color = RGB(226, 239, 218)
        wsInput.Cells(displayRow, "Q").Interior.Color = RGB(255, 242, 204)
    End If

End Sub

Private Sub HRS_MarkDisplayedProductConfirmed( _
    ByVal productCode As String, _
    ByVal productName As String)

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim r As Long
    Dim cacheRow As Long
    Dim displayedRowFound As Boolean

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PRODUCT_CACHE)

    cacheRow = HRS_FindProductCacheRow(productCode, productName)
    If cacheRow > 0 Then wsCache.Cells(cacheRow, 8).value = "■"

    For r = ITEM_TOP To ITEM_BOTTOM
        If HRS_NormalizeCode(wsInput.Cells(r, "A").value) = HRS_NormalizeCode(productCode) And _
           HRS_ProductNameMatches(CStr(wsInput.Cells(r, "B").value), productName) Then
            wsInput.Cells(r, "I").value = "■"
            displayedRowFound = True
            HRS_UpdateDisplayedProductFill r
            If HRS_GetDeleteItemStatus(productName) = 1 Then
                wsInput.Cells(r, "B").Font.Color = RGB(255, 0, 0)
                wsInput.Cells(r, "B").Font.Bold = True
            End If
            Exit For
        End If
    Next r

    ' 商品コードの表記差などで一致しなかった場合は、商品名だけで再検索する。
    If Not displayedRowFound Then
        For r = ITEM_TOP To ITEM_BOTTOM
            If HRS_ProductNameMatches(CStr(wsInput.Cells(r, "B").value), productName) Then
                wsInput.Cells(r, "I").value = "■"
                HRS_UpdateDisplayedProductFill r
                Exit For
            End If
        Next r
    End If

End Sub

Private Sub HRS_UpdateDisplayedProductFill(ByVal displayRow As Long)

    Dim wsInput As Worksheet
    Dim wsCache As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim hasDistribution As Boolean
    Dim distributionValue As Variant
    Dim productName As String
    Dim deleteStatus As Long

    Set wsInput = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsCache = ThisWorkbook.Worksheets(SH_PREVIEW_CACHE)

    productName = Trim$(CStr(wsInput.Cells(displayRow, "B").value))
    deleteStatus = HRS_GetDeleteItemStatus(productName)

    lastRow = HRS_LastRow(wsCache, 2)
    For r = 2 To lastRow
        distributionValue = wsCache.Cells(r, 6).value
        If Trim$(CStr(distributionValue)) <> "" And IsNumeric(distributionValue) Then
            If CDbl(distributionValue) <> 0 Then
                hasDistribution = True
                Exit For
            End If
        End If
    Next r

    wsInput.Range("A" & displayRow & ":I" & displayRow).Interior.ColorIndex = xlNone
    wsInput.Cells(displayRow, "F").Interior.Color = RGB(255, 242, 204)

    '削除項目シートの商品は、配分後の入力を消しても元の表示を維持する。
    If deleteStatus > 0 Then
        wsInput.Range("A" & displayRow & ":I" & displayRow).Interior.Color = RGB(217, 217, 217)

        If deleteStatus = 1 Then
            wsInput.Cells(displayRow, "B").Font.Color = RGB(255, 0, 0)
            wsInput.Cells(displayRow, "B").Font.Bold = True
        End If
    ElseIf hasDistribution Then
        wsInput.Range("A" & displayRow & ":I" & displayRow).Interior.Color = RGB(217, 217, 217)
    End If

End Sub

Private Function HRS_NormalizeText(ByVal sourceText As String) As String

    Dim resultText As String

    resultText = Trim$(sourceText)
    resultText = Replace(resultText, "　", "")
    resultText = Replace(resultText, " ", "")
    resultText = Replace(resultText, vbCr, "")
    resultText = Replace(resultText, vbLf, "")
    resultText = Replace(resultText, "Ｌ", "L")
    resultText = Replace(resultText, "ｌ", "L")

    HRS_NormalizeText = LCase$(resultText)

End Function

Private Function HRS_NormalizeCode(ByVal sourceValue As Variant) As String

    Dim resultText As String

    If IsError(sourceValue) Or IsEmpty(sourceValue) Then Exit Function

    resultText = Trim$(CStr(sourceValue))
    resultText = Replace(resultText, " ", "")
    resultText = Replace(resultText, "　", "")

    HRS_NormalizeCode = resultText

End Function

Private Function HRS_OnlyDigits(ByVal sourceText As String) As String

    Dim i As Long
    Dim ch As String
    Dim resultText As String

    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        If ch >= "0" And ch <= "9" Then resultText = resultText & ch
    Next i

    HRS_OnlyDigits = resultText

End Function

Private Sub HRS_ClearBelowHeader(ByVal ws As Worksheet, _
                                 ByVal firstCol As String, _
                                 ByVal lastCol As String)

    Dim lastRow As Long

    lastRow = HRS_LastRow(ws, 0)

    If lastRow >= 2 Then
        ws.Range(firstCol & "2:" & lastCol & lastRow).ClearContents
    End If

End Sub

Private Sub HRS_ClearSelectedPanels()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SH_INPUT)

    ws.Range("K8:S20").ClearContents
    ws.Range("K8:S20").Font.Strikethrough = False
    ws.Range("K8:S20").Font.Size = 11
    ws.Range("K8:S20").Interior.Color = RGB(226, 239, 218)
    ws.Range("L8:L20").Font.Size = 11
    ws.Range("Q8:Q20").Interior.Color = RGB(255, 242, 204)
    ws.Range("Q8:Q20").NumberFormat = "General"

    ws.Range("A23").value = _
        "商品を選択すると内容を表示します。"

    ws.Range("A26:H30").ClearContents
    ws.Range("J23:Q27").ClearContents
    ws.Range("J28:Q52").Clear
    HRS_UpdateMilkCookingDetail ""
    ws.Range("K5").value = "□ 全取消"

End Sub

Private Function HRS_FindOpenWorkbook(ByVal bookName As String) As Workbook

    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, bookName, vbTextCompare) = 0 Then
            Set HRS_FindOpenWorkbook = wb
            Exit Function
        End If
    Next wb

End Function

Private Sub HRS_WriteLog(ByVal processName As String, _
                         ByVal detailText As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    If Not HRS_SheetExists(SH_LOG, ThisWorkbook) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    nextRow = HRS_LastRow(ws, 1) + 1

    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = processName
    ws.Cells(nextRow, 3).value = detailText
    ws.Cells(nextRow, 4).value = Environ$("Username")

End Sub

Private Sub HRS_ShowError(ByVal processName As String, _
                          ByVal errorNumber As Long, _
                          ByVal errorDescription As String)

    HRS_WriteLog "ERROR:" & processName, _
                 CStr(errorNumber) & " : " & errorDescription

    MsgBox processName & "でエラーが発生しました。" & vbCrLf & _
           errorNumber & " : " & errorDescription, _
           vbCritical, APP_NAME

End Sub
Public Sub HRS_ApplyLayout_Ver240()

    On Error GoTo ErrHandler

    HRS_CreateDatabaseSheets
    HRS_ApplyLayout_Ver1360
    HRS_CreateDatabaseSheets

    MsgBox "Ver2.4.0の設定を反映しました。" & vbCrLf & vbCrLf & _
           "・読込設定シート" & vbCrLf & _
           "・発注原票履歴DB" & vbCrLf & _
           "・履歴保存確認" & vbCrLf & _
           "・読込前の最終削除確認", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    HRS_ShowError "HRS_ApplyLayout_Ver240", Err.Number, Err.Description

End Sub



'=========================================================
' Ver2.2からの互換用
'=========================================================
Public Sub HRS_ApplyLayout_Ver230()

    HRS_ApplyLayout_Ver240

End Sub

Public Sub HRS_ApplyLayout_Ver220()

    HRS_ApplyLayout_Ver240

End Sub
