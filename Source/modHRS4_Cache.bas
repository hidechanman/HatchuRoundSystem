Attribute VB_Name = "modHRS4_Cache"
Option Explicit

'=========================================================
' 発注まるめシステム Ver4.0 Part1
' 高速完成キャッシュ作成
'
' 発注原票DBの列
' A 登録日時 / B 取込元ブック / C シート名
' D 業者コード / E 業者名 / F 商品番号 / G 商品名
' H 単位 / I 納品日 / J 使用日 / K 区分
' L 使用数量 / M 元行 / N 元列 / O セル番地
' P 特殊商品 / Q 備考
'=========================================================

Private mNormalIndex As Object
Private mAggregateIndex As Object
Private mWriteBackIndex As Object
Private mSessionIndex As Object
Private mProductIndex As Object
Private mCacheLoaded As Boolean

Public Sub HRS4_InstallPart1()

    Dim startTime As Double
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ErrHandler

    startTime = Timer
    HRS4_BeginFast "Ver4.0 Part1を初期化しています..."

    HRS4_CreateCacheSheets
    HRS4_BuildAllCaches

ExitHandler:
    HRS4_EndFast

    If errorNumber = 0 Then
        MsgBox "Ver4.0 Part1の初期化が完了しました。" & vbCrLf & _
               "通常・集約・書戻し・セッションの" & vbCrLf & _
               "完成キャッシュを作成しました。" & vbCrLf & vbCrLf & _
               "処理時間: " & _
               Format$(HRS4_ElapsedSeconds(startTime), "0.000") & "秒", _
               vbInformation, "発注まるめシステム"
    Else
        MsgBox "Ver4.0 Part1の初期化に失敗しました。" & vbCrLf & _
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

Public Sub HRS4_BuildAllCaches()

    Dim totalStart As Double
    Dim wsRaw As Worksheet
    Dim lastRow As Long
    Dim rawData As Variant

    Dim normalData As Variant
    Dim aggregateData As Variant
    Dim writeBackData As Variant
    Dim sessionData As Variant
    Dim productData As Variant

    Dim normalCount As Long
    Dim aggregateCount As Long
    Dim writeBackCount As Long
    Dim sessionCount As Long
    Dim productCount As Long

    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ErrHandler

    totalStart = Timer
    HRS4_BeginFast "発注原票DBを配列へ読み込んでいます..."

    HRS4_CreateCacheSheets

    If Not HRS4_SheetExists(HRS4_SH_RAW) Then
        Err.Raise vbObjectError + 4001, _
                  "HRS4_BuildAllCaches", _
                  "発注原票DBが見つかりません。"
    End If

    Set wsRaw = ThisWorkbook.Worksheets(HRS4_SH_RAW)
    lastRow = HRS4_LastRow(wsRaw, 1)

    If lastRow < 2 Then
        HRS4_ClearAllCaches
        HRS4_ClearMemoryCaches
        HRS4_WritePerformanceLog _
            "V4全キャッシュ作成", _
            HRS4_ElapsedSeconds(totalStart), _
            0, _
            "発注原票DBにデータなし"
        GoTo ExitHandler
    End If

    rawData = wsRaw.Range("A2:Q" & lastRow).Value2

    Application.StatusBar = "通常表示キャッシュを作成しています..."
    normalData = HRS4_CreateNormalCache(rawData, normalCount)

    Application.StatusBar = "集約表示キャッシュを作成しています..."
    aggregateData = HRS4_CreateAggregateCache(rawData, aggregateCount)

    Application.StatusBar = "書戻しキャッシュを作成しています..."
    writeBackData = HRS4_CreateWriteBackCache(rawData, writeBackCount)

    Application.StatusBar = "セッションキャッシュを作成しています..."
    sessionData = HRS4_CreateSessionCache( _
        normalData, normalCount, sessionCount)

    Application.StatusBar = "商品索引を作成しています..."
    productData = HRS4_CreateProductIndex(rawData, productCount)

    Application.StatusBar = "キャッシュを一括保存しています..."
    HRS4_WriteCacheArray _
        HRS4_SH_NORMAL, normalData, normalCount, 20
    HRS4_WriteCacheArray _
        HRS4_SH_AGGREGATE, aggregateData, aggregateCount, 20
    HRS4_WriteCacheArray _
        HRS4_SH_WRITEBACK, writeBackData, writeBackCount, 15
    HRS4_WriteCacheArray _
        HRS4_SH_SESSION, sessionData, sessionCount, 18
    HRS4_WriteCacheArray _
        HRS4_SH_INDEX, productData, productCount, 9

    Application.StatusBar = "メモリ索引を構築しています..."
    HRS4_LoadMemoryCaches

    HRS4_WritePerformanceLog _
        "V4全キャッシュ作成", _
        HRS4_ElapsedSeconds(totalStart), _
        lastRow - 1, _
        "通常=" & normalCount & _
        " 集約=" & aggregateCount & _
        " 商品=" & productCount

ExitHandler:
    HRS4_EndFast

    If errorNumber <> 0 Then
        MsgBox "キャッシュ作成中にエラーが発生しました。" & vbCrLf & _
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

Private Sub HRS4_CreateCacheSheets()

    Dim ws As Worksheet

    Set ws = HRS4_GetOrCreateSheet(HRS4_SH_NORMAL, True)
    HRS4_SetHeaders ws, Array( _
        "キャッシュキー", "業者コード", "業者名", _
        "商品番号", "商品名", "単位", "納品日", _
        "使用日", "区分", "使用数量", "取消状態", _
        "配分後", "発注書", "シート名", "元行", _
        "元列", "セル番地", "特殊商品", "備考", "原票DB行")

    Set ws = HRS4_GetOrCreateSheet(HRS4_SH_AGGREGATE, True)
    HRS4_SetHeaders ws, Array( _
        "キャッシュキー", "業者コード", "業者名", _
        "商品番号", "商品名", "単位", "納品日", _
        "最初使用日", "最後使用日", "表示区分", _
        "使用数量合計", "取消状態", "配分後", _
        "書戻対象原票DB行", "明細原票DB行一覧", _
        "朝行あり", "発注書", "特殊商品", "備考", "集約順")

    Set ws = HRS4_GetOrCreateSheet(HRS4_SH_WRITEBACK, True)
    HRS4_SetHeaders ws, Array( _
        "キャッシュキー", "業者名", "商品番号", "商品名", _
        "納品日", "使用日", "区分", "原票DB行", _
        "取込元ブック", "シート名", "元行", "元列", _
        "セル番地", "書戻値", "変更済")

    Set ws = HRS4_GetOrCreateSheet(HRS4_SH_SESSION, True)
    HRS4_SetHeaders ws, Array( _
        "キャッシュキー", "業者名", "商品番号", "商品名", _
        "納品日", "使用日", "区分", "使用数量", _
        "取消状態", "配分後", "単位", "発注書", _
        "シート名", "セル番地", "原票DB行", _
        "変更済", "更新日時", "表示モード")

    Set ws = HRS4_GetOrCreateSheet(HRS4_SH_INDEX, True)
    HRS4_SetHeaders ws, Array( _
        "商品キー", "業者コード", "業者名", "商品番号", _
        "商品名", "単位", "通常件数", "集約件数", "表示順")

    Set ws = HRS4_GetOrCreateSheet(HRS4_SH_PERFORMANCE, False)
    HRS4_SetHeaders ws, Array( _
        "日時", "処理", "秒数", "件数", "備考")

End Sub

Private Function HRS4_CreateNormalCache( _
    ByVal rawData As Variant, _
    ByRef outputCount As Long) As Variant

    Dim rowCount As Long
    Dim resultData() As Variant
    Dim r As Long
    Dim keyText As String
    Dim productIdentity As String
    Dim deliveryKey As String
    Dim useKey As String

    rowCount = UBound(rawData, 1)
    ReDim resultData(1 To rowCount, 1 To 20)

    For r = 1 To rowCount
        If Trim$(CStr(rawData(r, 7))) <> "" Then

            outputCount = outputCount + 1

            productIdentity = HRS4_ProductIdentity( _
                rawData(r, 6), rawData(r, 7))
            deliveryKey = HRS4_DateKey(rawData(r, 9))
            useKey = HRS4_DateKey(rawData(r, 10))

            keyText = HRS4_MakeDetailKey( _
                rawData(r, 5), productIdentity, deliveryKey, _
                useKey, rawData(r, 11), r + 1)

            resultData(outputCount, 1) = keyText
            resultData(outputCount, 2) = rawData(r, 4)
            resultData(outputCount, 3) = rawData(r, 5)
            resultData(outputCount, 4) = _
                HRS4_NormalizeCode(rawData(r, 6))
            resultData(outputCount, 5) = rawData(r, 7)
            resultData(outputCount, 6) = rawData(r, 8)
            resultData(outputCount, 7) = rawData(r, 9)
            resultData(outputCount, 8) = rawData(r, 10)
            resultData(outputCount, 9) = rawData(r, 11)
            resultData(outputCount, 10) = _
                HRS4_NumberOrZero(rawData(r, 12))
            resultData(outputCount, 11) = _
                HRS4_DefaultCancelState(rawData(r, 12))
            resultData(outputCount, 12) = Empty
            resultData(outputCount, 13) = rawData(r, 2)
            resultData(outputCount, 14) = rawData(r, 3)
            resultData(outputCount, 15) = rawData(r, 13)
            resultData(outputCount, 16) = rawData(r, 14)
            resultData(outputCount, 17) = rawData(r, 15)
            resultData(outputCount, 18) = rawData(r, 16)
            resultData(outputCount, 19) = rawData(r, 17)
            resultData(outputCount, 20) = r + 1
        End If
    Next r

    HRS4_CreateNormalCache = resultData

End Function

Private Function HRS4_CreateAggregateCache( _
    ByVal rawData As Variant, _
    ByRef outputCount As Long) As Variant

    Dim rowCount As Long
    Dim dict As Object
    Dim groupNumber As Long
    Dim r As Long
    Dim groupIndex As Long
    Dim keyText As String
    Dim productIdentity As String
    Dim deliveryKey As String

    Dim vendorCodes() As Variant
    Dim vendorNames() As Variant
    Dim productCodes() As Variant
    Dim productNames() As Variant
    Dim units() As Variant
    Dim deliveryDates() As Variant
    Dim firstUseDates() As Variant
    Dim lastUseDates() As Variant
    Dim usageTotals() As Double
    Dim cancelStates() As String
    Dim targetRawRows() As Long
    Dim rawRowLists() As String
    Dim morningFlags() As Boolean
    Dim orderBooks() As Variant
    Dim specialFlags() As Variant
    Dim notes() As Variant
    Dim resultData() As Variant

    rowCount = UBound(rawData, 1)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    ReDim vendorCodes(1 To rowCount)
    ReDim vendorNames(1 To rowCount)
    ReDim productCodes(1 To rowCount)
    ReDim productNames(1 To rowCount)
    ReDim units(1 To rowCount)
    ReDim deliveryDates(1 To rowCount)
    ReDim firstUseDates(1 To rowCount)
    ReDim lastUseDates(1 To rowCount)
    ReDim usageTotals(1 To rowCount)
    ReDim cancelStates(1 To rowCount)
    ReDim targetRawRows(1 To rowCount)
    ReDim rawRowLists(1 To rowCount)
    ReDim morningFlags(1 To rowCount)
    ReDim orderBooks(1 To rowCount)
    ReDim specialFlags(1 To rowCount)
    ReDim notes(1 To rowCount)

    For r = 1 To rowCount
        If Trim$(CStr(rawData(r, 7))) <> "" Then

            productIdentity = HRS4_ProductIdentity( _
                rawData(r, 6), rawData(r, 7))
            deliveryKey = HRS4_DateKey(rawData(r, 9))

            If deliveryKey = "" Then
                deliveryKey = "BLANK|" & CStr(r + 1)
            End If

            keyText = UCase$(Trim$(CStr(rawData(r, 5)))) & "|" & _
                      productIdentity & "|" & deliveryKey

            If Not dict.Exists(keyText) Then
                groupNumber = groupNumber + 1
                groupIndex = groupNumber
                dict.Add keyText, groupIndex

                vendorCodes(groupIndex) = rawData(r, 4)
                vendorNames(groupIndex) = rawData(r, 5)
                productCodes(groupIndex) = _
                    HRS4_NormalizeCode(rawData(r, 6))
                productNames(groupIndex) = rawData(r, 7)
                units(groupIndex) = rawData(r, 8)
                deliveryDates(groupIndex) = rawData(r, 9)
                firstUseDates(groupIndex) = rawData(r, 10)
                lastUseDates(groupIndex) = rawData(r, 10)
                targetRawRows(groupIndex) = r + 1
                rawRowLists(groupIndex) = CStr(r + 1)
                morningFlags(groupIndex) = _
                    HRS4_IsMorning(rawData(r, 11))
                orderBooks(groupIndex) = rawData(r, 2)
                specialFlags(groupIndex) = rawData(r, 16)
                notes(groupIndex) = rawData(r, 17)
                cancelStates(groupIndex) = "○"
            Else
                groupIndex = CLng(dict(keyText))
                lastUseDates(groupIndex) = rawData(r, 10)
                rawRowLists(groupIndex) = _
                    rawRowLists(groupIndex) & "," & CStr(r + 1)

                If HRS4_IsMorning(rawData(r, 11)) Then
                    If Not morningFlags(groupIndex) Then
                        morningFlags(groupIndex) = True
                        targetRawRows(groupIndex) = r + 1
                    End If
                End If
            End If

            usageTotals(groupIndex) = usageTotals(groupIndex) + _
                HRS4_NumberOrZero(rawData(r, 12))

            If HRS4_DefaultCancelState(rawData(r, 12)) = "●" Then
                cancelStates(groupIndex) = "●"
            End If
        End If
    Next r

    outputCount = groupNumber

    If groupNumber = 0 Then
        ReDim resultData(1 To 1, 1 To 20)
        HRS4_CreateAggregateCache = resultData
        Exit Function
    End If

    ReDim resultData(1 To groupNumber, 1 To 20)

    For groupIndex = 1 To groupNumber
        keyText = UCase$(Trim$(CStr(vendorNames(groupIndex)))) & "|" & _
                  HRS4_ProductIdentity( _
                      productCodes(groupIndex), _
                      productNames(groupIndex)) & "|" & _
                  HRS4_DateKey(deliveryDates(groupIndex))

        resultData(groupIndex, 1) = keyText
        resultData(groupIndex, 2) = vendorCodes(groupIndex)
        resultData(groupIndex, 3) = vendorNames(groupIndex)
        resultData(groupIndex, 4) = productCodes(groupIndex)
        resultData(groupIndex, 5) = productNames(groupIndex)
        resultData(groupIndex, 6) = units(groupIndex)
        resultData(groupIndex, 7) = deliveryDates(groupIndex)
        resultData(groupIndex, 8) = firstUseDates(groupIndex)
        resultData(groupIndex, 9) = lastUseDates(groupIndex)
        resultData(groupIndex, 10) = "朝へ集約"
        resultData(groupIndex, 11) = usageTotals(groupIndex)
        resultData(groupIndex, 12) = cancelStates(groupIndex)
        resultData(groupIndex, 13) = Empty
        resultData(groupIndex, 14) = targetRawRows(groupIndex)
        resultData(groupIndex, 15) = rawRowLists(groupIndex)
        resultData(groupIndex, 16) = _
            IIf(morningFlags(groupIndex), "TRUE", "FALSE")
        resultData(groupIndex, 17) = orderBooks(groupIndex)
        resultData(groupIndex, 18) = specialFlags(groupIndex)
        resultData(groupIndex, 19) = notes(groupIndex)
        resultData(groupIndex, 20) = groupIndex
    Next groupIndex

    HRS4_CreateAggregateCache = resultData

End Function

Private Function HRS4_CreateWriteBackCache( _
    ByVal rawData As Variant, _
    ByRef outputCount As Long) As Variant

    Dim rowCount As Long
    Dim resultData() As Variant
    Dim r As Long
    Dim productIdentity As String
    Dim deliveryKey As String
    Dim useKey As String
    Dim keyText As String

    rowCount = UBound(rawData, 1)
    ReDim resultData(1 To rowCount, 1 To 15)

    For r = 1 To rowCount
        If Trim$(CStr(rawData(r, 7))) <> "" Then

            outputCount = outputCount + 1

            productIdentity = HRS4_ProductIdentity( _
                rawData(r, 6), rawData(r, 7))
            deliveryKey = HRS4_DateKey(rawData(r, 9))
            useKey = HRS4_DateKey(rawData(r, 10))

            keyText = HRS4_MakeDetailKey( _
                rawData(r, 5), productIdentity, deliveryKey, _
                useKey, rawData(r, 11), r + 1)

            resultData(outputCount, 1) = keyText
            resultData(outputCount, 2) = rawData(r, 5)
            resultData(outputCount, 3) = _
                HRS4_NormalizeCode(rawData(r, 6))
            resultData(outputCount, 4) = rawData(r, 7)
            resultData(outputCount, 5) = rawData(r, 9)
            resultData(outputCount, 6) = rawData(r, 10)
            resultData(outputCount, 7) = rawData(r, 11)
            resultData(outputCount, 8) = r + 1
            resultData(outputCount, 9) = rawData(r, 2)
            resultData(outputCount, 10) = rawData(r, 3)
            resultData(outputCount, 11) = rawData(r, 13)
            resultData(outputCount, 12) = rawData(r, 14)
            resultData(outputCount, 13) = rawData(r, 15)
            resultData(outputCount, 14) = Empty
            resultData(outputCount, 15) = "FALSE"
        End If
    Next r

    HRS4_CreateWriteBackCache = resultData

End Function

Private Function HRS4_CreateSessionCache( _
    ByVal normalData As Variant, _
    ByVal normalCount As Long, _
    ByRef outputCount As Long) As Variant

    Dim resultData() As Variant
    Dim r As Long

    If normalCount <= 0 Then
        ReDim resultData(1 To 1, 1 To 18)
        HRS4_CreateSessionCache = resultData
        Exit Function
    End If

    ReDim resultData(1 To normalCount, 1 To 18)

    For r = 1 To normalCount
        outputCount = outputCount + 1

        resultData(r, 1) = normalData(r, 1)
        resultData(r, 2) = normalData(r, 3)
        resultData(r, 3) = normalData(r, 4)
        resultData(r, 4) = normalData(r, 5)
        resultData(r, 5) = normalData(r, 7)
        resultData(r, 6) = normalData(r, 8)
        resultData(r, 7) = normalData(r, 9)
        resultData(r, 8) = normalData(r, 10)
        resultData(r, 9) = normalData(r, 11)
        resultData(r, 10) = normalData(r, 12)
        resultData(r, 11) = normalData(r, 6)
        resultData(r, 12) = normalData(r, 13)
        resultData(r, 13) = normalData(r, 14)
        resultData(r, 14) = normalData(r, 17)
        resultData(r, 15) = normalData(r, 20)
        resultData(r, 16) = "FALSE"
        resultData(r, 17) = Empty
        resultData(r, 18) = "通常"
    Next r

    HRS4_CreateSessionCache = resultData

End Function

Private Function HRS4_CreateProductIndex( _
    ByVal rawData As Variant, _
    ByRef outputCount As Long) As Variant

    Dim rowCount As Long
    Dim dict As Object
    Dim r As Long
    Dim indexNumber As Long
    Dim productKey As String
    Dim productIdentity As String
    Dim resultData() As Variant
    Dim keys As Variant
    Dim itemData As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    rowCount = UBound(rawData, 1)

    For r = 1 To rowCount
        If Trim$(CStr(rawData(r, 7))) <> "" Then

            productIdentity = HRS4_ProductIdentity( _
                rawData(r, 6), rawData(r, 7))
            productKey = UCase$(Trim$(CStr(rawData(r, 5)))) & "|" & _
                         productIdentity

            If Not dict.Exists(productKey) Then
                itemData = Array( _
                    rawData(r, 4), _
                    rawData(r, 5), _
                    HRS4_NormalizeCode(rawData(r, 6)), _
                    rawData(r, 7), _
                    rawData(r, 8), _
                    1, _
                    0)
                dict.Add productKey, itemData
            Else
                itemData = dict(productKey)
                itemData(5) = CLng(itemData(5)) + 1
                dict(productKey) = itemData
            End If
        End If
    Next r

    outputCount = dict.count

    If outputCount = 0 Then
        ReDim resultData(1 To 1, 1 To 9)
        HRS4_CreateProductIndex = resultData
        Exit Function
    End If

    ReDim resultData(1 To outputCount, 1 To 9)
    keys = dict.keys

    For indexNumber = 0 To dict.count - 1
        productKey = CStr(keys(indexNumber))
        itemData = dict(productKey)

        resultData(indexNumber + 1, 1) = productKey
        resultData(indexNumber + 1, 2) = itemData(0)
        resultData(indexNumber + 1, 3) = itemData(1)
        resultData(indexNumber + 1, 4) = itemData(2)
        resultData(indexNumber + 1, 5) = itemData(3)
        resultData(indexNumber + 1, 6) = itemData(4)
        resultData(indexNumber + 1, 7) = itemData(5)
        resultData(indexNumber + 1, 8) = 0
        resultData(indexNumber + 1, 9) = indexNumber + 1
    Next indexNumber

    HRS4_CreateProductIndex = resultData

End Function

Private Sub HRS4_WriteCacheArray( _
    ByVal sheetName As String, _
    ByVal cacheData As Variant, _
    ByVal recordCount As Long, _
    ByVal columnCount As Long)

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(sheetName)

    lastRow = HRS4_LastRow(ws, 1)

    If lastRow >= 2 Then
        ws.Range( _
            ws.Cells(2, 1), _
            ws.Cells(lastRow, columnCount)).ClearContents
    End If

    If recordCount > 0 Then
        ws.Cells(2, 1).Resize( _
            recordCount, columnCount).Value2 = cacheData
    End If

    ws.Columns(1).NumberFormat = "@"
    ws.Visible = xlSheetVeryHidden

End Sub

Public Sub HRS4_LoadMemoryCaches()

    Dim startTime As Double

    startTime = Timer

    Set mNormalIndex = HRS4_BuildSheetIndex( _
        HRS4_SH_NORMAL, 1)
    Set mAggregateIndex = HRS4_BuildSheetIndex( _
        HRS4_SH_AGGREGATE, 1)
    Set mWriteBackIndex = HRS4_BuildSheetIndex( _
        HRS4_SH_WRITEBACK, 1)
    Set mSessionIndex = HRS4_BuildSheetIndex( _
        HRS4_SH_SESSION, 1)
    Set mProductIndex = HRS4_BuildSheetIndex( _
        HRS4_SH_INDEX, 1)

    mCacheLoaded = True

    HRS4_WritePerformanceLog _
        "V4メモリ索引読込", _
        HRS4_ElapsedSeconds(startTime), _
        mProductIndex.count, _
        "Dictionary構築完了"

End Sub

Private Function HRS4_BuildSheetIndex( _
    ByVal sheetName As String, _
    ByVal keyColumn As Long) As Object

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim data As Variant
    Dim dict As Object
    Dim rowList As Collection
    Dim r As Long
    Dim keyText As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = HRS4_LastRow(ws, 1)
    lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    If lastRow < 2 Then
        Set HRS4_BuildSheetIndex = dict
        Exit Function
    End If

    data = ws.Range( _
        ws.Cells(2, 1), _
        ws.Cells(lastRow, lastColumn)).Value2

    For r = 1 To UBound(data, 1)
        keyText = CStr(data(r, keyColumn))

        If keyText <> "" Then
            If Not dict.Exists(keyText) Then
                Set rowList = New Collection
                dict.Add keyText, rowList
            End If

            Set rowList = dict(keyText)
            rowList.Add r + 1
        End If
    Next r

    Set HRS4_BuildSheetIndex = dict

End Function

Public Function HRS4_GetProductRows( _
    ByVal vendorName As String, _
    ByVal productCode As String, _
    ByVal productName As String, _
    ByVal aggregateMode As Boolean) As Variant

    Dim startTime As Double
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim data As Variant
    Dim resultData() As Variant
    Dim resultCount As Long
    Dim r As Long
    Dim identityText As String
    Dim rowIdentity As String
    Dim sourceColumnCount As Long

    startTime = Timer

    If Not mCacheLoaded Then
        HRS4_LoadMemoryCaches
    End If

    identityText = HRS4_ProductIdentity(productCode, productName)

    If aggregateMode Then
        Set ws = ThisWorkbook.Worksheets(HRS4_SH_AGGREGATE)
        sourceColumnCount = 20
    Else
        Set ws = ThisWorkbook.Worksheets(HRS4_SH_NORMAL)
        sourceColumnCount = 20
    End If

    lastRow = HRS4_LastRow(ws, 1)

    If lastRow < 2 Then
        HRS4_GetProductRows = Empty
        Exit Function
    End If

    data = ws.Cells(2, 1).Resize( _
        lastRow - 1, sourceColumnCount).Value2
    ReDim resultData(1 To UBound(data, 1), 1 To sourceColumnCount)

    For r = 1 To UBound(data, 1)
        rowIdentity = HRS4_ProductIdentity(data(r, 4), data(r, 5))

        If StrComp(Trim$(CStr(data(r, 3))), _
                   Trim$(vendorName), vbTextCompare) = 0 Then
            If StrComp(rowIdentity, identityText, vbTextCompare) = 0 Then

                resultCount = resultCount + 1
                HRS4_CopyArrayRow _
                    data, r, resultData, resultCount, sourceColumnCount
            End If
        End If
    Next r

    If resultCount = 0 Then
        HRS4_GetProductRows = Empty
    Else
        HRS4_GetProductRows = HRS4_Trim2DArray( _
            resultData, resultCount, sourceColumnCount)
    End If

    HRS4_WritePerformanceLog _
        "V4商品キャッシュ取得", _
        HRS4_ElapsedSeconds(startTime), _
        resultCount, _
        IIf(aggregateMode, "集約", "通常")

End Function

Public Sub HRS4_TestCurrentProductLookup()

    Dim vendorName As String
    Dim productCode As String
    Dim productName As String
    Dim aggregateMode As Boolean
    Dim resultData As Variant
    Dim recordCount As Long
    Dim startTime As Double

    vendorName = InputBox("業者名を入力してください。")
    If vendorName = "" Then Exit Sub

    productCode = InputBox( _
        "商品番号を入力してください。" & vbCrLf & _
        "番号がない場合は空欄で構いません。")

    productName = InputBox("商品名を入力してください。")
    If productCode = "" And productName = "" Then Exit Sub

    aggregateMode = _
        (MsgBox("集約表示をテストしますか？", _
                vbYesNo + vbQuestion) = vbYes)

    startTime = Timer
    resultData = HRS4_GetProductRows( _
        vendorName, productCode, productName, aggregateMode)

    If IsEmpty(resultData) Then
        recordCount = 0
    Else
        recordCount = UBound(resultData, 1)
    End If

    MsgBox "取得が完了しました。" & vbCrLf & _
           "件数: " & CStr(recordCount) & vbCrLf & _
           "時間: " & _
           Format$(HRS4_ElapsedSeconds(startTime), "0.000") & "秒", _
           vbInformation, "Ver4.0 キャッシュ取得テスト"

End Sub

Public Sub HRS4_ShowCacheSheets()

    ThisWorkbook.Worksheets(HRS4_SH_NORMAL).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(HRS4_SH_AGGREGATE).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(HRS4_SH_WRITEBACK).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(HRS4_SH_SESSION).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(HRS4_SH_INDEX).Visible = xlSheetVisible

    MsgBox "Ver4.0のキャッシュシートを表示しました。", _
           vbInformation, "発注まるめシステム"

End Sub

Public Sub HRS4_HideCacheSheets()

    ThisWorkbook.Worksheets(HRS4_SH_NORMAL).Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets(HRS4_SH_AGGREGATE).Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets(HRS4_SH_WRITEBACK).Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets(HRS4_SH_SESSION).Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets(HRS4_SH_INDEX).Visible = xlSheetVeryHidden

    MsgBox "Ver4.0のキャッシュシートを非表示にしました。", _
           vbInformation, "発注まるめシステム"

End Sub

Public Sub HRS4_ClearMemoryCaches()

    Set mNormalIndex = Nothing
    Set mAggregateIndex = Nothing
    Set mWriteBackIndex = Nothing
    Set mSessionIndex = Nothing
    Set mProductIndex = Nothing

    mCacheLoaded = False

End Sub

Private Sub HRS4_ClearAllCaches()

    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long

    sheetNames = Array( _
        HRS4_SH_NORMAL, HRS4_SH_AGGREGATE, _
        HRS4_SH_WRITEBACK, HRS4_SH_SESSION, HRS4_SH_INDEX)

    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        lastRow = HRS4_LastRow(ws, 1)
        lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

        If lastRow >= 2 Then
            ws.Range( _
                ws.Cells(2, 1), _
                ws.Cells(lastRow, lastColumn)).ClearContents
        End If
    Next sheetName

End Sub

Private Sub HRS4_SetHeaders( _
    ByVal ws As Worksheet, _
    ByVal headers As Variant)

    Dim columnCount As Long

    columnCount = UBound(headers) - LBound(headers) + 1
    ws.Cells(1, 1).Resize(1, columnCount).value = headers
    ws.Rows(1).Font.Bold = True

End Sub

Private Function HRS4_SheetExists( _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    HRS4_SheetExists = Not ws Is Nothing

End Function

Private Function HRS4_MakeDetailKey( _
    ByVal vendorName As Variant, _
    ByVal productIdentity As String, _
    ByVal deliveryKey As String, _
    ByVal useKey As String, _
    ByVal mealText As Variant, _
    ByVal rawRow As Long) As String

    HRS4_MakeDetailKey = _
        UCase$(Trim$(CStr(vendorName))) & "|" & _
        productIdentity & "|" & _
        deliveryKey & "|" & _
        useKey & "|" & _
        UCase$(Trim$(CStr(mealText))) & "|" & _
        CStr(rawRow)

End Function

Private Function HRS4_NumberOrZero(ByVal value As Variant) As Double

    If IsError(value) Or IsEmpty(value) Then Exit Function

    If Trim$(CStr(value)) = "" Then Exit Function

    If IsNumeric(value) Then
        HRS4_NumberOrZero = CDbl(value)
    End If

End Function

Private Function HRS4_DefaultCancelState( _
    ByVal usageValue As Variant) As String

    If IsNumeric(usageValue) Then
        If Abs(CDbl(usageValue) - 0.1) < 0.0000001 Then
            HRS4_DefaultCancelState = "●"
            Exit Function
        End If
    End If

    HRS4_DefaultCancelState = "○"

End Function

Private Sub HRS4_CopyArrayRow( _
    ByVal sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByRef targetData As Variant, _
    ByVal targetRow As Long, _
    ByVal columnCount As Long)

    Dim c As Long

    For c = 1 To columnCount
        targetData(targetRow, c) = sourceData(sourceRow, c)
    Next c

End Sub

Private Function HRS4_Trim2DArray( _
    ByVal sourceData As Variant, _
    ByVal rowCount As Long, _
    ByVal columnCount As Long) As Variant

    Dim resultData() As Variant
    Dim r As Long
    Dim c As Long

    ReDim resultData(1 To rowCount, 1 To columnCount)

    For r = 1 To rowCount
        For c = 1 To columnCount
            resultData(r, c) = sourceData(r, c)
        Next c
    Next r

    HRS4_Trim2DArray = resultData

End Function
