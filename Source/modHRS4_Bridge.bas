Attribute VB_Name = "modHRS4_Bridge"
Option Explicit

'=========================================================
' Ver4.0 Part1 既存読込処理との接続口
'=========================================================

Public Sub HRS4_AfterImport()

    '既存の発注書読込が完了し、
    '発注原票DBへの登録が終わった直後に呼び出します。
    HRS4_BuildAllCaches

End Sub

Public Sub HRS4_RebuildCacheOnly()

    HRS4_BuildAllCaches

    MsgBox "Ver4.0の完成キャッシュを再作成しました。", _
           vbInformation, "発注まるめシステム"

End Sub

Public Sub HRS4_CheckCacheStatus()

    Dim messageText As String

    messageText = "Ver4.0 Part1 キャッシュ保存先" & vbCrLf & vbCrLf
    messageText = messageText & "通常表示: " & _
                  HRS4_SH_NORMAL & vbCrLf
    messageText = messageText & "集約表示: " & _
                  HRS4_SH_AGGREGATE & vbCrLf
    messageText = messageText & "書戻し: " & _
                  HRS4_SH_WRITEBACK & vbCrLf
    messageText = messageText & "セッション: " & _
                  HRS4_SH_SESSION & vbCrLf
    messageText = messageText & "商品索引: " & _
                  HRS4_SH_INDEX & vbCrLf
    messageText = messageText & "速度ログ: " & _
                  HRS4_SH_PERFORMANCE

    MsgBox messageText, _
           vbInformation, "発注まるめシステム"

End Sub
