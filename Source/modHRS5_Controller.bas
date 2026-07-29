Attribute VB_Name = "modHRS5_Controller"
Option Explicit

'============================================================
' 発注まるめシステム Ver5.0
' Controller 基盤
' Ver5.0.1 Part1 Fix
' ParamArrayを使用しない固定引数方式
'============================================================

Private Const HRS5_VERSION As String = "Ver5.0.1 Part1 Fix"
Private Const PROC_AFTER_IMPORT As String = "HRS4_AfterImport"
Private Const PROC_LOAD_SELECTED As String = "HRS4_LoadSelectedProductToLegacyPreview"
Private Const PROC_SESSION_INIT As String = "HRS4S_Initialize"
Private Const PROC_SESSION_SAVE As String = "HRS4S_SaveCurrentPreview"
Private Const PROC_SESSION_RESTORE As String = "HRS4S_RestoreCurrentPreview"
Private Const PROC_SESSION_FLUSH As String = "HRS4S_FlushCache"
Private Const PROC_DISTRIBUTE As String = "HRS4D_DistributePreview"
Private Const PROC_CANCEL_CURRENT As String = "HRS4_CancelCurrentProduct"
Private Const PROC_CANCEL_ALL As String = "HRS4_CancelAll"
Private Const PROC_CANCEL_POINT_ONE As String = "HRS4_CancelPointOne"
Private Const PROC_CLEAR_CANCEL As String = "HRS4_ClearCancel"
Private Const PROC_UPDATE_CANCEL_DISPLAY As String = "HRS4C_UpdateCancelDisplay"

Private mBusy As Boolean
Private mInitialized As Boolean
Private mCurrentProductCode As String
Private mCurrentVendorCode As String
Private mLastError As String

Public Sub HRS5_ShowVersion()
    MsgBox "発注まるめシステム " & HRS5_VERSION, vbInformation
End Sub

Public Function HRS5_IsBusy() As Boolean
    HRS5_IsBusy = mBusy
End Function

Public Function HRS5_IsInitialized() As Boolean
    HRS5_IsInitialized = mInitialized
End Function

Public Function HRS5_GetCurrentProductCode() As String
    HRS5_GetCurrentProductCode = mCurrentProductCode
End Function

Public Function HRS5_GetLastError() As String
    HRS5_GetLastError = mLastError
End Function

Public Sub HRS5_Initialize()
    If Not HRS5_Enter("初期化") Then Exit Sub
    On Error GoTo EH

    HRS5_RunOptional1 PROC_SESSION_INIT, True
    mInitialized = True

Done:
    HRS5_Leave
    Exit Sub
EH:
    HRS5_RecordError "HRS5_Initialize", Err.Number, Err.Description
    Resume Done
End Sub

Public Sub HRS5_AfterImport()
    If Not HRS5_Enter("読込後処理") Then Exit Sub
    On Error GoTo EH

    HRS5_RunRequired0 PROC_AFTER_IMPORT
    HRS5_RunOptional1 PROC_SESSION_INIT, True
    mInitialized = True

Done:
    HRS5_Leave
    Exit Sub
EH:
    HRS5_RecordError "HRS5_AfterImport", Err.Number, Err.Description
    Resume Done
End Sub

Public Sub HRS5_SelectProduct(ByVal productCode As String, Optional ByVal vendorCode As String = "")
    If Len(Trim$(productCode)) = 0 Then Exit Sub
    If Not HRS5_Enter("商品選択") Then Exit Sub
    On Error GoTo EH

    If Len(mCurrentProductCode) > 0 Then
        HRS5_RunOptional2 PROC_SESSION_SAVE, mCurrentProductCode, mCurrentVendorCode
    End If

    mCurrentProductCode = Trim$(productCode)
    mCurrentVendorCode = Trim$(vendorCode)

    HRS5_RunRequired0 PROC_LOAD_SELECTED
    HRS5_RunOptional2 PROC_SESSION_RESTORE, mCurrentProductCode, mCurrentVendorCode
    HRS5_RunOptional0 PROC_UPDATE_CANCEL_DISPLAY

Done:
    HRS5_Leave
    Exit Sub
EH:
    HRS5_RecordError "HRS5_SelectProduct", Err.Number, Err.Description
    Resume Done
End Sub

Public Sub HRS5_Distribute(ByVal orderQty As Double, Optional ByVal includeCancelled As Boolean = False)
    If orderQty < 0 Then
        MsgBox "発注数は0以上で入力してください。", vbExclamation
        Exit Sub
    End If
    If Not HRS5_Enter("発注配分") Then Exit Sub
    On Error GoTo EH

    HRS5_RunRequired3 PROC_DISTRIBUTE, orderQty, includeCancelled, True
    HRS5_RunOptional2 PROC_SESSION_SAVE, mCurrentProductCode, mCurrentVendorCode
    HRS5_RunOptional0 PROC_UPDATE_CANCEL_DISPLAY

Done:
    HRS5_Leave
    Exit Sub
EH:
    HRS5_RecordError "HRS5_Distribute", Err.Number, Err.Description
    Resume Done
End Sub

Public Sub HRS5_CancelCurrentProduct(Optional ByVal turnOn As Boolean = True)
    If Not HRS5_Enter("商品取消") Then Exit Sub
    On Error GoTo EH

    HRS5_RunRequired1 PROC_CANCEL_CURRENT, turnOn
    HRS5_RunOptional2 PROC_SESSION_SAVE, mCurrentProductCode, mCurrentVendorCode

Done:
    HRS5_Leave
    Exit Sub
EH:
    HRS5_RecordError "HRS5_CancelCurrentProduct", Err.Number, Err.Description
    Resume Done
End Sub

Public Sub HRS5_CancelAll()
    If Not HRS5_Enter("全取消") Then Exit Sub
    On Error GoTo EH

    HRS5_RunRequired0 PROC_CANCEL_ALL
    HRS5_RunOptional0 PROC_SESSION_FLUSH

Done:
    HRS5_Leave
    Exit Sub
EH:
    HRS5_RecordError "HRS5_CancelAll", Err.Number, Err.Description
    Resume Done
End Sub

Public Sub HRS5_CancelPointOne()
    If Not HRS5_Enter("0.1一括取消") Then Exit Sub
    On Error GoTo EH

    HRS5_RunRequired0 PROC_CANCEL_POINT_ONE
    HRS5_RunOptional2 PROC_SESSION_SAVE, mCurrentProductCode, mCurrentVendorCode

Done:
    HRS5_Leave
    Exit Sub
EH:
    HRS5_RecordError "HRS5_CancelPointOne", Err.Number, Err.Description
    Resume Done
End Sub

Public Sub HRS5_ClearCancel()
    If Not HRS5_Enter("取消解除") Then Exit Sub
    On Error GoTo EH

    HRS5_RunRequired0 PROC_CLEAR_CANCEL
    HRS5_RunOptional2 PROC_SESSION_SAVE, mCurrentProductCode, mCurrentVendorCode

Done:
    HRS5_Leave
    Exit Sub
EH:
    HRS5_RecordError "HRS5_ClearCancel", Err.Number, Err.Description
    Resume Done
End Sub

Public Sub HRS5_SaveSession()
    If Not HRS5_Enter("セッション保存") Then Exit Sub
    On Error GoTo EH

    HRS5_RunOptional2 PROC_SESSION_SAVE, mCurrentProductCode, mCurrentVendorCode
    HRS5_RunOptional0 PROC_SESSION_FLUSH

Done:
    HRS5_Leave
    Exit Sub
EH:
    HRS5_RecordError "HRS5_SaveSession", Err.Number, Err.Description
    Resume Done
End Sub

Public Sub HRS5_Close()
    On Error Resume Next
    HRS5_SaveSession
    mInitialized = False
    mCurrentProductCode = vbNullString
    mCurrentVendorCode = vbNullString
    On Error GoTo 0
End Sub

Private Function HRS5_Enter(ByVal operationName As String) As Boolean
    If mBusy Then
        MsgBox "別の処理を実行中です。完了後にもう一度実行してください。", vbExclamation
        Exit Function
    End If

    mBusy = True
    mLastError = vbNullString
    Application.StatusBar = "発注まるめシステム Ver5: " & operationName & "..."
    HRS5_Enter = True
End Function

Private Sub HRS5_Leave()
    Application.StatusBar = False
    mBusy = False
End Sub

Private Sub HRS5_RecordError(ByVal procedureName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    mLastError = procedureName & " / " & CStr(errorNumber) & " / " & errorDescription
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & vbCrLf & mLastError, vbCritical
End Sub

'============================================================
' 必須処理
'============================================================
Private Sub HRS5_RunRequired0(ByVal procedureName As String)
    On Error GoTo EH
    Application.Run procedureName
    Exit Sub
EH:
    HRS5_RaiseRequiredError procedureName, Err.Number, Err.Description
End Sub

Private Sub HRS5_RunRequired1(ByVal procedureName As String, ByVal arg1 As Variant)
    On Error GoTo EH
    Application.Run procedureName, arg1
    Exit Sub
EH:
    HRS5_RaiseRequiredError procedureName, Err.Number, Err.Description
End Sub

Private Sub HRS5_RunRequired2(ByVal procedureName As String, ByVal arg1 As Variant, ByVal arg2 As Variant)
    On Error GoTo EH
    Application.Run procedureName, arg1, arg2
    Exit Sub
EH:
    HRS5_RaiseRequiredError procedureName, Err.Number, Err.Description
End Sub

Private Sub HRS5_RunRequired3(ByVal procedureName As String, ByVal arg1 As Variant, ByVal arg2 As Variant, ByVal arg3 As Variant)
    On Error GoTo EH
    Application.Run procedureName, arg1, arg2, arg3
    Exit Sub
EH:
    HRS5_RaiseRequiredError procedureName, Err.Number, Err.Description
End Sub

Private Sub HRS5_RaiseRequiredError(ByVal procedureName As String, ByVal sourceNumber As Long, ByVal sourceDescription As String)
    Dim messageText As String

    messageText = "必要な処理が見つからないか、実行に失敗しました: " & procedureName
    If Len(sourceDescription) > 0 Then
        messageText = messageText & vbCrLf & "元のエラー: " & CStr(sourceNumber) & " / " & sourceDescription
    End If

    Err.Raise vbObjectError + 5501, "modHRS5_Controller", messageText
End Sub

'============================================================
' 任意処理
' 存在しない場合や実行エラーの場合はController処理を停止しません
'============================================================
Private Sub HRS5_RunOptional0(ByVal procedureName As String)
    On Error Resume Next
    Application.Run procedureName
    On Error GoTo 0
End Sub

Private Sub HRS5_RunOptional1(ByVal procedureName As String, ByVal arg1 As Variant)
    On Error Resume Next
    Application.Run procedureName, arg1
    On Error GoTo 0
End Sub

Private Sub HRS5_RunOptional2(ByVal procedureName As String, ByVal arg1 As Variant, ByVal arg2 As Variant)
    On Error Resume Next
    Application.Run procedureName, arg1, arg2
    On Error GoTo 0
End Sub

Private Sub HRS5_RunOptional3(ByVal procedureName As String, ByVal arg1 As Variant, ByVal arg2 As Variant, ByVal arg3 As Variant)
    On Error Resume Next
    Application.Run procedureName, arg1, arg2, arg3
    On Error GoTo 0
End Sub
