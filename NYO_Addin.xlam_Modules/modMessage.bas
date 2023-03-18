Attribute VB_Name = "modMessage"
Option Explicit

'当モジュールのPublic関数を「マクロの実行」に表示しないようにする
Option Private Module

'＜メッセージ関連モジュール＞

'##########
' 関数
'##########

'*****************************************************************************
'[ 関数名 ] ErrorMessage
'[ 概  要 ] エラーメッセージダイアログを表示する
'[ 引  数 ] メッセージ
'           エラー情報（省略可）
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub ErrorMessage(ByVal message As String, Optional ByVal errInfo As ErrObject = Nothing)
    Call CommonMessage(message, vbOKOnly + vbExclamation, errInfo)
End Sub

'*****************************************************************************
'[ 関数名 ] SystemErrorMessage
'[ 概  要 ] システムエラーメッセージダイアログを表示する
'[ 引  数 ] メッセージ
'           エラー情報（省略可）
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub SystemErrorMessage(ByVal message As String, _
                                Optional ByVal errInfo As ErrObject = Nothing)
    Call CommonMessage(message, vbOKOnly + vbCritical, errInfo)
End Sub

'*****************************************************************************
'[ 関数名 ] InfoMessage
'[ 概  要 ] 情報メッセージダイアログを表示する
'[ 引  数 ] メッセージ
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub InfoMessage(ByVal message As String)
    Call CommonMessage(message, vbOKOnly + vbInformation)
End Sub

'*****************************************************************************
'[ 関数名 ] ConfirmMessage
'[ 概  要 ] 確認メッセージダイアログを表示する
'[ 引  数 ] メッセージ
'           メッセージボックススタイル（省略時YesNo）
'[ 戻り値 ] 結果コード(vbYes,vbNo,...)
'*****************************************************************************
Public Function ConfirmMessage(ByVal message As String, _
            Optional ByVal opt As VbMsgBoxStyle = VbMsgBoxStyle.vbYesNo) As VbMsgBoxResult
    ConfirmMessage = CommonMessage(message, opt + vbQuestion)
End Function

'*****************************************************************************
'[ 関数名 ] CommonMessage
'[ 概  要 ] メッセージダイアログを表示する
'[ 引  数 ] メッセージ
'           メッセージボックススタイル
'           エラー情報（省略可）
'[ 戻り値 ] 結果コード(vbYes,vbNo)
'*****************************************************************************
Public Function CommonMessage(ByVal message As String, _
                                ByVal buttons As VbMsgBoxStyle, _
                                Optional ByVal errInfo As ErrObject = Nothing) As VbMsgBoxResult
    Dim messageText As String
    
    If errInfo Is Nothing Then
        messageText = message & "　　"
        '                       ↑メッセージダイアログの右余白を少し空ける
    Else
        messageText = message & vbCrLf _
                & FormatErrorInfo(errInfo)
        Call DebugPrintErr(errInfo)
    End If
    
    Dim bakScreenUpdating As Boolean
    Dim bakCursor As XlMousePointer
    
    '現在の状態を退避
    bakScreenUpdating = Application.ScreenUpdating
    bakCursor = Application.Cursor
    
    '一時的に解除
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    
    'メッセージ表示
    If errInfo Is Nothing Then
        CommonMessage = MsgBox(messageText, buttons, C_TOOL_NAME)
    Else
        CommonMessage = MsgBox(messageText, buttons, C_TOOL_NAME, errInfo.HelpFile, errInfo.HelpContext)
    End If
    
    '状態を元に戻す
    Application.Cursor = bakCursor
    Application.ScreenUpdating = bakScreenUpdating
End Function


'*****************************************************************************
'[ 関数名 ] DebugPrintErr
'[ 概  要 ] エラー情報をイミディエイトウィンドウにログ出力する
'[ 引  数 ] エラー情報
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub DebugPrintErr(ByVal errInfo As ErrObject)
    
    If errInfo.Number = 0 Then Exit Sub
    
    Debug.Print "===="
    Debug.Print FormatErrorInfo(errInfo, True)
    Debug.Print "===="

End Sub

'*****************************************************************************
'[ 関数名 ] ErrorInfoToString
'[ 概  要 ] エラー情報を文字列に整形する
'[ 引  数 ] エラー情報
'           詳細モード（省略時False）
'[ 戻り値 ] エラー情報テキスト
'*****************************************************************************
Private Function FormatErrorInfo(ByVal errInfo As ErrObject, Optional ByVal detailMode As Boolean = False)
    
    Dim errText As String
    errText = "実行時エラー'" & CStr(Err.Number) & "':"
    
    errText = errText & vbCrLf & "エラー番号：0x" & Hex(errInfo.Number)
    errText = errText & vbCrLf & "エラー詳細：" & errInfo.Description
    
    '簡易モードの場合はここで終了
    If detailMode = False Then GoTo Finally
    
    '詳細モードの場合は、追加情報を出力
    errText = errText & vbCrLf & "Source：" & errInfo.Source
    errText = errText & vbCrLf & "HelpFile：" & errInfo.HelpFile
    errText = errText & vbCrLf & "HelpContext：" & errInfo.HelpContext
    If errInfo.LastDllError <> 0 Then
        errText = errText & vbCrLf & "LastDllError：" & errInfo.LastDllError
    End If
    
Finally:
    FormatErrorInfo = errText
    Exit Function

End Function

