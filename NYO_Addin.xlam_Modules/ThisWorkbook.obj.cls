VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'##########
' 変数
'##########
Public WithEvents xlApp As Excel.Application   '他ブックのイベントを取得するために使用
Attribute xlApp.VB_VarHelpID = -1

Private lstReadOnlyAlerted As New Dictionary    '読み取り専用アラート済みリスト

'##########
' イベント
'##########

'*****************************************************************************
'[ イベント名 ]  Addinインストール時
'*****************************************************************************
Private Sub Workbook_AddinInstall()
    On Error Resume Next
    ' ツールバー追加
    Call modMenu.SetMenu
End Sub

'*****************************************************************************
'[ イベント名 ]  Addinアンインストール時
'*****************************************************************************
Private Sub Workbook_AddinUninstall()
    ' ツールバー削除
    Call modMenu.DelMenu
End Sub

'*****************************************************************************
'[ イベント名 ]  Addinブックオープン時
'*****************************************************************************
Private Sub Workbook_Open()
    '他ブックのイベントを取得するためにExcelアプリケーションオブジェクトを取得
    Set xlApp = Excel.Application
    
    'ショートカットキー割り当て
    Call modMenu.RegistShortcutKey
    
    '右クリックメニュー追加
    Call modMenu.AddContextMenu
End Sub

'*****************************************************************************
'[ イベント名 ]  Addinブック保存時
'*****************************************************************************
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    'Addinブックオープン時に取得したExcelアプリケーションオブジェクトが、
    'マクロを編集するとクリアされるため、Addin保存時に再取得
    Set xlApp = Excel.Application
    
    'なぜかここで処理中モードになってしまうことがあるので、強制的に解除
    Call modCommon.ProcessModeEnd
End Sub

'*****************************************************************************
'[ イベント名 ]  Addinブッククローズ時
'*****************************************************************************
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'マクロを編集しているのに閉じようとした場合、Addinの場合でも警告する
    If Not ThisWorkbook.Saved Then
        If ConfirmMessage(ThisWorkbook.Name & "が未保存です。保存しますか？") = vbYes Then
            ThisWorkbook.Save
        End If
    End If
End Sub

'*****************************************************************************
'[ イベント名 ]  他ブックのシートダブルクリック時
'*****************************************************************************
Private Sub xlApp_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    
    '元のブックのイベントハンドラーでキャンセル済みの場合は、何もしない
    If Cancel Then Exit Sub
    
    '入力補助
    '(入力補助にて処理された場合は、ダブルクリックイベントをキャンセルする)
    Cancel = InputAssist(Target.Cells(1, 1))
End Sub

'*****************************************************************************
'[ イベント名 ]  他ブックのオープン時
'*****************************************************************************
Private Sub xlApp_WorkbookOpen(ByVal Wb As Workbook)
    
    Dim wbk As Workbook
    Set wbk = Wb
    
    '念のため、NULL対策
    If lstReadOnlyAlerted Is Nothing Then
        Set lstReadOnlyAlerted = New Dictionary
    End If
        
    '開き直した場合に、読み取り専用アラート済みリストから除去されるようにする
    If lstReadOnlyAlerted.Exists(wbk.FullName) Then
        Call lstReadOnlyAlerted.Remove(wbk.FullName)
    End If
End Sub

'*****************************************************************************
'[ イベント名 ]  他ブックのシート変更時
'*****************************************************************************
Private Sub xlApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    Dim wbk As Workbook
    Set wbk = Sh.Parent
    
    '読み取り専用の場合、初回のみアラート
    If wbk.ReadOnly Then
        '念のため、NULL対策
        If lstReadOnlyAlerted Is Nothing Then
            Set lstReadOnlyAlerted = New Dictionary
        End If
        
        If Not lstReadOnlyAlerted.Exists(wbk.FullName) Then
            Call modMessage.ErrorMessage("このブックは読み取り専用です。ご注意下さい。（以降は警告しません。）")
            Call lstReadOnlyAlerted.Add(wbk.FullName, True)
        End If
    End If

End Sub




''*****************************************************************************
''[ イベント名 ]  他ブックのクローズ前イベント
''*****************************************************************************
'Private Sub xlApp_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
'
'On Error GoTo Catch
'    'Excel2007以降の場合は、元々バックアップ機能があるので、何もしない
'    If Is2007orLater Then Exit Sub
'
'    '自動バックアップ
'    '(未保存の場合は、この後保存して閉じるかどうかに拘わらず、バックアップを作成する)
'    If Not wb.Saved Then
'        Call BackupBook(wb)
'    End If
'    Exit Sub
'
'Catch:
'    'エラーが発生しても、何もしない
'    Exit Sub
'
'End Sub
'
''*****************************************************************************
''[ イベント名 ]  他ブックの保存前イベント
''*****************************************************************************
'Private Sub xlApp_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    '保存する場合は、バックアップを削除しておく
'    Call DeleteBackupBook(wb)
'End Sub

'*****************************************************************************
'[ イベント名 ]  CtrlEnter時
'*****************************************************************************
Private Sub CtrlEnter()
    '入力補助
    Call InputAssist(ActiveCell)
End Sub

'*****************************************************************************
'[ イベント名 ]  Ctrl+Alt+";"時
'*****************************************************************************
Public Sub CtrlAltSemicolon()
    ' 入力補助(日付)
    Call InputAssistDate(ActiveCell)
End Sub

'*****************************************************************************
'[ イベント名 ]  Ctrl+Alt+":"時
'*****************************************************************************
Public Sub CtrlAltColon()
    ' 入力補助(時刻)
    Call InputAssistTime(ActiveCell)
End Sub

