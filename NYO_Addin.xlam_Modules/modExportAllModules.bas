Attribute VB_Name = "modExportAllModules"
Option Explicit

'以下を参照設定に追加して下さい。
'・Microsoft Visual Basic for Applications Extensibility
'（VBComponentクラスなど）

'*****************************************************************************
' VBAのすべてのモジュール等をエクスポート
'*****************************************************************************
Public Sub ExportAllModules()

    Dim intSheet    As Integer  'シート用ループ変数
    Dim varFolder   As Variant  '出力フォルダの格納先フォルダ
    Dim strFolder   As String   '出力フォルダ
    Dim strExt      As String   '拡張子
    Dim objVBC      As VBComponent   'VBA Component Object
    
    
    'セキュリティ設定チェック
    If Not CheckMacroSecurityWithMsg() Then
        Exit Sub
    End If
    
    '対象ブック選択
    Dim wbkTarget   As Workbook '対象ブック
    Dim varMode As Variant
    varMode = Application.InputBox( _
            Title:=C_TOOL_NAME, _
            Prompt:="VBAをエクスポートする対象のブックを選択して下さい。" & vbCrLf & _
                " [0]:ThisWorkbook[" & ThisWorkbook.Name & "]" & vbCrLf & _
                " [1]:ActiveWorkbook[" & ActiveWorkbook.Name & "]", _
            Default:="0", _
            Type:=1)    '数値のみ受け取る
    If VarType(varMode) = vbBoolean And varMode = False Then Exit Sub 'キャンセル時は何もせず終了
    
    Select Case varMode
        Case 0
            Set wbkTarget = ThisWorkbook
        Case 1
            Set wbkTarget = ActiveWorkbook
        Case Else
            Call ErrorMessage("入力値が範囲外です。")
            Exit Sub
    End Select
    
    'プロジェクトがロックされている場合は不可
    If wbkTarget.VBProject.Protection = VBIDE.vbext_ProjectProtection.vbext_pp_locked Then
        Call modMessage.ErrorMessage("プロジェクトがロックされています。解除してから再度実行して下さい。")
        Exit Sub
    End If
    
    'フォルダを選択(Defaultは本体と同パス)
    varFolder = GetSelectFolderPath(wbkTarget.Path)
    
    'キャンセル時は終了
    If varFolder = False Then
        Exit Sub
    End If
    
    '出力フォルダを作成(拡張子を含めたファイル名+_Modules)
    strFolder = varFolder & "\" & wbkTarget.Name & "_Modules"
    If FolderExists(strFolder) Then
        If modMessage.ConfirmMessage( _
                "前回出力フォルダが存在します。一旦削除して続行しますか？" & vbCrLf & _
                strFolder) <> VbMsgBoxResult.vbYes Then
            Exit Sub
        End If
        'フォルダ削除
        If Not RemoveFolder(strFolder) Then
            Call modMessage.ErrorMessage("前回出力フォルダを削除できませんでした。")
            Exit Sub
        End If
    End If
    Call MkDir(strFolder)
    
    'VBAコンポーネント
    For Each objVBC In wbkTarget.VBProject.VBComponents
        Select Case objVBC.Type
            Case 1
                '標準モジュール
                strExt = ".bas"
            Case 2
                'クラスモジュール
                strExt = ".cls"
            Case 3
                'フォームモジュール
                strExt = ".frm"
            Case 100
                'ThisWorkbook or Sheet
                strExt = ".obj.cls"
            Case Else
                'その他
                strExt = ".obj.cls"
        End Select
        
        If objVBC.CodeModule.CountOfLines > 1 Then
            '指定フォルダにエクスポート
            objVBC.Export strFolder & "\" & objVBC.Name & strExt
        End If
    Next objVBC
    
    Call modMessage.InfoMessage("完了しました。")
    
End Sub

'*****************************************************************************
'[ 関数名 ] GetSelectFolderPath
'[ 概  要 ] フォルダを選択しパスを取得する
'[ 引  数 ] フォルダ選択ダイアログの初期表示パス(省略可)
'[ 戻り値 ] 選択フォルダパス or False
'*****************************************************************************
Private Function GetSelectFolderPath(Optional ByVal strDefaultPath As String = "") As Variant
    Dim varFolderPath As Variant '選択フォルダパス
    Dim strDialogTitle As String 'ダイアログのタイトル
    
    'デフォルトパスが存在しない or 空白なら
    If FolderExists(strDefaultPath) = False Or Len(strDefaultPath) = 0 Then
        '自身のパスを設定
        strDefaultPath = ActiveWorkbook.Path & "\"
    End If
    
    'フォルダ選択ウィンドウを表示
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        '初期フォルダ設定
        .InitialFileName = strDefaultPath
        'ダイアログタイトルの設定
        .Title = "フォルダを選択してください"
        
        '複数選択禁止
        .AllowMultiSelect = False
        
        If .Show = True Then
            '選択された場合、フォルダパス
            varFolderPath = .SelectedItems(1)
        Else
            '選択されなかった（キャンセル等）場合、False
            varFolderPath = False
        End If
    End With
    
    '上記ダイアログにより、勝手にカレントディレクトリが変更されるので、デフォルトパスまたはブックのパスに変更する
On Error GoTo Catch
    Call ChDir(strDefaultPath)
    GoTo Finally
Catch:
    Dim resetPath As String
    resetPath = ActiveWorkbook.Path
    If resetPath <> Empty Then    ' ※但し、未保存のブックの場合は諦める
        Call ChDir(resetPath)
    End If
Finally:
On Error GoTo 0
    
    GetSelectFolderPath = varFolderPath
End Function

'*****************************************************************************
' フォルダ存在チェック
'*****************************************************************************
Private Function FolderExists(strPath) As Boolean
    
    Dim result As Boolean
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    result = FSO.FolderExists(strPath)
    Set FSO = Nothing
    
    FolderExists = result
    
End Function

'*****************************************************************************
' フォルダ削除
'*****************************************************************************
Private Function RemoveFolder(ByVal strFolderPath As String, Optional ByVal force As Boolean = True) As Boolean

On Error GoTo Catch
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Call FSO.DeleteFolder(strFolderPath)
    Set FSO = Nothing
    RemoveFolder = True
    
Finally:
    Exit Function
    
Catch:
    Call modMessage.DebugPrintErr(Err)
    Debug.Print strFolderPath
    
    RemoveFolder = False
    Resume Finally
    
End Function
    
