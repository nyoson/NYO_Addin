Attribute VB_Name = "modFile"
Option Explicit

'*****************************************************************************
' ひとつ上のフォルダを開く
'*****************************************************************************
Public Sub OpenDir()
    Dim WSHShell As Object
    Set WSHShell = CreateObject("Wscript.Shell")
    WSHShell.Exec ("explorer /select," + ActiveWorkbook.FullName)
    Set WSHShell = Nothing
End Sub

'*****************************************************************************
' ファイル名をクリップボードにコピー
'*****************************************************************************
Public Sub CopyFileName()
    Dim data As New DataObject
    
    Call data.SetText(ActiveWorkbook.Name)
    Call data.PutInClipboard
End Sub

'*****************************************************************************
' ActiveSheetをコピーして別のワークブック（CSVファイル）として保存
'*****************************************************************************
Public Sub ExportCSV()
    
    '対象シート
    Dim shCopy As Worksheet
    Set shCopy = ActiveSheet
    
    '新しいファイルのファイル名
    Dim newFileName As String
    newFileName = shCopy.Name
    
    '保存先フォルダパス
    Dim newFileFolder As String
    newFileFolder = shCopy.Parent.Path
    
    '出力ファイルパス
    Dim newFile As String
    newFile = newFileFolder & "\" & newFileName & ".csv"
    
    '生成しようとしているファイルが書き込みできるか確認
On Error Resume Next
    Open newFile For Append As #1
    Close #1
    If Err.Number > 0 Then
        Call modMessage.ErrorMessage("書き込みできません。" & vbCrLf _
                & newFile)
        Exit Sub
    End If
On Error GoTo 0
    
    '新しいファイルと同名のファイルが開かれていないか確認
On Error Resume Next
    Dim wbkTest As Workbook
    Set wbkTest = Workbooks(newFileName & ".csv")
    If Not wbkTest Is Nothing Then
        'あればアクティブにした上で、エラーメッセージを表示
        wbkTest.Activate
        Call modMessage.ErrorMessage( _
            "生成するCSVと同名のファイルがすでに開かれています。" & vbCrLf _
            & wbkTest.FullName & "を閉じてやり直してください。")
        Exit Sub
    End If
On Error GoTo 0
    
    '警告表示を一時的に抑制
    Application.DisplayAlerts = False
    
    '対象シートをコピー
    shCopy.Copy
    
    'コピーしたシートを新しいファイルとしてCSV形式で保存
    ActiveWorkbook.SaveAs FileName:=newFile, FileFormat:=xlCSV
    
    '閉じる
    ActiveWindow.Close
    
    '警告表示の抑制を解除
    Application.DisplayAlerts = True
    
    '終了メッセージ
    Call modMessage.InfoMessage("ファイルの作成に成功しました。")
    
End Sub

