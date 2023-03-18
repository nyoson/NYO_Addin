Attribute VB_Name = "modBackUp"
Option Explicit

'Excel2007以降、バックアップ機能が実装されたので､不要

''##########
'' 定数
''##########
'Private Const BAKDIR As String = "C:\Temp\ExcelBak"
'Private Const TMP_FILE_NAME As String = "tmpBak.xls"
'
''##########
'' 関数
''##########
'
''*****************************************************************************
''[ 関数名 ]  BackupBook
''[ 概  要 ]　ブックのバックアップ
''[ 引  数 ]　対象のブック
''[ 戻り値 ]　なし
''*****************************************************************************
'Public Sub BackupBook(ByVal wbkTarget As Workbook)
'
'    If Dir(BAKDIR, vbDirectory) = Empty Then
'        'Call MkDir(BAKDIR)
'    End If
'    Call MakeDir(BAKDIR)
'
'    'バックアップのファイル名
'    Dim strBookName As String
'    strBookName = GetBackupFileName(wbkTarget)
'
'On Error GoTo Catch
'    'コピーを保存
'    Call wbkTarget.SaveCopyAs(BAKDIR & "\" & strBookName)
'    GoTo Finally
'Catch:
'    '【ToDo】CSVは「コピーを保存」が出来ないっぽい…？
'    'Call wbkTarget.SaveCopyAs(BAKDIR & "\" & strBookName & ".csv")
'Finally:
'
'End Sub
'
''*****************************************************************************
''[ 関数名 ]  DeleteBackupBook
''[ 概  要 ]　過去のバックアップファイルを削除
''[ 引  数 ]　対象のブック
''[ 戻り値 ]　なし
''*****************************************************************************
'Public Sub DeleteBackupBook(ByVal wbkTarget As Workbook)
'    '保存する場合は、バックアップを削除しておく
'    Dim strBakPath As String
'    strBakPath = BAKDIR & "\" & GetBackupFileName(wbkTarget)
'    If Dir(strBakPath) <> Empty Then
'        Call Kill(strBakPath)
'    End If
'End Sub
'
''*****************************************************************************
''[ 関数名 ]  GetBackupFileName
''[ 概  要 ]　ブックのバックアップ時のファイル名を取得
''[ 引  数 ]　対象のブック
''[ 戻り値 ]　ファイル名
''*****************************************************************************
'Private Function GetBackupFileName(ByVal wbkTarget As Workbook)
'
'    'バックアップのファイル名
'    Dim strBookName As String
'    strBookName = wbkTarget.Name
'
'    '一度も保存していない場合(Book1など)は、拡張子を付加
'    If wbkTarget.Path = Empty Then
'        'strBookName = strBookName & "_" & Format(Date, "yyyy_mmdd") & ".xls"
'        strBookName = strBookName & ".xls"
'    End If
'
'    'strBookName = Replace(Replace(strBookName, "[", "［"), "]", "］")
'
'    GetBackupFileName = strBookName
'
'End Function
