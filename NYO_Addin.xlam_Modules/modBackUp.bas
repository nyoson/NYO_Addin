Attribute VB_Name = "modBackUp"
Option Explicit

'Excel2007�ȍ~�A�o�b�N�A�b�v�@�\���������ꂽ�̂Ť�s�v

''##########
'' �萔
''##########
'Private Const BAKDIR As String = "C:\Temp\ExcelBak"
'Private Const TMP_FILE_NAME As String = "tmpBak.xls"
'
''##########
'' �֐�
''##########
'
''*****************************************************************************
''[ �֐��� ]  BackupBook
''[ �T  �v ]�@�u�b�N�̃o�b�N�A�b�v
''[ ��  �� ]�@�Ώۂ̃u�b�N
''[ �߂�l ]�@�Ȃ�
''*****************************************************************************
'Public Sub BackupBook(ByVal wbkTarget As Workbook)
'
'    If Dir(BAKDIR, vbDirectory) = Empty Then
'        'Call MkDir(BAKDIR)
'    End If
'    Call MakeDir(BAKDIR)
'
'    '�o�b�N�A�b�v�̃t�@�C����
'    Dim strBookName As String
'    strBookName = GetBackupFileName(wbkTarget)
'
'On Error GoTo Catch
'    '�R�s�[��ۑ�
'    Call wbkTarget.SaveCopyAs(BAKDIR & "\" & strBookName)
'    GoTo Finally
'Catch:
'    '�yToDo�zCSV�́u�R�s�[��ۑ��v���o���Ȃ����ۂ��c�H
'    'Call wbkTarget.SaveCopyAs(BAKDIR & "\" & strBookName & ".csv")
'Finally:
'
'End Sub
'
''*****************************************************************************
''[ �֐��� ]  DeleteBackupBook
''[ �T  �v ]�@�ߋ��̃o�b�N�A�b�v�t�@�C�����폜
''[ ��  �� ]�@�Ώۂ̃u�b�N
''[ �߂�l ]�@�Ȃ�
''*****************************************************************************
'Public Sub DeleteBackupBook(ByVal wbkTarget As Workbook)
'    '�ۑ�����ꍇ�́A�o�b�N�A�b�v���폜���Ă���
'    Dim strBakPath As String
'    strBakPath = BAKDIR & "\" & GetBackupFileName(wbkTarget)
'    If Dir(strBakPath) <> Empty Then
'        Call Kill(strBakPath)
'    End If
'End Sub
'
''*****************************************************************************
''[ �֐��� ]  GetBackupFileName
''[ �T  �v ]�@�u�b�N�̃o�b�N�A�b�v���̃t�@�C�������擾
''[ ��  �� ]�@�Ώۂ̃u�b�N
''[ �߂�l ]�@�t�@�C����
''*****************************************************************************
'Private Function GetBackupFileName(ByVal wbkTarget As Workbook)
'
'    '�o�b�N�A�b�v�̃t�@�C����
'    Dim strBookName As String
'    strBookName = wbkTarget.Name
'
'    '��x���ۑ����Ă��Ȃ��ꍇ(Book1�Ȃ�)�́A�g���q��t��
'    If wbkTarget.Path = Empty Then
'        'strBookName = strBookName & "_" & Format(Date, "yyyy_mmdd") & ".xls"
'        strBookName = strBookName & ".xls"
'    End If
'
'    'strBookName = Replace(Replace(strBookName, "[", "�m"), "]", "�n")
'
'    GetBackupFileName = strBookName
'
'End Function
