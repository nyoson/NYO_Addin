Attribute VB_Name = "modFile"
Option Explicit

'*****************************************************************************
' �ЂƂ�̃t�H���_���J��
'*****************************************************************************
Public Sub OpenDir()
    Dim WSHShell As Object
    Set WSHShell = CreateObject("Wscript.Shell")
    WSHShell.Exec ("explorer /select," + ActiveWorkbook.FullName)
    Set WSHShell = Nothing
End Sub

'*****************************************************************************
' �t�@�C�������N���b�v�{�[�h�ɃR�s�[
'*****************************************************************************
Public Sub CopyFileName()
    Dim data As New DataObject
    
    Call data.SetText(ActiveWorkbook.Name)
    Call data.PutInClipboard
End Sub

'*****************************************************************************
' ActiveSheet���R�s�[���ĕʂ̃��[�N�u�b�N�iCSV�t�@�C���j�Ƃ��ĕۑ�
'*****************************************************************************
Public Sub ExportCSV()
    
    '�ΏۃV�[�g
    Dim shCopy As Worksheet
    Set shCopy = ActiveSheet
    
    '�V�����t�@�C���̃t�@�C����
    Dim newFileName As String
    newFileName = shCopy.Name
    
    '�ۑ���t�H���_�p�X
    Dim newFileFolder As String
    newFileFolder = shCopy.Parent.Path
    
    '�o�̓t�@�C���p�X
    Dim newFile As String
    newFile = newFileFolder & "\" & newFileName & ".csv"
    
    '�������悤�Ƃ��Ă���t�@�C�����������݂ł��邩�m�F
On Error Resume Next
    Open newFile For Append As #1
    Close #1
    If Err.Number > 0 Then
        Call modMessage.ErrorMessage("�������݂ł��܂���B" & vbCrLf _
                & newFile)
        Exit Sub
    End If
On Error GoTo 0
    
    '�V�����t�@�C���Ɠ����̃t�@�C�����J����Ă��Ȃ����m�F
On Error Resume Next
    Dim wbkTest As Workbook
    Set wbkTest = Workbooks(newFileName & ".csv")
    If Not wbkTest Is Nothing Then
        '����΃A�N�e�B�u�ɂ�����ŁA�G���[���b�Z�[�W��\��
        wbkTest.Activate
        Call modMessage.ErrorMessage( _
            "��������CSV�Ɠ����̃t�@�C�������łɊJ����Ă��܂��B" & vbCrLf _
            & wbkTest.FullName & "����Ă�蒼���Ă��������B")
        Exit Sub
    End If
On Error GoTo 0
    
    '�x���\�����ꎞ�I�ɗ}��
    Application.DisplayAlerts = False
    
    '�ΏۃV�[�g���R�s�[
    shCopy.Copy
    
    '�R�s�[�����V�[�g��V�����t�@�C���Ƃ���CSV�`���ŕۑ�
    ActiveWorkbook.SaveAs FileName:=newFile, FileFormat:=xlCSV
    
    '����
    ActiveWindow.Close
    
    '�x���\���̗}��������
    Application.DisplayAlerts = True
    
    '�I�����b�Z�[�W
    Call modMessage.InfoMessage("�t�@�C���̍쐬�ɐ������܂����B")
    
End Sub

