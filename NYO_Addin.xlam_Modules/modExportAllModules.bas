Attribute VB_Name = "modExportAllModules"
Option Explicit

'�ȉ����Q�Ɛݒ�ɒǉ����ĉ������B
'�EMicrosoft Visual Basic for Applications Extensibility
'�iVBComponent�N���X�Ȃǁj

'*****************************************************************************
' VBA�̂��ׂẴ��W���[�������G�N�X�|�[�g
'*****************************************************************************
Public Sub ExportAllModules()

    Dim intSheet    As Integer  '�V�[�g�p���[�v�ϐ�
    Dim varFolder   As Variant  '�o�̓t�H���_�̊i�[��t�H���_
    Dim strFolder   As String   '�o�̓t�H���_
    Dim strExt      As String   '�g���q
    Dim objVBC      As VBComponent   'VBA Component Object
    
    
    '�Z�L�����e�B�ݒ�`�F�b�N
    If Not CheckMacroSecurityWithMsg() Then
        Exit Sub
    End If
    
    '�Ώۃu�b�N�I��
    Dim wbkTarget   As Workbook '�Ώۃu�b�N
    Dim varMode As Variant
    varMode = Application.InputBox( _
            Title:=C_TOOL_NAME, _
            Prompt:="VBA���G�N�X�|�[�g����Ώۂ̃u�b�N��I�����ĉ������B" & vbCrLf & _
                " [0]:ThisWorkbook[" & ThisWorkbook.Name & "]" & vbCrLf & _
                " [1]:ActiveWorkbook[" & ActiveWorkbook.Name & "]", _
            Default:="0", _
            Type:=1)    '���l�̂ݎ󂯎��
    If VarType(varMode) = vbBoolean And varMode = False Then Exit Sub '�L�����Z�����͉��������I��
    
    Select Case varMode
        Case 0
            Set wbkTarget = ThisWorkbook
        Case 1
            Set wbkTarget = ActiveWorkbook
        Case Else
            Call ErrorMessage("���͒l���͈͊O�ł��B")
            Exit Sub
    End Select
    
    '�v���W�F�N�g�����b�N����Ă���ꍇ�͕s��
    If wbkTarget.VBProject.Protection = VBIDE.vbext_ProjectProtection.vbext_pp_locked Then
        Call modMessage.ErrorMessage("�v���W�F�N�g�����b�N����Ă��܂��B�������Ă���ēx���s���ĉ������B")
        Exit Sub
    End If
    
    '�t�H���_��I��(Default�͖{�̂Ɠ��p�X)
    varFolder = GetSelectFolderPath(wbkTarget.Path)
    
    '�L�����Z�����͏I��
    If varFolder = False Then
        Exit Sub
    End If
    
    '�o�̓t�H���_���쐬(�g���q���܂߂��t�@�C����+_Modules)
    strFolder = varFolder & "\" & wbkTarget.Name & "_Modules"
    If FolderExists(strFolder) Then
        If modMessage.ConfirmMessage( _
                "�O��o�̓t�H���_�����݂��܂��B��U�폜���đ��s���܂����H" & vbCrLf & _
                strFolder) <> VbMsgBoxResult.vbYes Then
            Exit Sub
        End If
        '�t�H���_�폜
        If Not RemoveFolder(strFolder) Then
            Call modMessage.ErrorMessage("�O��o�̓t�H���_���폜�ł��܂���ł����B")
            Exit Sub
        End If
    End If
    Call MkDir(strFolder)
    
    'VBA�R���|�[�l���g
    For Each objVBC In wbkTarget.VBProject.VBComponents
        Select Case objVBC.Type
            Case 1
                '�W�����W���[��
                strExt = ".bas"
            Case 2
                '�N���X���W���[��
                strExt = ".cls"
            Case 3
                '�t�H�[�����W���[��
                strExt = ".frm"
            Case 100
                'ThisWorkbook or Sheet
                strExt = ".obj.cls"
            Case Else
                '���̑�
                strExt = ".obj.cls"
        End Select
        
        If objVBC.CodeModule.CountOfLines > 1 Then
            '�w��t�H���_�ɃG�N�X�|�[�g
            objVBC.Export strFolder & "\" & objVBC.Name & strExt
        End If
    Next objVBC
    
    Call modMessage.InfoMessage("�������܂����B")
    
End Sub

'*****************************************************************************
'[ �֐��� ] GetSelectFolderPath
'[ �T  �v ] �t�H���_��I�����p�X���擾����
'[ ��  �� ] �t�H���_�I���_�C�A���O�̏����\���p�X(�ȗ���)
'[ �߂�l ] �I���t�H���_�p�X or False
'*****************************************************************************
Private Function GetSelectFolderPath(Optional ByVal strDefaultPath As String = "") As Variant
    Dim varFolderPath As Variant '�I���t�H���_�p�X
    Dim strDialogTitle As String '�_�C�A���O�̃^�C�g��
    
    '�f�t�H���g�p�X�����݂��Ȃ� or �󔒂Ȃ�
    If FolderExists(strDefaultPath) = False Or Len(strDefaultPath) = 0 Then
        '���g�̃p�X��ݒ�
        strDefaultPath = ActiveWorkbook.Path & "\"
    End If
    
    '�t�H���_�I���E�B���h�E��\��
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        '�����t�H���_�ݒ�
        .InitialFileName = strDefaultPath
        '�_�C�A���O�^�C�g���̐ݒ�
        .Title = "�t�H���_��I�����Ă�������"
        
        '�����I���֎~
        .AllowMultiSelect = False
        
        If .Show = True Then
            '�I�����ꂽ�ꍇ�A�t�H���_�p�X
            varFolderPath = .SelectedItems(1)
        Else
            '�I������Ȃ������i�L�����Z�����j�ꍇ�AFalse
            varFolderPath = False
        End If
    End With
    
    '��L�_�C�A���O�ɂ��A����ɃJ�����g�f�B���N�g�����ύX�����̂ŁA�f�t�H���g�p�X�܂��̓u�b�N�̃p�X�ɕύX����
On Error GoTo Catch
    Call ChDir(strDefaultPath)
    GoTo Finally
Catch:
    Dim resetPath As String
    resetPath = ActiveWorkbook.Path
    If resetPath <> Empty Then    ' ���A���A���ۑ��̃u�b�N�̏ꍇ�͒��߂�
        Call ChDir(resetPath)
    End If
Finally:
On Error GoTo 0
    
    GetSelectFolderPath = varFolderPath
End Function

'*****************************************************************************
' �t�H���_���݃`�F�b�N
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
' �t�H���_�폜
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
    
