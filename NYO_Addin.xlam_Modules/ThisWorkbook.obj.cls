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
' �ϐ�
'##########
Public WithEvents xlApp As Excel.Application   '���u�b�N�̃C�x���g���擾���邽�߂Ɏg�p
Attribute xlApp.VB_VarHelpID = -1

Private lstReadOnlyAlerted As New Dictionary    '�ǂݎ���p�A���[�g�ς݃��X�g

'##########
' �C�x���g
'##########

'*****************************************************************************
'[ �C�x���g�� ]  Addin�C���X�g�[����
'*****************************************************************************
Private Sub Workbook_AddinInstall()
    On Error Resume Next
    ' �c�[���o�[�ǉ�
    Call modMenu.SetMenu
End Sub

'*****************************************************************************
'[ �C�x���g�� ]  Addin�A���C���X�g�[����
'*****************************************************************************
Private Sub Workbook_AddinUninstall()
    ' �c�[���o�[�폜
    Call modMenu.DelMenu
End Sub

'*****************************************************************************
'[ �C�x���g�� ]  Addin�u�b�N�I�[�v����
'*****************************************************************************
Private Sub Workbook_Open()
    '���u�b�N�̃C�x���g���擾���邽�߂�Excel�A�v���P�[�V�����I�u�W�F�N�g���擾
    Set xlApp = Excel.Application
    
    '�V���[�g�J�b�g�L�[���蓖��
    Call modMenu.RegistShortcutKey
    
    '�E�N���b�N���j���[�ǉ�
    Call modMenu.AddContextMenu
End Sub

'*****************************************************************************
'[ �C�x���g�� ]  Addin�u�b�N�ۑ���
'*****************************************************************************
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    'Addin�u�b�N�I�[�v�����Ɏ擾����Excel�A�v���P�[�V�����I�u�W�F�N�g���A
    '�}�N����ҏW����ƃN���A����邽�߁AAddin�ۑ����ɍĎ擾
    Set xlApp = Excel.Application
    
    '�Ȃ��������ŏ��������[�h�ɂȂ��Ă��܂����Ƃ�����̂ŁA�����I�ɉ���
    Call modCommon.ProcessModeEnd
End Sub

'*****************************************************************************
'[ �C�x���g�� ]  Addin�u�b�N�N���[�Y��
'*****************************************************************************
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    '�}�N����ҏW���Ă���̂ɕ��悤�Ƃ����ꍇ�AAddin�̏ꍇ�ł��x������
    If Not ThisWorkbook.Saved Then
        If ConfirmMessage(ThisWorkbook.Name & "�����ۑ��ł��B�ۑ����܂����H") = vbYes Then
            ThisWorkbook.Save
        End If
    End If
End Sub

'*****************************************************************************
'[ �C�x���g�� ]  ���u�b�N�̃V�[�g�_�u���N���b�N��
'*****************************************************************************
Private Sub xlApp_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    
    '���̃u�b�N�̃C�x���g�n���h���[�ŃL�����Z���ς݂̏ꍇ�́A�������Ȃ�
    If Cancel Then Exit Sub
    
    '���͕⏕
    '(���͕⏕�ɂď������ꂽ�ꍇ�́A�_�u���N���b�N�C�x���g���L�����Z������)
    Cancel = InputAssist(Target.Cells(1, 1))
End Sub

'*****************************************************************************
'[ �C�x���g�� ]  ���u�b�N�̃I�[�v����
'*****************************************************************************
Private Sub xlApp_WorkbookOpen(ByVal Wb As Workbook)
    
    Dim wbk As Workbook
    Set wbk = Wb
    
    '�O�̂��߁ANULL�΍�
    If lstReadOnlyAlerted Is Nothing Then
        Set lstReadOnlyAlerted = New Dictionary
    End If
        
    '�J���������ꍇ�ɁA�ǂݎ���p�A���[�g�ς݃��X�g���珜�������悤�ɂ���
    If lstReadOnlyAlerted.Exists(wbk.FullName) Then
        Call lstReadOnlyAlerted.Remove(wbk.FullName)
    End If
End Sub

'*****************************************************************************
'[ �C�x���g�� ]  ���u�b�N�̃V�[�g�ύX��
'*****************************************************************************
Private Sub xlApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    Dim wbk As Workbook
    Set wbk = Sh.Parent
    
    '�ǂݎ���p�̏ꍇ�A����̂݃A���[�g
    If wbk.ReadOnly Then
        '�O�̂��߁ANULL�΍�
        If lstReadOnlyAlerted Is Nothing Then
            Set lstReadOnlyAlerted = New Dictionary
        End If
        
        If Not lstReadOnlyAlerted.Exists(wbk.FullName) Then
            Call modMessage.ErrorMessage("���̃u�b�N�͓ǂݎ���p�ł��B�����Ӊ������B�i�ȍ~�͌x�����܂���B�j")
            Call lstReadOnlyAlerted.Add(wbk.FullName, True)
        End If
    End If

End Sub




''*****************************************************************************
''[ �C�x���g�� ]  ���u�b�N�̃N���[�Y�O�C�x���g
''*****************************************************************************
'Private Sub xlApp_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
'
'On Error GoTo Catch
'    'Excel2007�ȍ~�̏ꍇ�́A���X�o�b�N�A�b�v�@�\������̂ŁA�������Ȃ�
'    If Is2007orLater Then Exit Sub
'
'    '�����o�b�N�A�b�v
'    '(���ۑ��̏ꍇ�́A���̌�ۑ����ĕ��邩�ǂ����ɍS��炸�A�o�b�N�A�b�v���쐬����)
'    If Not wb.Saved Then
'        Call BackupBook(wb)
'    End If
'    Exit Sub
'
'Catch:
'    '�G���[���������Ă��A�������Ȃ�
'    Exit Sub
'
'End Sub
'
''*****************************************************************************
''[ �C�x���g�� ]  ���u�b�N�̕ۑ��O�C�x���g
''*****************************************************************************
'Private Sub xlApp_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    '�ۑ�����ꍇ�́A�o�b�N�A�b�v���폜���Ă���
'    Call DeleteBackupBook(wb)
'End Sub

'*****************************************************************************
'[ �C�x���g�� ]  CtrlEnter��
'*****************************************************************************
Private Sub CtrlEnter()
    '���͕⏕
    Call InputAssist(ActiveCell)
End Sub

'*****************************************************************************
'[ �C�x���g�� ]  Ctrl+Alt+";"��
'*****************************************************************************
Public Sub CtrlAltSemicolon()
    ' ���͕⏕(���t)
    Call InputAssistDate(ActiveCell)
End Sub

'*****************************************************************************
'[ �C�x���g�� ]  Ctrl+Alt+":"��
'*****************************************************************************
Public Sub CtrlAltColon()
    ' ���͕⏕(����)
    Call InputAssistTime(ActiveCell)
End Sub
