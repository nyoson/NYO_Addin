VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInputDate 
   Caption         =   "���t���͕⏕"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   OleObjectBlob   =   "frmInputDate.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmInputDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' �y�萔�z
'=============================================================================
' ���̓L�[��ID
Private Const KeyPAGE_UP As Integer = 33    ' PageUp�L�[
Private Const KeyPAGE_DOWN As Integer = 34  ' PageDown�L�[
Private Const KeyLEFT As Integer = 37       ' Left(��)�L�[
Private Const KeyUP As Integer = 38         ' Up(��)�L�[
Private Const KeyRIGHT As Integer = 39      ' Right(��)�L�[
Private Const KeyDOWN As Integer = 40       ' Down(��)�L�[
Private Const KeyESC As Integer = 27        ' �G�X�P�[�v(Esc)�L�[

'=============================================================================
' �y�ϐ��z
'=============================================================================
Private datInputDate As Date                ' ���͒l(���t)
Private isInputCancel As Boolean            ' ���̓L�����Z���L��

'=============================================================================
' �y�C�x���g�z
'=============================================================================

' �t�H�[��������
Private Sub UserForm_Initialize()

    isInputCancel = True
    cmdOk.SetFocus
    
End Sub

' �L�[������
Private Sub cmdOk_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    'Debug.Print KeyCode.Value
    
    ' �\���l��ϐ��ɐݒ�
    Call GetTextTargetDate
    
    Select Case KeyCode
        Case KeyLEFT
            ReduceDay1
        Case KeyUP
            ReduceWeek1
        Case KeyRIGHT
            AddDay1
        Case KeyDOWN
            AddWeek1
            
        Case KeyPAGE_UP
            ' Ctl�L�[�Ɠ������͂̏ꍇ�A�N�����Z
            ' Ctl�L�[�����͂���Ă��Ȃ��ꍇ�A�������Z
            If Shift = 2 Then
                ReduceYear1
            Else
                ReduceMonth1
            End If
            
        Case KeyPAGE_DOWN
            ' Ctl�L�[�Ɠ������͂̏ꍇ�A�N�����Z
            ' Ctl�L�[�����͂���Ă��Ȃ��ꍇ�A�������Z
            If Shift = 2 Then
                AddYear1
            Else
                AddMonth1
            End If

        Case KeyESC
            ' �G�X�P�[�v�L�[�������͏I��
            isInputCancel = True
            Me.Hide

        Case Else
    
    End Select

    ' �ϐ��̒l�����x���ɕ\��
    Call SetTextTargetDate

    cmdOk.SetFocus

End Sub

' OK������
'  �\�����Ă�����t����͕⏕�̕ϐ��ɐݒ�
Private Sub cmdOk_Click()
    Call GetTextTargetDate
    isInputCancel = False
    Me.Hide
End Sub

'=============================================================================
'�yPublic�֐��z
'=============================================================================

'*****************************************************************************
'[ �֐��� ]�@GetInputDate
'[ �T  �v ]�@���͒l(���t�^)��Ԃ�
'[ ��  �� ]�@-
'[ �߂�l ]  ���͒l(���t�^)
'*****************************************************************************
Public Function GetInputDate()
    GetInputDate = datInputDate
End Function

'*****************************************************************************
'[ �֐��� ]�@SetInputDate
'[ �T  �v ]�@���t����͒l�ɐݒ�A�\������
'[ ��  �� ]�@���t�^�̃f�[�^
'[ �߂�l ]�@-
'*****************************************************************************
Public Function SetInputDate(ByVal datTargetDate As Date)
    datInputDate = datTargetDate
    
    ' �ϐ��̒l�����x���ɕ\��
    Call SetTextTargetDate
End Function

'*****************************************************************************
'[ �֐��� ]�@GetInputCancel
'[ �T  �v ]�@���̓L�����Z���L����Ԃ�
'[ ��  �� ]�@-
'[ �߂�l ]  ���̓L�����Z�� [�L:True ��:False]
'*****************************************************************************
Public Function GetInputCancel()
    GetInputCancel = isInputCancel
End Function

'=============================================================================
'�yPrivate�֐��z
'=============================================================================

'*****************************************************************************
'[ �֐��� ]�@GetTextTargetDate
'[ �T  �v ]�@���ݕ\���l���擾(���͒l�ɐݒ�)����
'            - �\���l���ُ�̏ꍇ�A���ݓ��t��ݒ�
'[ ��  �R ]�@�ϐ��Ƃ̕s�������Ȃ�����
'[ ��  �� ]�@-
'[ �߂�l ]�@-
'*****************************************************************************
Private Sub GetTextTargetDate()

    If IsDate(lblTargetDate.Caption) Then
        datInputDate = CDate(lblTargetDate.Caption)
    Else
        datInputDate = Date
        Call SetTextTargetDate
    End If
    
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetTextTargetDate
'[ �T  �v ]�@���͒l�Ɠ��͒l�ɑ΂���j����\������
'[ ��  �� ]�@-
'[ �߂�l ]�@-
'*****************************************************************************
Private Sub SetTextTargetDate()
    ' ���t�\��
    lblTargetDate.Caption = Format(datInputDate, "yyyy/mm/dd")
    ' �j���\��
    lblDayOfWeek.Caption = Format(datInputDate, "(aaa)")
End Sub

'*****************************************************************************
'[ �֐��� ]�@AddDay1
'[ �T  �v ]�@���͒l�ƕ\���l(���t)�ɂP�����Z����
'[ ��  �� ]�@-
'[ �߂�l ]�@-
'*****************************************************************************
Private Sub AddDay1()
    datInputDate = DateAdd("d", 1, datInputDate)
End Sub

'*****************************************************************************
'[ �֐��� ]�@ReduceDay1
'[ �T  �v ]�@���͒l�ƕ\���l(���t)����P�����Z����
'[ ��  �� ]�@-
'[ �߂�l ]�@-
'*****************************************************************************
Private Sub ReduceDay1()
    datInputDate = DateAdd("d", -1, datInputDate)
End Sub

'*****************************************************************************
'[ �֐��� ]�@AddWeek1
'[ �T  �v ]�@���͒l�ƕ\���l(���t)�ɂP�T��(�V��)���Z����
'[ ��  �� ]�@-
'[ �߂�l ]�@-
'*****************************************************************************
Private Sub AddWeek1()
    datInputDate = DateAdd("d", 7, datInputDate)
End Sub

'*****************************************************************************
'[ �֐��� ]�@ReduceWeek1
'[ �T  �v ]�@���͒l�ƕ\���l(���t)����P�T��(�V��)���Z����
'[ ��  �� ]�@-
'[ �߂�l ]�@-
'*****************************************************************************
Private Sub ReduceWeek1()
    datInputDate = DateAdd("d", -7, datInputDate)
End Sub

'*****************************************************************************
'[ �֐��� ]�@AddMonth1
'[ �T  �v ]�@���͒l�ƕ\���l(���t)�ɂP�������Z����
'[ ��  �� ]�@-
'[ �߂�l ]�@-
'*****************************************************************************
Private Sub AddMonth1()
    datInputDate = DateAdd("m", 1, datInputDate)
End Sub

'*****************************************************************************
'[ �֐��� ]�@ReduceMonth1
'[ �T  �v ]�@���͒l�ƕ\���l(���t)����P�������Z����
'[ ��  �� ]�@-
'[ �߂�l ]�@-
'*****************************************************************************
Private Sub ReduceMonth1()
    datInputDate = DateAdd("m", -1, datInputDate)
End Sub

'*****************************************************************************
'[ �֐��� ]�@AddYear1
'[ �T  �v ]�@���͒l�ƕ\���l(���t)�ɂP�N���Z����
'[ ��  �� ]�@-
'[ �߂�l ]�@-
'*****************************************************************************
Private Sub AddYear1()
    datInputDate = DateAdd("yyyy", 1, datInputDate)
End Sub

'*****************************************************************************
'[ �֐��� ]�@ReduceYear1
'[ �T  �v ]�@���͒l�ƕ\���l(���t)����P�N���Z����
'[ ��  �� ]�@-
'[ �߂�l ]�@-
'*****************************************************************************
Private Sub ReduceYear1()
    datInputDate = DateAdd("yyyy", -1, datInputDate)
End Sub

