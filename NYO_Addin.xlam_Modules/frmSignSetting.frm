VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSignSetting 
   Caption         =   "�d�q��-�ݒ�"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   OleObjectBlob   =   "frmSignSetting.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSignSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'##########
' �C�x���g
'##########

'*****************************************************************************
'[ �C�x���g�� ] �t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()

    '�d�q��ݒ�擾
    Dim objSignData As SignData
    Call GetSignSetting(objSignData)
    
    '��ʂɔ��f
    txtNameTop.Text = objSignData.Name1
    txtNameBottom = objSignData.Name2
    txtSignDate = Format(objSignData.SignDate, "yy/mm/dd")
    
End Sub

'*****************************************************************************
'[ �C�x���g�� ] �W���ɐݒ� �{�^��������
'*****************************************************************************
Private Sub btnSave_Click()
    
    '��ʂ̓��͓��e����d�q��f�[�^����
    Dim objSignData As SignData
    Call MakeSignData(objSignData)
    
    '�d�q��ݒ�ۑ�
    Call SetSignSetting(objSignData)
    
End Sub

'*****************************************************************************
'[ �C�x���g�� ] �����ɉ��� �{�^��������
'*****************************************************************************
Private Sub btnTmpSign_Click()
    
    '�d�q��ݒ�擾
    Dim objSignData As SignData
    Call MakeSignData(objSignData)
    
    '�d�q����쐬
    Call MakeSign(objSignData)
    
End Sub

'*****************************************************************************
'[ �C�x���g�� ] ���� �{�^��������
'*****************************************************************************
Private Sub btnClose_Click()
    
    Call Me.Hide
    
End Sub

'##########
' �֐�
'##########

'*****************************************************************************
'[ �֐��� ] MakeSignData
'[ �T  �v ] ��ʂ̓��͓��e����d�q��f�[�^�𐶐�����
'[ ��  �� ] [Ref]�d�q��f�[�^
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Private Sub MakeSignData(ByRef objSignData As SignData)

    objSignData.Name1 = txtNameTop.Text
    objSignData.Name2 = txtNameBottom.Text
    objSignData.SignDate = CDate(txtSignDate.Text)
    
End Sub

