VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchNext 
   Caption         =   "��������"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   OleObjectBlob   =   "frmSearchNext.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSearchNext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'�L�����Z���t���O
Public isCancel As Boolean

'�t�H�[�����[�h��
Private Sub UserForm_Initialize()
    isCancel = True
End Sub

'�����{�^��������
Private Sub btnSearch_Click()
    isCancel = False
    Call Me.Hide
End Sub

'�L�����Z���{�^��������
Private Sub btnCancel_Click()
    Call Me.Hide
End Sub


