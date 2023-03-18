Attribute VB_Name = "modSheet"
Option Explicit

'##########
' �֐�
'##########

'*****************************************************************************
'[ �֐��� ]�@SheetDispChanger
'[ �T  �v ]�@�V�[�g�̕\��/��\����؂�ւ���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SheetDispChanger()

On Error GoTo Catch
    Load frmSheetDispChanger
    frmSheetDispChanger.Show
    GoTo Finally
Catch:
    Call modMessage.ErrorMessage("��ʋN�����ɃG���[���������܂����B", Err)
Finally:
    Unload frmSheetDispChanger
    
End Sub

'*****************************************************************************
' R1C1�`���؂�ւ�
'*****************************************************************************
Public Sub SwitchR1C1()
    If Application.ReferenceStyle <> xlR1C1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub

'*****************************************************************************
' �A�E�g���C���̏W�v�s�����C�W�v�����ɂ���
'*****************************************************************************
Sub OutlineConfigAboveLeft()
    With ActiveSheet.Outline
        .SummaryRow = xlAbove
        .SummaryColumn = xlLeft
    End With
End Sub

