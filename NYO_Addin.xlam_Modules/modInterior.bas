Attribute VB_Name = "modInterior"
Option Explicit

'*****************************************************************************
' �I���Z���̔w�i�F�����F�܂��͂Ȃ��ɂ���
'*****************************************************************************
Public Sub SwitchInteriorYellow()
    
    ' �Z���ȊO�͔�Ή�
    ' ���I�[�g�V�F�C�v�̏ꍇ�A���ɉ��F�̎��ɁA�w�i�F�Ȃ�(����)�ɂ��邩���ɂ��邩�����Ȃ���
    If Not TypeOf Selection Is Range Then
        Call modMessage.ErrorMessage("�Z����I�����ĉ�����")
        Exit Sub
    End If
    
    Dim itr As Interior
    Set itr = Selection.Interior
    
    If itr.Color <> RGB(255, 255, 0) Then
        itr.Color = RGB(255, 255, 0)
    Else
        '���ɉ��F�̏ꍇ�́A�w�i�F�Ȃ��ɂ���
        itr.ColorIndex = XlColorIndex.xlColorIndexNone
    End If
    
End Sub

