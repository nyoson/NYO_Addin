Attribute VB_Name = "modA1"
Option Explicit

'*****************************************************************************
'[ �֐��� ]�@ActivateAllA1Cell
'[ �T  �v ]�@���ׂẴV�[�g��A1�Z����I����Ԃɂ���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub ActivateAllA1Cell()
    
    '���߂̕\���V�[�g
    Dim shtFirstVisibleSheet As Worksheet
    Set shtFirstVisibleSheet = Nothing
    
    Dim shtCurSheet As Worksheet
    For Each shtCurSheet In ActiveWorkbook.Worksheets
        
        '��\���V�[�g�̓X�L�b�v����
        If shtCurSheet.Visible <> xlSheetVisible Then GoTo NEXT_SHEET
            
        '���߂̕\���V�[�g�̃V�[�g�I�u�W�F�N�g���o���Ă���
        If shtFirstVisibleSheet Is Nothing Then
            Set shtFirstVisibleSheet = shtCurSheet
        End If
        
        '�V�[�g���A�N�e�B�u�ɂ���
        shtCurSheet.Activate
        
        '���Z�����ň�ԍ���Z����I����Ԃɂ���(A1�Z���Ƃ͌���Ȃ�)
        '���E�B���h�E�g�Œ莞�ȂǁACtrl+Home�̌��ʂƈقȂ�P�[�X����
        shtCurSheet.Cells.SpecialCells(xlVisible).Item(1).Activate
        
        '�X�N���[���ʒu����ԍ���ɂ���
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        
NEXT_SHEET:
    Next
    
    '�O�̂���null�`�F�b�N
    If shtFirstVisibleSheet Is Nothing Then Exit Sub
    
    '�ŏ��̕\���V�[�g���A�N�e�B�u�ɂ��ďI������
    Call shtFirstVisibleSheet.Activate
    
End Sub

