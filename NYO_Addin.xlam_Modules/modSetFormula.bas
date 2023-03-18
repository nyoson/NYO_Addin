Attribute VB_Name = "modSetFormula"
Option Explicit

'�V�[�g������9����(��FYYYY_MMDD)���擾���鐔��
Private Const FORMULA_GET_SHEET_NAME_LAST9 As String = "RIGHT(CELL(""filename"", $A$1),9)"

'*****************************************************************************
'[ �֐��� ]�@InputSheetName
'[ �T  �v ]�@�V�[�g�����A�N�e�B�u�Z���ɃZ�b�g����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub InputSheetName()
    Select Case SelectFormulaMode()
        Case VbMsgBoxResult.vbYes
            ActiveCell.Formula = "=MID(CELL(""filename"", $A$1),SEARCH(""]"",CELL(""filename"", $A$1))+1,255)"
        Case VbMsgBoxResult.vbNo
            ActiveCell.Value = ActiveCell.Worksheet.Name
        Case Else
            
    End Select
    
End Sub

'*****************************************************************************
'[ �֐��� ]�@InputSheetName_YYYY_MMDD
'[ �T  �v ]�@�V�[�g��(YYYY_MMDD�`��)���擾���鐔�����A�N�e�B�u�Z���ɃZ�b�g����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub InputSheetName_YYYY_MMDD()
    ActiveCell.Formula = "=" & FORMULA_GET_SHEET_NAME_LAST9
End Sub

'*****************************************************************************
'[ �֐��� ]�@InputSheetName_YYYY_MMDD_AsDate
'[ �T  �v ]�@�V�[�g��(YYYY_MMDD�`��)����t�l�Ƃ��Ď擾���鐔�����A�N�e�B�u�Z���ɃZ�b�g����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub InputSheetName_YYYY_MMDD_AsDate()
    ActiveCell.Formula = "=TEXT(REPLACE(" & FORMULA_GET_SHEET_NAME_LAST9 & ",5,1,""""), ""0000!/00!/00"")-0"
End Sub

'*****************************************************************************
'[ �֐��� ]�@InputSeqNo
'[ �T  �v ]�@�A�Ԃ��擾���鐔�����A�N�e�B�u�Z���ɃZ�b�g����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub InputSeqNo()
    
    Dim rngCell As Range
    
    If TypeName(Selection) <> "Range" Then
        Call modMessage.ErrorMessage("�Z����I�����ĉ�����")
        Exit Sub
    End If
    
    Select Case SelectFormulaMode()
        Case VbMsgBoxResult.vbYes
            For Each rngCell In Selection
                '�����Z���̍���Z���ȊO�̓X�L�b�v���A�������Z�b�g
                If rngCell.Address = rngCell.MergeArea.Cells(1, 1).Address Then
                    'rngCell.Formula = "=N(INDIRECT(""R[-1]C"",FALSE))+1"
                    rngCell.Formula = "=N(OFFSET(" & rngCell.Address(False, False) & ",-1,0))+1"
                End If
            Next
            
        Case VbMsgBoxResult.vbNo
            Dim lngCnt As Long
            lngCnt = 0
            For Each rngCell In Selection
                '�����Z���̍���Z���ȊO�̓X�L�b�v���A�J�E���g�A�b�v
                If rngCell.Address = rngCell.MergeArea.Cells(1, 1).Address Then
                    lngCnt = lngCnt + 1
                    rngCell.Value = lngCnt
                End If
            Next
            
        Case Else
            
    End Select
    
End Sub

'*****************************************************************************
'[ �֐��� ]�@SelectFormulaMode
'[ �T  �v ]�@�����œ��͂��邩�m�F���b�Z�[�W��\������
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@vbYes/vbNo/vbCancel
'*****************************************************************************
Private Function SelectFormulaMode() As VbMsgBoxResult
    SelectFormulaMode = modMessage.ConfirmMessage( _
            "�����œ��͂��܂����H�i��������I������ƒl�Ƃ��ē��͂���܂��B�j", _
            vbYesNoCancel)
End Function
