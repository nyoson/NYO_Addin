Attribute VB_Name = "modRow"
Option Explicit

'##########
' �֐�
'##########

'*****************************************************************************
'[ �֐��� ]�@AddCopyRow
'[ �T  �v ]�@�I���Z���̍s���R�s�[���ĉ��ɒǉ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]  �Ȃ�
'*****************************************************************************
Public Sub AddCopyRow()
    
    Dim rngSrc As Range
    Set rngSrc = ActiveCell.EntireRow
    
    Call rngSrc.Copy
    Call rngSrc.Offset(1, 0).Insert(XlInsertShiftDirection.xlShiftDown)
    
    Application.CutCopyMode = False
    
    Call ActiveCell.Offset(1, 0).Activate
    
End Sub

'*****************************************************************************
'[ �֐��� ]�@DelRow
'[ �T  �v ]�@�I���Z���̍s���폜
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]  �Ȃ�
'*****************************************************************************
Public Sub DelRow()
    
    Dim rngSrc As Range
    Set rngSrc = ActiveWindow.RangeSelection

    Dim rngRows As Range
    Set rngRows = rngSrc.EntireRow
    
    Call rngRows.Delete(XlDirection.xlUp)
    
    Call ActiveWindow.RangeSelection.Cells(1, 1).Select
    
End Sub

'*****************************************************************************
'[ �֐��� ]�@UpRow
'[ �T  �v ]�@�I���Z���̍s����Ɉړ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]  �Ȃ�
'*****************************************************************************
Public Sub UpRow()
    
    Dim rngSrc As Range
    Set rngSrc = ActiveWindow.RangeSelection

    Dim rngRows As Range
    Set rngRows = rngSrc.EntireRow
    
    '�擪�s�̏ꍇ�͉��������I��
    If rngRows.Row <= 1 Then Exit Sub
    
    Call rngRows.Cut
    Call rngRows.Offset(-1, 0).Insert
    
    Call rngRows.Select
    
End Sub

'*****************************************************************************
'[ �֐��� ]�@DownRow
'[ �T  �v ]�@�I���Z���̍s�����Ɉړ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]  �Ȃ�
'*****************************************************************************
Public Sub DownRow()
    
    Dim rngSrc As Range
    Set rngSrc = ActiveWindow.RangeSelection

    Dim rngRows As Range
    Set rngRows = rngSrc.EntireRow
    
    '�ŏI�s�̏ꍇ�͉��������I��
    If rngRows.Row + rngRows.Rows.Count - 1 >= ActiveSheet.Rows.Count Then Exit Sub
    
    Call rngRows.Cut
    Call rngRows.Offset(rngRows.Rows.Count + 1, 0).Insert
    
    Call rngRows.Select
    
End Sub

