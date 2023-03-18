Attribute VB_Name = "modSQL"
Option Explicit

'*****************************************************************************
'[ �֐��� ] PasteSQLGrid
'[ �T  �v ] SQL Management Studio �Ȃǂ̃O���b�h�f�[�^�𐮌`���ē\��t����
'[ ��  �� ] �Ȃ�
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub PasteSQLGrid()
    
    If Application.CutCopyMode <> 0 Then
        Call modMessage.ErrorMessage("Excel����̓\��t���͏o���܂���B")
        Exit Sub
    End If
    
On Error GoTo Catch
    
    '���C������
    PasteSQLGridMain
    
    GoTo Finally
    
Catch:
    Call modMessage.ErrorMessage("���s���܂����B", Err)
    
Finally:
    Exit Sub
    
End Sub

'���C������
Private Sub PasteSQLGridMain()
    
    '���݂̃V�[�g��Ώۂɂ���
    Dim shtTarget As Worksheet
    Set shtTarget = ActiveSheet
    
    '�܂��\��t��
    shtTarget.Paste
    
    '�\��t���͈͂��擾
    Dim rngTable As Range
    Set rngTable = Selection
    
    '�\���`�����u������v�ɐݒ�
    rngTable.NumberFormatLocal = "@"
    
    '�ēx�\��t��
    shtTarget.Paste
    
    '�i�q��Ɍr��������
    With rngTable
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    '�擪�s�̔w�i�F��ݒ�
    With rngTable.Rows(1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(204, 255, 204) 'CCFFCC: ����
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Trim
    Dim rngCell As Range
    For Each rngCell In rngTable
        Dim strValueOrg As String
        Dim strValue As String
        
        strValueOrg = rngCell.Text
        strValue = Trim(strValueOrg)
        
        If strValue <> strValueOrg Then
            rngCell.Value = strValue
        End If
        
    Next
    
    '�񕝂���������
    rngTable.EntireColumn.AutoFit

End Sub

