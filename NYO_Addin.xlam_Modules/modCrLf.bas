Attribute VB_Name = "modCrLf"
Option Explicit

'*****************************************************************************
' �I��͈͂̊e�Z���̖����ɉ��s��ǉ�����i���ɂ���Ή������Ȃ��j
'*****************************************************************************
Public Sub AddCrLf()
    Call EditCrLf(True)
End Sub

'*****************************************************************************
' �I��͈͂̊e�Z���̖����̉��s���폜����i���ɂȂ���Ή������Ȃ��j
'*****************************************************************************
Public Sub RemoveCrLf()
    Call EditCrLf(False)
End Sub

'���s����
Private Sub EditCrLf(ByVal flgAdd As Boolean)
    Dim rngCell As Range
    
    'UsedRange�ŏI�Z���ʒu���擾
    Dim rngLastCell As Range
    Set rngLastCell = modCommon.GetLastCell(ActiveSheet)
    Dim lngLastRow As Long
    lngLastRow = rngLastCell.Row
    Dim lngLastCol As Long
    lngLastCol = rngLastCell.Column
    
    For Each rngCell In Selection
        
        'UsedRange�𒴂����璆�f
        If rngCell.Row > lngLastRow Or _
           rngCell.Column > lngLastCol Then
            Call rngCell.Activate
            Call modMessage.InfoMessage("UsedRange�𒴂����̂Œ��f���܂����B")
            Call rngCell.Activate
            Exit For
        End If
        
        '�����Z���̏ꍇ�A����Z���ȊO�̓X�L�b�v
        If rngCell.MergeArea(1, 1).Address <> rngCell.Address Then
            GoTo NEXT_CELL
        End If
        
        '�Z���̒l���󕶎��Ȃ�X�L�b�v
        If rngCell.Text = "" Then
            GoTo NEXT_CELL
        End If
        
        '������1����
        Dim lastChar As String
        lastChar = Right(rngCell.Value, 1)
        
        If flgAdd Then
            '�ǉ����[�h�F�����ɉ��s��ǉ�����i���ɂ���Ή������Ȃ��j
            If lastChar <> vbLf Then
                rngCell.Value = rngCell.Value + vbLf
            End If
        Else
            '�폜���[�h�F�����̉��s���폜����i���ɂȂ���Ή������Ȃ��j
            If lastChar = vbLf Then
                rngCell.Value = Left(rngCell.Value, Len(rngCell.Value) - 1)
            End If
        End If

NEXT_CELL:
    Next

End Sub
