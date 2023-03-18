Attribute VB_Name = "modCond"
Option Explicit

'�ȉ��̃T�C�g����p�N���������t�������𐮗���������}�N��
'https://excel-ubara.com/excelvba5/EXCELVBA262.html

'�����t���������i�[����\����
Private Type tFormat
    AppliesTo As String '�K�p�͈�
    Formula1 As String '����1
    Formula2 As String '����2
    Operator As String '���Z�q
    NumberFormat As String '�\���`��
    FontBold As String '����
    FontColor As String '�����F
    InteriorColor As String '�h��Ԃ��F
    '�ǉ����肵�����v���p�e�B�͂����ɒǉ�
End Type

'�A�N�e�B�u�V�[�g�̏����t�������𐮗�����
Public Sub �A�N�e�B�u�V�[�g�̏����t�������𐮗�����()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
  
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Call UnionFormatConditions(ws)
  
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'�u�b�N�S�ẴV�[�g�̏����t�������𐮗���������
Public Sub �u�b�N�S�ẴV�[�g�̏����t�������𐮗���������()
    Dim FileName As Variant
    Dim Wb As Workbook
    Dim ws As Worksheet
  
    FileName = Application.GetOpenFilename(FileFilter:="Excel�t�@�C��, *.xls*")
    If FileName = False Then
        Exit Sub
    End If
  
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
  
    Set Wb = Workbooks.Open(FileName:=FileName, UpdateLinks:=0, ReadOnly:=True)
  
    For Each ws In Wb.Worksheets
        Call UnionFormatConditions(ws)
    Next
  
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
  
    FileName = Application.GetSaveAsFilename(InitialFileName:=Wb.Name, _
                                             FileFilter:="Excel�t�@�C��,*.xls*")
    If FileName = False Then
        Exit Sub
    End If
    Wb.SaveAs FileName
    Wb.Close SaveChanges:=True
End Sub

'�����t�������𐮗���������
Private Sub UnionFormatConditions(ByVal ws As Worksheet, _
                                 Optional ByVal NewName As String = "")
    '�����t���������i�[����\���̔z��
    Dim fAry() As tFormat
  
    '�����t�������������ꍇ�͏I��
    If ws.Cells.FormatConditions.Count = 0 Then Exit Sub
   
    '�I�v�V�����ɂ�茳�V�[�g���R�s�[
    If NewName <> "" Then
        ws.Copy After:=ws
        Set ws = ActiveSheet
        ws.Name = NewName '�V�[�g���̃`�F�b�N�͏ȗ����Ă��܂��B
    End If
  
    '�����t���������\���̔z��֊i�[
    Call SetFormatToType(fAry, ws)
  
    '��������t�������̌����F�z����ŃZ���͈͎w�蕶���������
    Call JoinAppliesTo(fAry, ws)
  
    '�����t�������̓����F�z�����Applies��FormatCondition�ɓK�p
    Call ModifyApplies(fAry, ws)
End Sub

'�����t���������\���̔z��֊i�[
Private Sub SetFormatToType(ByRef fAry() As tFormat, _
                            ByVal ws As Worksheet)
    Dim i As Long
    Dim fObj As FormatCondition
    On Error Resume Next '.Formula2���擾�ł��Ȃ��ꍇ�̑Ώ�
  
    ReDim fAry(ws.Cells.FormatConditions.Count)
    For i = 1 To ws.Cells.FormatConditions.Count
        Set fObj = ws.Cells.FormatConditions(i)
        fAry(i).AppliesTo = fObj.AppliesTo.Address
        fAry(i).Formula1 = fObj.Formula1
        fAry(i).Formula2 = fObj.Formula2
        fAry(i).Operator = fObj.Operator
        fAry(i).NumberFormat = fObj.NumberFormat
        fAry(i).FontBold = fObj.Font.Bold
        fAry(i).FontColor = fObj.Font.Color
        fAry(i).InteriorColor = fObj.Interior.Color
        '�ǉ����肵�����v���p�e�B�͂����ɒǉ�
    
        '�����G���[�̏����t�������͍폜������
        If isErrorFormula(fAry(i).Formula1) Or _
           isErrorFormula(fAry(i).Formula1) Then
            fAry(i).AppliesTo = ""
        End If
    Next
End Sub

'�����t�������̐����G���[����
Private Function isErrorFormula(ByVal sFormula As String) As Boolean
    If IsError(Evaluate(sFormula)) Then
        isErrorFormula = True
    Else
        isErrorFormula = False
    End If
End Function

'��������t�������̌����F�z����ŃZ���͈͎w�蕶���������
Private Sub JoinAppliesTo(ByRef fAry() As tFormat, _
                           ByVal ws As Worksheet)
    Dim i1 As Long, i2 As Long
    For i1 = 1 To UBound(fAry)
        For i2 = 1 To i1 - 1
            '�v�Z��1,2�A�����F�A�h��Ԃ��̈�v����
            If isMatchFormat(fAry(i1), fAry(i2), ws) Then
                fAry(i2).AppliesTo = Union(Range(fAry(i2).AppliesTo), _
                                           Range(fAry(i1).AppliesTo)).Address
                fAry(i1).AppliesTo = ""
                Exit For
            End If
        Next
    Next
End Sub

'�v�Z��1,2�A���Z�q�A�����F�A�h��Ԃ��̈�v����
Private Function isMatchFormat(ByRef fAry1 As tFormat, _
                               ByRef fAry2 As tFormat, _
                               ByVal ws As Worksheet) As Boolean
    If fAry1.AppliesTo = "" Or _
       fAry2.AppliesTo = "" Then
        Exit Function
    End If
  
    Dim sFormula1 As String, sFormula2 As String
    isMatchFormat = True
  
    '�v�Z��1
    sFormula1 = ToR1C1(fAry1.Formula1, fAry1.AppliesTo)
    sFormula2 = ToR1C1(fAry2.Formula1, fAry2.AppliesTo)
    If sFormula1 <> sFormula2 Then isMatchFormat = False
  
    '�v�Z��2
    sFormula1 = ToR1C1(fAry1.Formula2, fAry1.AppliesTo)
    sFormula2 = ToR1C1(fAry2.Formula2, fAry2.AppliesTo)
    If sFormula1 <> sFormula2 Then isMatchFormat = False
  
    '���Z�q
    If fAry1.Operator <> fAry2.Operator Then isMatchFormat = False
  
    '�\���`��
    If fAry1.NumberFormat <> fAry2.NumberFormat Then isMatchFormat = False
  
    '����
    If fAry1.FontBold <> fAry2.FontBold Then isMatchFormat = False
  
    '�����F
    If fAry1.FontColor <> fAry2.FontColor Then isMatchFormat = False
  
    '�h��Ԃ�
    If fAry1.InteriorColor <> fAry2.InteriorColor Then isMatchFormat = False
  
    '�ǉ����肵�����v���p�e�B�͂����ɒǉ�
End Function

'A1�`����R1C1�`���ɕϊ�
Private Function ToR1C1(ByVal sFormula As String, _
                              ByVal sAppliesTo As String)
    If sFormula = "" Then Exit Function
    Dim rng As Range
    Set rng = Range(sAppliesTo)
    ToR1C1 = Application.ConvertFormula(sFormula, xlA1, xlR1C1, , rng.Item(1))
End Function

'�����t�������̓����F�z�����Applies��FormatCondition�ɓK�p
Private Sub ModifyApplies(ByRef fAry() As tFormat, _
                          ByVal ws As Worksheet)
    Dim i As Long
    Dim fObj As Object
    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        Set fObj = ws.Cells.FormatConditions(i)
        If fAry(i).AppliesTo = "" Then
            fObj.Delete
        Else
            If fObj.AppliesTo.Address <> ws.Range(fAry(i).AppliesTo).Address Then
                fObj.ModifyAppliesToRange ws.Range(fAry(i).AppliesTo)
            End If
        End If
    Next
End Sub

