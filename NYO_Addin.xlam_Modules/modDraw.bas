Attribute VB_Name = "modDraw"
Option Explicit

'*****************************************************************************
' �~���`�悷��
'*****************************************************************************
Public Sub DrawX()
    
    Dim rngCur1 As Range
    Dim rngCur2 As Range
    
    Dim shpList As Shapes
    Dim shpLines(1 To 2) As Shape
    Dim varLineNames(1 To 2) As Variant
    
    Dim i As Long
    
    '�Z���ȊO��I�����Ă���ꍇ�́A�I������
    ActiveCell.Activate
    
    '�I���Z���͈͂��擾(A1:B2�̏ꍇ�́AA1��C3���擾)
    Set rngCur1 = Selection.Areas(1)
    Set rngCur2 = rngCur1.Offset(rngCur1.Rows.Count, rngCur1.Columns.Count)
    Set rngCur1 = rngCur1.Cells(1, 1)
    
    '�I��͈͂̍���Z���݂̂�I��������
    rngCur1.Select
    
    '�V�F�C�v�I�u�W�F�N�g���X�g
    Set shpList = ActiveSheet.Shapes
    
    '���P�i���ォ��E���j���쐬
    Set shpLines(1) = shpList.AddLine( _
                            rngCur1.Left, rngCur1.Top, _
                            rngCur2.Left, rngCur2.Top)
    '���Q�i�E�ォ�獶���j���쐬
    Set shpLines(2) = shpList.AddLine( _
                            rngCur2.Left, rngCur1.Top, _
                            rngCur1.Left, rngCur2.Top)
    
    '���P�����Q���A2pt�̐Ԑ��ɂ���
    For i = 1 To 2
        shpLines(i).Line.ForeColor.RGB = vbRed
        shpLines(i).Line.Weight = 2
        
        ' ���łɃI�u�W�F�N�g���̔z���variant�̔z��Ƃ��č쐬
        varLineNames(i) = shpLines(i).Name
    Next
    
    '�Q�{�̐����O���[�v�����A�I����Ԃɂ��ďI���
    shpList.Range(varLineNames).Group.Select
    
End Sub

'*****************************************************************************
' �Ԙg��`�悷��
'*****************************************************************************
Public Sub DrawRedFrame()
    
    Dim rngCur1 As Range
    Dim rngCur2 As Range
    
    Dim shpList As Shapes
    Dim shpFrame As Shape
    
    '�Z���ȊO��I�����Ă���ꍇ�́A�I������
    ActiveCell.Activate
    
    '�I���Z���͈͂��擾(A1:B2�̏ꍇ�́AA1��C3���擾)
    Set rngCur1 = Selection.Areas(1)
    Set rngCur2 = rngCur1.Offset(rngCur1.Rows.Count, rngCur1.Columns.Count)
    Set rngCur1 = rngCur1.Cells(1, 1)
    
    '�I��͈͂̍���Z���݂̂�I��������
    rngCur1.Select
    
    '�V�F�C�v�I�u�W�F�N�g���X�g
    Set shpList = ActiveSheet.Shapes
    
    '�l�p��`��
    Set shpFrame = shpList.AddShape(Type:=msoShapeRectangle, _
                                      Left:=rngCur1.Left, _
                                      Top:=rngCur1.Top, _
                                      Width:=Abs(rngCur2.Left - rngCur1.Left), _
                                      Height:=Abs(rngCur2.Top - rngCur1.Top) _
                                      )
    
    '�h��Ԃ��Ȃ���2pt�̐Ԑ��ɂ���
    shpFrame.Fill.Visible = msoFalse
    shpFrame.Line.ForeColor.RGB = vbRed
    shpFrame.Line.Weight = 2
    
    '�I����Ԃɂ��ďI���
    shpFrame.Select
    
End Sub

'*****************************************************************************
' 50%�Ƀ��T�C�Y����
'*****************************************************************************
Public Sub SetZoomRate_50per()
    Call SetZoomRate(50)
End Sub

'*****************************************************************************
' ���T�C�Y����
'*****************************************************************************
Private Sub SetZoomRate(ByVal rate As Double)
    
    ' �Z�����I������Ă���ꍇ�́A�������Ȃ�
    If TypeOf Selection Is Range Then Exit Sub
    
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    
    Dim idxMax As Integer
    idxMax = sr.Count
    If idxMax > 0 Then
        Dim idx As Integer
        For idx = 1 To idxMax
            Dim shp As Shape
            Set shp = sr(idx)
            
            ' �c����Œ����U�������Ă��烊�T�C�Y
            shp.LockAspectRatio = msoFalse
            Call shp.ScaleHeight(rate / 100, msoTrue)
            Call shp.ScaleWidth(rate / 100, msoTrue)
            shp.LockAspectRatio = msoTrue
            
        Next
    End If
End Sub

