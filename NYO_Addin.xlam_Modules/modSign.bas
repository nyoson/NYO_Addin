Attribute VB_Name = "modSign"
Option Explicit

'##########
' �萔
'##########

Private Const SIGN_SIZE As Single = 50#
Private Const TEXT_HEIGHT As Single = 14#
Private Const DATE_WIDTH As Single = 48#
Private Const NAME_WIDTH As Single = 30#
Private Const DATE_FONT_SIZE As Double = 10#

'##########
' �\����
'##########

'�d�q��f�[�^
Public Type SignData
    Name1 As String
    Name2 As String
    SignDate As Date
End Type

'##########
' �֐�
'##########

'*****************************************************************************
'[ �֐��� ] ShowSignSetting
'[ �T  �v ] �d�q��ݒ��ʂ�\������
'[ ��  �� ] �Ȃ�
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub ShowSignSetting()
    
    Call Load(frmSignSetting)
    Call frmSignSetting.Show
    Call Unload(frmSignSetting)
    
End Sub

'*****************************************************************************
'[ �֐��� ] GetSignSetting
'[ �T  �v ] ���W�X�g������d�q��ݒ���擾����
'[ ��  �� ] [Out]�d�q��ݒ�
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub GetSignSetting(ByRef objSignData As SignData)
    
    '�d�q��ݒ�����W�X�g������擾
    objSignData.Name1 = GetSetting(C_TOOLBAR_NAME, "Sign", "Name1")
    objSignData.Name2 = GetSetting(C_TOOLBAR_NAME, "Sign", "Name2")
    
    '�{�����t���Z�b�g
    objSignData.SignDate = Date
    
End Sub

'*****************************************************************************
'[ �֐��� ] SetSignSetting
'[ �T  �v ] ���W�X�g���ɓd�q��ݒ��ۑ�����
'[ ��  �� ] �d�q��ݒ�
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub SetSignSetting(ByRef objSignData As SignData)
    
    '�d�q��ݒ�����W�X�g���ɕۑ�
    Call SaveSetting(C_TOOLBAR_NAME, "Sign", "Name1", objSignData.Name1)
    Call SaveSetting(C_TOOLBAR_NAME, "Sign", "Name2", objSignData.Name2)
    
End Sub

'*****************************************************************************
'[ �֐��� ] Sign
'[ �T  �v ] �W���ݒ�őI���Z���ɓd�q����쐬����
'[ ��  �� ] �Ȃ�
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub Sign()
    
    '�d�q��ݒ�擾
    Dim objSignData As SignData
    Call GetSignSetting(objSignData)
    
    '�I���Z���ɓd�q����쐬����
    Call MakeSign(objSignData)
    
End Sub

'*****************************************************************************
'[ �֐��� ] MakeSign
'[ �T  �v ] �w�肳�ꂽ�ݒ�őI���Z���ɓd�q����쐬����
'[ ��  �� ] �d�q��ݒ�
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub MakeSign(ByRef objSignData As SignData)
    
    '�ΏۃV�[�g
    Dim shtTarget As Worksheet
    Set shtTarget = ActiveSheet
    
    '�V�F�C�v���X�g
    Dim shpList As Shapes
    Set shpList = shtTarget.Shapes
    
    '�~
    Dim shpCircle As Shape
    Set shpCircle = shpList.AddShape(msoShapeOval, _
                                0, _
                                0, _
                                SIGN_SIZE, SIGN_SIZE)
    shpCircle.Line.ForeColor.RGB = vbRed
    shpCircle.Line.Weight = 1           '���̑����𖾎��I�Ɏw��
    'shpCircle.Fill.Visible = msoFalse   '�h��Ԃ��Ȃ��𖾎��I�Ɏw��
    '���F�œh��Ԃ�
    shpCircle.Fill.Visible = msoTrue
    shpCircle.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    '���t�e�L�X�g
    Dim shpDate As Shape
    Set shpDate = shpList.AddTextbox(msoTextOrientationHorizontal, _
                                1, _
                                SIGN_SIZE / 2 - TEXT_HEIGHT / 2, _
                                DATE_WIDTH, TEXT_HEIGHT)
    shpDate.Line.Visible = msoFalse
    shpDate.TextFrame.HorizontalAlignment = xlHAlignCenter
    shpDate.TextFrame.VerticalAlignment = xlVAlignCenter
    shpDate.TextFrame.Characters.Text = Format(objSignData.SignDate, "yy/mm/dd")
    shpDate.TextFrame.Characters.Font.Color = vbRed
    shpDate.TextFrame.Characters.Font.Size = DATE_FONT_SIZE
    shpDate.TextFrame.MarginLeft = 0
    shpDate.TextFrame.MarginRight = 0
    shpDate.TextFrame.MarginTop = 0
    shpDate.TextFrame.MarginBottom = 0
    shpDate.Fill.Visible = msoFalse
    
    '���t�̏㉺�̐�
    Dim shpLine1, shpLine2 As Shape
    Set shpLine1 = shpList.AddLine( _
                                (SIGN_SIZE - DATE_WIDTH) / 2, _
                                SIGN_SIZE / 2 - TEXT_HEIGHT / 2, _
                                SIGN_SIZE - (SIGN_SIZE - DATE_WIDTH) / 2, _
                                SIGN_SIZE / 2 - TEXT_HEIGHT / 2)
    shpLine1.Line.ForeColor.RGB = vbRed
    Set shpLine2 = shpList.AddLine( _
                                (SIGN_SIZE - DATE_WIDTH) / 2, _
                                SIGN_SIZE / 2 + TEXT_HEIGHT / 2, _
                                SIGN_SIZE - (SIGN_SIZE - DATE_WIDTH) / 2, _
                                SIGN_SIZE / 2 + TEXT_HEIGHT / 2)
    shpLine2.Line.ForeColor.RGB = vbRed
    
    '���O�̃t�H���g�T�C�Y�Z�o
    Dim intNameCharWidth As Integer '���p�������������߂�
    intNameCharWidth = WorksheetFunction.Max( _
        LenB(StrConv(objSignData.Name1, vbFromUnicode)), _
        LenB(StrConv(objSignData.Name2, vbFromUnicode)))
    Dim sglFontSize As Single
    sglFontSize = 11# * 4 / intNameCharWidth    '���p4������11.0���œK�˂��̔䗦�ŋ��߂�
    
    '���O1
    Dim shpName1 As Shape
    Set shpName1 = shpList.AddTextbox(msoTextOrientationHorizontal, _
                                (SIGN_SIZE - NAME_WIDTH) / 2, _
                                SIGN_SIZE / 2 - TEXT_HEIGHT / 2 - TEXT_HEIGHT, _
                                NAME_WIDTH, TEXT_HEIGHT)
    shpName1.Line.Visible = msoFalse
    shpName1.TextFrame.HorizontalAlignment = xlHAlignCenter
    shpName1.TextFrame.VerticalAlignment = xlVAlignCenter
    shpName1.TextFrame.Characters.Text = objSignData.Name1
    shpName1.TextFrame.Characters.Font.Color = vbRed
    shpName1.TextFrame.Characters.Font.Size = sglFontSize
    shpName1.TextFrame.MarginLeft = 0
    shpName1.TextFrame.MarginRight = 0
    shpName1.TextFrame.MarginTop = 0
    shpName1.TextFrame.MarginBottom = 0
    shpName1.Fill.Visible = msoFalse
    
    '���O2
    Dim shpName2 As Shape
    Set shpName2 = shpList.AddTextbox(msoTextOrientationHorizontal, _
                                (SIGN_SIZE - NAME_WIDTH) / 2, _
                                SIGN_SIZE / 2 + TEXT_HEIGHT / 2, _
                                NAME_WIDTH, TEXT_HEIGHT)
    shpName2.Line.Visible = msoFalse
    shpName2.TextFrame.HorizontalAlignment = xlHAlignCenter
    shpName2.TextFrame.VerticalAlignment = xlVAlignCenter
    shpName2.TextFrame.Characters.Text = objSignData.Name2
    shpName2.TextFrame.Characters.Font.Color = vbRed
    shpName2.TextFrame.Characters.Font.Size = sglFontSize
    shpName2.TextFrame.MarginLeft = 0
    shpName2.TextFrame.MarginRight = 0
    shpName2.TextFrame.MarginTop = 0
    shpName2.TextFrame.MarginBottom = 0
    shpName2.Fill.Visible = msoFalse
    
    Dim varShapeNameList(1 To 6) As Variant
    varShapeNameList(1) = shpCircle.Name
    varShapeNameList(2) = shpDate.Name
    varShapeNameList(3) = shpLine1.Name
    varShapeNameList(4) = shpLine2.Name
    varShapeNameList(5) = shpName1.Name
    varShapeNameList(6) = shpName2.Name
    
    Dim shpSign As Shape
    Set shpSign = shpList.Range(varShapeNameList).Group
    
    Call shpSign.Copy
    
    
    '�I���Z���͈̓G���A���X�g
    Dim rangeAreaList As Areas
    Set rangeAreaList = ActiveWindow.RangeSelection.Areas
    
    '�I���Z���͈̓G���A���Ƃɍ쐬
    Dim rangeArea As Range
    For Each rangeArea In rangeAreaList
        
        '�I���Z���͈̓G���A�̍ŏ��ƍŌ�̃Z��(�̉E���Z��)���擾
        '(�I���Z���͈͂̒����ɔz�u���邽��)
        Dim rngTarget1 As Range
        Set rngTarget1 = rangeArea.Item(1)
        Dim rngTarget2 As Range
        Set rngTarget2 = rangeArea.Item(rangeArea.Count).Offset(1, 1)
        
        '�d�q����摜�\��t��
        Call shtTarget.PasteSpecial(Format:="�} (GIF)")
        'ActiveCell.Offset(5, 0).Activate
        'Call shtTarget.PasteSpecial(Format:="�} (PNG)")    'GIF�Ɠ��ɍ��Ȃ�
        'ActiveCell.Offset(5, 0).Activate
        'Call shtTarget.PasteSpecial(Format:="�} (�g�����^�t�@�C��)") '�����ڃC�}�C�`
        'ActiveCell.Offset(5, 0).Activate
        'Call shtTarget.PasteSpecial(Format:="MS Office �`��I�u�W�F�N�g")   '�I�[�g�V�F�C�v�Ƃ��ăR�s�[����邽��NG
        'ActiveCell.Offset(5, 0).Activate
        'Call shtTarget.PasteSpecial(Format:="�} (JPEG)") '���߂���Ȃ�����NG
        
        '�d�q��摜
        Dim picSign As Picture
        Set picSign = Selection
        
        '�I���Z���̒����ɔz�u
        picSign.Left = ((rngTarget1.Left + rngTarget2.Left - picSign.Width) / 2)
        picSign.Top = ((rngTarget1.Top + rngTarget2.Top - picSign.Height) / 2)
        
        '�Z���ɍ��킹�Ĉړ��C�T�C�Y�ύX����ݒ�ɕύX����
        picSign.Placement = XlPlacement.xlMoveAndSize
    
    Next
    
    '�ŏI�I�ɒ��ԃf�[�^�I�u�W�F�N�g���폜
    Call shpSign.Delete
    
    '�I����ԉ���
    rngTarget1.Activate
    
End Sub
