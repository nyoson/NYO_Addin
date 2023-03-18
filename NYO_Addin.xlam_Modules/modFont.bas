Attribute VB_Name = "modFont"
Option Explicit

'*****************************************************************************
' �I��͈͂̕�����Ԏ��܂��̓f�t�H���g�ɂ���
'*****************************************************************************
Public Sub SwitchRed()
    Dim fnt As Font
    Set fnt = Selection.Font
    
    If fnt.Color <> RGB(255, 0, 0) Then
        fnt.Color = RGB(255, 0, 0)
    Else
        fnt.ColorIndex = xlAutomatic
    End If
    
    'Debug.Print "SwitchRed()"
    
End Sub

'*****************************************************************************
' �I��͈͂̕����������N���A����
'*****************************************************************************
Public Sub ClearFont()
    
    '' �Z���ȊO�͔�Ή�
    'If Not TypeOf Selection Is Range Then
    '    Call modMessage.ErrorMessage("�Z����I�����ĉ�����")
    '    Exit Sub
    'End If
    '����Ή��Ȃ̂̓t�H���g��ނ݂̂Ȃ̂ŁA�R�����g�A�E�g
    
    Dim fnt As Font
    Set fnt = Selection.Font
    
    fnt.ColorIndex = xlAutomatic
    fnt.FontStyle = "�W��"
    'fnt.Bold = False       'FontStyle = "�W��" �ɂ��A�N���A�����̂ŕs�v
    'fnt.Italic = False     'FontStyle = "�W��" �ɂ��A�N���A�����̂ŕs�v
    fnt.Name = Application.StandardFont '[�I�v�V����]�̕W���t�H���g���̗p����
    'fnt.OutlineFont        'Windows�ł͖���
    'fnt.Shadow             'Windows�ł͖���
    fnt.Size = Application.StandardFontSize '[�I�v�V����]�̕W���t�H���g�T�C�Y���̗p����
    fnt.Strikethrough = False
    fnt.Subscript = False
    fnt.Superscript = False
    fnt.Underline = XlUnderlineStyle.xlUnderlineStyleNone
    
End Sub

'*****************************************************************************
' �n�C�p�[�����N�X�^�C���ύX
' �i�n�C�p�[�����N�������̃t�H���g���ƃT�C�Y���A�A�N�e�B�u�Z���Ɠ����ɂ���j
'*****************************************************************************
Public Sub SetHyperLinkStyle()
    
    Dim rngCell As Range
    Set rngCell = ActiveCell
    
    Dim styleHyperLink As Style
    Set styleHyperLink = ActiveWorkbook.Styles("Hyperlink")
    
    styleHyperLink.IncludeFont = True
    styleHyperLink.Font.Name = rngCell.Font.Name
    styleHyperLink.Font.Size = rngCell.Font.Size
    
End Sub

