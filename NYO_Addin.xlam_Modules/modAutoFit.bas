Attribute VB_Name = "modAutoFit"
Option Explicit

'*****************************************************************************
'[ �֐��� ]�@AutoFit
'[ �T  �v ]�@�I�[�g�V�F�C�v�̎����T�C�Y�����ݒ��؂�ւ���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub AutoFit()
    ' �G���[�͖�������
    On Error Resume Next
    
    Dim idxMax As Integer
    Dim idx As Integer
    Dim flg As Boolean
    Dim button As CommandBarButton
    Set button = Application.CommandBars(C_TOOLBAR_NAME).Controls("�I�[�g�V�F�C�v�̎����T�C�Y����")
    
    ' �R�}���h�o�[�{�^���̃X�e�[�^�X�����Z�b�g���Ă���
    button.State = msoButtonUp

    ' �Z�����I������Ă���ꍇ�́A�������Ȃ�
    If TypeOf Selection Is Range Then Exit Sub
    
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    idxMax = sr.Count
    If idxMax > 0 Then
        ' �ЂƂڂ̃I�[�g�V�F�C�v�̎����T�C�Y�����t���O�𔽓]���A�I�����ꂽ���ׂẴI�[�g�V�F�C�v�ɔ��f
        flg = Not sr(1).DrawingObject.AutoSize
        For idx = 1 To idxMax
            ' �����o�����̒���������ɕς���Ă��܂��̂�h�����߂ɁA
            ' �����o�����̒������Œ�ɂ��Ă����ˌ����Ȃ��c�y�ۑ�z
            Dim calloutFmt As CalloutFormat
            Set calloutFmt = sr(idx).Callout
            calloutFmt.CustomLength calloutFmt.Length
            
            sr(idx).DrawingObject.AutoSize = flg
            
            If flg Then
                ' �e�L�X�g��}�`����͂ݏo���ĕ\������
                Dim tf As TextFrame
                Set tf = sr(idx).TextFrame
                tf.HorizontalOverflow = xlOartHorizontalOverflowOverflow '��������
                tf.VerticalOverflow = xlOartVerticalOverflowOverflow '��������
            
            End If
            
            ' �����o�����̒����������ɂ���ˁy�ۑ�z
            calloutFmt.AutomaticLength
            
        Next
        
        ' �R�}���h�o�[�{�^���̃X�e�[�^�X��ύX�˖{����SelectionChange�C�x���g�ł����{���������A���@���s���c�y�ۑ�z
        If flg Then
            button.State = msoButtonDown
        Else
            button.State = msoButtonUp
        End If
        
    End If
    
    On Error GoTo 0
End Sub

