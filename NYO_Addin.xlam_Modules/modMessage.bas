Attribute VB_Name = "modMessage"
Option Explicit

'�����W���[����Public�֐����u�}�N���̎��s�v�ɕ\�����Ȃ��悤�ɂ���
Option Private Module

'�����b�Z�[�W�֘A���W���[����

'##########
' �֐�
'##########

'*****************************************************************************
'[ �֐��� ] ErrorMessage
'[ �T  �v ] �G���[���b�Z�[�W�_�C�A���O��\������
'[ ��  �� ] ���b�Z�[�W
'           �G���[���i�ȗ��j
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub ErrorMessage(ByVal message As String, Optional ByVal errInfo As ErrObject = Nothing)
    Call CommonMessage(message, vbOKOnly + vbExclamation, errInfo)
End Sub

'*****************************************************************************
'[ �֐��� ] SystemErrorMessage
'[ �T  �v ] �V�X�e���G���[���b�Z�[�W�_�C�A���O��\������
'[ ��  �� ] ���b�Z�[�W
'           �G���[���i�ȗ��j
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub SystemErrorMessage(ByVal message As String, _
                                Optional ByVal errInfo As ErrObject = Nothing)
    Call CommonMessage(message, vbOKOnly + vbCritical, errInfo)
End Sub

'*****************************************************************************
'[ �֐��� ] InfoMessage
'[ �T  �v ] ��񃁃b�Z�[�W�_�C�A���O��\������
'[ ��  �� ] ���b�Z�[�W
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub InfoMessage(ByVal message As String)
    Call CommonMessage(message, vbOKOnly + vbInformation)
End Sub

'*****************************************************************************
'[ �֐��� ] ConfirmMessage
'[ �T  �v ] �m�F���b�Z�[�W�_�C�A���O��\������
'[ ��  �� ] ���b�Z�[�W
'           ���b�Z�[�W�{�b�N�X�X�^�C���i�ȗ���YesNo�j
'[ �߂�l ] ���ʃR�[�h(vbYes,vbNo,...)
'*****************************************************************************
Public Function ConfirmMessage(ByVal message As String, _
            Optional ByVal opt As VbMsgBoxStyle = VbMsgBoxStyle.vbYesNo) As VbMsgBoxResult
    ConfirmMessage = CommonMessage(message, opt + vbQuestion)
End Function

'*****************************************************************************
'[ �֐��� ] CommonMessage
'[ �T  �v ] ���b�Z�[�W�_�C�A���O��\������
'[ ��  �� ] ���b�Z�[�W
'           ���b�Z�[�W�{�b�N�X�X�^�C��
'           �G���[���i�ȗ��j
'[ �߂�l ] ���ʃR�[�h(vbYes,vbNo)
'*****************************************************************************
Public Function CommonMessage(ByVal message As String, _
                                ByVal buttons As VbMsgBoxStyle, _
                                Optional ByVal errInfo As ErrObject = Nothing) As VbMsgBoxResult
    Dim messageText As String
    
    If errInfo Is Nothing Then
        messageText = message & "�@�@"
        '                       �����b�Z�[�W�_�C�A���O�̉E�]���������󂯂�
    Else
        messageText = message & vbCrLf _
                & FormatErrorInfo(errInfo)
        Call DebugPrintErr(errInfo)
    End If
    
    Dim bakScreenUpdating As Boolean
    Dim bakCursor As XlMousePointer
    
    '���݂̏�Ԃ�ޔ�
    bakScreenUpdating = Application.ScreenUpdating
    bakCursor = Application.Cursor
    
    '�ꎞ�I�ɉ���
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    
    '���b�Z�[�W�\��
    If errInfo Is Nothing Then
        CommonMessage = MsgBox(messageText, buttons, C_TOOL_NAME)
    Else
        CommonMessage = MsgBox(messageText, buttons, C_TOOL_NAME, errInfo.HelpFile, errInfo.HelpContext)
    End If
    
    '��Ԃ����ɖ߂�
    Application.Cursor = bakCursor
    Application.ScreenUpdating = bakScreenUpdating
End Function


'*****************************************************************************
'[ �֐��� ] DebugPrintErr
'[ �T  �v ] �G���[�����C�~�f�B�G�C�g�E�B���h�E�Ƀ��O�o�͂���
'[ ��  �� ] �G���[���
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub DebugPrintErr(ByVal errInfo As ErrObject)
    
    If errInfo.Number = 0 Then Exit Sub
    
    Debug.Print "===="
    Debug.Print FormatErrorInfo(errInfo, True)
    Debug.Print "===="

End Sub

'*****************************************************************************
'[ �֐��� ] ErrorInfoToString
'[ �T  �v ] �G���[���𕶎���ɐ��`����
'[ ��  �� ] �G���[���
'           �ڍ׃��[�h�i�ȗ���False�j
'[ �߂�l ] �G���[���e�L�X�g
'*****************************************************************************
Private Function FormatErrorInfo(ByVal errInfo As ErrObject, Optional ByVal detailMode As Boolean = False)
    
    Dim errText As String
    errText = "���s���G���['" & CStr(Err.Number) & "':"
    
    errText = errText & vbCrLf & "�G���[�ԍ��F0x" & Hex(errInfo.Number)
    errText = errText & vbCrLf & "�G���[�ڍׁF" & errInfo.Description
    
    '�ȈՃ��[�h�̏ꍇ�͂����ŏI��
    If detailMode = False Then GoTo Finally
    
    '�ڍ׃��[�h�̏ꍇ�́A�ǉ������o��
    errText = errText & vbCrLf & "Source�F" & errInfo.Source
    errText = errText & vbCrLf & "HelpFile�F" & errInfo.HelpFile
    errText = errText & vbCrLf & "HelpContext�F" & errInfo.HelpContext
    If errInfo.LastDllError <> 0 Then
        errText = errText & vbCrLf & "LastDllError�F" & errInfo.LastDllError
    End If
    
Finally:
    FormatErrorInfo = errText
    Exit Function

End Function

