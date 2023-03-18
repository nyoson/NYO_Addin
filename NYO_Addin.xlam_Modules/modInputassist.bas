Attribute VB_Name = "modInputassist"
Option Explicit

'*****************************************************************************
'[ �֐��� ]�@InputAssist
'[ �T  �v ]�@���͕⏕
'[ ��  �� ]�@�ΏۃZ��
'[ �߂�l ]�@���͕⏕���ŏ������������߁A�C�x���g���L�����Z�����ׂ��ꍇ��True��Ԃ�
'*****************************************************************************
Public Function InputAssist(ByVal Target As Range) As Boolean
    
    Dim Cancel As Boolean   '�C�x���g�L�����Z���t���O
    
    Cancel = False
    
    '�����Z���̏ꍇ��1�Z����
    Dim rng As Range
    Set rng = Target.Cells(1, 1)
    
    '���������͂���Ă��邩�`�F�b�N���Ă���
    Dim isFormula As Boolean
    isFormula = False
    If Len(rng.Formula) > 1 Then
        If Left(rng.Formula, 1) = "=" Then
            isFormula = True
        End If
    End If
    
    '�ЂƂ߂̕\���`�����擾
    Dim strNumFmt As String
    strNumFmt = rng.NumberFormat
    strNumFmt = Split(strNumFmt, ";", 2, vbBinaryCompare)(0)
    '�i�Ⴆ�΁A�umm";"dd�v�̗l�ȏꍇ�ɕ��f����Ă��܂����A���邱�Ƃ͍l���ɂ����̂ł��̂܂܂Ƃ���j
    
    Select Case strNumFmt
        
        '�������͕⏕
        Case "hh:mm", "h:mm", "h:m"
            '���������͂���Ă���ꍇ�͊m�FMSG
            If isFormula = True Then
                If WarnFormula <> True Then
                    InputAssist = False
                    Exit Function
                End If
            End If

            ' �������͕⏕����
            Call InputAssistTime(rng)
            '�C�x���g�̓L�����Z������
            Cancel = True

        '���t���͕⏕
        Case "m""��""d""��""", "m/d/yyyy", "yyyy/mm/dd", "m/dd/yyyy", "mm/dd", "m/d", "m/dd"
            
            '���������͂���Ă���ꍇ�͊m�FMSG
            If isFormula = True Then
                If WarnFormula <> True Then
                    InputAssist = False
                    Exit Function
                End If
            End If

            ' ���t���͕⏕����
            Call InputAssistDate(rng)
            ' �C�x���g�̓L�����Z������
            Cancel = True
            
        Case Else
            ' �������Ȃ�

    End Select

    
    '�`�F�b�N�{�b�N�X��؂�ւ���
    If Not Cancel Then Cancel = InputAssistRotateStatus(rng, Array("��", "��"))
    If Not Cancel Then Cancel = InputAssistRotateStatus(rng, Array("��", "�~", "��"))
    
    InputAssist = Cancel

End Function

'*****************************************************************************
'[ �֐��� ]�@WarnFormula
'[ �T  �v ]�@���������͂���Ă���|�̌x�������s�m�FMSG
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]  ���sOK:True/���f:False
'*****************************************************************************
Private Function WarnFormula()
    
    If ConfirmMessage("���������͂���Ă��܂��B���͕⏕�𑱍s���Ă���낵���ł����H", vbOKCancel) = vbOK Then
        WarnFormula = True
    Else
        WarnFormula = False
    End If
    Exit Function
    
End Function

'*****************************************************************************
'[ �֐��� ]�@InputAssistRotateStatus
'[ �T  �v ]�@��ԃ��[�e�[�g���͕⏕
'[ ��  �� ]�@�ΏۃZ��, ��ԃ��X�g
'[ �߂�l ]  �C�x���g�L�����Z���t���O
'*****************************************************************************
Private Function InputAssistRotateStatus(ByRef rng As Range, ByRef statusTextArray As Variant) As Boolean
    Dim Cancel As Boolean   '�C�x���g�L�����Z���t���O
    
    Dim strText As String
    strText = rng.Text

    '��ԃ��X�g�̂����ꂩ�Ɉ�v�����玟�̏�Ԃɐ؂�ւ���
    Dim i As Long
    For i = LBound(statusTextArray) To UBound(statusTextArray)
        If statusTextArray(i) = strText Then
            Dim lngIdx As Long
            lngIdx = i - LBound(statusTextArray)
            Dim lngLen As Long
            lngLen = UBound(statusTextArray) - LBound(statusTextArray) + 1
            Dim lngNextIdx As Long
            lngNextIdx = ((lngIdx + 1) Mod lngLen)
            Dim lngNextNo As Long
            lngNextNo = lngNextIdx + LBound(statusTextArray)
            rng.Value = statusTextArray(lngNextNo)
            ' �C�x���g�̓L�����Z������
            Cancel = True
            
            Exit For    '�����I��
        End If
    Next i

    InputAssistRotateStatus = Cancel
End Function

'*****************************************************************************
'[ �֐��� ]�@InputAssistDate
'[ �T  �v ]�@���t���͕⏕
'[ ��  �� ]�@�ΏۃZ��
'[ �߂�l ]�@-
'*****************************************************************************
Public Sub InputAssistDate(ByVal Target As Range)
    Dim datTargetDate As Date
    Dim isInputCancel As Boolean
                
    '�Z���̌��̒l���擾
    datTargetDate = ConvertToDate(Target.Value)
    
    Load frmInputDate
    
    ' �t�H�[���̓��͒l�ɃZ���̒l��ݒ�
    Call frmInputDate.SetInputDate(datTargetDate)
    
    frmInputDate.Show
    
    ' �t�H�[���̓��͒l���擾
    datTargetDate = frmInputDate.GetInputDate
    ' ���̓L�����Z���L�����擾
    isInputCancel = frmInputDate.GetInputCancel
    
    Unload frmInputDate
    
    ' �t�H�[���ł̓��͊m���A�ΏۃZ���ɒl��ݒ肷��
    If isInputCancel <> True Then
        If Not IsDate(Target.Value) Then
            Call SendKeysWithRetry(CStr(datTargetDate), True, 2)
        ElseIf CDate(Target.Value) <> datTargetDate Then
            Call SendKeysWithRetry(CStr(datTargetDate), True, 2)
        End If
    End If

End Sub

'*****************************************************************************
'[ �֐��� ]�@InputAssistTime
'[ �T  �v ]�@�������͕⏕
'[ ��  �� ]�@�ΏۃZ��
'[ �߂�l ]�@-
'*****************************************************************************
Public Sub InputAssistTime(ByVal Target As Range)
    Dim varInput As Variant '���͒l
    
    '�Z���̌��̒l���擾
    varInput = ConvertToHHMM(Target.Value)
    
    '����InputBox�ō��E�L�[�ɂ��Z���ړ���h�����߂ɁAF2�L�[�𑗂��ĕҏW���[�h�ɂ���
    Call Application.SendKeys("{F2}")
    
    '�������͉�ʕ\��
    varInput = Application.InputBox("������hhmm�`��(�R��������)�œ��͂��ĉ������B", "�������͕⏕", varInput)
    '����InputBox�֐����ƁA�L�����Z�����󕶎������ʕs�\�Ȃ̂ŁAInputBox���\�b�h���g�p
    
    Call Sleep(100)
    
    '�L�����Z���{�^���������ȊO�Ȃ�A�������{
    If varInput <> False Then
        
        Select Case Len(varInput)
            Case 4:
                varInput = Format(varInput, "@@:@@")
                'Call Application.SendKeys(varInput & "{Tab}+{Tab}")
                Call SendKeysWithRetry(varInput, True, 2)
            Case 3:
                varInput = Format(varInput, "@:@@")
                'Call Application.SendKeys(varInput & "{Tab}+{Tab}")
                Call SendKeysWithRetry(varInput, True, 2)
            Case 0:
                '�N���A
                'Call Application.SendKeys("{Del}")
                Call SendKeysWithRetry("{Del}", False, 2)
            Case Else:
                '�G���[MSG
                Call modMessage.ErrorMessage("hhmm�܂���hmm�̌`���œ��͂��ĉ������B")
        End Select
        
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@ConvertToHHMM
'[ �T  �v ]�@"HHMM"�`���Ƀt�H�[�}�b�g�ϊ�
'[ ��  �� ]�@���͒l
'[ �߂�l ]  �ϊ���̕�����i�G���[���͋��Ԃ��j
'*****************************************************************************
Private Function ConvertToHHMM(ByVal varInput As Variant) As String
    
    Dim strResult As String
    
On Error GoTo Catch
    strResult = Format(varInput, "hhmm")
    GoTo Finally

Catch:
    '�t�H�[�}�b�g�G���[���͋󕶎���Ƃ���
    strResult = Empty
    
Finally:
    On Error GoTo 0
    
    ConvertToHHMM = strResult
End Function

'*****************************************************************************
'[ �֐��� ]�@ConvertToDate
'[ �T  �v ]�@���t�^�ɕϊ�
'[ ��  �� ]�@���͒l
'[ �߂�l ]  �ϊ���̓��t�i���͒l���� ���� �G���[���ɂ͌��ݓ��t��Ԃ��j
'*****************************************************************************
Private Function ConvertToDate(ByVal varInput As Variant) As Date
    
    Dim datResult As Date
    
On Error GoTo Catch
    If varInput <> Empty Then
        datResult = CDate(varInput)
    Else
        datResult = Date
    End If
    
    GoTo Finally

Catch:
    '�t�H�[�}�b�g�G���[���͌��ݓ��t�Ƃ���
    datResult = Date
    
Finally:
    On Error GoTo 0
    
    ConvertToDate = datResult
End Function

'*****************************************************************************
'[ �֐��� ]�@SendKeysWithRetry
'[ �T  �v ]�@SendKey���s���A�Z���ɓ��͂��s���B�Z���̒��g���ς���Ă��Ȃ��ꍇ�́A���g���C����
'[ ��  �� ]�@���̓L�[������
'            ActiveCell�����̏�ɗ��߂邽�߂ɁA�Ō�� {Tab}��Shift + {Tab} �𑗂邩�ǂ���
'            ���g���C��(�ȗ���=1:���g���C�Ȃ�)
'[ �߂�l ]  �Ȃ�
'*****************************************************************************
Private Sub SendKeysWithRetry(ByVal strSendKeysOrg As String, ByVal withStay As Boolean, Optional ByVal intRetry As Integer = 1)
    
    Dim rngCurCell As Range
    Set rngCurCell = ActiveCell
    
    Dim strCurVal As String
    strCurVal = rngCurCell.Text
    
    Dim strSendKeys As String
    strSendKeys = strSendKeysOrg
    
    'ActiveCell�����̏�ɗ��߂�ꍇ�́A�Ō�� {Tab}��Shift + {Tab} �𑗂�
    If withStay Then
        strSendKeys = strSendKeys & "{Tab}+{Tab}"
    End If
    
    'IME��OFF�ɂ���
    Call SetIMEMode(False)
    
    Do
        '�L�[�𑗏o
        Call Application.SendKeys(strSendKeys)
        
        '���g���C�J�E���^��0�Ȃ烊�g���C���Ȃ�
        intRetry = intRetry - 1
        If intRetry <= 0 Then Exit Do
        
        '200ms�҂�
        'Call Sleep(200)
        
        '����ActiveCell�̈ʒu���ς���Ă��܂��Ă����烊�g���C���Ȃ�
        If rngCurCell.Address <> ActiveCell.Address Then
            Debug.Print "<SendKeysWithRetry>ActiveCell Moved:" & rngCurCell.Address(False, False) & "��" & ActiveCell.Address(False, False)
            'Call rngCurCell.Activate
            Exit Do
        End If
        
        '�Z���̒l���ς�����ꍇ�͐������Ă���̂Ń��g���C���Ȃ�
        If ActiveCell.Text <> strCurVal Then Exit Do
        
        '�Z���̒l���ς���Ă��Ȃ��ꍇ�͉������̂Ń��g���C
        Debug.Print "<SendKeysWithRetry>Retry:" & intRetry & " CellValue=" & ActiveCell.Text & "��" & strSendKeysOrg
    Loop
    
End Sub

