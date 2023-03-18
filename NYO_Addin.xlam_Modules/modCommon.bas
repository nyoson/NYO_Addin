Attribute VB_Name = "modCommon"
Option Explicit

'�����W���[����Public�֐����u�}�N���̎��s�v�ɕ\�����Ȃ��悤�ɂ���
Option Private Module

'�����ʏ������W���[����

'�ȉ����Q�Ɛݒ�ɒǉ����ĉ������B
'�EMicrosoft Visual Basic for Applications Extensibility
'�iVBProject�N���X�Ȃǁj
'�EMicrosoft Scripting Runtime
'�iFileSystemObject�N���X�Ȃǁj

'##########
' �萔
'##########
Public Const C_TOOL_NAME As String = "NYO_Addin"
Public Const C_TOOLBAR_NAME As String = "NYO"

'##########
' API�֐�
'##########
' Sleep�֐�(�X���[�v����[ms])
Public Declare PtrSafe Sub Sleep Lib "KERNEL32.dll" (ByVal dwMilliseconds As Long)

' IME����
Private Declare PtrSafe Function ImmGetContext Lib "imm32.dll" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function ImmSetOpenStatus Lib "imm32.dll" (ByVal himc As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function ImmReleaseContext Lib "imm32.dll" (ByVal hWnd As Long, ByVal himc As Long) As Long

'##########
' �񋓑�
'##########
' ExcelVersion
Public Enum ExcelMajorVersion
    Ver2003 = 11
    Ver2007 = 12
    Ver2010 = 14
    Ver2013 = 15
    Ver2016 = 16
End Enum

'##########
' �ϐ�
'##########
' ExcelVersion
Private enmExcelMajorVersion As ExcelMajorVersion

'##########
' �֐�
'##########

'*****************************************************************************
'[ �֐��� ]�@ProcessModeEnd
'[ �T  �v ]�@���������[�h����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]  �Ȃ�
'*****************************************************************************
Public Sub ProcessModeEnd()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.Cursor = xlDefault
    Application.EnableEvents = True
End Sub

'*****************************************************************************
'[ �֐��� ] Is2016orLater
'[ �T  �v ] Excel2016�ȍ~���ǂ�����Ԃ�
'[ ��  �� ] �Ȃ�
'[ �߂�l ] True�FExcel2016�ȍ~�^False�F������O
'*****************************************************************************
Public Function Is2016orLater() As Boolean
    Dim intVersion As Integer
    intVersion = GetExcelMajorVersion()
    If intVersion >= ExcelMajorVersion.Ver2016 Then
        Is2016orLater = True
    Else
        Is2016orLater = False
    End If
End Function

'*****************************************************************************
'[ �֐��� ] Is2007orLater
'[ �T  �v ] Excel2007�ȍ~���ǂ�����Ԃ�
'[ ��  �� ] �Ȃ�
'[ �߂�l ] True�FExcel2007�ȍ~�^False�F������O
'*****************************************************************************
Public Function Is2007orLater() As Boolean
    Dim intVersion As Integer
    intVersion = GetExcelMajorVersion()
    If intVersion >= ExcelMajorVersion.Ver2007 Then
        Is2007orLater = True
    Else
        Is2007orLater = False
    End If
End Function

'*****************************************************************************
'[ �֐��� ] GetExcelMajorVersion
'[ �T  �v ] Excel�̓����o�[�W�������擾����
'[ ��  �� ] �Ȃ�
'[ �߂�l ] 12�FExcel2007�^14�FExcel2010�^15�FExcel2013�^16�FExcel2016�ȍ~(2019�܂�)
'           �y�Q�l�zhttps://minimashia.net/vba-excel-check/
'*****************************************************************************
Public Function GetExcelMajorVersion() As ExcelMajorVersion
    If Not enmExcelMajorVersion > 0 Then
        enmExcelMajorVersion = CInt(Split(Application.Version, ".", 2, vbBinaryCompare)(0))
    End If
    GetExcelMajorVersion = enmExcelMajorVersion
End Function


'*****************************************************************************
'[ �֐��� ] SetIMEMode
'[ �T  �v ] IME���[�h��ON/OFF��ύX����
'[ ��  �� ] True:IME-ON / False:IME-OFF
'           �ΏۃR���g���[���̃n���h��(ex: txtInput.hWnd) ���Z���Ȃǂ̏ꍇ�͏ȗ���
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Function SetIMEMode(ByVal IMEMode As Boolean, Optional ByVal hWnd As Long = 0)
    
    '�ΏۃR���g���[����hWnd�̎w�肪�Ȃ���΁AExcelApp��hWnd���擾
    If hWnd = 0 Then hWnd = Application.hWnd
    
    'IME��On
    Dim himc As Long
    himc = ImmGetContext(hWnd)
    Call ImmSetOpenStatus(himc, IIf(IMEMode, 1, 0))
    Call ImmReleaseContext(hWnd, himc)
    
End Function

'*****************************************************************************
'[ �֐��� ] CheckMacroSecurityWithMsg
'[ �T  �v ] VBA�v���W�F�N�g�ւ̃A�N�Z�X���̃`�F�b�N���s���A�G���[���̓��b�Z�[�W���\������
'[ ��  �� ] �Ȃ�
'[ �߂�l ] �A�N�Z�X��
'*****************************************************************************
Public Function CheckMacroSecurityWithMsg() As Boolean
    
    Dim ret As Boolean
    
    '�Z�L�����e�B�ݒ�`�F�b�N
    ret = CheckMacroSecurity()
    
    If Not ret Then
        Dim strMsg As String
        If Is2016orLater Then
            'Excel 2016�ȍ~
            strMsg = "���݂̃Z�L�����e�B�ݒ�ł͎��s�ł��܂���B" & vbCrLf _
                    & "Excel��[�t�@�C��]���{��->[���̑�...]->[�I�v�V����]->" & vbCrLf _
                    & "[�g���X�g �Z���^�[]�^�u->�u�g���X�g �Z���^�[�̐ݒ�...�v�{�^��->" & vbCrLf _
                    & "[�}�N���̐ݒ�]�^�u->�uVBA �v���W�F�N�g �I�u�W�F�N�g ���f���ւ̃A�N�Z�X��M������v�Ƀ`�F�b�N���ĉ������B"
        ElseIf Is2007orLater Then
            'Excel 2007�ȍ~
            strMsg = "���݂̃Z�L�����e�B�ݒ�ł͎��s�ł��܂���B" & vbCrLf _
                    & "Excel��[�t�@�C��]���{��->[�I�v�V����]->" & vbCrLf _
                    & "[�Z�L�����e�B �Z���^�[]�^�u->�u�Z�L�����e�B �Z���^�[�̐ݒ�v�{�^��->" & vbCrLf _
                    & "[�}�N���̐ݒ�]�^�u->�uVBA �v���W�F�N�g �I�u�W�F�N�g ���f���ւ̃A�N�Z�X��M������v���`�F�b�N���ĉ������B"
        Else
            'Excel 2003�ȑO
            strMsg = "���݂̃Z�L�����e�B�ݒ�ł͎��s�ł��܂���B" & vbCrLf _
                    & "Excel��[�c�[��]->[�}�N��]->[�Z�L�����e�B]->[�M���̂����锭�s��]�^�u�́A" & vbCrLf _
                    & "[Visual Basic�v���W�F�N�g�ւ̃A�N�Z�X��M������]���I���ɂ��ĉ������B"
        End If
        Call modMessage.ErrorMessage(strMsg)
    End If
    
    CheckMacroSecurityWithMsg = ret
    
End Function

'*****************************************************************************
'[ �֐��� ] CheckMacroSecurity
'[ �T  �v ] VBA�v���W�F�N�g�ւ̃A�N�Z�X���̃`�F�b�N���s��
'[ ��  �� ] �Ȃ�
'[ �߂�l ] �A�N�Z�X��
'*****************************************************************************
Public Function CheckMacroSecurity() As Boolean
    
    Dim ret As Boolean
    Dim test As Object 'VBProject
    
On Error GoTo Catch
    '�����ɁA���A�h�C����VB�R���|�[�l���g�̐����擾
    Set test = ThisWorkbook.VBProject
    ret = True
    
    GoTo Fainally
    
Catch:
    ret = False
    
Fainally:
    Set test = Nothing
    CheckMacroSecurity = ret
    
End Function

'*****************************************************************************
'[ �֐��� ] MakeDir
'[ �T  �v ] �T�u�t�H���_���܂߂ăt�H���_���쐬����
'[ ��  �� ] �쐬����t�H���_�̃t���p�X
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub MakeDir(ByVal strPath As String)
    
    'FileSystemObject
    Dim objFso As FileSystemObject
    Set objFso = New FileSystemObject
    '���ɑ��݂��Ă���΁A���������I��
    If objFso.FolderExists(strPath) Then
        Exit Sub
    End If
    
    '�e�t�H���_���Ȃ���΍쐬(�ċA�Ăяo��)
    Call MakeDir(objFso.GetParentFolderName(strPath))
    
    '�ړI�̃t�H���_���쐬
    Call objFso.CreateFolder(strPath)
    
End Sub

'*****************************************************************************
'[ �֐��� ] GetLastCell
'[ �T  �v ] �V�[�g���̍ŏI�Z�����擾
'[ ��  �� ] �ΏۃV�[�g
'[ �߂�l ] �ŏI�Z��
'*****************************************************************************
Public Function GetLastCell(ByVal shtTarget As Worksheet) As Range
    Dim rngLastCell As Range
    Set rngLastCell = shtTarget.UsedRange
    Set rngLastCell = rngLastCell.Cells(rngLastCell.Rows.Count, _
                                        rngLastCell.Columns.Count)
    Set GetLastCell = rngLastCell
End Function

'*****************************************************************************
'[ �֐��� ] ScrollTo
'[ �T  �v ] �ΏۃZ���܂ŃW�����v
'[ ��  �� ] �ΏۃZ��, �X�N���[���s�I�t�Z�b�g�i�ȗ�����-5�s�j
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Public Sub ScrollTo(ByRef rng As Range, Optional ByVal scrollRowOffset As Long = -5)

    '�ΏۃZ���̃V�[�g���A�N�e�B�u�ɂ���
    Call rng.Parent.Activate
    
    '�ΏۃZ����菭����̍s�ɃX�N���[���i�A���AA1�Z������ɂȂ�Ȃ��悤�ɃK�[�h�j
    If rng.Row + scrollRowOffset < 1 Then
        scrollRowOffset = 1 - rng.Row
    End If
    Call Application.GoTo(rng.Offset(scrollRowOffset, 1 - rng.Column), True)
    
    '�ΏۃZ���ɃJ�[�\���ړ�
    Call rng.Activate
    
End Sub

