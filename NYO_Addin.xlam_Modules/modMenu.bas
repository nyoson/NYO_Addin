Attribute VB_Name = "modMenu"
Option Explicit

'*****************************************************************************
'[ �֐��� ]�@SetMenu
'[ �T  �v ]�@�c�[���o�[��ݒ肷��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetMenu()
    Dim objMenu As CommandBar           ' �R�}���h�o�[
    Dim objControl As CommandBarControl ' �R�}���h�o�[���R���g���[��
    Dim Exists As Boolean               ' �������q�b�g�t���O
    Dim i As Integer

    On Error Resume Next
    Set objMenu = Application.CommandBars(C_TOOLBAR_NAME)
    On Error GoTo 0
    If objMenu Is Nothing Then
        ' �c�[���o�[��V�K�쐬
        Set objMenu = Application.CommandBars.Add(Name:=C_TOOLBAR_NAME)
    Else
        ' ���ɑ��݂�����N���A�m�F
        If modMessage.ConfirmMessage("���Ƀc�[���o�[" & C_TOOLBAR_NAME & "�͑��݂��܂��B�N���A���܂����H") <> vbYes Then
            Exit Sub
        End If
        
        ' �A�C�e���N���A
        For Each objControl In objMenu.Controls
            objControl.Delete
        Next
    End If
    
    '�N�C�b�N�c�[���o�[�Ɉڍs�ς݂̂��߁AExcel2007�ȍ~�Ȃ�ǉ����Ȃ�
    If Not modCommon.Is2007orLater Then
        ' �{�^����ǉ��F�ǂݎ���p�̐ݒ�/����
        With objMenu.Controls.Add(Type:=msoControlButton, ID:=456)
        End With
        
        ' �{�^����ǉ��F�t�@�C���̍X�V
        With objMenu.Controls.Add(Type:=msoControlButton, ID:=455)
        End With
        
        ' �{�^����ǉ��F�l�\�t
        With objMenu.Controls.Add(Type:=msoControlButton, ID:=370)
            .Style = msoButtonCaption
            .Caption = "�l(&P)"
        End With
        
        ' �{�^����ǉ��FUnFilter
        With objMenu.Controls.Add(Type:=msoControlButton, ID:=900)
            .Caption = "&UnFilter"
            .TooltipText = "�t�B���^����"
        End With
    End If
    
    ' �{�^����ǉ��FxPaste
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=5837)
        .Style = msoButtonCaption
        .Caption = "&xPaste"
        .TooltipText = "�r���ȊO��\��t��"
    End With
    
    ' �{�^����ǉ��F�A�E�g���C���O���[�v��
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=3159)
    End With
    
    ' �{�^����ǉ��F�A�E�g���C���O���[�v������
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=3160)
    End With
    
    ' �{�^����ǉ��F�E�B���h�E�g�̌Œ�
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=443)
    End With
    
    ' �{�^����ǉ��FOpenDir
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=2950)
        .Style = msoButtonCaption
        .Caption = "Open&Dir"
        .TooltipText = "�ЂƂ���J��"
        .OnAction = "OpenDir"   'modFile
    End With
    
    ' �{�^����ǉ��FAddCopyRow
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=2950)
        .Style = msoButtonCaption
        .Caption = "AddCopy&Row"
        .TooltipText = "�s����(CS++)"
        .OnAction = "AddCopyRow"    'modRow
    End With
    
    ' �{�^����ǉ��F�I�[�g�V�F�C�v�̎����T�C�Y����
    With objMenu.Controls.Add(Type:=msoControlButton)
        .FaceId = 5866
        .Style = msoButtonIcon
        .Caption = "�I�[�g�V�F�C�v�̎����T�C�Y����"
        .TooltipText = .Caption
        .OnAction = "AutoFit"   'modAutoFit
    End With
        
    ' �{�^����ǉ��F�O���b�h���̕\��/��\��
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=485)
        .TooltipText = "�O���b�h���̕\��/��\��"
    End With
    
    ' �{�^����ǉ��FR1C1�`���؂�ւ�
    With objMenu.Controls.Add(Type:=msoControlButton)
        .FaceId = 52
        .Style = msoButtonIcon
        .Caption = "R1C1�`���؂�ւ�"
        .TooltipText = .Caption
        .OnAction = "SwitchR1C1"    'modSheet
    End With
    
    ' �{�^����ǉ��F�V�[�g�\���؂�ւ�
    With objMenu.Controls.Add(Type:=msoControlButton)
        .FaceId = 461
        .Style = msoButtonIcon
        .Caption = "�V�[�g�\���؂�ւ�"
        .TooltipText = .Caption
        .OnAction = "SheetDispChanger"    'modSheet
    End With
    
    '
    ' �{�^����ǉ��FResize50%
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=2950)
        .Style = msoButtonCaption
        .Caption = "Resize50%"
        .TooltipText = "50%�Ƀ��T�C�Y����"
        .OnAction = "SetZoomRate_50per" 'modDraw
    End With
    
    ' ���j���[��ǉ��FELSE
    Dim objMenuElse As CommandBarPopup
    Set objMenuElse = objMenu.Controls.Add(Type:=msoControlPopup)
    With objMenuElse
        .Caption = "E&LSE"
        .TooltipText = .Caption
        
        ' ���j���[��ǉ��FmodFormula
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "��������"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�V�[�g��"
                .OnAction = "InputSheetName"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�V�[�g��(YYYY_MMDD�`��)"
                .OnAction = "InputSheetName_YYYY_MMDD"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�V�[�g��(YYYY_MMDD�`��)=>���t�l�ϊ�"
                .OnAction = "InputSheetName_YYYY_MMDD_AsDate"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�A��"
                .OnAction = "InputSeqNo"
            End With
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "A1�Z����Active�ɂ���"
            .OnAction = "ActivateAllA1Cell" 'modA1
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "�A�E�g���C���̏W�v����������ɂ���"
            .OnAction = "OutlineConfigAboveLeft"    'modSheet
            .TooltipText = "�A�E�g���C���̏W�v�s�����C�W�v�����ɂ���"
        End With
        
        ' ���j���[��ǉ��FmodCrlf
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "���s����"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�I��͈͂̊e�Z���ɉ��s��ǉ�����"
                .OnAction = "AddCrLf"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�I��͈͂̊e�Z���̉��s���폜����"
                .OnAction = "RemoveCrLf"
            End With
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "SQL�O���b�h�\��t��(CS+Q)"
            .OnAction = "PasteSQLGrid"  'modSQL
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "���W���[����S�ăG�N�X�|�[�g"
            .OnAction = "ExportAllModules"  'modExportAllModules
        End With
        
        ' ���j���[��ǉ��FmodFile
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "�t�@�C���n"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�t�@�C�������N���b�v�{�[�h�ɃR�s�["
                .OnAction = "CopyFileName"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "ExportCSV"
                .OnAction = "ExportCSV"
            End With
            
        End With
        
        ' ���j���[��ǉ��FmodDraw
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "�`��n"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�~���`�悷��"
                .OnAction = "DrawX"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�Ԙg��`�悷��"
                .OnAction = "DrawRedFrame"
            End With
            
        End With
        
        ' ���j���[��ǉ��FmodSign
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "�d�q��n"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�d�q��"
                .OnAction = "Sign"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�d�q��ݒ�"
                .OnAction = "ShowSignSetting"
            End With
            
        End With
        
        ' ���j���[��ǉ��FmodWindow
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "�E�B���h�E����"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�V����Excel�ŊJ������"
                .OnAction = "OpenAsNewWindow"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "���E�ɕ��ׂĔ�r"
                .OnAction = "CompareWindowInVertical"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�E�B���h�E�����C���f�B�X�v���C�Ɉړ�"
                .OnAction = "MoveWindowToMainDisplay"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�E�B���h�E���T�u�f�B�X�v���C�Ɉړ�"
                .OnAction = "MoveWindowToSubDisplay"
            End With
            
        End With
        
        ' ���j���[��ǉ��FmodFont
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "���������n"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�Ԏ�(CS+R)"
                .TooltipText = "�����F�́u�ԁv/�u�ʏ�v��؂�ւ���"
                .OnAction = "SwitchRed"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�t�H���g�N���A"
                .TooltipText = "�I���Z���̕����������N���A����"
                .OnAction = "ClearFont"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�n�C�p�[�����N�X�^�C���ύX"
                .TooltipText = "�n�C�p�[�����N�������̃t�H���g���ƃT�C�Y���A�A�N�e�B�u�Z���Ɠ����ɂ���"
                .OnAction = "SetHyperLinkStyle"
            End With
            
        End With
        
        ' ���j���[��ǉ��FmodRecover
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "�C���n"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "���������[�h����"
                .OnAction = "ProcessModeEnd"    'modCommon
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�����t���������ĕ`�悳��Ȃ��s��̉���"
                .OnAction = "FormatConditionsRedrawBugFix"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�����Ȗ��O�t����`���폜"
                .TooltipText = "#REF!�ɂȂ��Ă��閼�O�t����`���폜"
                .OnAction = "DeleteInvaridNameDef"
            End With

            With .Controls.Add(Type:=msoControlButton)
                .Caption = "�J���[�p���b�g�����Z�b�g"
                .TooltipText = "�u�b�N�̃J���[�p���b�g��W���ɖ߂�"
                .OnAction = "ResetColorPalette"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "<DEBUG>�E�N���b�N���j���[�̕���"
                .OnAction = "AddContextMenu"    'modMenu
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "<DEBUG>" & C_TOOL_NAME & "�A�h�C���̍ē���"
                .OnAction = "RestoreNYOAddin"
            End With
            
        End With
        
    End With
    
    ' �c�[���o�[�̈ʒu�i��j
    objMenu.Position = msoBarTop
    
    ' �c�[���o�[��\��
    objMenu.Visible = True
    
    'Excel2003�ȑO�̏ꍇ�̂�
    If Not Is2007orLater Then
        
        ' ���C�����j���[��[�A�h���X]��ǉ�
        Exists = False
        Set objMenu = Application.CommandBars("Worksheet Menu Bar")
        For Each objControl In objMenu.Controls
            ' ���ɓo�^�ς݂łȂ����T��
            If objControl.ID = 1740 Then
                Exists = True
                Exit For
            End If
        Next objControl
        If Exists <> True Then
            ' �R���{�{�b�N�X��ǉ��F�A�h���X
            With objMenu.Controls.Add(Type:=msoControlComboBox, ID:=1740)
                .BeginGroup = True
                .Width = 250
            End With
        End If
        
        ' �����c�[���o�[��[��������],[�㑵��],[�㉺��������],[�������Ɍ���]��ǉ�
        Set objMenu = Application.CommandBars("Formatting")
        objMenu.Visible = True
        objMenu.Reset   ' ��U���Z�b�g
        For i = 1 To objMenu.Controls.Count
            Set objControl = objMenu.Controls(i)
            
            Select Case objControl.ID
                ' �����{�^��
                Case 115:
                    '' ���ɉE���Ɏ�������������΁A�������Ȃ�
                    'If objMenu.Controls(i + 1).ID = 290 Then
                    '    Exit For
                    'End If
                    
                    ' �E���Ƀ{�^����ǉ��F��������
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=290, Before:=i)
                    End With
                    
                ' �E�����{�^��
                Case 121:
                    ' �E���Ƀ{�^����ǉ��F�㑵��
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=2600, Before:=i)
                    End With
                    
                    ' �E���Ƀ{�^����ǉ��F�㉺��������
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=6542, Before:=i)
                    End With
                    
                ' �Z�����������Ē��������{�^��
                Case 402:
                    ' �E���Ƀ{�^����ǉ��F�Z���̌���
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=798, Before:=i)
                    End With
                    
                    ' �E���Ƀ{�^����ǉ��F�������Ɍ���
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=1742, Before:=i)
                    End With
                    
                    ' �E���Ƀ{�^����ǉ��F�Z�������̉���
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=800, Before:=i)
                    End With
                    
                    ' �I��
                    Exit For
                    
            End Select
    
        Next i
        
        ' �W���c�[���o�[��\��
        Set objMenu = Application.CommandBars("Standard")
        objMenu.Visible = True
        
        ' �r���c�[���o�[��\��
        Set objMenu = Application.CommandBars("Borders")
        objMenu.Visible = True
        
    End If
    
End Sub

'*****************************************************************************
'[ �֐��� ]�@DelMenu
'[ �T  �v ]�@�c�[���o�[���폜����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub DelMenu()
    
    '�G���[���͉����������s
    On Error Resume Next

    ' �c�[���o�[�폜
    Call Application.CommandBars(C_TOOLBAR_NAME).Delete

End Sub

'*****************************************************************************
'[ �֐��� ]�@AddContextMenu
'[ �T  �v ]�@�E�N���b�N���j���[�ɃR�}���h��ǉ�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub AddContextMenu()
    
    Debug.Print "�yDEBUG�zAddContextMenu"
    
    '��Addin���ǂݎ���p���ǂ����̃��x��
    Dim strThisAddinReadOnlyText As String
    If ThisWorkbook.ReadOnly Then
        strThisAddinReadOnlyText = "<ReadOnly>"
    Else
        strThisAddinReadOnlyText = "<Editable>"
    End If
    
    '�E�N���b�N���j���[�ɒǉ�
    '(�W�����[�h�Ɖ��y�[�W�v���r���[���[�h�̗����ɒǉ�����)
    Dim objMenu As CommandBar           '�E�N���b�N���j���[
    Dim objControl As CommandBarControl '�E�N���b�N���j���[���R���g���[��
    
    Dim lngCmdNo As Long
    For lngCmdNo = 1 To Application.CommandBars.Count
        Set objMenu = Application.CommandBars(lngCmdNo)
        
        '�S�R�}���h�o�[�̓��A�ȉ��̖��̂̂��́i���Z���̉E�N���b�N���j���[�j���J�X�^�}�C�Y
        '��Cell�Ȃǂ͕W������y�[�W�v���r���[�ȂǕ������݂��邽�߃��[�v����
        '��List Range Popup�̓e�[�u���Ƃ��ď����ݒ肳�ꂽ�͈͓��̉E�N���b�N���j���[
        If objMenu.Name = "Cell" Or _
           objMenu.Name = "List Range Popup" Then
            
            'DEBUG:��U���Z�b�g
            Call objMenu.Reset
            '�������̃A�h�C���ɂ��ύX��������̂Ŗ{���̓R�����g�A�E�g����
            
            'DEBUG:LOG
            Debug.Print objMenu.Index & ":" & objMenu.Name
            
            Set objControl = objMenu.Controls.Add(Temporary:=True)
            '��Temporary:=True�ɂ��A�u�b�N�N���[�Y���Ɏ����I�ɍ폜�����
            With objControl
                .Caption = "�y" & C_TOOL_NAME & "�z" & strThisAddinReadOnlyText
                .Enabled = False
            End With
            Set objControl = Nothing
            
            Set objControl = objMenu.Controls.Add(Temporary:=True)
            '��Temporary:=True�ɂ��A�u�b�N�N���[�Y���Ɏ����I�ɍ폜�����
            With objControl
                .Caption = "�d�q��"
                .OnAction = "Sign"
                .BeginGroup = False
            End With
            Set objControl = Nothing
            
            Set objControl = objMenu.Controls.Add(Temporary:=True)
            '��Temporary:=True�ɂ��A�u�b�N�N���[�Y���Ɏ����I�ɍ폜�����
            With objControl
                .Caption = "�d�q��ݒ�"
                .OnAction = "ShowSignSetting"
                .BeginGroup = False
            End With
            Set objControl = Nothing
            
            Set objControl = objMenu.Controls.Add(Temporary:=True)
            '��Temporary:=True�ɂ��A�u�b�N�N���[�Y���Ɏ����I�ɍ폜�����
            With objControl
                .Caption = "<�����>�c�[���o�[" & C_TOOLBAR_NAME & "�̕���"
                .OnAction = "SetMenu"
                .BeginGroup = False
            End With
            Set objControl = Nothing
            
        End If
    Next
End Sub

'*****************************************************************************
'[ �֐��� ]�@RegistShortcutKey
'[ �T  �v ]�@�V���[�g�J�b�g�L�[���蓖��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub RegistShortcutKey()
    '�y�Q�l�zhttp://excel-ubara.com/excelvba1/EXCELVBA421.html
    
    Application.OnKey "^{Return}", "ThisWorkbook.CtrlEnter"             'Ctrl + Enter            �˓��͕⏕
    Application.OnKey "^{Enter}", "ThisWorkbook.CtrlEnter"              'Ctrl + Enter(Ten�L�[��) �˓��͕⏕
    Application.OnKey "^%;", "ThisWorkbook.CtrlAltSemicolon"            'Ctrl + Alt + ";"        �˓��͕⏕�i���t�j
    Application.OnKey "^%:", "ThisWorkbook.CtrlAltColon"                'Ctrl + Alt + ":"        �˓��͕⏕�i�����j
    Application.OnKey "^{F3}", "modSearchNext.SearchNextThisCellText"   'Ctrl + F3               �˃Z�����e�Ō�����ʂ��J�� �������@�\���㏑��
    Application.OnKey "{F3}", "modSearchNext.SearchNextForward"         'F3                      �ˎ�������
    Application.OnKey "+{F3}", "modSearchNext.SearchNextPrevious"       'Shift + F3              �ˑO������
    Application.OnKey "^+Q", "modSQL.PasteSQLGrid"                       'Ctrl + Shift + Q        ��SQL�O���b�h�\��t��
    Application.OnKey "^+R", "modFont.SwitchRed"                        'Ctrl + Shift + R        �ːԎ��؂�ւ�
    Application.OnKey "^+Y", "modInterior.SwitchInteriorYellow"         'Ctrl + Shift + Y        �ˉ��F�w�i�F�؂�ւ�
    Application.OnKey "^+{+}", "modRow.AddCopyRow"                      'Ctrl + Shift + "+"            �ˑI���s���R�s�[���ĉ��ɒǉ�
    Application.OnKey "^+{107}", "modRow.AddCopyRow"                    'Ctrl + Shift + "+"(Ten�L�[��) �ˑI���s���R�s�[���ĉ��ɒǉ�
    Application.OnKey "^+{-}", "modRow.DelRow"                          'Ctrl + Shift + "-"            �ˑI���s���폜
    Application.OnKey "^+{109}", "modRow.DelRow"                        'Ctrl + Shift + "-"(Ten�L�[��) �ˑI���s���폜
    Application.OnKey "^%{Up}", "modRow.UpRow"                          'Ctrl + Alt + "��"             �ˑI���s����ړ�
    Application.OnKey "^%{Down}", "modRow.DownRow"                      'Ctrl + Alt + "��"             �ˑI���s�����ړ�
    
    Application.OnKey "{F1}", ""                                        'F1                      �˃w���v�ւ̃V���[�g�J�b�g�@�\�𖳌���
    
End Sub

