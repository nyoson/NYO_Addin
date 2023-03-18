Attribute VB_Name = "modWindow"
Option Explicit

'*****************************************************************************
' �V�����E�B���h�E�ŊJ������
' �}�N���쐬�� : 2013/05/20  ���[�U�[�� : nyoson
'*****************************************************************************
Public Sub OpenAsNewWindow()
    
    Dim wbkTarget As Workbook
    Set wbkTarget = ActiveWorkbook
    
    With CreateObject("Scripting.FileSystemObject")
        If .GetExtensionName(wbkTarget.FullName) = Empty Then
            Call modMessage.ErrorMessage("�ۑ�����Ă��Ȃ����߁A���s�ł��܂���B")
            Exit Sub
        End If
    End With
    
    Dim newExcelApp As Excel.Application
    Set newExcelApp = CreateObject("Excel.Application")
    newExcelApp.Visible = True
    Call newExcelApp.Workbooks.Open(FileName:=wbkTarget.FullName, ReadOnly:=True)
    
    If modMessage.ConfirmMessage("���̃u�b�N����܂����H") <> vbYes Then
        Exit Sub
    End If
    
    Call wbkTarget.Close
    
End Sub

'*****************************************************************************
' ���E�ɕ��ׂĔ�r
' �}�N���쐬�� : 2013/06/10  ���[�U�[�� : nyoson
'*****************************************************************************
Public Sub CompareWindowInVertical()
    
    Dim winCurrent As Window
    Dim winTarget As Window
    
    If Windows.Count < 2 Then
        Call modMessage.ErrorMessage("�q�E�B���h�E�̐�������܂���B")
        Exit Sub
    End If
    
    Set winCurrent = ActiveWindow
    For Each winTarget In Windows
        If winCurrent.Caption <> winTarget.Caption Then
            Exit For
        End If
    Next
    
    '���ׂĔ�r
    Windows.CompareSideBySideWith winTarget.Caption
    
    '���E�ɕ��ׂ�
    Windows.Arrange ArrangeStyle:=xlVertical
    
End Sub

'*****************************************************************************
' �E�B���h�E�����C���f�B�X�v���C�Ɉړ�
' �}�N���쐬�� : 2013/07/09  ���[�U�[�� : nyoson
'*****************************************************************************
Public Sub MoveWindowToMainDisplay()
    
    '�ő剻��ŏ���������
    Application.WindowState = xlNormal
    
    '�E�B���h�E�T�C�Y�ƈʒu��ύX
    Application.Height = 621.75
    Application.Width = 937.5
    Application.Top = 2.25
    Application.Left = 7
    
End Sub

'*****************************************************************************
' �E�B���h�E���T�u�f�B�X�v���C�Ɉړ�
' �}�N���쐬�� : 2013/06/26  ���[�U�[�� : nyoson
'*****************************************************************************
Public Sub MoveWindowToSubDisplay()
    
    '�ő剻��ŏ���������
    Application.WindowState = xlNormal
    
    '�E�B���h�E�T�C�Y�ƈʒu��ύX
    Application.Height = 621.75
    Application.Width = 937.5
    Application.Top = -295.5
    Application.Left = 1212.25
    
End Sub

'*****************************************************************************
' �E�B���h�E�T�C�Y���f�o�b�O�o�͂���
'*****************************************************************************
Private Sub GetWindowState()
    
    'ActiveWindow�̃E�B���h�E�T�C�Y���o��
    Debug.Print "==========================="
    Debug.Print "Application.Height = " & Application.Height; ""
    Debug.Print "Application.Width = " & Application.Width; ""
    Debug.Print "Application.Top = " & Application.Top; ""
    Debug.Print "Application.Left = " & Application.Left
    
End Sub

