VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSheetDispChanger 
   Caption         =   "SheetDispChanger"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   OleObjectBlob   =   "frmSheetDispChanger.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSheetDispChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'##########
' API
'##########

'API�֐��i�t�H�[���T�C�Y���ςɂ��邽�߁j
Private Declare PtrSafe Function DrawMenuBar Lib "user32" _
    (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'API�萔
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000

'##########
' �萔
'##########

'�t�H�[���̍ŏ��T�C�Y
Private Const WIN_MIN_SIZE_W As Double = 347
Private Const WIN_MIN_SIZE_H As Double = 150

'�A���J�[���
Private Enum Anchor
    Top = 1
    Left = 2
    Right = 4
    Bottom = 8
    
    TopLeft = 3
    TopLeftRight = 7
    LeftBottom = 10
    RightBottom = 12
    
    All = 15
End Enum

'�R���g���[�����
Private Type ControlInfo
    objCtrl As Control
    enmAnchor As Anchor
    diffRight As Double
    diffBottom As Double
    diffWidth As Double
    diffHeight As Double
End Type

'##########
' �ϐ�
'##########

'�R���g���[����񃊃X�g
Private ctrlList() As ControlInfo

'##########
' �C�x���g
'##########

'*****************************************************************************
'[ �C�x���g�� ] �t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()
    
    'DEBUG�p�̃��x�����\���ɂ���
    lblDebug.Visible = False
    
    '�e�R���g���[���z�u�̃A���J�[�ݒ���s��
    ReDim ctrlList(0) As ControlInfo
    Call AddControlInfo(lstSheets, Anchor.All)
    Call AddControlInfo(btnSelectAll, Anchor.LeftBottom)
    Call AddControlInfo(btnSelectNone, Anchor.LeftBottom)
    Call AddControlInfo(btnGetSheetName, Anchor.LeftBottom)
    Call AddControlInfo(btnMakeIndexSheet, Anchor.LeftBottom)
    Call AddControlInfo(btnApply, Anchor.RightBottom)
    Call AddControlInfo(btnClose, Anchor.RightBottom)
    
    lstSheets.Clear
    lstSheets.ListStyle = fmListStyleOption
    
    Dim i As Long
    For i = 1 To Worksheets.Count
        lstSheets.AddItem Worksheets(i).Name, i - 1
        lstSheets.Selected(i - 1) = IIf(Worksheets(i).Visible = XlSheetVisibility.xlSheetVisible, True, False)
    Next i
End Sub

'*****************************************************************************
'[ �C�x���g�� ] �t�H�[���\����
'*****************************************************************************
Private Sub UserForm_Activate()
    Dim hWnd As Long
    Dim lngStyle As Long
    
    '�t�H�[�����T�C�Y�ύX�\�ɂ���
    hWnd = GetActiveWindow
    lngStyle = GetWindowLong(hWnd, GWL_STYLE)
    lngStyle = lngStyle Or WS_THICKFRAME
    Call SetWindowLong(hWnd, GWL_STYLE, lngStyle)
    
    '���j���[�o�[���ĕ`��(�~�{�^���������̂�)
    Call DrawMenuBar(hWnd)
End Sub

'*****************************************************************************
'[ �C�x���g�� ] �t�H�[�����T�C�Y��
'*****************************************************************************
Private Sub UserForm_Resize()
    
    '�t�H�[���̍ŏ��T�C�Y�𐧌�
    Me.Width = Max(Me.Width, WIN_MIN_SIZE_W)
    Me.Height = Max(Me.Height, WIN_MIN_SIZE_H)
    
    '�e�R���g���[�����Ǐ]
    Dim i As Long
    For i = 1 To UBound(ctrlList)
        Call RepositionControl(ctrlList(i))
    Next
    
End Sub

'*****************************************************************************
'[ �C�x���g�� ] �S�I�� �{�^��������
'*****************************************************************************
Private Sub btnSelectAll_Click()
    Call SetSelectedAll(True)
End Sub

'*****************************************************************************
'[ �C�x���g�� ] �I������ �{�^��������
'*****************************************************************************
Private Sub btnSelectNone_Click()
    Call SetSelectedAll(False)
End Sub

'*****************************************************************************
'[ �C�x���g�� ] �V�[�g�� �{�^��������
'*****************************************************************************
Private Sub btnGetSheetName_Click()
    Call GetSheetName
End Sub

'*****************************************************************************
'[ �C�x���g�� ] �ڎ��V�[�g�쐬 �{�^��������
'*****************************************************************************
Private Sub btnMakeIndexSheet_Click()
    Call MakeIndexSheet
End Sub


'*****************************************************************************
'[ �C�x���g�� ] �K�p �{�^��������
'*****************************************************************************
Private Sub btnApply_Click()
    
    Dim strPwd As String
    
    If RetryUnprotectBookForStruct(ActiveWorkbook, strPwd) = False Then
        Call modMessage.ErrorMessage("���f���܂���")
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To Worksheets.Count
        
        Dim shtTarget As Worksheet
        Set shtTarget = Worksheets(i)
        
        If RetryUnprotectSheet(shtTarget, strPwd) = False Then
            Call modMessage.ErrorMessage("���f���܂���")
            Exit Sub
        End If
        shtTarget.Visible = lstSheets.Selected(i - 1)
        
    Next i
    'Me.Hide
End Sub

'*****************************************************************************
'[ �C�x���g�� ] �L�����Z�� �{�^��������
'*****************************************************************************
Private Sub btnClose_Click()
    Me.Hide
End Sub

'##########
' �֐�
'##########
'*****************************************************************************
'[ �֐��� ] AddControlInfo
'[ �T  �v ] �R���g���[������ǉ�
'[ ��  �� ] �R���g���[���C�A���J�[���
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Private Sub AddControlInfo(ByVal objCtrl As Control, ByVal enmAnchor As Anchor)
    
    Dim info As ControlInfo
    
    Set info.objCtrl = objCtrl
    info.enmAnchor = enmAnchor
    
    If enmAnchor And Anchor.Right Then
        If enmAnchor And Anchor.Left Then
            info.diffWidth = Me.Width - objCtrl.Width
        Else
            info.diffRight = Me.Width - objCtrl.Left
        End If
    End If
    
    If enmAnchor And Anchor.Bottom Then
        If enmAnchor And Anchor.Top Then
            info.diffHeight = Me.Height - objCtrl.Height
        Else
            info.diffBottom = Me.Height - objCtrl.Top
        End If
    End If
    
    '�R���g���[����񃊃X�g�ɒǉ�
    Dim lngCount As Long
    lngCount = UBound(ctrlList) + 1
    ReDim Preserve ctrlList(lngCount) As ControlInfo
    ctrlList(lngCount) = info
    
End Sub

'*****************************************************************************
'[ �֐��� ] RepositionControl
'[ �T  �v ] �R���g���[���������ɁA�R���g���[���̈ʒu�ƃT�C�Y���ĕ`��
'[ ��  �� ] �R���g���[���C�A���J�[���
'[ �߂�l ] �R���g���[�����
'*****************************************************************************
Private Sub RepositionControl(ByRef info As ControlInfo)
    
    Dim objCtrl As Control
    Set objCtrl = info.objCtrl
    
    If objCtrl Is Nothing Then Exit Sub
    
    If info.enmAnchor And Anchor.Right Then
        If info.enmAnchor And Anchor.Left Then
            objCtrl.Width = Max(Me.Width - info.diffWidth, 0)
        Else
            objCtrl.Left = Me.Width - info.diffRight
        End If
    End If
    
    If info.enmAnchor And Anchor.Bottom Then
        If info.enmAnchor And Anchor.Top Then
            objCtrl.Height = Max(Me.Height - info.diffHeight, 0)
        Else
            objCtrl.Top = Me.Height - info.diffBottom
        End If
    End If
    
End Sub

'*****************************************************************************
'[ �֐��� ] Max
'[ �T  �v ] �w�肳�ꂽ�l�̂����A�傫������Ԃ�
'[ ��  �� ] �l�P�C�l�Q
'[ �߂�l ] �ő�l
'*****************************************************************************
Private Function Max(ByVal dblVal1 As Double, ByVal dblVal2 As Double) As Double
    If dblVal1 < dblVal2 Then
        Max = dblVal2
    Else
        Max = dblVal1
    End If
End Function

'*****************************************************************************
'[ �֐��� ] Min
'[ �T  �v ] �w�肳�ꂽ�l�̂����A����������Ԃ�
'[ ��  �� ] �l�P�C�l�Q
'[ �߂�l ] �ŏ��l
'*****************************************************************************
Private Function Min(ByVal dblVal1 As Double, ByVal dblVal2 As Double) As Double
    Min = -Max(-dblVal1, -dblVal2)
End Function

'*****************************************************************************
'[ �֐��� ] SetSelectedAll
'[ �T  �v ] �S�Ẵm�[�h�̃`�F�b�N��Ԃ�ύX����
'[ ��  �� ] �`�F�b�N���
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Private Sub SetSelectedAll(flg As Boolean)
    Dim i As Long
    For i = 1 To lstSheets.ListCount
        lstSheets.Selected(i - 1) = flg
    Next i
End Sub

'*****************************************************************************
'[ �֐��� ] GetSheetName
'[ �T  �v ] �I���V�[�g�����R�s�[����
'[ ��  �� ] �Ȃ�
'[ �߂�l ] �Ȃ�
'*****************************************************************************
Private Sub GetSheetName()
    Dim data As New DataObject
    Dim strText As String
    Dim i As Long
    
    strText = ""
    For i = 1 To lstSheets.ListCount
        If lstSheets.Selected(i - 1) Then
            strText = strText & lstSheets.List(i - 1) & vbCrLf
        End If
    Next
    
    '�N���b�v�{�[�h�ɃR�s�[
    Call data.SetText(strText)
    Call data.PutInClipboard
End Sub

'*****************************************************************************
'[ �֐��� ]�@MakeIndexSheet
'[ �T  �v ]�@�ڎ��V�[�g�쐬
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]  �Ȃ�
'*****************************************************************************
Private Sub MakeIndexSheet()
    
    Const ROW_START As Long = 2
    
    Const COL As Long = 2
    Const IDX_SHT_NAME As String = "�ڎ�"
    
    Dim wbkActive As Workbook
    Set wbkActive = ActiveWorkbook
    
    Dim shtCur As Worksheet
    Dim shtIndex As Worksheet
    
    '�ڎ��V�[�g����
    Set shtIndex = Nothing
    For Each shtCur In wbkActive.Worksheets
        If shtCur.Name = IDX_SHT_NAME Then
            Set shtIndex = shtCur
            Exit For
        End If
    Next
    
    '�ڎ��V�[�g�����ɂ���΍폜���Ă���
    If Not shtIndex Is Nothing Then
        shtIndex.Activate
        If ConfirmMessage(IDX_SHT_NAME & "�V�[�g�����ɂ���܂��B��U�폜���Ă�낵���ł����H") _
           <> VbMsgBoxResult.vbYes Then
            Exit Sub
        End If
        Application.DisplayAlerts = False
        shtIndex.Delete
        Application.DisplayAlerts = True
    End If
    
    '�ڎ��V�[�g����
    Set shtIndex = wbkActive.Worksheets.Add(Before:=wbkActive.Worksheets(1))
    shtIndex.Name = IDX_SHT_NAME
    
    Dim lngRow As Long
    lngRow = ROW_START
    
    For Each shtCur In wbkActive.Worksheets
        
        '��\���V�[�g�̓X�L�b�v
        If shtCur.Visible <> xlSheetVisible Then GoTo NEXT_SHEET
        
        Dim rngItem As Range
        Set rngItem = shtIndex.Cells(lngRow, COL)
        
        If shtCur Is shtIndex Then
            '�w�b�_�[�s�o��
            rngItem.Value = "����"
        Else
            '���׍s�o��
            rngItem.Value = shtCur.Name
            Call shtIndex.Hyperlinks.Add(rngItem, "#" & shtCur.Name & "!$A$1")
            rngItem.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone
        End If
        
        lngRow = lngRow + 1
        
NEXT_SHEET:
    Next
    
    '�S�̂̌r��
    Dim rngTable As Range
    Set rngTable = Range(shtIndex.Cells(ROW_START, COL), shtIndex.Cells(lngRow, COL))
    With rngTable.Borders
        .LineStyle = XlLineStyle.xlDot
        .ColorIndex = xlColorIndexAutomatic
    End With
    Call rngTable.BorderAround(XlLineStyle.xlContinuous, xlThin, XlColorIndex.xlColorIndexAutomatic)
    
    '�w�b�_�[�s�̌r��
    Dim rngHeadr As Range
    Set rngHeadr = rngTable.Resize(1, rngTable.Columns.Count)
    
    '�w�b�_�[�s�̐F
    rngHeadr.Interior.Color = RGB(0, 112, 192)
    rngHeadr.Font.Color = ColorConstants.vbWhite
    
    '�񕝎�������
    Set rngTable = shtIndex.Columns(COL)
    Call rngTable.AutoFit
    
End Sub

'*****************************************************************************
'[ �֐��� ] RetryUnprotectBookForStruct
'[ �T  �v ] �i�V�[�g�\���́j�u�b�N�ی�������ł���܂�PWD���͂��J��Ԃ�
'[ ��  �� ] �Ώۃu�b�N�C�p�X���[�h
'[ �߂�l ] True:����/False:���s
'*****************************************************************************
Private Function RetryUnprotectBookForStruct(ByVal wbkTarget As Workbook, ByRef strPwd As String) As Boolean
    Dim varPwd As Variant
    
    Do While TryUnprotectBookForStruct(wbkTarget, strPwd) = False
        varPwd = Application.InputBox(Title:="�V�[�g�ی�����p�X���[�h����", Prompt:="�p�X���[�h����͂��ĉ������B", Default:=strPwd)
        If varPwd = False Then
            RetryUnprotectBookForStruct = False
            Exit Function
        End If
        
        strPwd = varPwd
    Loop
    
    RetryUnprotectBookForStruct = True
    
End Function

'*****************************************************************************
'[ �֐��� ] TryUnprotectBookForStruct
'[ �T  �v ] �i�V�[�g�\���́j�u�b�N�ی����������
'[ ��  �� ] �Ώۃu�b�N�C�p�X���[�h
'[ �߂�l ] True:����/False:���s
'*****************************************************************************
Private Function TryUnprotectBookForStruct(ByVal wbkTarget As Workbook, ByVal strPwd As String) As Boolean

On Error GoTo Catch
    Call wbkTarget.Unprotect(strPwd)
    TryUnprotectBookForStruct = True
    Exit Function
    
Catch:
    TryUnprotectBookForStruct = False
    
End Function

'*****************************************************************************
'[ �֐��� ] RetryUnprotectSheet
'[ �T  �v ] �V�[�g�ی�������ł���܂�PWD���͂��J��Ԃ�
'[ ��  �� ] �ΏۃV�[�g�C�p�X���[�h
'[ �߂�l ] True:����/False:���s
'*****************************************************************************
Private Function RetryUnprotectSheet(ByVal shtTarget As Worksheet, ByRef strPwd As String) As Boolean
    Dim varPwd As Variant
    
    Do While TryUnprotectSheet(shtTarget, strPwd) = False
        varPwd = Application.InputBox(Title:="�V�[�g�ی�����p�X���[�h����", Prompt:="�p�X���[�h����͂��ĉ������B", Default:=strPwd)
        If varPwd = False Then
            RetryUnprotectSheet = False
            Exit Function
        End If
        
        strPwd = varPwd
    Loop
    
    RetryUnprotectSheet = True
    
End Function

'*****************************************************************************
'[ �֐��� ] TryUnprotectSheet
'[ �T  �v ] �V�[�g�ی����������
'[ ��  �� ] �ΏۃV�[�g�C�p�X���[�h
'[ �߂�l ] True:����/False:���s
'*****************************************************************************
Private Function TryUnprotectSheet(ByVal shtTarget As Worksheet, ByVal strPwd As String) As Boolean
    
On Error GoTo Catch
    Call shtTarget.Unprotect(strPwd)
    TryUnprotectSheet = True
    Exit Function
    
Catch:
    TryUnprotectSheet = False
    
End Function

