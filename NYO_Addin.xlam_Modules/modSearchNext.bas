Attribute VB_Name = "modSearchNext"
Option Explicit

'TODO:
'(1)A1�Z������u�O�������v�ɖ��Ή��i���l�ɍŏI�Z������u���������v���j
'   �Ή�����Ȃ�ΏۃV�[�g���X�g���Ɏ擾���Ă����K�v����
'   �����ɍ���Ȃ��̂Ń��b�Z�[�W�\���Ή��݂̂��Ă����āA�D��x��
'(2)�ȉ��̃��[�h���ςɂ���d�g�݂��~�����i�ݒ��ʁH�j
'�E�����ΏہF�l/�\���e�L�X�g/����
'�E��v����F������v/���S��v
'�E�啶��/������/�S�p/���p�̋��
'  ���u��ʂ��Ȃ��v�̏ꍇ�A�u�́v����������Ɓu�΁v��u�ρv���q�b�g����B
'    ���Ȃ݂Ɂu�΁v����������Ɓu�́v��u�ρv�̓q�b�g���Ȃ��B
'    �iInStr�������I�ɔ��p�ϊ����Ă��猟�����Ă�H�j

'##########
' �ϐ�
'##########

'����������
Private SearchText As String
'����������(�啶��/�S�p�ϊ���)
Private SearchTextU As String

'�啶��/������/�S�p/���p����ʂ��邩�H
Private CompareMethod As VbCompareMethod

'������v�����Ƃ��邩�H
Private PartialMatching As Boolean

'��������
Private SearchDirection As XlSearchDirection


'##########
' �֐�
'##########

'�����������ݒ肵�āi�f�t�H���g�F���݃Z���e�L�X�g�j��������
Public Sub SearchNextThisCellText()
    
    Dim strSearchText As String
    Dim isCaseSensitive As Boolean
    Dim isPartialMatching As Boolean
    Dim isCancel As Boolean
    isCancel = False
    
    '�t�H�[���\��
    Call Load(frmSearchNext)
    frmSearchNext.txtSearchText = CStr(ActiveCell.Value)
    If frmSearchNext.txtSearchText = Empty Then
        frmSearchNext.txtSearchText = SearchText
    End If
    frmSearchNext.txtSearchText.SelStart = 0    '�e�L�X�g�S�I����Ԃɂ���
    frmSearchNext.txtSearchText.SelLength = 256 '
    
    frmSearchNext.Show
    isCancel = frmSearchNext.isCancel
    strSearchText = frmSearchNext.txtSearchText.Text
    isCaseSensitive = frmSearchNext.chkCaseSensitive.Value
    isPartialMatching = frmSearchNext.chkPartialMatching.Value
    Call Unload(frmSearchNext)
    
    '�L�����Z�����͏I��
    If isCancel Then
        Exit Sub
    End If
    
    SearchText = strSearchText
    SearchTextU = ConvUpperWideHiragana(SearchText)
    
    '�����͎��͏I��
    If SearchText = Empty Then
        Call modMessage.InfoMessage("���������񂪓��͂���Ă��܂���B")
        Exit Sub
    End If
    
    '�啶��/������/�S�p/���p����ʂ��邩�ǂ���
    If isCaseSensitive Then
        CompareMethod = VbCompareMethod.vbBinaryCompare
    Else
        CompareMethod = VbCompareMethod.vbTextCompare
    End If
    
    '������v�Ƃ��邩�ǂ���
    PartialMatching = isPartialMatching
    
    '��������
    Call SearchNextForward
    
End Sub

'��������
Public Sub SearchNextForward()
    SearchDirection = xlNext
    Call SearchNext_Main
End Sub

'�O������
Public Sub SearchNextPrevious()
    SearchDirection = xlPrevious
    Call SearchNext_Main
End Sub

'�u���������v���C������
Private Sub SearchNext_Main()
    
    If SearchText = "" Then
        Call modMessage.ErrorMessage("���������񂪐ݒ肳��Ă��܂���B" & vbCrLf & _
                    "��U�ACtrl+F3�����s���ĉ������B")
        Exit Sub
    End If
    
    '���݃V�[�g
    Dim shtNext As Worksheet
    Set shtNext = ActiveSheet
    
    '���Ɍ�������Z��
    Dim rngNextCell As Range
    
    '�u�����v�̏ꍇ:
    If SearchDirection = xlNext Then
        '�ŏI��łȂ���΁A�ЂƂE�̃Z��
        If ActiveCell.Column < shtNext.Columns.Count Then
            Set rngNextCell = ActiveCell.Offset(0, 1)
        '�ŏI�񂾂��A�ŏI�s�ł͂Ȃ��ꍇ�́A���s��1��ڂ̃Z��
        ElseIf ActiveCell.Row < shtNext.Rows.Count Then
            Set rngNextCell = shtNext.Cells(ActiveCell.Row + 1, 1)
        '�ŏI��ōŏI�s�̏ꍇ:
        Else
            '�蔲��
            Call modMessage.ErrorMessage( _
                    "�ŏI�Z������́u���֌����v�͖��Ή��ł�m(__)m" & vbCrLf & _
                    "���̃Z���ɃJ�[�\�����ړ����Ă���A���s���ĉ������B")
            Exit Sub
        End If

    '�u�O���v�̏ꍇ:
    Else
        '1��ڂłȂ���΂ЂƂ��̃Z��
        If ActiveCell.Column > 1 Then
            Set rngNextCell = ActiveCell.Offset(0, -1)
        '1��ڂ����A1�s�ڂł͂Ȃ��ꍇ�́A�O�s�̍ŏI�Z��
        ElseIf ActiveCell.Row > 1 Then
            Set rngNextCell = shtNext.Cells(ActiveCell.Row - 1, _
                                            GetLastCol(ActiveCell))
        '1�s�ڂ�1��ځi��A1�Z���j�̏ꍇ:
        Else
            '�蔲��
            Call modMessage.ErrorMessage( _
                    "A1�Z������́u�O�֌����v�͖��Ή��ł�m(__)m" & vbCrLf & _
                    "���̃Z���ɃJ�[�\�����ړ����Ă���A���s���ĉ������B")
            Exit Sub
        End If
    End If
    
    '�u�b�N������
    Call SearchNext_Book(rngNextCell)
    
End Sub

'�u�b�N������
Private Sub SearchNext_Book(ByRef rngNextCell As Range)
    
    '�u�b�N���̍ŏI�Z���ɓ��B�������ǂ���
    Dim hasReachedLast As Boolean
    hasReachedLast = False
    
    '���Ɍ�������V�[�g
    Dim shtNext As Worksheet
    Set shtNext = rngNextCell.Parent
    
    '�Ώۃu�b�N
    Dim wbkTarget As Workbook
    Set wbkTarget = shtNext.Parent
    
    '�����ΏۃV�[�g���X�g�̌���Index
    Dim lngSheetIndex As Long
    lngSheetIndex = 0
    
    '�����ΏۃV�[�g�̃��X�g�𐶐��i����\���V�[�g�ȊO�j
    Dim shtTarget As Worksheet
    Dim shtTargetList() As Worksheet
    Dim lngTargetListCount As Long
    lngTargetListCount = 0
    ReDim shtTargetList(0 To 0)
    For Each shtTarget In wbkTarget.Worksheets
        If shtTarget.Visible = xlSheetVisible Then
            'Call shtTargetList.Add(shtTarget)
            lngTargetListCount = lngTargetListCount + 1
            ReDim Preserve shtTargetList(0 To lngTargetListCount)
            Set shtTargetList(lngTargetListCount) = shtTarget
        End If
        '���łɁA��L���X�g�ɂ����錻�V�[�g��Index���擾
        If shtTarget Is shtNext Then
            lngSheetIndex = lngTargetListCount
        End If
    Next
    
    Do
        '�V�[�g���ŏI�Z�����擾
        Dim rngLastCell As Range
        Set rngLastCell = modCommon.GetLastCell(shtNext)
        
        '�V�[�g���̍ŏI�s�ƍŏI����擾
        Dim lngMaxRow As Long
        Dim lngMaxCol As Long
        lngMaxRow = rngLastCell.Row
        lngMaxCol = rngLastCell.Column
        
        '�V�[�g��������
        If SearchNext_Sheet(rngNextCell, lngMaxCol, lngMaxRow) Then
            '���������̂Ȃ�I��
            Exit Sub
        End If
        
        '������Ȃ���΁A���̃V�[�g�̔�����s��
        
        '�u�����v�����̏ꍇ�F
        If SearchDirection = xlNext Then
            If lngSheetIndex < lngTargetListCount Then
                '�ŏI�V�[�g�łȂ���΁A���̃V�[�g
                lngSheetIndex = lngSheetIndex + 1
            Else
                '���ɍŌ�ɓ��B���Ă�����A����ȏ�T���Ă�������Ȃ��̂ŏI��
                If hasReachedLast = True Then
                    Call modMessage.InfoMessage("������܂���ł����B")
                    Exit Sub
                End If
                hasReachedLast = True
                
                '�Ō�̃V�[�g�܂Ō�����Ȃ�������A�ŏ��̃V�[�g�ɖ߂邩�m�F
                If modMessage.ConfirmMessage( _
                            "�Ō�̃V�[�g�ɓ��B���܂����B" & vbCrLf & _
                            "�ŏ��̃V�[�g�ɖ߂��Č������܂����H") <> vbYes Then
                    Exit Sub
                End If
                lngSheetIndex = 1
            End If
            
            '���̃V�[�g��A1�Z�����猟��
            Set shtNext = shtTargetList(lngSheetIndex)
            Set rngNextCell = shtNext.Cells(1, 1)
        
        '�u�O���v�����̏ꍇ�F
        Else
            If lngSheetIndex > 1 Then
                '�ŏ��̃V�[�g�łȂ���΁A�O�̃V�[�g
                lngSheetIndex = lngSheetIndex - 1
            Else
                '���ɍŏ��ɓ��B���Ă�����A����ȏ�T���Ă�������Ȃ��̂ŏI��
                If hasReachedLast = True Then
                    Call modMessage.InfoMessage("������܂���ł����B")
                    Exit Sub
                End If
                hasReachedLast = True
                
                '�ŏ��̃V�[�g�܂Ō�����Ȃ�������A�Ō�̃V�[�g�ɖ߂邩�m�F
                If modMessage.ConfirmMessage( _
                            "�ŏ��̃V�[�g�ɓ��B���܂����B" & vbCrLf & _
                            "�Ō�̃V�[�g�ɖ߂��Č������܂����H") <> vbYes Then
                    Exit Sub
                End If
                lngSheetIndex = lngTargetListCount
            End If
            
            '�O�̃V�[�g�̍ŏI�Z�����猟��
            Set shtNext = shtTargetList(lngSheetIndex)
            Set rngNextCell = modCommon.GetLastCell(shtNext)
        
        End If
        
    Loop
    
End Sub

'�V�[�g������
Private Function SearchNext_Sheet(ByRef rngNextCell As Range, ByVal lngMaxCol As Long, ByVal lngMaxRow As Long) As Boolean
    
    Do
        '�s��������
        If SearchNext_Row(rngNextCell, lngMaxCol) Then
            SearchNext_Sheet = True
            Exit Function
        End If
        
        '�u�����v�����̏ꍇ�F
        If SearchDirection = xlNext Then
        
            If rngNextCell.Row >= lngMaxRow Then Exit Do
            
            '�ЂƂ��̍s�̍ō��Z��
            Set rngNextCell = rngNextCell.Offset(1, 0).EntireRow _
                                .Cells(1, 1)
            
        '�u�O���v�����̏ꍇ�F
        Else
            
            If rngNextCell.Row <= 1 Then Exit Do
            
            '�ЂƂ�̍s�̍ŏI�Z��
            Set rngNextCell = rngNextCell.Offset(-1, 0).EntireRow _
                                .Cells(1, GetLastCol(rngNextCell))
        
        End If
    Loop
    
    SearchNext_Sheet = False
    
End Function

'�s������
Private Function SearchNext_Row(ByRef rngNextCell As Range, ByVal lngMaxCol As Long) As Boolean
    
    Do
        If IsHit(rngNextCell) Then
            
            '�Z����\��
            Call Application.GoTo(rngNextCell, False)
            
            SearchNext_Row = True
            Exit Function
        End If
        
        '�u�����v�����̏ꍇ�F
        If SearchDirection = xlNext Then
            
            If rngNextCell.Column >= lngMaxCol Then Exit Do
            
            '�ЂƂE�̃Z����
            Set rngNextCell = rngNextCell.Offset(0, 1)
            
        '�u�O���v�����̏ꍇ�F
        Else
            
            If rngNextCell.Column <= 1 Then Exit Do
            
            '�ЂƂ��̃Z����
            Set rngNextCell = rngNextCell.Offset(0, -1)
            
        End If
    Loop
    
    SearchNext_Row = False
    
End Function

'�q�b�g����
Private Function IsHit(ByRef rngNextCell As Range) As Boolean
    
    '�ԋp�l������
    IsHit = False
    
    '�Z�����e�L�X�g
    Dim strText As String
    strText = CStr(rngNextCell.Value)
    
    '��̃Z���̓X�L�b�v
    If strText = "" Then
        Exit Function
    End If
    
    '������v���邩�H
    If PartialMatching Then
        If InStr(1, strText, SearchText, CompareMethod) > 0 Then
            IsHit = True
        End If
    Else
        If strText = SearchText Then
            IsHit = True
        Else
            If CompareMethod = vbTextCompare Then
                Dim strTextU As String
                strTextU = ConvUpperWideHiragana(strText)
                
                If strTextU = SearchTextU Then
                    IsHit = True
                End If
            End If
        End If
    End If
    
End Function

'�啶���S�p�Ђ炪�Ȃɕϊ�
Private Function ConvUpperWideHiragana(ByVal strText As String) As String
    ConvUpperWideHiragana = StrConv(strText, vbUpperCase + vbWide + vbHiragana)
End Function

'�ΏۃZ���̃V�[�g���̍ŏI����擾
Private Function GetLastCol(ByVal rngTarget As Range) As Long
    Dim shtTarget As Worksheet
    Set shtTarget = rngTarget.Parent
    
    Dim rngUsedRange As Range
    Set rngUsedRange = shtTarget.UsedRange
    
    GetLastCol = rngUsedRange.Column + rngUsedRange.Columns.Count - 1
End Function

