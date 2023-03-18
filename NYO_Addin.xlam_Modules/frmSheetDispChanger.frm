VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSheetDispChanger 
   Caption         =   "SheetDispChanger"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   OleObjectBlob   =   "frmSheetDispChanger.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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

'API関数（フォームサイズを可変にするため）
Private Declare PtrSafe Function DrawMenuBar Lib "user32" _
    (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'API定数
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000

'##########
' 定数
'##########

'フォームの最小サイズ
Private Const WIN_MIN_SIZE_W As Double = 347
Private Const WIN_MIN_SIZE_H As Double = 150

'アンカー種別
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

'コントロール情報
Private Type ControlInfo
    objCtrl As Control
    enmAnchor As Anchor
    diffRight As Double
    diffBottom As Double
    diffWidth As Double
    diffHeight As Double
End Type

'##########
' 変数
'##########

'コントロール情報リスト
Private ctrlList() As ControlInfo

'##########
' イベント
'##########

'*****************************************************************************
'[ イベント名 ] フォームロード時
'*****************************************************************************
Private Sub UserForm_Initialize()
    
    'DEBUG用のラベルを非表示にする
    lblDebug.Visible = False
    
    '各コントロール配置のアンカー設定を行う
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
'[ イベント名 ] フォーム表示時
'*****************************************************************************
Private Sub UserForm_Activate()
    Dim hWnd As Long
    Dim lngStyle As Long
    
    'フォームをサイズ変更可能にする
    hWnd = GetActiveWindow
    lngStyle = GetWindowLong(hWnd, GWL_STYLE)
    lngStyle = lngStyle Or WS_THICKFRAME
    Call SetWindowLong(hWnd, GWL_STYLE, lngStyle)
    
    'メニューバーを再描画(×ボタンがずれるので)
    Call DrawMenuBar(hWnd)
End Sub

'*****************************************************************************
'[ イベント名 ] フォームリサイズ時
'*****************************************************************************
Private Sub UserForm_Resize()
    
    'フォームの最小サイズを制限
    Me.Width = Max(Me.Width, WIN_MIN_SIZE_W)
    Me.Height = Max(Me.Height, WIN_MIN_SIZE_H)
    
    '各コントロールも追従
    Dim i As Long
    For i = 1 To UBound(ctrlList)
        Call RepositionControl(ctrlList(i))
    Next
    
End Sub

'*****************************************************************************
'[ イベント名 ] 全選択 ボタン押下時
'*****************************************************************************
Private Sub btnSelectAll_Click()
    Call SetSelectedAll(True)
End Sub

'*****************************************************************************
'[ イベント名 ] 選択解除 ボタン押下時
'*****************************************************************************
Private Sub btnSelectNone_Click()
    Call SetSelectedAll(False)
End Sub

'*****************************************************************************
'[ イベント名 ] シート名 ボタン押下時
'*****************************************************************************
Private Sub btnGetSheetName_Click()
    Call GetSheetName
End Sub

'*****************************************************************************
'[ イベント名 ] 目次シート作成 ボタン押下時
'*****************************************************************************
Private Sub btnMakeIndexSheet_Click()
    Call MakeIndexSheet
End Sub


'*****************************************************************************
'[ イベント名 ] 適用 ボタン押下時
'*****************************************************************************
Private Sub btnApply_Click()
    
    Dim strPwd As String
    
    If RetryUnprotectBookForStruct(ActiveWorkbook, strPwd) = False Then
        Call modMessage.ErrorMessage("中断しました")
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To Worksheets.Count
        
        Dim shtTarget As Worksheet
        Set shtTarget = Worksheets(i)
        
        If RetryUnprotectSheet(shtTarget, strPwd) = False Then
            Call modMessage.ErrorMessage("中断しました")
            Exit Sub
        End If
        shtTarget.Visible = lstSheets.Selected(i - 1)
        
    Next i
    'Me.Hide
End Sub

'*****************************************************************************
'[ イベント名 ] キャンセル ボタン押下時
'*****************************************************************************
Private Sub btnClose_Click()
    Me.Hide
End Sub

'##########
' 関数
'##########
'*****************************************************************************
'[ 関数名 ] AddControlInfo
'[ 概  要 ] コントロール情報を追加
'[ 引  数 ] コントロール，アンカー種別
'[ 戻り値 ] なし
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
    
    'コントロール情報リストに追加
    Dim lngCount As Long
    lngCount = UBound(ctrlList) + 1
    ReDim Preserve ctrlList(lngCount) As ControlInfo
    ctrlList(lngCount) = info
    
End Sub

'*****************************************************************************
'[ 関数名 ] RepositionControl
'[ 概  要 ] コントロール情報を元に、コントロールの位置とサイズを再描画
'[ 引  数 ] コントロール，アンカー種別
'[ 戻り値 ] コントロール情報
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
'[ 関数名 ] Max
'[ 概  要 ] 指定された値のうち、大きい方を返す
'[ 引  数 ] 値１，値２
'[ 戻り値 ] 最大値
'*****************************************************************************
Private Function Max(ByVal dblVal1 As Double, ByVal dblVal2 As Double) As Double
    If dblVal1 < dblVal2 Then
        Max = dblVal2
    Else
        Max = dblVal1
    End If
End Function

'*****************************************************************************
'[ 関数名 ] Min
'[ 概  要 ] 指定された値のうち、小さい方を返す
'[ 引  数 ] 値１，値２
'[ 戻り値 ] 最小値
'*****************************************************************************
Private Function Min(ByVal dblVal1 As Double, ByVal dblVal2 As Double) As Double
    Min = -Max(-dblVal1, -dblVal2)
End Function

'*****************************************************************************
'[ 関数名 ] SetSelectedAll
'[ 概  要 ] 全てのノードのチェック状態を変更する
'[ 引  数 ] チェック状態
'[ 戻り値 ] なし
'*****************************************************************************
Private Sub SetSelectedAll(flg As Boolean)
    Dim i As Long
    For i = 1 To lstSheets.ListCount
        lstSheets.Selected(i - 1) = flg
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ] GetSheetName
'[ 概  要 ] 選択シート名をコピーする
'[ 引  数 ] なし
'[ 戻り値 ] なし
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
    
    'クリップボードにコピー
    Call data.SetText(strText)
    Call data.PutInClipboard
End Sub

'*****************************************************************************
'[ 関数名 ]　MakeIndexSheet
'[ 概  要 ]　目次シート作成
'[ 引  数 ]　なし
'[ 戻り値 ]  なし
'*****************************************************************************
Private Sub MakeIndexSheet()
    
    Const ROW_START As Long = 2
    
    Const COL As Long = 2
    Const IDX_SHT_NAME As String = "目次"
    
    Dim wbkActive As Workbook
    Set wbkActive = ActiveWorkbook
    
    Dim shtCur As Worksheet
    Dim shtIndex As Worksheet
    
    '目次シート検索
    Set shtIndex = Nothing
    For Each shtCur In wbkActive.Worksheets
        If shtCur.Name = IDX_SHT_NAME Then
            Set shtIndex = shtCur
            Exit For
        End If
    Next
    
    '目次シートが既にあれば削除しておく
    If Not shtIndex Is Nothing Then
        shtIndex.Activate
        If ConfirmMessage(IDX_SHT_NAME & "シートが既にあります。一旦削除してよろしいですか？") _
           <> VbMsgBoxResult.vbYes Then
            Exit Sub
        End If
        Application.DisplayAlerts = False
        shtIndex.Delete
        Application.DisplayAlerts = True
    End If
    
    '目次シート生成
    Set shtIndex = wbkActive.Worksheets.Add(Before:=wbkActive.Worksheets(1))
    shtIndex.Name = IDX_SHT_NAME
    
    Dim lngRow As Long
    lngRow = ROW_START
    
    For Each shtCur In wbkActive.Worksheets
        
        '非表示シートはスキップ
        If shtCur.Visible <> xlSheetVisible Then GoTo NEXT_SHEET
        
        Dim rngItem As Range
        Set rngItem = shtIndex.Cells(lngRow, COL)
        
        If shtCur Is shtIndex Then
            'ヘッダー行出力
            rngItem.Value = "名称"
        Else
            '明細行出力
            rngItem.Value = shtCur.Name
            Call shtIndex.Hyperlinks.Add(rngItem, "#" & shtCur.Name & "!$A$1")
            rngItem.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone
        End If
        
        lngRow = lngRow + 1
        
NEXT_SHEET:
    Next
    
    '全体の罫線
    Dim rngTable As Range
    Set rngTable = Range(shtIndex.Cells(ROW_START, COL), shtIndex.Cells(lngRow, COL))
    With rngTable.Borders
        .LineStyle = XlLineStyle.xlDot
        .ColorIndex = xlColorIndexAutomatic
    End With
    Call rngTable.BorderAround(XlLineStyle.xlContinuous, xlThin, XlColorIndex.xlColorIndexAutomatic)
    
    'ヘッダー行の罫線
    Dim rngHeadr As Range
    Set rngHeadr = rngTable.Resize(1, rngTable.Columns.Count)
    
    'ヘッダー行の色
    rngHeadr.Interior.Color = RGB(0, 112, 192)
    rngHeadr.Font.Color = ColorConstants.vbWhite
    
    '列幅自動調整
    Set rngTable = shtIndex.Columns(COL)
    Call rngTable.AutoFit
    
End Sub

'*****************************************************************************
'[ 関数名 ] RetryUnprotectBookForStruct
'[ 概  要 ] （シート構成の）ブック保護を解除できるまでPWD入力を繰り返す
'[ 引  数 ] 対象ブック，パスワード
'[ 戻り値 ] True:成功/False:失敗
'*****************************************************************************
Private Function RetryUnprotectBookForStruct(ByVal wbkTarget As Workbook, ByRef strPwd As String) As Boolean
    Dim varPwd As Variant
    
    Do While TryUnprotectBookForStruct(wbkTarget, strPwd) = False
        varPwd = Application.InputBox(Title:="シート保護解除パスワード入力", Prompt:="パスワードを入力して下さい。", Default:=strPwd)
        If varPwd = False Then
            RetryUnprotectBookForStruct = False
            Exit Function
        End If
        
        strPwd = varPwd
    Loop
    
    RetryUnprotectBookForStruct = True
    
End Function

'*****************************************************************************
'[ 関数名 ] TryUnprotectBookForStruct
'[ 概  要 ] （シート構成の）ブック保護解除を試す
'[ 引  数 ] 対象ブック，パスワード
'[ 戻り値 ] True:成功/False:失敗
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
'[ 関数名 ] RetryUnprotectSheet
'[ 概  要 ] シート保護を解除できるまでPWD入力を繰り返す
'[ 引  数 ] 対象シート，パスワード
'[ 戻り値 ] True:成功/False:失敗
'*****************************************************************************
Private Function RetryUnprotectSheet(ByVal shtTarget As Worksheet, ByRef strPwd As String) As Boolean
    Dim varPwd As Variant
    
    Do While TryUnprotectSheet(shtTarget, strPwd) = False
        varPwd = Application.InputBox(Title:="シート保護解除パスワード入力", Prompt:="パスワードを入力して下さい。", Default:=strPwd)
        If varPwd = False Then
            RetryUnprotectSheet = False
            Exit Function
        End If
        
        strPwd = varPwd
    Loop
    
    RetryUnprotectSheet = True
    
End Function

'*****************************************************************************
'[ 関数名 ] TryUnprotectSheet
'[ 概  要 ] シート保護解除を試す
'[ 引  数 ] 対象シート，パスワード
'[ 戻り値 ] True:成功/False:失敗
'*****************************************************************************
Private Function TryUnprotectSheet(ByVal shtTarget As Worksheet, ByVal strPwd As String) As Boolean
    
On Error GoTo Catch
    Call shtTarget.Unprotect(strPwd)
    TryUnprotectSheet = True
    Exit Function
    
Catch:
    TryUnprotectSheet = False
    
End Function

