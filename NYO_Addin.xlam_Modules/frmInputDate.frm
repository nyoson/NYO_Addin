VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInputDate 
   Caption         =   "日付入力補助"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   OleObjectBlob   =   "frmInputDate.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmInputDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' 【定数】
'=============================================================================
' 入力キーのID
Private Const KeyPAGE_UP As Integer = 33    ' PageUpキー
Private Const KeyPAGE_DOWN As Integer = 34  ' PageDownキー
Private Const KeyLEFT As Integer = 37       ' Left(←)キー
Private Const KeyUP As Integer = 38         ' Up(↑)キー
Private Const KeyRIGHT As Integer = 39      ' Right(→)キー
Private Const KeyDOWN As Integer = 40       ' Down(↓)キー
Private Const KeyESC As Integer = 27        ' エスケープ(Esc)キー

'=============================================================================
' 【変数】
'=============================================================================
Private datInputDate As Date                ' 入力値(日付)
Private isInputCancel As Boolean            ' 入力キャンセル有無

'=============================================================================
' 【イベント】
'=============================================================================

' フォーム生成時
Private Sub UserForm_Initialize()

    isInputCancel = True
    cmdOk.SetFocus
    
End Sub

' キー押下時
Private Sub cmdOk_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    'Debug.Print KeyCode.Value
    
    ' 表示値を変数に設定
    Call GetTextTargetDate
    
    Select Case KeyCode
        Case KeyLEFT
            ReduceDay1
        Case KeyUP
            ReduceWeek1
        Case KeyRIGHT
            AddDay1
        Case KeyDOWN
            AddWeek1
            
        Case KeyPAGE_UP
            ' Ctlキーと同時入力の場合、年を減算
            ' Ctlキーが入力されていない場合、月を減算
            If Shift = 2 Then
                ReduceYear1
            Else
                ReduceMonth1
            End If
            
        Case KeyPAGE_DOWN
            ' Ctlキーと同時入力の場合、年を加算
            ' Ctlキーが入力されていない場合、月を加算
            If Shift = 2 Then
                AddYear1
            Else
                AddMonth1
            End If

        Case KeyESC
            ' エスケープキー押下時は終了
            isInputCancel = True
            Me.Hide

        Case Else
    
    End Select

    ' 変数の値をラベルに表示
    Call SetTextTargetDate

    cmdOk.SetFocus

End Sub

' OK押下時
'  表示している日付を入力補助の変数に設定
Private Sub cmdOk_Click()
    Call GetTextTargetDate
    isInputCancel = False
    Me.Hide
End Sub

'=============================================================================
'【Public関数】
'=============================================================================

'*****************************************************************************
'[ 関数名 ]　GetInputDate
'[ 概  要 ]　入力値(日付型)を返す
'[ 引  数 ]　-
'[ 戻り値 ]  入力値(日付型)
'*****************************************************************************
Public Function GetInputDate()
    GetInputDate = datInputDate
End Function

'*****************************************************************************
'[ 関数名 ]　SetInputDate
'[ 概  要 ]　日付を入力値に設定、表示する
'[ 引  数 ]　日付型のデータ
'[ 戻り値 ]　-
'*****************************************************************************
Public Function SetInputDate(ByVal datTargetDate As Date)
    datInputDate = datTargetDate
    
    ' 変数の値をラベルに表示
    Call SetTextTargetDate
End Function

'*****************************************************************************
'[ 関数名 ]　GetInputCancel
'[ 概  要 ]　入力キャンセル有無を返す
'[ 引  数 ]　-
'[ 戻り値 ]  入力キャンセル [有:True 無:False]
'*****************************************************************************
Public Function GetInputCancel()
    GetInputCancel = isInputCancel
End Function

'=============================================================================
'【Private関数】
'=============================================================================

'*****************************************************************************
'[ 関数名 ]　GetTextTargetDate
'[ 概  要 ]　現在表示値を取得(入力値に設定)する
'            - 表示値が異常の場合、現在日付を設定
'[ 理  由 ]　変数との不整合をなくす為
'[ 引  数 ]　-
'[ 戻り値 ]　-
'*****************************************************************************
Private Sub GetTextTargetDate()

    If IsDate(lblTargetDate.Caption) Then
        datInputDate = CDate(lblTargetDate.Caption)
    Else
        datInputDate = Date
        Call SetTextTargetDate
    End If
    
End Sub

'*****************************************************************************
'[ 関数名 ]　SetTextTargetDate
'[ 概  要 ]　入力値と入力値に対する曜日を表示する
'[ 引  数 ]　-
'[ 戻り値 ]　-
'*****************************************************************************
Private Sub SetTextTargetDate()
    ' 日付表示
    lblTargetDate.Caption = Format(datInputDate, "yyyy/mm/dd")
    ' 曜日表示
    lblDayOfWeek.Caption = Format(datInputDate, "(aaa)")
End Sub

'*****************************************************************************
'[ 関数名 ]　AddDay1
'[ 概  要 ]　入力値と表示値(日付)に１日加算する
'[ 引  数 ]　-
'[ 戻り値 ]　-
'*****************************************************************************
Private Sub AddDay1()
    datInputDate = DateAdd("d", 1, datInputDate)
End Sub

'*****************************************************************************
'[ 関数名 ]　ReduceDay1
'[ 概  要 ]　入力値と表示値(日付)から１日減算する
'[ 引  数 ]　-
'[ 戻り値 ]　-
'*****************************************************************************
Private Sub ReduceDay1()
    datInputDate = DateAdd("d", -1, datInputDate)
End Sub

'*****************************************************************************
'[ 関数名 ]　AddWeek1
'[ 概  要 ]　入力値と表示値(日付)に１週間(７日)加算する
'[ 引  数 ]　-
'[ 戻り値 ]　-
'*****************************************************************************
Private Sub AddWeek1()
    datInputDate = DateAdd("d", 7, datInputDate)
End Sub

'*****************************************************************************
'[ 関数名 ]　ReduceWeek1
'[ 概  要 ]　入力値と表示値(日付)から１週間(７日)減算する
'[ 引  数 ]　-
'[ 戻り値 ]　-
'*****************************************************************************
Private Sub ReduceWeek1()
    datInputDate = DateAdd("d", -7, datInputDate)
End Sub

'*****************************************************************************
'[ 関数名 ]　AddMonth1
'[ 概  要 ]　入力値と表示値(日付)に１ヶ月加算する
'[ 引  数 ]　-
'[ 戻り値 ]　-
'*****************************************************************************
Private Sub AddMonth1()
    datInputDate = DateAdd("m", 1, datInputDate)
End Sub

'*****************************************************************************
'[ 関数名 ]　ReduceMonth1
'[ 概  要 ]　入力値と表示値(日付)から１ヶ月減算する
'[ 引  数 ]　-
'[ 戻り値 ]　-
'*****************************************************************************
Private Sub ReduceMonth1()
    datInputDate = DateAdd("m", -1, datInputDate)
End Sub

'*****************************************************************************
'[ 関数名 ]　AddYear1
'[ 概  要 ]　入力値と表示値(日付)に１年加算する
'[ 引  数 ]　-
'[ 戻り値 ]　-
'*****************************************************************************
Private Sub AddYear1()
    datInputDate = DateAdd("yyyy", 1, datInputDate)
End Sub

'*****************************************************************************
'[ 関数名 ]　ReduceYear1
'[ 概  要 ]　入力値と表示値(日付)から１年減算する
'[ 引  数 ]　-
'[ 戻り値 ]　-
'*****************************************************************************
Private Sub ReduceYear1()
    datInputDate = DateAdd("yyyy", -1, datInputDate)
End Sub

