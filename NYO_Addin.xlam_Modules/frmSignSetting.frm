VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSignSetting 
   Caption         =   "電子印-設定"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   OleObjectBlob   =   "frmSignSetting.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSignSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'##########
' イベント
'##########

'*****************************************************************************
'[ イベント名 ] フォームロード時
'*****************************************************************************
Private Sub UserForm_Initialize()

    '電子印設定取得
    Dim objSignData As SignData
    Call GetSignSetting(objSignData)
    
    '画面に反映
    txtNameTop.Text = objSignData.Name1
    txtNameBottom = objSignData.Name2
    txtSignDate = Format(objSignData.SignDate, "yy/mm/dd")
    
End Sub

'*****************************************************************************
'[ イベント名 ] 標準に設定 ボタン押下時
'*****************************************************************************
Private Sub btnSave_Click()
    
    '画面の入力内容から電子印データ生成
    Dim objSignData As SignData
    Call MakeSignData(objSignData)
    
    '電子印設定保存
    Call SetSignSetting(objSignData)
    
End Sub

'*****************************************************************************
'[ イベント名 ] 試しに押印 ボタン押下時
'*****************************************************************************
Private Sub btnTmpSign_Click()
    
    '電子印設定取得
    Dim objSignData As SignData
    Call MakeSignData(objSignData)
    
    '電子印を作成
    Call MakeSign(objSignData)
    
End Sub

'*****************************************************************************
'[ イベント名 ] 閉じる ボタン押下時
'*****************************************************************************
Private Sub btnClose_Click()
    
    Call Me.Hide
    
End Sub

'##########
' 関数
'##########

'*****************************************************************************
'[ 関数名 ] MakeSignData
'[ 概  要 ] 画面の入力内容から電子印データを生成する
'[ 引  数 ] [Ref]電子印データ
'[ 戻り値 ] なし
'*****************************************************************************
Private Sub MakeSignData(ByRef objSignData As SignData)

    objSignData.Name1 = txtNameTop.Text
    objSignData.Name2 = txtNameBottom.Text
    objSignData.SignDate = CDate(txtSignDate.Text)
    
End Sub

