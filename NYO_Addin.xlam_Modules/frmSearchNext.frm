VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchNext 
   Caption         =   "次を検索"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   OleObjectBlob   =   "frmSearchNext.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSearchNext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'キャンセルフラグ
Public isCancel As Boolean

'フォームロード時
Private Sub UserForm_Initialize()
    isCancel = True
End Sub

'検索ボタン押下時
Private Sub btnSearch_Click()
    isCancel = False
    Call Me.Hide
End Sub

'キャンセルボタン押下時
Private Sub btnCancel_Click()
    Call Me.Hide
End Sub


