Attribute VB_Name = "modSheet"
Option Explicit

'##########
' 関数
'##########

'*****************************************************************************
'[ 関数名 ]　SheetDispChanger
'[ 概  要 ]　シートの表示/非表示を切り替える
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SheetDispChanger()

On Error GoTo Catch
    Load frmSheetDispChanger
    frmSheetDispChanger.Show
    GoTo Finally
Catch:
    Call modMessage.ErrorMessage("画面起動中にエラーが発生しました。", Err)
Finally:
    Unload frmSheetDispChanger
    
End Sub

'*****************************************************************************
' R1C1形式切り替え
'*****************************************************************************
Public Sub SwitchR1C1()
    If Application.ReferenceStyle <> xlR1C1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub

'*****************************************************************************
' アウトラインの集計行を左，集計列を上にする
'*****************************************************************************
Sub OutlineConfigAboveLeft()
    With ActiveSheet.Outline
        .SummaryRow = xlAbove
        .SummaryColumn = xlLeft
    End With
End Sub

