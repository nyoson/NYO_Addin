Attribute VB_Name = "modRow"
Option Explicit

'##########
' 関数
'##########

'*****************************************************************************
'[ 関数名 ]　AddCopyRow
'[ 概  要 ]　選択セルの行をコピーして下に追加
'[ 引  数 ]　なし
'[ 戻り値 ]  なし
'*****************************************************************************
Public Sub AddCopyRow()
    
    Dim rngSrc As Range
    Set rngSrc = ActiveCell.EntireRow
    
    Call rngSrc.Copy
    Call rngSrc.Offset(1, 0).Insert(XlInsertShiftDirection.xlShiftDown)
    
    Application.CutCopyMode = False
    
    Call ActiveCell.Offset(1, 0).Activate
    
End Sub

'*****************************************************************************
'[ 関数名 ]　DelRow
'[ 概  要 ]　選択セルの行を削除
'[ 引  数 ]　なし
'[ 戻り値 ]  なし
'*****************************************************************************
Public Sub DelRow()
    
    Dim rngSrc As Range
    Set rngSrc = ActiveWindow.RangeSelection

    Dim rngRows As Range
    Set rngRows = rngSrc.EntireRow
    
    Call rngRows.Delete(XlDirection.xlUp)
    
    Call ActiveWindow.RangeSelection.Cells(1, 1).Select
    
End Sub

'*****************************************************************************
'[ 関数名 ]　UpRow
'[ 概  要 ]　選択セルの行を上に移動
'[ 引  数 ]　なし
'[ 戻り値 ]  なし
'*****************************************************************************
Public Sub UpRow()
    
    Dim rngSrc As Range
    Set rngSrc = ActiveWindow.RangeSelection

    Dim rngRows As Range
    Set rngRows = rngSrc.EntireRow
    
    '先頭行の場合は何もせず終了
    If rngRows.Row <= 1 Then Exit Sub
    
    Call rngRows.Cut
    Call rngRows.Offset(-1, 0).Insert
    
    Call rngRows.Select
    
End Sub

'*****************************************************************************
'[ 関数名 ]　DownRow
'[ 概  要 ]　選択セルの行を下に移動
'[ 引  数 ]　なし
'[ 戻り値 ]  なし
'*****************************************************************************
Public Sub DownRow()
    
    Dim rngSrc As Range
    Set rngSrc = ActiveWindow.RangeSelection

    Dim rngRows As Range
    Set rngRows = rngSrc.EntireRow
    
    '最終行の場合は何もせず終了
    If rngRows.Row + rngRows.Rows.Count - 1 >= ActiveSheet.Rows.Count Then Exit Sub
    
    Call rngRows.Cut
    Call rngRows.Offset(rngRows.Rows.Count + 1, 0).Insert
    
    Call rngRows.Select
    
End Sub

