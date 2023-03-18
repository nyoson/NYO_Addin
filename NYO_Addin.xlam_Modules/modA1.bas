Attribute VB_Name = "modA1"
Option Explicit

'*****************************************************************************
'[ 関数名 ]　ActivateAllA1Cell
'[ 概  要 ]　すべてのシートでA1セルを選択状態にする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub ActivateAllA1Cell()
    
    '初めの表示シート
    Dim shtFirstVisibleSheet As Worksheet
    Set shtFirstVisibleSheet = Nothing
    
    Dim shtCurSheet As Worksheet
    For Each shtCurSheet In ActiveWorkbook.Worksheets
        
        '非表示シートはスキップする
        If shtCurSheet.Visible <> xlSheetVisible Then GoTo NEXT_SHEET
            
        '初めの表示シートのシートオブジェクトを覚えておく
        If shtFirstVisibleSheet Is Nothing Then
            Set shtFirstVisibleSheet = shtCurSheet
        End If
        
        'シートをアクティブにする
        shtCurSheet.Activate
        
        '可視セル内で一番左上セルを選択状態にする(A1セルとは限らない)
        '※ウィンドウ枠固定時など、Ctrl+Homeの結果と異なるケースあり
        shtCurSheet.Cells.SpecialCells(xlVisible).Item(1).Activate
        
        'スクロール位置を一番左上にする
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        
NEXT_SHEET:
    Next
    
    '念のためnullチェック
    If shtFirstVisibleSheet Is Nothing Then Exit Sub
    
    '最初の表示シートをアクティブにして終了する
    Call shtFirstVisibleSheet.Activate
    
End Sub

