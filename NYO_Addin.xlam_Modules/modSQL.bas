Attribute VB_Name = "modSQL"
Option Explicit

'*****************************************************************************
'[ 関数名 ] PasteSQLGrid
'[ 概  要 ] SQL Management Studio などのグリッドデータを整形して貼り付ける
'[ 引  数 ] なし
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub PasteSQLGrid()
    
    If Application.CutCopyMode <> 0 Then
        Call modMessage.ErrorMessage("Excelからの貼り付けは出来ません。")
        Exit Sub
    End If
    
On Error GoTo Catch
    
    'メイン処理
    PasteSQLGridMain
    
    GoTo Finally
    
Catch:
    Call modMessage.ErrorMessage("失敗しました。", Err)
    
Finally:
    Exit Sub
    
End Sub

'メイン処理
Private Sub PasteSQLGridMain()
    
    '現在のシートを対象にする
    Dim shtTarget As Worksheet
    Set shtTarget = ActiveSheet
    
    'まず貼り付け
    shtTarget.Paste
    
    '貼り付け範囲を取得
    Dim rngTable As Range
    Set rngTable = Selection
    
    '表示形式を「文字列」に設定
    rngTable.NumberFormatLocal = "@"
    
    '再度貼り付け
    shtTarget.Paste
    
    '格子状に罫線を引く
    With rngTable
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    '先頭行の背景色を設定
    With rngTable.Rows(1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(204, 255, 204) 'CCFFCC: 薄緑
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Trim
    Dim rngCell As Range
    For Each rngCell In rngTable
        Dim strValueOrg As String
        Dim strValue As String
        
        strValueOrg = rngCell.Text
        strValue = Trim(strValueOrg)
        
        If strValue <> strValueOrg Then
            rngCell.Value = strValue
        End If
        
    Next
    
    '列幅を自動調整
    rngTable.EntireColumn.AutoFit

End Sub

