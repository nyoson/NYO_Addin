Attribute VB_Name = "modCrLf"
Option Explicit

'*****************************************************************************
' 選択範囲の各セルの末尾に改行を追加する（既にあれば何もしない）
'*****************************************************************************
Public Sub AddCrLf()
    Call EditCrLf(True)
End Sub

'*****************************************************************************
' 選択範囲の各セルの末尾の改行を削除する（既になければ何もしない）
'*****************************************************************************
Public Sub RemoveCrLf()
    Call EditCrLf(False)
End Sub

'改行操作
Private Sub EditCrLf(ByVal flgAdd As Boolean)
    Dim rngCell As Range
    
    'UsedRange最終セル位置を取得
    Dim rngLastCell As Range
    Set rngLastCell = modCommon.GetLastCell(ActiveSheet)
    Dim lngLastRow As Long
    lngLastRow = rngLastCell.Row
    Dim lngLastCol As Long
    lngLastCol = rngLastCell.Column
    
    For Each rngCell In Selection
        
        'UsedRangeを超えたら中断
        If rngCell.Row > lngLastRow Or _
           rngCell.Column > lngLastCol Then
            Call rngCell.Activate
            Call modMessage.InfoMessage("UsedRangeを超えたので中断しました。")
            Call rngCell.Activate
            Exit For
        End If
        
        '結合セルの場合、左上セル以外はスキップ
        If rngCell.MergeArea(1, 1).Address <> rngCell.Address Then
            GoTo NEXT_CELL
        End If
        
        'セルの値が空文字ならスキップ
        If rngCell.Text = "" Then
            GoTo NEXT_CELL
        End If
        
        '末尾の1文字
        Dim lastChar As String
        lastChar = Right(rngCell.Value, 1)
        
        If flgAdd Then
            '追加モード：末尾に改行を追加する（既にあれば何もしない）
            If lastChar <> vbLf Then
                rngCell.Value = rngCell.Value + vbLf
            End If
        Else
            '削除モード：末尾の改行を削除する（既になければ何もしない）
            If lastChar = vbLf Then
                rngCell.Value = Left(rngCell.Value, Len(rngCell.Value) - 1)
            End If
        End If

NEXT_CELL:
    Next

End Sub
