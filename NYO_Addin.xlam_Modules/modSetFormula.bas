Attribute VB_Name = "modSetFormula"
Option Explicit

'シート名末尾9文字(例：YYYY_MMDD)を取得する数式
Private Const FORMULA_GET_SHEET_NAME_LAST9 As String = "RIGHT(CELL(""filename"", $A$1),9)"

'*****************************************************************************
'[ 関数名 ]　InputSheetName
'[ 概  要 ]　シート名をアクティブセルにセットする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub InputSheetName()
    Select Case SelectFormulaMode()
        Case VbMsgBoxResult.vbYes
            ActiveCell.Formula = "=MID(CELL(""filename"", $A$1),SEARCH(""]"",CELL(""filename"", $A$1))+1,255)"
        Case VbMsgBoxResult.vbNo
            ActiveCell.Value = ActiveCell.Worksheet.Name
        Case Else
            
    End Select
    
End Sub

'*****************************************************************************
'[ 関数名 ]　InputSheetName_YYYY_MMDD
'[ 概  要 ]　シート名(YYYY_MMDD形式)を取得する数式をアクティブセルにセットする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub InputSheetName_YYYY_MMDD()
    ActiveCell.Formula = "=" & FORMULA_GET_SHEET_NAME_LAST9
End Sub

'*****************************************************************************
'[ 関数名 ]　InputSheetName_YYYY_MMDD_AsDate
'[ 概  要 ]　シート名(YYYY_MMDD形式)を日付値として取得する数式をアクティブセルにセットする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub InputSheetName_YYYY_MMDD_AsDate()
    ActiveCell.Formula = "=TEXT(REPLACE(" & FORMULA_GET_SHEET_NAME_LAST9 & ",5,1,""""), ""0000!/00!/00"")-0"
End Sub

'*****************************************************************************
'[ 関数名 ]　InputSeqNo
'[ 概  要 ]　連番を取得する数式をアクティブセルにセットする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub InputSeqNo()
    
    Dim rngCell As Range
    
    If TypeName(Selection) <> "Range" Then
        Call modMessage.ErrorMessage("セルを選択して下さい")
        Exit Sub
    End If
    
    Select Case SelectFormulaMode()
        Case VbMsgBoxResult.vbYes
            For Each rngCell In Selection
                '結合セルの左上セル以外はスキップし、数式をセット
                If rngCell.Address = rngCell.MergeArea.Cells(1, 1).Address Then
                    'rngCell.Formula = "=N(INDIRECT(""R[-1]C"",FALSE))+1"
                    rngCell.Formula = "=N(OFFSET(" & rngCell.Address(False, False) & ",-1,0))+1"
                End If
            Next
            
        Case VbMsgBoxResult.vbNo
            Dim lngCnt As Long
            lngCnt = 0
            For Each rngCell In Selection
                '結合セルの左上セル以外はスキップし、カウントアップ
                If rngCell.Address = rngCell.MergeArea.Cells(1, 1).Address Then
                    lngCnt = lngCnt + 1
                    rngCell.Value = lngCnt
                End If
            Next
            
        Case Else
            
    End Select
    
End Sub

'*****************************************************************************
'[ 関数名 ]　SelectFormulaMode
'[ 概  要 ]　数式で入力するか確認メッセージを表示する
'[ 引  数 ]　なし
'[ 戻り値 ]　vbYes/vbNo/vbCancel
'*****************************************************************************
Private Function SelectFormulaMode() As VbMsgBoxResult
    SelectFormulaMode = modMessage.ConfirmMessage( _
            "数式で入力しますか？（いいえを選択すると値として入力されます。）", _
            vbYesNoCancel)
End Function
