Attribute VB_Name = "modSearchNext"
Option Explicit

'TODO:
'(1)A1セルから「前を検索」に未対応（同様に最終セルから「次を検索」も）
'   対応するなら対象シートリストを先に取得しておく必要あり
'   →特に困らないのでメッセージ表示対応のみしておいて、優先度低
'(2)以下のモードを可変にする仕組みが欲しい（設定画面？）
'・検索対象：値/表示テキスト/数式
'・一致判定：部分一致/完全一致
'・大文字/小文字/全角/半角の区別
'  ↑「区別しない」の場合、「は」を検索すると「ば」や「ぱ」もヒットする。
'    ちなみに「ば」を検索すると「は」や「ぱ」はヒットしない。
'    （InStrが内部的に半角変換してから検索してる？）

'##########
' 変数
'##########

'検索文字列
Private SearchText As String
'検索文字列(大文字/全角変換後)
Private SearchTextU As String

'大文字/小文字/全角/半角を区別するか？
Private CompareMethod As VbCompareMethod

'部分一致検索とするか？
Private PartialMatching As Boolean

'検索方向
Private SearchDirection As XlSearchDirection


'##########
' 関数
'##########

'検索文字列を設定して（デフォルト：現在セルテキスト）次を検索
Public Sub SearchNextThisCellText()
    
    Dim strSearchText As String
    Dim isCaseSensitive As Boolean
    Dim isPartialMatching As Boolean
    Dim isCancel As Boolean
    isCancel = False
    
    'フォーム表示
    Call Load(frmSearchNext)
    frmSearchNext.txtSearchText = CStr(ActiveCell.Value)
    If frmSearchNext.txtSearchText = Empty Then
        frmSearchNext.txtSearchText = SearchText
    End If
    frmSearchNext.txtSearchText.SelStart = 0    'テキスト全選択状態にする
    frmSearchNext.txtSearchText.SelLength = 256 '
    
    frmSearchNext.Show
    isCancel = frmSearchNext.isCancel
    strSearchText = frmSearchNext.txtSearchText.Text
    isCaseSensitive = frmSearchNext.chkCaseSensitive.Value
    isPartialMatching = frmSearchNext.chkPartialMatching.Value
    Call Unload(frmSearchNext)
    
    'キャンセル時は終了
    If isCancel Then
        Exit Sub
    End If
    
    SearchText = strSearchText
    SearchTextU = ConvUpperWideHiragana(SearchText)
    
    '未入力時は終了
    If SearchText = Empty Then
        Call modMessage.InfoMessage("検索文字列が入力されていません。")
        Exit Sub
    End If
    
    '大文字/小文字/全角/半角を区別するかどうか
    If isCaseSensitive Then
        CompareMethod = VbCompareMethod.vbBinaryCompare
    Else
        CompareMethod = VbCompareMethod.vbTextCompare
    End If
    
    '部分一致とするかどうか
    PartialMatching = isPartialMatching
    
    '次を検索
    Call SearchNextForward
    
End Sub

'次を検索
Public Sub SearchNextForward()
    SearchDirection = xlNext
    Call SearchNext_Main
End Sub

'前を検索
Public Sub SearchNextPrevious()
    SearchDirection = xlPrevious
    Call SearchNext_Main
End Sub

'「次を検索」メイン処理
Private Sub SearchNext_Main()
    
    If SearchText = "" Then
        Call modMessage.ErrorMessage("検索文字列が設定されていません。" & vbCrLf & _
                    "一旦、Ctrl+F3を実行して下さい。")
        Exit Sub
    End If
    
    '現在シート
    Dim shtNext As Worksheet
    Set shtNext = ActiveSheet
    
    '次に検索するセル
    Dim rngNextCell As Range
    
    '「次を」の場合:
    If SearchDirection = xlNext Then
        '最終列でなければ、ひとつ右のセル
        If ActiveCell.Column < shtNext.Columns.Count Then
            Set rngNextCell = ActiveCell.Offset(0, 1)
        '最終列だが、最終行ではない場合は、次行の1列目のセル
        ElseIf ActiveCell.Row < shtNext.Rows.Count Then
            Set rngNextCell = shtNext.Cells(ActiveCell.Row + 1, 1)
        '最終列で最終行の場合:
        Else
            '手抜き
            Call modMessage.ErrorMessage( _
                    "最終セルからの「次へ検索」は未対応ですm(__)m" & vbCrLf & _
                    "他のセルにカーソルを移動してから、実行して下さい。")
            Exit Sub
        End If

    '「前を」の場合:
    Else
        '1列目でなければひとつ左のセル
        If ActiveCell.Column > 1 Then
            Set rngNextCell = ActiveCell.Offset(0, -1)
        '1列目だが、1行目ではない場合は、前行の最終セル
        ElseIf ActiveCell.Row > 1 Then
            Set rngNextCell = shtNext.Cells(ActiveCell.Row - 1, _
                                            GetLastCol(ActiveCell))
        '1行目で1列目（＝A1セル）の場合:
        Else
            '手抜き
            Call modMessage.ErrorMessage( _
                    "A1セルからの「前へ検索」は未対応ですm(__)m" & vbCrLf & _
                    "他のセルにカーソルを移動してから、実行して下さい。")
            Exit Sub
        End If
    End If
    
    'ブック内検索
    Call SearchNext_Book(rngNextCell)
    
End Sub

'ブック内検索
Private Sub SearchNext_Book(ByRef rngNextCell As Range)
    
    'ブック内の最終セルに到達したかどうか
    Dim hasReachedLast As Boolean
    hasReachedLast = False
    
    '次に検索するシート
    Dim shtNext As Worksheet
    Set shtNext = rngNextCell.Parent
    
    '対象ブック
    Dim wbkTarget As Workbook
    Set wbkTarget = shtNext.Parent
    
    '検索対象シートリストの現在Index
    Dim lngSheetIndex As Long
    lngSheetIndex = 0
    
    '検索対象シートのリストを生成（＝非表示シート以外）
    Dim shtTarget As Worksheet
    Dim shtTargetList() As Worksheet
    Dim lngTargetListCount As Long
    lngTargetListCount = 0
    ReDim shtTargetList(0 To 0)
    For Each shtTarget In wbkTarget.Worksheets
        If shtTarget.Visible = xlSheetVisible Then
            'Call shtTargetList.Add(shtTarget)
            lngTargetListCount = lngTargetListCount + 1
            ReDim Preserve shtTargetList(0 To lngTargetListCount)
            Set shtTargetList(lngTargetListCount) = shtTarget
        End If
        'ついでに、上記リストにおける現シートのIndexを取得
        If shtTarget Is shtNext Then
            lngSheetIndex = lngTargetListCount
        End If
    Next
    
    Do
        'シート内最終セルを取得
        Dim rngLastCell As Range
        Set rngLastCell = modCommon.GetLastCell(shtNext)
        
        'シート内の最終行と最終列を取得
        Dim lngMaxRow As Long
        Dim lngMaxCol As Long
        lngMaxRow = rngLastCell.Row
        lngMaxCol = rngLastCell.Column
        
        'シート内を検索
        If SearchNext_Sheet(rngNextCell, lngMaxCol, lngMaxRow) Then
            '見つかったのなら終了
            Exit Sub
        End If
        
        '見つからなければ、次のシートの判定を行う
        
        '「次を」検索の場合：
        If SearchDirection = xlNext Then
            If lngSheetIndex < lngTargetListCount Then
                '最終シートでなければ、次のシート
                lngSheetIndex = lngSheetIndex + 1
            Else
                '既に最後に到達していたら、これ以上探しても見つからないので終了
                If hasReachedLast = True Then
                    Call modMessage.InfoMessage("見つかりませんでした。")
                    Exit Sub
                End If
                hasReachedLast = True
                
                '最後のシートまで見つからなかったら、最初のシートに戻るか確認
                If modMessage.ConfirmMessage( _
                            "最後のシートに到達しました。" & vbCrLf & _
                            "最初のシートに戻って検索しますか？") <> vbYes Then
                    Exit Sub
                End If
                lngSheetIndex = 1
            End If
            
            '次のシートのA1セルから検索
            Set shtNext = shtTargetList(lngSheetIndex)
            Set rngNextCell = shtNext.Cells(1, 1)
        
        '「前を」検索の場合：
        Else
            If lngSheetIndex > 1 Then
                '最初のシートでなければ、前のシート
                lngSheetIndex = lngSheetIndex - 1
            Else
                '既に最初に到達していたら、これ以上探しても見つからないので終了
                If hasReachedLast = True Then
                    Call modMessage.InfoMessage("見つかりませんでした。")
                    Exit Sub
                End If
                hasReachedLast = True
                
                '最初のシートまで見つからなかったら、最後のシートに戻るか確認
                If modMessage.ConfirmMessage( _
                            "最初のシートに到達しました。" & vbCrLf & _
                            "最後のシートに戻って検索しますか？") <> vbYes Then
                    Exit Sub
                End If
                lngSheetIndex = lngTargetListCount
            End If
            
            '前のシートの最終セルから検索
            Set shtNext = shtTargetList(lngSheetIndex)
            Set rngNextCell = modCommon.GetLastCell(shtNext)
        
        End If
        
    Loop
    
End Sub

'シート内検索
Private Function SearchNext_Sheet(ByRef rngNextCell As Range, ByVal lngMaxCol As Long, ByVal lngMaxRow As Long) As Boolean
    
    Do
        '行内を検索
        If SearchNext_Row(rngNextCell, lngMaxCol) Then
            SearchNext_Sheet = True
            Exit Function
        End If
        
        '「次を」検索の場合：
        If SearchDirection = xlNext Then
        
            If rngNextCell.Row >= lngMaxRow Then Exit Do
            
            'ひとつ下の行の最左セル
            Set rngNextCell = rngNextCell.Offset(1, 0).EntireRow _
                                .Cells(1, 1)
            
        '「前を」検索の場合：
        Else
            
            If rngNextCell.Row <= 1 Then Exit Do
            
            'ひとつ上の行の最終セル
            Set rngNextCell = rngNextCell.Offset(-1, 0).EntireRow _
                                .Cells(1, GetLastCol(rngNextCell))
        
        End If
    Loop
    
    SearchNext_Sheet = False
    
End Function

'行内検索
Private Function SearchNext_Row(ByRef rngNextCell As Range, ByVal lngMaxCol As Long) As Boolean
    
    Do
        If IsHit(rngNextCell) Then
            
            'セルを表示
            Call Application.GoTo(rngNextCell, False)
            
            SearchNext_Row = True
            Exit Function
        End If
        
        '「次を」検索の場合：
        If SearchDirection = xlNext Then
            
            If rngNextCell.Column >= lngMaxCol Then Exit Do
            
            'ひとつ右のセルへ
            Set rngNextCell = rngNextCell.Offset(0, 1)
            
        '「前を」検索の場合：
        Else
            
            If rngNextCell.Column <= 1 Then Exit Do
            
            'ひとつ左のセルへ
            Set rngNextCell = rngNextCell.Offset(0, -1)
            
        End If
    Loop
    
    SearchNext_Row = False
    
End Function

'ヒット判定
Private Function IsHit(ByRef rngNextCell As Range) As Boolean
    
    '返却値初期化
    IsHit = False
    
    'セル内テキスト
    Dim strText As String
    strText = CStr(rngNextCell.Value)
    
    '空のセルはスキップ
    If strText = "" Then
        Exit Function
    End If
    
    '部分一致するか？
    If PartialMatching Then
        If InStr(1, strText, SearchText, CompareMethod) > 0 Then
            IsHit = True
        End If
    Else
        If strText = SearchText Then
            IsHit = True
        Else
            If CompareMethod = vbTextCompare Then
                Dim strTextU As String
                strTextU = ConvUpperWideHiragana(strText)
                
                If strTextU = SearchTextU Then
                    IsHit = True
                End If
            End If
        End If
    End If
    
End Function

'大文字全角ひらがなに変換
Private Function ConvUpperWideHiragana(ByVal strText As String) As String
    ConvUpperWideHiragana = StrConv(strText, vbUpperCase + vbWide + vbHiragana)
End Function

'対象セルのシート内の最終列を取得
Private Function GetLastCol(ByVal rngTarget As Range) As Long
    Dim shtTarget As Worksheet
    Set shtTarget = rngTarget.Parent
    
    Dim rngUsedRange As Range
    Set rngUsedRange = shtTarget.UsedRange
    
    GetLastCol = rngUsedRange.Column + rngUsedRange.Columns.Count - 1
End Function

