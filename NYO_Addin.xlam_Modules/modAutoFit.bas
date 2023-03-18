Attribute VB_Name = "modAutoFit"
Option Explicit

'*****************************************************************************
'[ 関数名 ]　AutoFit
'[ 概  要 ]　オートシェイプの自動サイズ調整設定を切り替える
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub AutoFit()
    ' エラーは無視する
    On Error Resume Next
    
    Dim idxMax As Integer
    Dim idx As Integer
    Dim flg As Boolean
    Dim button As CommandBarButton
    Set button = Application.CommandBars(C_TOOLBAR_NAME).Controls("オートシェイプの自動サイズ調整")
    
    ' コマンドバーボタンのステータスをリセットしておく
    button.State = msoButtonUp

    ' セルが選択されている場合は、何もしない
    If TypeOf Selection Is Range Then Exit Sub
    
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    idxMax = sr.Count
    If idxMax > 0 Then
        ' ひとつ目のオートシェイプの自動サイズ調整フラグを反転し、選択されたすべてのオートシェイプに反映
        flg = Not sr(1).DrawingObject.AutoSize
        For idx = 1 To idxMax
            ' 引き出し線の長さが勝手に変わってしまうのを防ぐために、
            ' 引き出し線の長さを固定にしておく⇒効かない…【課題】
            Dim calloutFmt As CalloutFormat
            Set calloutFmt = sr(idx).Callout
            calloutFmt.CustomLength calloutFmt.Length
            
            sr(idx).DrawingObject.AutoSize = flg
            
            If flg Then
                ' テキストを図形からはみ出して表示する
                Dim tf As TextFrame
                Set tf = sr(idx).TextFrame
                tf.HorizontalOverflow = xlOartHorizontalOverflowOverflow '水平方向
                tf.VerticalOverflow = xlOartVerticalOverflowOverflow '垂直方向
            
            End If
            
            ' 引き出し線の長さを自動にする⇒【課題】
            calloutFmt.AutomaticLength
            
        Next
        
        ' コマンドバーボタンのステータスを変更⇒本当はSelectionChangeイベントでも実施したいが、方法が不明…【課題】
        If flg Then
            button.State = msoButtonDown
        Else
            button.State = msoButtonUp
        End If
        
    End If
    
    On Error GoTo 0
End Sub

