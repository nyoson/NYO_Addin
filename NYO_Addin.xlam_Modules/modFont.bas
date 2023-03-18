Attribute VB_Name = "modFont"
Option Explicit

'*****************************************************************************
' 選択範囲の文字を赤字またはデフォルトにする
'*****************************************************************************
Public Sub SwitchRed()
    Dim fnt As Font
    Set fnt = Selection.Font
    
    If fnt.Color <> RGB(255, 0, 0) Then
        fnt.Color = RGB(255, 0, 0)
    Else
        fnt.ColorIndex = xlAutomatic
    End If
    
    'Debug.Print "SwitchRed()"
    
End Sub

'*****************************************************************************
' 選択範囲の文字装飾をクリアする
'*****************************************************************************
Public Sub ClearFont()
    
    '' セル以外は非対応
    'If Not TypeOf Selection Is Range Then
    '    Call modMessage.ErrorMessage("セルを選択して下さい")
    '    Exit Sub
    'End If
    '→非対応なのはフォント種類のみなので、コメントアウト
    
    Dim fnt As Font
    Set fnt = Selection.Font
    
    fnt.ColorIndex = xlAutomatic
    fnt.FontStyle = "標準"
    'fnt.Bold = False       'FontStyle = "標準" により、クリアされるので不要
    'fnt.Italic = False     'FontStyle = "標準" により、クリアされるので不要
    fnt.Name = Application.StandardFont '[オプション]の標準フォントを採用する
    'fnt.OutlineFont        'Windowsでは無効
    'fnt.Shadow             'Windowsでは無効
    fnt.Size = Application.StandardFontSize '[オプション]の標準フォントサイズを採用する
    fnt.Strikethrough = False
    fnt.Subscript = False
    fnt.Superscript = False
    fnt.Underline = XlUnderlineStyle.xlUnderlineStyleNone
    
End Sub

'*****************************************************************************
' ハイパーリンクスタイル変更
' （ハイパーリンク生成時のフォント名とサイズを、アクティブセルと同じにする）
'*****************************************************************************
Public Sub SetHyperLinkStyle()
    
    Dim rngCell As Range
    Set rngCell = ActiveCell
    
    Dim styleHyperLink As Style
    Set styleHyperLink = ActiveWorkbook.Styles("Hyperlink")
    
    styleHyperLink.IncludeFont = True
    styleHyperLink.Font.Name = rngCell.Font.Name
    styleHyperLink.Font.Size = rngCell.Font.Size
    
End Sub

