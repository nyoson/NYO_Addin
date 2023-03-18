Attribute VB_Name = "modSign"
Option Explicit

'##########
' 定数
'##########

Private Const SIGN_SIZE As Single = 50#
Private Const TEXT_HEIGHT As Single = 14#
Private Const DATE_WIDTH As Single = 48#
Private Const NAME_WIDTH As Single = 30#
Private Const DATE_FONT_SIZE As Double = 10#

'##########
' 構造体
'##########

'電子印データ
Public Type SignData
    Name1 As String
    Name2 As String
    SignDate As Date
End Type

'##########
' 関数
'##########

'*****************************************************************************
'[ 関数名 ] ShowSignSetting
'[ 概  要 ] 電子印設定画面を表示する
'[ 引  数 ] なし
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub ShowSignSetting()
    
    Call Load(frmSignSetting)
    Call frmSignSetting.Show
    Call Unload(frmSignSetting)
    
End Sub

'*****************************************************************************
'[ 関数名 ] GetSignSetting
'[ 概  要 ] レジストリから電子印設定を取得する
'[ 引  数 ] [Out]電子印設定
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub GetSignSetting(ByRef objSignData As SignData)
    
    '電子印設定をレジストリから取得
    objSignData.Name1 = GetSetting(C_TOOLBAR_NAME, "Sign", "Name1")
    objSignData.Name2 = GetSetting(C_TOOLBAR_NAME, "Sign", "Name2")
    
    '本日日付をセット
    objSignData.SignDate = Date
    
End Sub

'*****************************************************************************
'[ 関数名 ] SetSignSetting
'[ 概  要 ] レジストリに電子印設定を保存する
'[ 引  数 ] 電子印設定
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub SetSignSetting(ByRef objSignData As SignData)
    
    '電子印設定をレジストリに保存
    Call SaveSetting(C_TOOLBAR_NAME, "Sign", "Name1", objSignData.Name1)
    Call SaveSetting(C_TOOLBAR_NAME, "Sign", "Name2", objSignData.Name2)
    
End Sub

'*****************************************************************************
'[ 関数名 ] Sign
'[ 概  要 ] 標準設定で選択セルに電子印を作成する
'[ 引  数 ] なし
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub Sign()
    
    '電子印設定取得
    Dim objSignData As SignData
    Call GetSignSetting(objSignData)
    
    '選択セルに電子印を作成する
    Call MakeSign(objSignData)
    
End Sub

'*****************************************************************************
'[ 関数名 ] MakeSign
'[ 概  要 ] 指定された設定で選択セルに電子印を作成する
'[ 引  数 ] 電子印設定
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub MakeSign(ByRef objSignData As SignData)
    
    '対象シート
    Dim shtTarget As Worksheet
    Set shtTarget = ActiveSheet
    
    'シェイプリスト
    Dim shpList As Shapes
    Set shpList = shtTarget.Shapes
    
    '円
    Dim shpCircle As Shape
    Set shpCircle = shpList.AddShape(msoShapeOval, _
                                0, _
                                0, _
                                SIGN_SIZE, SIGN_SIZE)
    shpCircle.Line.ForeColor.RGB = vbRed
    shpCircle.Line.Weight = 1           '線の太さを明示的に指定
    'shpCircle.Fill.Visible = msoFalse   '塗りつぶしなしを明示的に指定
    '白色で塗りつぶし
    shpCircle.Fill.Visible = msoTrue
    shpCircle.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    '日付テキスト
    Dim shpDate As Shape
    Set shpDate = shpList.AddTextbox(msoTextOrientationHorizontal, _
                                1, _
                                SIGN_SIZE / 2 - TEXT_HEIGHT / 2, _
                                DATE_WIDTH, TEXT_HEIGHT)
    shpDate.Line.Visible = msoFalse
    shpDate.TextFrame.HorizontalAlignment = xlHAlignCenter
    shpDate.TextFrame.VerticalAlignment = xlVAlignCenter
    shpDate.TextFrame.Characters.Text = Format(objSignData.SignDate, "yy/mm/dd")
    shpDate.TextFrame.Characters.Font.Color = vbRed
    shpDate.TextFrame.Characters.Font.Size = DATE_FONT_SIZE
    shpDate.TextFrame.MarginLeft = 0
    shpDate.TextFrame.MarginRight = 0
    shpDate.TextFrame.MarginTop = 0
    shpDate.TextFrame.MarginBottom = 0
    shpDate.Fill.Visible = msoFalse
    
    '日付の上下の線
    Dim shpLine1, shpLine2 As Shape
    Set shpLine1 = shpList.AddLine( _
                                (SIGN_SIZE - DATE_WIDTH) / 2, _
                                SIGN_SIZE / 2 - TEXT_HEIGHT / 2, _
                                SIGN_SIZE - (SIGN_SIZE - DATE_WIDTH) / 2, _
                                SIGN_SIZE / 2 - TEXT_HEIGHT / 2)
    shpLine1.Line.ForeColor.RGB = vbRed
    Set shpLine2 = shpList.AddLine( _
                                (SIGN_SIZE - DATE_WIDTH) / 2, _
                                SIGN_SIZE / 2 + TEXT_HEIGHT / 2, _
                                SIGN_SIZE - (SIGN_SIZE - DATE_WIDTH) / 2, _
                                SIGN_SIZE / 2 + TEXT_HEIGHT / 2)
    shpLine2.Line.ForeColor.RGB = vbRed
    
    '名前のフォントサイズ算出
    Dim intNameCharWidth As Integer '半角何文字分か求める
    intNameCharWidth = WorksheetFunction.Max( _
        LenB(StrConv(objSignData.Name1, vbFromUnicode)), _
        LenB(StrConv(objSignData.Name2, vbFromUnicode)))
    Dim sglFontSize As Single
    sglFontSize = 11# * 4 / intNameCharWidth    '半角4文字で11.0が最適⇒その比率で求める
    
    '名前1
    Dim shpName1 As Shape
    Set shpName1 = shpList.AddTextbox(msoTextOrientationHorizontal, _
                                (SIGN_SIZE - NAME_WIDTH) / 2, _
                                SIGN_SIZE / 2 - TEXT_HEIGHT / 2 - TEXT_HEIGHT, _
                                NAME_WIDTH, TEXT_HEIGHT)
    shpName1.Line.Visible = msoFalse
    shpName1.TextFrame.HorizontalAlignment = xlHAlignCenter
    shpName1.TextFrame.VerticalAlignment = xlVAlignCenter
    shpName1.TextFrame.Characters.Text = objSignData.Name1
    shpName1.TextFrame.Characters.Font.Color = vbRed
    shpName1.TextFrame.Characters.Font.Size = sglFontSize
    shpName1.TextFrame.MarginLeft = 0
    shpName1.TextFrame.MarginRight = 0
    shpName1.TextFrame.MarginTop = 0
    shpName1.TextFrame.MarginBottom = 0
    shpName1.Fill.Visible = msoFalse
    
    '名前2
    Dim shpName2 As Shape
    Set shpName2 = shpList.AddTextbox(msoTextOrientationHorizontal, _
                                (SIGN_SIZE - NAME_WIDTH) / 2, _
                                SIGN_SIZE / 2 + TEXT_HEIGHT / 2, _
                                NAME_WIDTH, TEXT_HEIGHT)
    shpName2.Line.Visible = msoFalse
    shpName2.TextFrame.HorizontalAlignment = xlHAlignCenter
    shpName2.TextFrame.VerticalAlignment = xlVAlignCenter
    shpName2.TextFrame.Characters.Text = objSignData.Name2
    shpName2.TextFrame.Characters.Font.Color = vbRed
    shpName2.TextFrame.Characters.Font.Size = sglFontSize
    shpName2.TextFrame.MarginLeft = 0
    shpName2.TextFrame.MarginRight = 0
    shpName2.TextFrame.MarginTop = 0
    shpName2.TextFrame.MarginBottom = 0
    shpName2.Fill.Visible = msoFalse
    
    Dim varShapeNameList(1 To 6) As Variant
    varShapeNameList(1) = shpCircle.Name
    varShapeNameList(2) = shpDate.Name
    varShapeNameList(3) = shpLine1.Name
    varShapeNameList(4) = shpLine2.Name
    varShapeNameList(5) = shpName1.Name
    varShapeNameList(6) = shpName2.Name
    
    Dim shpSign As Shape
    Set shpSign = shpList.Range(varShapeNameList).Group
    
    Call shpSign.Copy
    
    
    '選択セル範囲エリアリスト
    Dim rangeAreaList As Areas
    Set rangeAreaList = ActiveWindow.RangeSelection.Areas
    
    '選択セル範囲エリアごとに作成
    Dim rangeArea As Range
    For Each rangeArea In rangeAreaList
        
        '選択セル範囲エリアの最初と最後のセル(の右下セル)を取得
        '(選択セル範囲の中央に配置するため)
        Dim rngTarget1 As Range
        Set rngTarget1 = rangeArea.Item(1)
        Dim rngTarget2 As Range
        Set rngTarget2 = rangeArea.Item(rangeArea.Count).Offset(1, 1)
        
        '電子印を画像貼り付け
        Call shtTarget.PasteSpecial(Format:="図 (GIF)")
        'ActiveCell.Offset(5, 0).Activate
        'Call shtTarget.PasteSpecial(Format:="図 (PNG)")    'GIFと特に差なし
        'ActiveCell.Offset(5, 0).Activate
        'Call shtTarget.PasteSpecial(Format:="図 (拡張メタファイル)") '見た目イマイチ
        'ActiveCell.Offset(5, 0).Activate
        'Call shtTarget.PasteSpecial(Format:="MS Office 描画オブジェクト")   'オートシェイプとしてコピーされるためNG
        'ActiveCell.Offset(5, 0).Activate
        'Call shtTarget.PasteSpecial(Format:="図 (JPEG)") '透過されないためNG
        
        '電子印画像
        Dim picSign As Picture
        Set picSign = Selection
        
        '選択セルの中央に配置
        picSign.Left = ((rngTarget1.Left + rngTarget2.Left - picSign.Width) / 2)
        picSign.Top = ((rngTarget1.Top + rngTarget2.Top - picSign.Height) / 2)
        
        'セルに合わせて移動，サイズ変更する設定に変更する
        picSign.Placement = XlPlacement.xlMoveAndSize
    
    Next
    
    '最終的に中間データオブジェクトを削除
    Call shpSign.Delete
    
    '選択状態解除
    rngTarget1.Activate
    
End Sub
