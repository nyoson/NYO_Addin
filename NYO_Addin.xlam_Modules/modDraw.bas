Attribute VB_Name = "modDraw"
Option Explicit

'*****************************************************************************
' ×印を描画する
'*****************************************************************************
Public Sub DrawX()
    
    Dim rngCur1 As Range
    Dim rngCur2 As Range
    
    Dim shpList As Shapes
    Dim shpLines(1 To 2) As Shape
    Dim varLineNames(1 To 2) As Variant
    
    Dim i As Long
    
    'セル以外を選択している場合は、選択解除
    ActiveCell.Activate
    
    '選択セル範囲を取得(A1:B2の場合は、A1とC3を取得)
    Set rngCur1 = Selection.Areas(1)
    Set rngCur2 = rngCur1.Offset(rngCur1.Rows.Count, rngCur1.Columns.Count)
    Set rngCur1 = rngCur1.Cells(1, 1)
    
    '選択範囲の左上セルのみを選択し直す
    rngCur1.Select
    
    'シェイプオブジェクトリスト
    Set shpList = ActiveSheet.Shapes
    
    '線１（左上から右下）を作成
    Set shpLines(1) = shpList.AddLine( _
                            rngCur1.Left, rngCur1.Top, _
                            rngCur2.Left, rngCur2.Top)
    '線２（右上から左下）を作成
    Set shpLines(2) = shpList.AddLine( _
                            rngCur2.Left, rngCur1.Top, _
                            rngCur1.Left, rngCur2.Top)
    
    '線１も線２も、2ptの赤線にする
    For i = 1 To 2
        shpLines(i).Line.ForeColor.RGB = vbRed
        shpLines(i).Line.Weight = 2
        
        ' ついでにオブジェクト名の配列をvariantの配列として作成
        varLineNames(i) = shpLines(i).Name
    Next
    
    '２本の線をグループ化し、選択状態にして終わる
    shpList.Range(varLineNames).Group.Select
    
End Sub

'*****************************************************************************
' 赤枠を描画する
'*****************************************************************************
Public Sub DrawRedFrame()
    
    Dim rngCur1 As Range
    Dim rngCur2 As Range
    
    Dim shpList As Shapes
    Dim shpFrame As Shape
    
    'セル以外を選択している場合は、選択解除
    ActiveCell.Activate
    
    '選択セル範囲を取得(A1:B2の場合は、A1とC3を取得)
    Set rngCur1 = Selection.Areas(1)
    Set rngCur2 = rngCur1.Offset(rngCur1.Rows.Count, rngCur1.Columns.Count)
    Set rngCur1 = rngCur1.Cells(1, 1)
    
    '選択範囲の左上セルのみを選択し直す
    rngCur1.Select
    
    'シェイプオブジェクトリスト
    Set shpList = ActiveSheet.Shapes
    
    '四角を描画
    Set shpFrame = shpList.AddShape(Type:=msoShapeRectangle, _
                                      Left:=rngCur1.Left, _
                                      Top:=rngCur1.Top, _
                                      Width:=Abs(rngCur2.Left - rngCur1.Left), _
                                      Height:=Abs(rngCur2.Top - rngCur1.Top) _
                                      )
    
    '塗りつぶしなしの2ptの赤線にする
    shpFrame.Fill.Visible = msoFalse
    shpFrame.Line.ForeColor.RGB = vbRed
    shpFrame.Line.Weight = 2
    
    '選択状態にして終わる
    shpFrame.Select
    
End Sub

'*****************************************************************************
' 50%にリサイズする
'*****************************************************************************
Public Sub SetZoomRate_50per()
    Call SetZoomRate(50)
End Sub

'*****************************************************************************
' リサイズする
'*****************************************************************************
Private Sub SetZoomRate(ByVal rate As Double)
    
    ' セルが選択されている場合は、何もしない
    If TypeOf Selection Is Range Then Exit Sub
    
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    
    Dim idxMax As Integer
    idxMax = sr.Count
    If idxMax > 0 Then
        Dim idx As Integer
        For idx = 1 To idxMax
            Dim shp As Shape
            Set shp = sr(idx)
            
            ' 縦横比固定を一旦解除してからリサイズ
            shp.LockAspectRatio = msoFalse
            Call shp.ScaleHeight(rate / 100, msoTrue)
            Call shp.ScaleWidth(rate / 100, msoTrue)
            shp.LockAspectRatio = msoTrue
            
        Next
    End If
End Sub

