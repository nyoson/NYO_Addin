Attribute VB_Name = "modCond"
Option Explicit

'以下のサイトからパクった条件付き書式を整理統合するマクロ
'https://excel-ubara.com/excelvba5/EXCELVBA262.html

'条件付き書式を格納する構造体
Private Type tFormat
    AppliesTo As String '適用範囲
    Formula1 As String '数式1
    Formula2 As String '数式2
    Operator As String '演算子
    NumberFormat As String '表示形式
    FontBold As String '太字
    FontColor As String '文字色
    InteriorColor As String '塗りつぶし色
    '追加判定したいプロパティはここに追加
End Type

'アクティブシートの条件付き書式を整理統合
Public Sub アクティブシートの条件付き書式を整理統合()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
  
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Call UnionFormatConditions(ws)
  
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'ブック全てのシートの条件付き書式を整理統合する
Public Sub ブック全てのシートの条件付き書式を整理統合する()
    Dim FileName As Variant
    Dim Wb As Workbook
    Dim ws As Worksheet
  
    FileName = Application.GetOpenFilename(FileFilter:="Excelファイル, *.xls*")
    If FileName = False Then
        Exit Sub
    End If
  
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
  
    Set Wb = Workbooks.Open(FileName:=FileName, UpdateLinks:=0, ReadOnly:=True)
  
    For Each ws In Wb.Worksheets
        Call UnionFormatConditions(ws)
    Next
  
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
  
    FileName = Application.GetSaveAsFilename(InitialFileName:=Wb.Name, _
                                             FileFilter:="Excelファイル,*.xls*")
    If FileName = False Then
        Exit Sub
    End If
    Wb.SaveAs FileName
    Wb.Close SaveChanges:=True
End Sub

'条件付き書式を整理統合する
Private Sub UnionFormatConditions(ByVal ws As Worksheet, _
                                 Optional ByVal NewName As String = "")
    '条件付き書式を格納する構造体配列
    Dim fAry() As tFormat
  
    '条件付き書式が無い場合は終了
    If ws.Cells.FormatConditions.Count = 0 Then Exit Sub
   
    'オプションにより元シートをコピー
    If NewName <> "" Then
        ws.Copy After:=ws
        Set ws = ActiveSheet
        ws.Name = NewName 'シート名のチェックは省略しています。
    End If
  
    '条件付き書式を構造体配列へ格納
    Call SetFormatToType(fAry, ws)
  
    '同一条件付き書式の結合：配列内でセル範囲指定文字列を結合
    Call JoinAppliesTo(fAry, ws)
  
    '条件付き書式の統合：配列内のAppliesをFormatConditionに適用
    Call ModifyApplies(fAry, ws)
End Sub

'条件付き書式を構造体配列へ格納
Private Sub SetFormatToType(ByRef fAry() As tFormat, _
                            ByVal ws As Worksheet)
    Dim i As Long
    Dim fObj As FormatCondition
    On Error Resume Next '.Formula2が取得できない場合の対処
  
    ReDim fAry(ws.Cells.FormatConditions.Count)
    For i = 1 To ws.Cells.FormatConditions.Count
        Set fObj = ws.Cells.FormatConditions(i)
        fAry(i).AppliesTo = fObj.AppliesTo.Address
        fAry(i).Formula1 = fObj.Formula1
        fAry(i).Formula2 = fObj.Formula2
        fAry(i).Operator = fObj.Operator
        fAry(i).NumberFormat = fObj.NumberFormat
        fAry(i).FontBold = fObj.Font.Bold
        fAry(i).FontColor = fObj.Font.Color
        fAry(i).InteriorColor = fObj.Interior.Color
        '追加判定したいプロパティはここに追加
    
        '数式エラーの条件付き書式は削除をする
        If isErrorFormula(fAry(i).Formula1) Or _
           isErrorFormula(fAry(i).Formula1) Then
            fAry(i).AppliesTo = ""
        End If
    Next
End Sub

'条件付き書式の数式エラー判定
Private Function isErrorFormula(ByVal sFormula As String) As Boolean
    If IsError(Evaluate(sFormula)) Then
        isErrorFormula = True
    Else
        isErrorFormula = False
    End If
End Function

'同一条件付き書式の結合：配列内でセル範囲指定文字列を結合
Private Sub JoinAppliesTo(ByRef fAry() As tFormat, _
                           ByVal ws As Worksheet)
    Dim i1 As Long, i2 As Long
    For i1 = 1 To UBound(fAry)
        For i2 = 1 To i1 - 1
            '計算式1,2、文字色、塗りつぶしの一致判定
            If isMatchFormat(fAry(i1), fAry(i2), ws) Then
                fAry(i2).AppliesTo = Union(Range(fAry(i2).AppliesTo), _
                                           Range(fAry(i1).AppliesTo)).Address
                fAry(i1).AppliesTo = ""
                Exit For
            End If
        Next
    Next
End Sub

'計算式1,2、演算子、文字色、塗りつぶしの一致判定
Private Function isMatchFormat(ByRef fAry1 As tFormat, _
                               ByRef fAry2 As tFormat, _
                               ByVal ws As Worksheet) As Boolean
    If fAry1.AppliesTo = "" Or _
       fAry2.AppliesTo = "" Then
        Exit Function
    End If
  
    Dim sFormula1 As String, sFormula2 As String
    isMatchFormat = True
  
    '計算式1
    sFormula1 = ToR1C1(fAry1.Formula1, fAry1.AppliesTo)
    sFormula2 = ToR1C1(fAry2.Formula1, fAry2.AppliesTo)
    If sFormula1 <> sFormula2 Then isMatchFormat = False
  
    '計算式2
    sFormula1 = ToR1C1(fAry1.Formula2, fAry1.AppliesTo)
    sFormula2 = ToR1C1(fAry2.Formula2, fAry2.AppliesTo)
    If sFormula1 <> sFormula2 Then isMatchFormat = False
  
    '演算子
    If fAry1.Operator <> fAry2.Operator Then isMatchFormat = False
  
    '表示形式
    If fAry1.NumberFormat <> fAry2.NumberFormat Then isMatchFormat = False
  
    '太字
    If fAry1.FontBold <> fAry2.FontBold Then isMatchFormat = False
  
    '文字色
    If fAry1.FontColor <> fAry2.FontColor Then isMatchFormat = False
  
    '塗りつぶし
    If fAry1.InteriorColor <> fAry2.InteriorColor Then isMatchFormat = False
  
    '追加判定したいプロパティはここに追加
End Function

'A1形式をR1C1形式に変換
Private Function ToR1C1(ByVal sFormula As String, _
                              ByVal sAppliesTo As String)
    If sFormula = "" Then Exit Function
    Dim rng As Range
    Set rng = Range(sAppliesTo)
    ToR1C1 = Application.ConvertFormula(sFormula, xlA1, xlR1C1, , rng.Item(1))
End Function

'条件付き書式の統合：配列内のAppliesをFormatConditionに適用
Private Sub ModifyApplies(ByRef fAry() As tFormat, _
                          ByVal ws As Worksheet)
    Dim i As Long
    Dim fObj As Object
    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        Set fObj = ws.Cells.FormatConditions(i)
        If fAry(i).AppliesTo = "" Then
            fObj.Delete
        Else
            If fObj.AppliesTo.Address <> ws.Range(fAry(i).AppliesTo).Address Then
                fObj.ModifyAppliesToRange ws.Range(fAry(i).AppliesTo)
            End If
        End If
    Next
End Sub

