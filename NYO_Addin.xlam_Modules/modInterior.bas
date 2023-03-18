Attribute VB_Name = "modInterior"
Option Explicit

'*****************************************************************************
' 選択セルの背景色を黄色またはなしにする
'*****************************************************************************
Public Sub SwitchInteriorYellow()
    
    ' セル以外は非対応
    ' ※オートシェイプの場合、既に黄色の時に、背景色なし(透明)にするか白にするか微妙なため
    If Not TypeOf Selection Is Range Then
        Call modMessage.ErrorMessage("セルを選択して下さい")
        Exit Sub
    End If
    
    Dim itr As Interior
    Set itr = Selection.Interior
    
    If itr.Color <> RGB(255, 255, 0) Then
        itr.Color = RGB(255, 255, 0)
    Else
        '既に黄色の場合は、背景色なしにする
        itr.ColorIndex = XlColorIndex.xlColorIndexNone
    End If
    
End Sub

