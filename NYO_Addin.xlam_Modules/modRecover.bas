Attribute VB_Name = "modRecover"
Option Explicit

'*****************************************************************************
' 条件付き書式が再描画されない不具合の解消
'*****************************************************************************
Public Sub FormatConditionsRedrawBugFix()
    
    Dim shtTarget As Worksheet
    For Each shtTarget In ActiveWorkbook.Worksheets
        shtTarget.EnableFormatConditionsCalculation = True
    Next
    
End Sub

'*****************************************************************************
' 無効な名前付き定義を削除
'*****************************************************************************
Public Sub DeleteInvaridNameDef()
    
    '名前付き定義リスト
    Dim nameDefList As Names
    Set nameDefList = ActiveWorkbook.Names
    
    'リストを後ろから処理
    Dim idx As Long
    idx = nameDefList.Count
    Do While idx > 0
        Dim nameDefCur As Name
        Set nameDefCur = nameDefList(idx)
        
        '無効な名前付き定義なら削除
        If InStr(1, nameDefCur.RefersTo, "#REF!", vbBinaryCompare) > 0 Then
            Call nameDefCur.Delete
            GoTo Next_NameDef
        End If
        
        ''Print_*(Print_AreaやPrint_Titles)以外を削除
        'If InStr(1, nameDefCur.Name, "Print_", vbBinaryCompare) < 1 Then
        '    Call nameDefCur.Delete
        '    GoTo Next_NameDef
        'End If
        
Next_NameDef:
        idx = idx - 1
    Loop
    
End Sub

'*****************************************************************************
' カラーパレットをリセット
'*****************************************************************************
Public Sub ResetColorPalette()
    
    Dim wbkTarget As Workbook
    Set wbkTarget = ActiveWorkbook
    
    'wbkTarget.ResetColors
    
    '標準カラーインデックスにリセット
    wbkTarget.Colors(1) = RGB(0, 0, 0) '0
    wbkTarget.Colors(2) = RGB(255, 255, 255) '16777215
    wbkTarget.Colors(3) = RGB(255, 0, 0) '255
    wbkTarget.Colors(4) = RGB(0, 255, 0) '65280
    wbkTarget.Colors(5) = RGB(0, 0, 255) '16711680
    wbkTarget.Colors(6) = RGB(255, 255, 0) '65535
    wbkTarget.Colors(7) = RGB(255, 0, 255) '16711935
    wbkTarget.Colors(8) = RGB(0, 255, 255) '16776960
    wbkTarget.Colors(9) = RGB(128, 0, 0) '128
    wbkTarget.Colors(10) = RGB(0, 128, 0) '32768
    wbkTarget.Colors(11) = RGB(0, 0, 128) '8388608
    wbkTarget.Colors(12) = RGB(128, 128, 0) '32896
    wbkTarget.Colors(13) = RGB(128, 0, 128) '8388736
    wbkTarget.Colors(14) = RGB(0, 128, 128) '8421376
    wbkTarget.Colors(15) = RGB(192, 192, 192) '12632256
    wbkTarget.Colors(16) = RGB(128, 128, 128) '8421504
    wbkTarget.Colors(17) = RGB(153, 153, 255) '16751001
    wbkTarget.Colors(18) = RGB(153, 51, 102) '6697881
    wbkTarget.Colors(19) = RGB(255, 255, 204) '13434879
    wbkTarget.Colors(20) = RGB(204, 255, 255) '16777164
    wbkTarget.Colors(21) = RGB(102, 0, 102) '6684774
    wbkTarget.Colors(22) = RGB(255, 128, 128) '8421631
    wbkTarget.Colors(23) = RGB(0, 102, 204) '13395456
    wbkTarget.Colors(24) = RGB(204, 204, 255) '16764108
    wbkTarget.Colors(25) = RGB(0, 0, 128) '8388608
    wbkTarget.Colors(26) = RGB(255, 0, 255) '16711935
    wbkTarget.Colors(27) = RGB(255, 255, 0) '65535
    wbkTarget.Colors(28) = RGB(0, 255, 255) '16776960
    wbkTarget.Colors(29) = RGB(128, 0, 128) '8388736
    wbkTarget.Colors(30) = RGB(128, 0, 0) '128
    wbkTarget.Colors(31) = RGB(0, 128, 128) '8421376
    wbkTarget.Colors(32) = RGB(0, 0, 255) '16711680
    wbkTarget.Colors(33) = RGB(0, 204, 255) '16763904
    wbkTarget.Colors(34) = RGB(204, 255, 255) '16777164
    wbkTarget.Colors(35) = RGB(204, 255, 204) '13434828
    wbkTarget.Colors(36) = RGB(255, 255, 153) '10092543
    wbkTarget.Colors(37) = RGB(153, 204, 255) '16764057
    wbkTarget.Colors(38) = RGB(255, 153, 204) '13408767
    wbkTarget.Colors(39) = RGB(204, 153, 255) '16751052
    wbkTarget.Colors(40) = RGB(255, 204, 153) '10079487
    wbkTarget.Colors(41) = RGB(51, 102, 255) '16737843
    wbkTarget.Colors(42) = RGB(51, 204, 204) '13421619
    wbkTarget.Colors(43) = RGB(153, 204, 0) '52377
    wbkTarget.Colors(44) = RGB(255, 204, 0) '52479
    wbkTarget.Colors(45) = RGB(255, 153, 0) '39423
    wbkTarget.Colors(46) = RGB(255, 102, 0) '26367
    wbkTarget.Colors(47) = RGB(102, 102, 153) '10053222
    wbkTarget.Colors(48) = RGB(150, 150, 150) '9868950
    wbkTarget.Colors(49) = RGB(0, 51, 102) '6697728
    wbkTarget.Colors(50) = RGB(51, 153, 102) '6723891
    wbkTarget.Colors(51) = RGB(0, 51, 0) '13056
    wbkTarget.Colors(52) = RGB(51, 51, 0) '13107
    wbkTarget.Colors(53) = RGB(153, 51, 0) '13209
    wbkTarget.Colors(54) = RGB(153, 51, 102) '6697881
    wbkTarget.Colors(55) = RGB(51, 51, 153) '10040115
    wbkTarget.Colors(56) = RGB(51, 51, 51) '3355443
End Sub

'*****************************************************************************
'<DEBUG>NYOアドインの再導入
'*****************************************************************************
Public Sub RestoreNYOAddin()
    ' ツールバー削除
    Call DelMenu
    ' ツールバー追加
    Call SetMenu
    ' ショートカットキー割り当て
    Call RegistShortcutKey
    ' 右クリックメニュー追加
    Call AddContextMenu
End Sub
'ShortHand
Public Sub nyo()
    Call RestoreNYOAddin
End Sub


