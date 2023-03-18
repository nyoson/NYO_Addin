Attribute VB_Name = "modWindow"
Option Explicit

'*****************************************************************************
' 新しいウィンドウで開き直す
' マクロ作成日 : 2013/05/20  ユーザー名 : nyoson
'*****************************************************************************
Public Sub OpenAsNewWindow()
    
    Dim wbkTarget As Workbook
    Set wbkTarget = ActiveWorkbook
    
    With CreateObject("Scripting.FileSystemObject")
        If .GetExtensionName(wbkTarget.FullName) = Empty Then
            Call modMessage.ErrorMessage("保存されていないため、続行できません。")
            Exit Sub
        End If
    End With
    
    Dim newExcelApp As Excel.Application
    Set newExcelApp = CreateObject("Excel.Application")
    newExcelApp.Visible = True
    Call newExcelApp.Workbooks.Open(FileName:=wbkTarget.FullName, ReadOnly:=True)
    
    If modMessage.ConfirmMessage("元のブックを閉じますか？") <> vbYes Then
        Exit Sub
    End If
    
    Call wbkTarget.Close
    
End Sub

'*****************************************************************************
' 左右に並べて比較
' マクロ作成日 : 2013/06/10  ユーザー名 : nyoson
'*****************************************************************************
Public Sub CompareWindowInVertical()
    
    Dim winCurrent As Window
    Dim winTarget As Window
    
    If Windows.Count < 2 Then
        Call modMessage.ErrorMessage("子ウィンドウの数が足りません。")
        Exit Sub
    End If
    
    Set winCurrent = ActiveWindow
    For Each winTarget In Windows
        If winCurrent.Caption <> winTarget.Caption Then
            Exit For
        End If
    Next
    
    '並べて比較
    Windows.CompareSideBySideWith winTarget.Caption
    
    '左右に並べる
    Windows.Arrange ArrangeStyle:=xlVertical
    
End Sub

'*****************************************************************************
' ウィンドウをメインディスプレイに移動
' マクロ作成日 : 2013/07/09  ユーザー名 : nyoson
'*****************************************************************************
Public Sub MoveWindowToMainDisplay()
    
    '最大化や最小化を解除
    Application.WindowState = xlNormal
    
    'ウィンドウサイズと位置を変更
    Application.Height = 621.75
    Application.Width = 937.5
    Application.Top = 2.25
    Application.Left = 7
    
End Sub

'*****************************************************************************
' ウィンドウをサブディスプレイに移動
' マクロ作成日 : 2013/06/26  ユーザー名 : nyoson
'*****************************************************************************
Public Sub MoveWindowToSubDisplay()
    
    '最大化や最小化を解除
    Application.WindowState = xlNormal
    
    'ウィンドウサイズと位置を変更
    Application.Height = 621.75
    Application.Width = 937.5
    Application.Top = -295.5
    Application.Left = 1212.25
    
End Sub

'*****************************************************************************
' ウィンドウサイズをデバッグ出力する
'*****************************************************************************
Private Sub GetWindowState()
    
    'ActiveWindowのウィンドウサイズを出力
    Debug.Print "==========================="
    Debug.Print "Application.Height = " & Application.Height; ""
    Debug.Print "Application.Width = " & Application.Width; ""
    Debug.Print "Application.Top = " & Application.Top; ""
    Debug.Print "Application.Left = " & Application.Left
    
End Sub

