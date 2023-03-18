Attribute VB_Name = "modCommon"
Option Explicit

'当モジュールのPublic関数を「マクロの実行」に表示しないようにする
Option Private Module

'＜共通処理モジュール＞

'以下を参照設定に追加して下さい。
'・Microsoft Visual Basic for Applications Extensibility
'（VBProjectクラスなど）
'・Microsoft Scripting Runtime
'（FileSystemObjectクラスなど）

'##########
' 定数
'##########
Public Const C_TOOL_NAME As String = "NYO_Addin"
Public Const C_TOOLBAR_NAME As String = "NYO"

'##########
' API関数
'##########
' Sleep関数(スリープ時間[ms])
Public Declare PtrSafe Sub Sleep Lib "KERNEL32.dll" (ByVal dwMilliseconds As Long)

' IME制御
Private Declare PtrSafe Function ImmGetContext Lib "imm32.dll" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function ImmSetOpenStatus Lib "imm32.dll" (ByVal himc As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function ImmReleaseContext Lib "imm32.dll" (ByVal hWnd As Long, ByVal himc As Long) As Long

'##########
' 列挙体
'##########
' ExcelVersion
Public Enum ExcelMajorVersion
    Ver2003 = 11
    Ver2007 = 12
    Ver2010 = 14
    Ver2013 = 15
    Ver2016 = 16
End Enum

'##########
' 変数
'##########
' ExcelVersion
Private enmExcelMajorVersion As ExcelMajorVersion

'##########
' 関数
'##########

'*****************************************************************************
'[ 関数名 ]　ProcessModeEnd
'[ 概  要 ]　処理中モード解除
'[ 引  数 ]　なし
'[ 戻り値 ]  なし
'*****************************************************************************
Public Sub ProcessModeEnd()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.Cursor = xlDefault
    Application.EnableEvents = True
End Sub

'*****************************************************************************
'[ 関数名 ] Is2016orLater
'[ 概  要 ] Excel2016以降かどうかを返す
'[ 引  数 ] なし
'[ 戻り値 ] True：Excel2016以降／False：それより前
'*****************************************************************************
Public Function Is2016orLater() As Boolean
    Dim intVersion As Integer
    intVersion = GetExcelMajorVersion()
    If intVersion >= ExcelMajorVersion.Ver2016 Then
        Is2016orLater = True
    Else
        Is2016orLater = False
    End If
End Function

'*****************************************************************************
'[ 関数名 ] Is2007orLater
'[ 概  要 ] Excel2007以降かどうかを返す
'[ 引  数 ] なし
'[ 戻り値 ] True：Excel2007以降／False：それより前
'*****************************************************************************
Public Function Is2007orLater() As Boolean
    Dim intVersion As Integer
    intVersion = GetExcelMajorVersion()
    If intVersion >= ExcelMajorVersion.Ver2007 Then
        Is2007orLater = True
    Else
        Is2007orLater = False
    End If
End Function

'*****************************************************************************
'[ 関数名 ] GetExcelMajorVersion
'[ 概  要 ] Excelの内部バージョンを取得する
'[ 引  数 ] なし
'[ 戻り値 ] 12：Excel2007／14：Excel2010／15：Excel2013／16：Excel2016以降(2019含む)
'           【参考】https://minimashia.net/vba-excel-check/
'*****************************************************************************
Public Function GetExcelMajorVersion() As ExcelMajorVersion
    If Not enmExcelMajorVersion > 0 Then
        enmExcelMajorVersion = CInt(Split(Application.Version, ".", 2, vbBinaryCompare)(0))
    End If
    GetExcelMajorVersion = enmExcelMajorVersion
End Function


'*****************************************************************************
'[ 関数名 ] SetIMEMode
'[ 概  要 ] IMEモードのON/OFFを変更する
'[ 引  数 ] True:IME-ON / False:IME-OFF
'           対象コントロールのハンドル(ex: txtInput.hWnd) ※セルなどの場合は省略可
'[ 戻り値 ] なし
'*****************************************************************************
Public Function SetIMEMode(ByVal IMEMode As Boolean, Optional ByVal hWnd As Long = 0)
    
    '対象コントロールのhWndの指定がなければ、ExcelAppのhWndを取得
    If hWnd = 0 Then hWnd = Application.hWnd
    
    'IMEをOn
    Dim himc As Long
    himc = ImmGetContext(hWnd)
    Call ImmSetOpenStatus(himc, IIf(IMEMode, 1, 0))
    Call ImmReleaseContext(hWnd, himc)
    
End Function

'*****************************************************************************
'[ 関数名 ] CheckMacroSecurityWithMsg
'[ 概  要 ] VBAプロジェクトへのアクセス権のチェックを行い、エラー時はメッセージも表示する
'[ 引  数 ] なし
'[ 戻り値 ] アクセス可否
'*****************************************************************************
Public Function CheckMacroSecurityWithMsg() As Boolean
    
    Dim ret As Boolean
    
    'セキュリティ設定チェック
    ret = CheckMacroSecurity()
    
    If Not ret Then
        Dim strMsg As String
        If Is2016orLater Then
            'Excel 2016以降
            strMsg = "現在のセキュリティ設定では実行できません。" & vbCrLf _
                    & "Excelの[ファイル]リボン->[その他...]->[オプション]->" & vbCrLf _
                    & "[トラスト センター]タブ->「トラスト センターの設定...」ボタン->" & vbCrLf _
                    & "[マクロの設定]タブ->「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェックして下さい。"
        ElseIf Is2007orLater Then
            'Excel 2007以降
            strMsg = "現在のセキュリティ設定では実行できません。" & vbCrLf _
                    & "Excelの[ファイル]リボン->[オプション]->" & vbCrLf _
                    & "[セキュリティ センター]タブ->「セキュリティ センターの設定」ボタン->" & vbCrLf _
                    & "[マクロの設定]タブ->「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をチェックして下さい。"
        Else
            'Excel 2003以前
            strMsg = "現在のセキュリティ設定では実行できません。" & vbCrLf _
                    & "Excelの[ツール]->[マクロ]->[セキュリティ]->[信頼のおける発行元]タブの、" & vbCrLf _
                    & "[Visual Basicプロジェクトへのアクセスを信頼する]をオンにして下さい。"
        End If
        Call modMessage.ErrorMessage(strMsg)
    End If
    
    CheckMacroSecurityWithMsg = ret
    
End Function

'*****************************************************************************
'[ 関数名 ] CheckMacroSecurity
'[ 概  要 ] VBAプロジェクトへのアクセス権のチェックを行う
'[ 引  数 ] なし
'[ 戻り値 ] アクセス可否
'*****************************************************************************
Public Function CheckMacroSecurity() As Boolean
    
    Dim ret As Boolean
    Dim test As Object 'VBProject
    
On Error GoTo Catch
    '試しに、当アドインのVBコンポーネントの数を取得
    Set test = ThisWorkbook.VBProject
    ret = True
    
    GoTo Fainally
    
Catch:
    ret = False
    
Fainally:
    Set test = Nothing
    CheckMacroSecurity = ret
    
End Function

'*****************************************************************************
'[ 関数名 ] MakeDir
'[ 概  要 ] サブフォルダも含めてフォルダを作成する
'[ 引  数 ] 作成するフォルダのフルパス
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub MakeDir(ByVal strPath As String)
    
    'FileSystemObject
    Dim objFso As FileSystemObject
    Set objFso = New FileSystemObject
    '既に存在していれば、何もせず終了
    If objFso.FolderExists(strPath) Then
        Exit Sub
    End If
    
    '親フォルダがなければ作成(再帰呼び出し)
    Call MakeDir(objFso.GetParentFolderName(strPath))
    
    '目的のフォルダを作成
    Call objFso.CreateFolder(strPath)
    
End Sub

'*****************************************************************************
'[ 関数名 ] GetLastCell
'[ 概  要 ] シート内の最終セルを取得
'[ 引  数 ] 対象シート
'[ 戻り値 ] 最終セル
'*****************************************************************************
Public Function GetLastCell(ByVal shtTarget As Worksheet) As Range
    Dim rngLastCell As Range
    Set rngLastCell = shtTarget.UsedRange
    Set rngLastCell = rngLastCell.Cells(rngLastCell.Rows.Count, _
                                        rngLastCell.Columns.Count)
    Set GetLastCell = rngLastCell
End Function

'*****************************************************************************
'[ 関数名 ] ScrollTo
'[ 概  要 ] 対象セルまでジャンプ
'[ 引  数 ] 対象セル, スクロール行オフセット（省略時は-5行）
'[ 戻り値 ] なし
'*****************************************************************************
Public Sub ScrollTo(ByRef rng As Range, Optional ByVal scrollRowOffset As Long = -5)

    '対象セルのシートをアクティブにする
    Call rng.Parent.Activate
    
    '対象セルより少し上の行にスクロール（但し、A1セルより上にならないようにガード）
    If rng.Row + scrollRowOffset < 1 Then
        scrollRowOffset = 1 - rng.Row
    End If
    Call Application.GoTo(rng.Offset(scrollRowOffset, 1 - rng.Column), True)
    
    '対象セルにカーソル移動
    Call rng.Activate
    
End Sub

