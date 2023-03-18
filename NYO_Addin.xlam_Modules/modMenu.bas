Attribute VB_Name = "modMenu"
Option Explicit

'*****************************************************************************
'[ 関数名 ]　SetMenu
'[ 概  要 ]　ツールバーを設定する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetMenu()
    Dim objMenu As CommandBar           ' コマンドバー
    Dim objControl As CommandBarControl ' コマンドバー内コントロール
    Dim Exists As Boolean               ' 検索時ヒットフラグ
    Dim i As Integer

    On Error Resume Next
    Set objMenu = Application.CommandBars(C_TOOLBAR_NAME)
    On Error GoTo 0
    If objMenu Is Nothing Then
        ' ツールバーを新規作成
        Set objMenu = Application.CommandBars.Add(Name:=C_TOOLBAR_NAME)
    Else
        ' 既に存在したらクリア確認
        If modMessage.ConfirmMessage("既にツールバー" & C_TOOLBAR_NAME & "は存在します。クリアしますか？") <> vbYes Then
            Exit Sub
        End If
        
        ' アイテムクリア
        For Each objControl In objMenu.Controls
            objControl.Delete
        Next
    End If
    
    'クイックツールバーに移行済みのため、Excel2007以降なら追加しない
    If Not modCommon.Is2007orLater Then
        ' ボタンを追加：読み取り専用の設定/解除
        With objMenu.Controls.Add(Type:=msoControlButton, ID:=456)
        End With
        
        ' ボタンを追加：ファイルの更新
        With objMenu.Controls.Add(Type:=msoControlButton, ID:=455)
        End With
        
        ' ボタンを追加：値貼付
        With objMenu.Controls.Add(Type:=msoControlButton, ID:=370)
            .Style = msoButtonCaption
            .Caption = "値(&P)"
        End With
        
        ' ボタンを追加：UnFilter
        With objMenu.Controls.Add(Type:=msoControlButton, ID:=900)
            .Caption = "&UnFilter"
            .TooltipText = "フィルタ解除"
        End With
    End If
    
    ' ボタンを追加：xPaste
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=5837)
        .Style = msoButtonCaption
        .Caption = "&xPaste"
        .TooltipText = "罫線以外を貼り付け"
    End With
    
    ' ボタンを追加：アウトライングループ化
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=3159)
    End With
    
    ' ボタンを追加：アウトライングループ化解除
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=3160)
    End With
    
    ' ボタンを追加：ウィンドウ枠の固定
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=443)
    End With
    
    ' ボタンを追加：OpenDir
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=2950)
        .Style = msoButtonCaption
        .Caption = "Open&Dir"
        .TooltipText = "ひとつ上を開く"
        .OnAction = "OpenDir"   'modFile
    End With
    
    ' ボタンを追加：AddCopyRow
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=2950)
        .Style = msoButtonCaption
        .Caption = "AddCopy&Row"
        .TooltipText = "行複製(CS++)"
        .OnAction = "AddCopyRow"    'modRow
    End With
    
    ' ボタンを追加：オートシェイプの自動サイズ調整
    With objMenu.Controls.Add(Type:=msoControlButton)
        .FaceId = 5866
        .Style = msoButtonIcon
        .Caption = "オートシェイプの自動サイズ調整"
        .TooltipText = .Caption
        .OnAction = "AutoFit"   'modAutoFit
    End With
        
    ' ボタンを追加：グリッド線の表示/非表示
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=485)
        .TooltipText = "グリッド線の表示/非表示"
    End With
    
    ' ボタンを追加：R1C1形式切り替え
    With objMenu.Controls.Add(Type:=msoControlButton)
        .FaceId = 52
        .Style = msoButtonIcon
        .Caption = "R1C1形式切り替え"
        .TooltipText = .Caption
        .OnAction = "SwitchR1C1"    'modSheet
    End With
    
    ' ボタンを追加：シート表示切り替え
    With objMenu.Controls.Add(Type:=msoControlButton)
        .FaceId = 461
        .Style = msoButtonIcon
        .Caption = "シート表示切り替え"
        .TooltipText = .Caption
        .OnAction = "SheetDispChanger"    'modSheet
    End With
    
    '
    ' ボタンを追加：Resize50%
    With objMenu.Controls.Add(Type:=msoControlButton, ID:=2950)
        .Style = msoButtonCaption
        .Caption = "Resize50%"
        .TooltipText = "50%にリサイズする"
        .OnAction = "SetZoomRate_50per" 'modDraw
    End With
    
    ' メニューを追加：ELSE
    Dim objMenuElse As CommandBarPopup
    Set objMenuElse = objMenu.Controls.Add(Type:=msoControlPopup)
    With objMenuElse
        .Caption = "E&LSE"
        .TooltipText = .Caption
        
        ' メニューを追加：modFormula
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "数式入力"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "シート名"
                .OnAction = "InputSheetName"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "シート名(YYYY_MMDD形式)"
                .OnAction = "InputSheetName_YYYY_MMDD"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "シート名(YYYY_MMDD形式)=>日付値変換"
                .OnAction = "InputSheetName_YYYY_MMDD_AsDate"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "連番"
                .OnAction = "InputSeqNo"
            End With
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "A1セルをActiveにする"
            .OnAction = "ActivateAllA1Cell" 'modA1
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "アウトラインの集計方向を左上にする"
            .OnAction = "OutlineConfigAboveLeft"    'modSheet
            .TooltipText = "アウトラインの集計行を左，集計列を上にする"
        End With
        
        ' メニューを追加：modCrlf
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "改行操作"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "選択範囲の各セルに改行を追加する"
                .OnAction = "AddCrLf"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "選択範囲の各セルの改行を削除する"
                .OnAction = "RemoveCrLf"
            End With
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "SQLグリッド貼り付け(CS+Q)"
            .OnAction = "PasteSQLGrid"  'modSQL
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "モジュールを全てエクスポート"
            .OnAction = "ExportAllModules"  'modExportAllModules
        End With
        
        ' メニューを追加：modFile
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "ファイル系"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "ファイル名をクリップボードにコピー"
                .OnAction = "CopyFileName"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "ExportCSV"
                .OnAction = "ExportCSV"
            End With
            
        End With
        
        ' メニューを追加：modDraw
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "描画系"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "×印を描画する"
                .OnAction = "DrawX"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "赤枠を描画する"
                .OnAction = "DrawRedFrame"
            End With
            
        End With
        
        ' メニューを追加：modSign
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "電子印系"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "電子印"
                .OnAction = "Sign"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "電子印設定"
                .OnAction = "ShowSignSetting"
            End With
            
        End With
        
        ' メニューを追加：modWindow
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "ウィンドウ操作"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "新しいExcelで開き直す"
                .OnAction = "OpenAsNewWindow"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "左右に並べて比較"
                .OnAction = "CompareWindowInVertical"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "ウィンドウをメインディスプレイに移動"
                .OnAction = "MoveWindowToMainDisplay"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "ウィンドウをサブディスプレイに移動"
                .OnAction = "MoveWindowToSubDisplay"
            End With
            
        End With
        
        ' メニューを追加：modFont
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "文字装飾系"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "赤字(CS+R)"
                .TooltipText = "文字色の「赤」/「通常」を切り替える"
                .OnAction = "SwitchRed"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "フォントクリア"
                .TooltipText = "選択セルの文字装飾をクリアする"
                .OnAction = "ClearFont"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "ハイパーリンクスタイル変更"
                .TooltipText = "ハイパーリンク生成時のフォント名とサイズを、アクティブセルと同じにする"
                .OnAction = "SetHyperLinkStyle"
            End With
            
        End With
        
        ' メニューを追加：modRecover
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "修復系"
            .TooltipText = .Caption
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "処理中モード解除"
                .OnAction = "ProcessModeEnd"    'modCommon
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "条件付き書式が再描画されない不具合の解消"
                .OnAction = "FormatConditionsRedrawBugFix"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "無効な名前付き定義を削除"
                .TooltipText = "#REF!になっている名前付き定義を削除"
                .OnAction = "DeleteInvaridNameDef"
            End With

            With .Controls.Add(Type:=msoControlButton)
                .Caption = "カラーパレットをリセット"
                .TooltipText = "ブックのカラーパレットを標準に戻す"
                .OnAction = "ResetColorPalette"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "<DEBUG>右クリックメニューの復元"
                .OnAction = "AddContextMenu"    'modMenu
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "<DEBUG>" & C_TOOL_NAME & "アドインの再導入"
                .OnAction = "RestoreNYOAddin"
            End With
            
        End With
        
    End With
    
    ' ツールバーの位置（上）
    objMenu.Position = msoBarTop
    
    ' ツールバーを表示
    objMenu.Visible = True
    
    'Excel2003以前の場合のみ
    If Not Is2007orLater Then
        
        ' メインメニューに[アドレス]を追加
        Exists = False
        Set objMenu = Application.CommandBars("Worksheet Menu Bar")
        For Each objControl In objMenu.Controls
            ' 既に登録済みでないか探す
            If objControl.ID = 1740 Then
                Exists = True
                Exit For
            End If
        Next objControl
        If Exists <> True Then
            ' コンボボックスを追加：アドレス
            With objMenu.Controls.Add(Type:=msoControlComboBox, ID:=1740)
                .BeginGroup = True
                .Width = 250
            End With
        End If
        
        ' 書式ツールバーに[取り消し線],[上揃え],[上下中央揃え],[横方向に結合]を追加
        Set objMenu = Application.CommandBars("Formatting")
        objMenu.Visible = True
        objMenu.Reset   ' 一旦リセット
        For i = 1 To objMenu.Controls.Count
            Set objControl = objMenu.Controls(i)
            
            Select Case objControl.ID
                ' 下線ボタン
                Case 115:
                    '' 既に右側に取り消し線があれば、何もしない
                    'If objMenu.Controls(i + 1).ID = 290 Then
                    '    Exit For
                    'End If
                    
                    ' 右側にボタンを追加：取り消し線
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=290, Before:=i)
                    End With
                    
                ' 右揃えボタン
                Case 121:
                    ' 右側にボタンを追加：上揃え
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=2600, Before:=i)
                    End With
                    
                    ' 右側にボタンを追加：上下中央揃え
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=6542, Before:=i)
                    End With
                    
                ' セルを結合して中央揃えボタン
                Case 402:
                    ' 右側にボタンを追加：セルの結合
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=798, Before:=i)
                    End With
                    
                    ' 右側にボタンを追加：横方向に結合
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=1742, Before:=i)
                    End With
                    
                    ' 右側にボタンを追加：セル結合の解除
                    i = i + 1
                    With objMenu.Controls.Add(Type:=msoControlButton, ID:=800, Before:=i)
                    End With
                    
                    ' 終了
                    Exit For
                    
            End Select
    
        Next i
        
        ' 標準ツールバーを表示
        Set objMenu = Application.CommandBars("Standard")
        objMenu.Visible = True
        
        ' 罫線ツールバーを表示
        Set objMenu = Application.CommandBars("Borders")
        objMenu.Visible = True
        
    End If
    
End Sub

'*****************************************************************************
'[ 関数名 ]　DelMenu
'[ 概  要 ]　ツールバーを削除する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub DelMenu()
    
    'エラー時は何もせず続行
    On Error Resume Next

    ' ツールバー削除
    Call Application.CommandBars(C_TOOLBAR_NAME).Delete

End Sub

'*****************************************************************************
'[ 関数名 ]　AddContextMenu
'[ 概  要 ]　右クリックメニューにコマンドを追加する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub AddContextMenu()
    
    Debug.Print "【DEBUG】AddContextMenu"
    
    '当Addinが読み取り専用かどうかのラベル
    Dim strThisAddinReadOnlyText As String
    If ThisWorkbook.ReadOnly Then
        strThisAddinReadOnlyText = "<ReadOnly>"
    Else
        strThisAddinReadOnlyText = "<Editable>"
    End If
    
    '右クリックメニューに追加
    '(標準モードと改ページプレビューモードの両方に追加する)
    Dim objMenu As CommandBar           '右クリックメニュー
    Dim objControl As CommandBarControl '右クリックメニュー内コントロール
    
    Dim lngCmdNo As Long
    For lngCmdNo = 1 To Application.CommandBars.Count
        Set objMenu = Application.CommandBars(lngCmdNo)
        
        '全コマンドバーの内、以下の名称のもの（＝セルの右クリックメニュー）をカスタマイズ
        '※Cellなどは標準や改ページプレビューなど複数存在するためループ処理
        '※List Range Popupはテーブルとして書式設定された範囲内の右クリックメニュー
        If objMenu.Name = "Cell" Or _
           objMenu.Name = "List Range Popup" Then
            
            'DEBUG:一旦リセット
            Call objMenu.Reset
            '※↑他のアドインによる変更が消えるので本来はコメントアウト推奨
            
            'DEBUG:LOG
            Debug.Print objMenu.Index & ":" & objMenu.Name
            
            Set objControl = objMenu.Controls.Add(Temporary:=True)
            '↑Temporary:=Trueにより、ブッククローズ時に自動的に削除される
            With objControl
                .Caption = "【" & C_TOOL_NAME & "】" & strThisAddinReadOnlyText
                .Enabled = False
            End With
            Set objControl = Nothing
            
            Set objControl = objMenu.Controls.Add(Temporary:=True)
            '↑Temporary:=Trueにより、ブッククローズ時に自動的に削除される
            With objControl
                .Caption = "電子印"
                .OnAction = "Sign"
                .BeginGroup = False
            End With
            Set objControl = Nothing
            
            Set objControl = objMenu.Controls.Add(Temporary:=True)
            '↑Temporary:=Trueにより、ブッククローズ時に自動的に削除される
            With objControl
                .Caption = "電子印設定"
                .OnAction = "ShowSignSetting"
                .BeginGroup = False
            End With
            Set objControl = Nothing
            
            Set objControl = objMenu.Controls.Add(Temporary:=True)
            '↑Temporary:=Trueにより、ブッククローズ時に自動的に削除される
            With objControl
                .Caption = "<回避策>ツールバー" & C_TOOLBAR_NAME & "の復元"
                .OnAction = "SetMenu"
                .BeginGroup = False
            End With
            Set objControl = Nothing
            
        End If
    Next
End Sub

'*****************************************************************************
'[ 関数名 ]　RegistShortcutKey
'[ 概  要 ]　ショートカットキー割り当て
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub RegistShortcutKey()
    '【参考】http://excel-ubara.com/excelvba1/EXCELVBA421.html
    
    Application.OnKey "^{Return}", "ThisWorkbook.CtrlEnter"             'Ctrl + Enter            ⇒入力補助
    Application.OnKey "^{Enter}", "ThisWorkbook.CtrlEnter"              'Ctrl + Enter(Tenキー側) ⇒入力補助
    Application.OnKey "^%;", "ThisWorkbook.CtrlAltSemicolon"            'Ctrl + Alt + ";"        ⇒入力補助（日付）
    Application.OnKey "^%:", "ThisWorkbook.CtrlAltColon"                'Ctrl + Alt + ":"        ⇒入力補助（時刻）
    Application.OnKey "^{F3}", "modSearchNext.SearchNextThisCellText"   'Ctrl + F3               ⇒セル内容で検索画面を開く ※既存機能を上書き
    Application.OnKey "{F3}", "modSearchNext.SearchNextForward"         'F3                      ⇒次を検索
    Application.OnKey "+{F3}", "modSearchNext.SearchNextPrevious"       'Shift + F3              ⇒前を検索
    Application.OnKey "^+Q", "modSQL.PasteSQLGrid"                       'Ctrl + Shift + Q        ⇒SQLグリッド貼り付け
    Application.OnKey "^+R", "modFont.SwitchRed"                        'Ctrl + Shift + R        ⇒赤字切り替え
    Application.OnKey "^+Y", "modInterior.SwitchInteriorYellow"         'Ctrl + Shift + Y        ⇒黄色背景色切り替え
    Application.OnKey "^+{+}", "modRow.AddCopyRow"                      'Ctrl + Shift + "+"            ⇒選択行をコピーして下に追加
    Application.OnKey "^+{107}", "modRow.AddCopyRow"                    'Ctrl + Shift + "+"(Tenキー側) ⇒選択行をコピーして下に追加
    Application.OnKey "^+{-}", "modRow.DelRow"                          'Ctrl + Shift + "-"            ⇒選択行を削除
    Application.OnKey "^+{109}", "modRow.DelRow"                        'Ctrl + Shift + "-"(Tenキー側) ⇒選択行を削除
    Application.OnKey "^%{Up}", "modRow.UpRow"                          'Ctrl + Alt + "↑"             ⇒選択行を上移動
    Application.OnKey "^%{Down}", "modRow.DownRow"                      'Ctrl + Alt + "↓"             ⇒選択行を下移動
    
    Application.OnKey "{F1}", ""                                        'F1                      ⇒ヘルプへのショートカット機能を無効化
    
End Sub

