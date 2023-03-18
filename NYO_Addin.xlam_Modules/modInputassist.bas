Attribute VB_Name = "modInputassist"
Option Explicit

'*****************************************************************************
'[ 関数名 ]　InputAssist
'[ 概  要 ]　入力補助
'[ 引  数 ]　対象セル
'[ 戻り値 ]　入力補助側で処理をしたため、イベントをキャンセルすべき場合にTrueを返す
'*****************************************************************************
Public Function InputAssist(ByVal Target As Range) As Boolean
    
    Dim Cancel As Boolean   'イベントキャンセルフラグ
    
    Cancel = False
    
    '複数セルの場合は1セル目
    Dim rng As Range
    Set rng = Target.Cells(1, 1)
    
    '数式が入力されているかチェックしておく
    Dim isFormula As Boolean
    isFormula = False
    If Len(rng.Formula) > 1 Then
        If Left(rng.Formula, 1) = "=" Then
            isFormula = True
        End If
    End If
    
    'ひとつめの表示形式を取得
    Dim strNumFmt As String
    strNumFmt = rng.NumberFormat
    strNumFmt = Split(strNumFmt, ";", 2, vbBinaryCompare)(0)
    '（例えば、「mm";"dd」の様な場合に分断されてしまうが、困ることは考えにくいのでこのままとする）
    
    Select Case strNumFmt
        
        '時刻入力補助
        Case "hh:mm", "h:mm", "h:m"
            '数式が入力されている場合は確認MSG
            If isFormula = True Then
                If WarnFormula <> True Then
                    InputAssist = False
                    Exit Function
                End If
            End If

            ' 時刻入力補助処理
            Call InputAssistTime(rng)
            'イベントはキャンセルする
            Cancel = True

        '日付入力補助
        Case "m""月""d""日""", "m/d/yyyy", "yyyy/mm/dd", "m/dd/yyyy", "mm/dd", "m/d", "m/dd"
            
            '数式が入力されている場合は確認MSG
            If isFormula = True Then
                If WarnFormula <> True Then
                    InputAssist = False
                    Exit Function
                End If
            End If

            ' 日付入力補助処理
            Call InputAssistDate(rng)
            ' イベントはキャンセルする
            Cancel = True
            
        Case Else
            ' 何もしない

    End Select

    
    'チェックボックスを切り替える
    If Not Cancel Then Cancel = InputAssistRotateStatus(rng, Array("■", "□"))
    If Not Cancel Then Cancel = InputAssistRotateStatus(rng, Array("○", "×", "△"))
    
    InputAssist = Cancel

End Function

'*****************************************************************************
'[ 関数名 ]　WarnFormula
'[ 概  要 ]　数式が入力されている旨の警告＆続行確認MSG
'[ 引  数 ]　なし
'[ 戻り値 ]  続行OK:True/中断:False
'*****************************************************************************
Private Function WarnFormula()
    
    If ConfirmMessage("数式が入力されています。入力補助を続行してもよろしいですか？", vbOKCancel) = vbOK Then
        WarnFormula = True
    Else
        WarnFormula = False
    End If
    Exit Function
    
End Function

'*****************************************************************************
'[ 関数名 ]　InputAssistRotateStatus
'[ 概  要 ]　状態ローテート入力補助
'[ 引  数 ]　対象セル, 状態リスト
'[ 戻り値 ]  イベントキャンセルフラグ
'*****************************************************************************
Private Function InputAssistRotateStatus(ByRef rng As Range, ByRef statusTextArray As Variant) As Boolean
    Dim Cancel As Boolean   'イベントキャンセルフラグ
    
    Dim strText As String
    strText = rng.Text

    '状態リストのいずれかに一致したら次の状態に切り替える
    Dim i As Long
    For i = LBound(statusTextArray) To UBound(statusTextArray)
        If statusTextArray(i) = strText Then
            Dim lngIdx As Long
            lngIdx = i - LBound(statusTextArray)
            Dim lngLen As Long
            lngLen = UBound(statusTextArray) - LBound(statusTextArray) + 1
            Dim lngNextIdx As Long
            lngNextIdx = ((lngIdx + 1) Mod lngLen)
            Dim lngNextNo As Long
            lngNextNo = lngNextIdx + LBound(statusTextArray)
            rng.Value = statusTextArray(lngNextNo)
            ' イベントはキャンセルする
            Cancel = True
            
            Exit For    '検索終了
        End If
    Next i

    InputAssistRotateStatus = Cancel
End Function

'*****************************************************************************
'[ 関数名 ]　InputAssistDate
'[ 概  要 ]　日付入力補助
'[ 引  数 ]　対象セル
'[ 戻り値 ]　-
'*****************************************************************************
Public Sub InputAssistDate(ByVal Target As Range)
    Dim datTargetDate As Date
    Dim isInputCancel As Boolean
                
    'セルの元の値を取得
    datTargetDate = ConvertToDate(Target.Value)
    
    Load frmInputDate
    
    ' フォームの入力値にセルの値を設定
    Call frmInputDate.SetInputDate(datTargetDate)
    
    frmInputDate.Show
    
    ' フォームの入力値を取得
    datTargetDate = frmInputDate.GetInputDate
    ' 入力キャンセル有無を取得
    isInputCancel = frmInputDate.GetInputCancel
    
    Unload frmInputDate
    
    ' フォームでの入力確定後、対象セルに値を設定する
    If isInputCancel <> True Then
        If Not IsDate(Target.Value) Then
            Call SendKeysWithRetry(CStr(datTargetDate), True, 2)
        ElseIf CDate(Target.Value) <> datTargetDate Then
            Call SendKeysWithRetry(CStr(datTargetDate), True, 2)
        End If
    End If

End Sub

'*****************************************************************************
'[ 関数名 ]　InputAssistTime
'[ 概  要 ]　時刻入力補助
'[ 引  数 ]　対象セル
'[ 戻り値 ]　-
'*****************************************************************************
Public Sub InputAssistTime(ByVal Target As Range)
    Dim varInput As Variant '入力値
    
    'セルの元の値を取得
    varInput = ConvertToHHMM(Target.Value)
    
    '次のInputBoxで左右キーによるセル移動を防ぐために、F2キーを送って編集モードにする
    Call Application.SendKeys("{F2}")
    
    '時刻入力画面表示
    varInput = Application.InputBox("時刻をhhmm形式(コロン無し)で入力して下さい。", "時刻入力補助", varInput)
    '※↑InputBox関数だと、キャンセルか空文字か判別不可能なので、InputBoxメソッドを使用
    
    Call Sleep(100)
    
    'キャンセルボタン押下時以外なら、処理実施
    If varInput <> False Then
        
        Select Case Len(varInput)
            Case 4:
                varInput = Format(varInput, "@@:@@")
                'Call Application.SendKeys(varInput & "{Tab}+{Tab}")
                Call SendKeysWithRetry(varInput, True, 2)
            Case 3:
                varInput = Format(varInput, "@:@@")
                'Call Application.SendKeys(varInput & "{Tab}+{Tab}")
                Call SendKeysWithRetry(varInput, True, 2)
            Case 0:
                'クリア
                'Call Application.SendKeys("{Del}")
                Call SendKeysWithRetry("{Del}", False, 2)
            Case Else:
                'エラーMSG
                Call modMessage.ErrorMessage("hhmmまたはhmmの形式で入力して下さい。")
        End Select
        
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　ConvertToHHMM
'[ 概  要 ]　"HHMM"形式にフォーマット変換
'[ 引  数 ]　入力値
'[ 戻り値 ]  変換後の文字列（エラー時は空を返す）
'*****************************************************************************
Private Function ConvertToHHMM(ByVal varInput As Variant) As String
    
    Dim strResult As String
    
On Error GoTo Catch
    strResult = Format(varInput, "hhmm")
    GoTo Finally

Catch:
    'フォーマットエラー時は空文字列とする
    strResult = Empty
    
Finally:
    On Error GoTo 0
    
    ConvertToHHMM = strResult
End Function

'*****************************************************************************
'[ 関数名 ]　ConvertToDate
'[ 概  要 ]　日付型に変換
'[ 引  数 ]　入力値
'[ 戻り値 ]  変換後の日付（入力値が空白 又は エラー時には現在日付を返す）
'*****************************************************************************
Private Function ConvertToDate(ByVal varInput As Variant) As Date
    
    Dim datResult As Date
    
On Error GoTo Catch
    If varInput <> Empty Then
        datResult = CDate(varInput)
    Else
        datResult = Date
    End If
    
    GoTo Finally

Catch:
    'フォーマットエラー時は現在日付とする
    datResult = Date
    
Finally:
    On Error GoTo 0
    
    ConvertToDate = datResult
End Function

'*****************************************************************************
'[ 関数名 ]　SendKeysWithRetry
'[ 概  要 ]　SendKeyを行い、セルに入力を行う。セルの中身が変わっていない場合は、リトライする
'[ 引  数 ]　入力キー文字列
'            ActiveCellをその場に留めるために、最後に {Tab}→Shift + {Tab} を送るかどうか
'            リトライ回数(省略時=1:リトライなし)
'[ 戻り値 ]  なし
'*****************************************************************************
Private Sub SendKeysWithRetry(ByVal strSendKeysOrg As String, ByVal withStay As Boolean, Optional ByVal intRetry As Integer = 1)
    
    Dim rngCurCell As Range
    Set rngCurCell = ActiveCell
    
    Dim strCurVal As String
    strCurVal = rngCurCell.Text
    
    Dim strSendKeys As String
    strSendKeys = strSendKeysOrg
    
    'ActiveCellをその場に留める場合は、最後に {Tab}→Shift + {Tab} を送る
    If withStay Then
        strSendKeys = strSendKeys & "{Tab}+{Tab}"
    End If
    
    'IMEをOFFにする
    Call SetIMEMode(False)
    
    Do
        'キーを送出
        Call Application.SendKeys(strSendKeys)
        
        'リトライカウンタが0ならリトライしない
        intRetry = intRetry - 1
        If intRetry <= 0 Then Exit Do
        
        '200ms待つ
        'Call Sleep(200)
        
        'もしActiveCellの位置が変わってしまっていたらリトライしない
        If rngCurCell.Address <> ActiveCell.Address Then
            Debug.Print "<SendKeysWithRetry>ActiveCell Moved:" & rngCurCell.Address(False, False) & "⇒" & ActiveCell.Address(False, False)
            'Call rngCurCell.Activate
            Exit Do
        End If
        
        'セルの値が変わった場合は成功しているのでリトライしない
        If ActiveCell.Text <> strCurVal Then Exit Do
        
        'セルの値が変わっていない場合は怪しいのでリトライ
        Debug.Print "<SendKeysWithRetry>Retry:" & intRetry & " CellValue=" & ActiveCell.Text & "←" & strSendKeysOrg
    Loop
    
End Sub

