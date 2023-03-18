Attribute VB_Name = "modTest"
Option Explicit


'====================================================================================
Public Sub test_ScrollTo()
    Call modCommon.ScrollTo(Range("A1"))
End Sub



Public Sub OutputCommandIDList()
    Dim shtTarget As Worksheet
    Set shtTarget = ActiveSheet
    Dim i As Long
    
    Call shtTarget.Range("A:B").Delete
    With CommandBars("Standard")
        For i = 1 To .Controls.Count
            shtTarget.Cells(i, 1) = .Controls(i).Caption
            shtTarget.Cells(i, 2) = .Controls(i).ID
        Next i
    End With
    Call shtTarget.Range("A:B").AutoFit
End Sub


Public Sub testCopyScrollAndFreeze()
    
    Dim lngFreeze As Long
    Dim lngScroll As Long
    
    'Win1側
    If ActiveWindow.FreezePanes Then
        Call ActiveWindow.SmallScroll(Up:=65535)
        
        lngFreeze = ActiveWindow.ScrollRow
        
        ActiveWindow.FreezePanes = False
        lngScroll = ActiveWindow.ScrollRow
        
    End If
    
    '新しいウィンドウを開く
    
    'Win2側
    If lngFreeze <> 0 Then
        ActiveWindow.ScrollRow = lngScroll
        
        ActiveWindow.ActiveSheet.Cells(lngFreeze, 1).Activate
        ActiveWindow.FreezePanes = True
    End If

End Sub


Public Sub GetFreezePosition()
    
    Dim shtCur As Worksheet
    Dim lngRow As Long
    Dim lngCol As Long
    Dim rngPos As Range
    
    If ActiveWindow.FreezePanes Then
        lngRow = ActiveWindow.SplitRow
        lngCol = ActiveWindow.SplitColumn
    End If
    Set shtCur = ActiveSheet
    Set rngPos = shtCur.Cells(lngRow + 1, lngCol + 1)
    
    Call modMessage.InfoMessage( _
            "Address: " & rngPos.Address(False, False) & vbCrLf & _
            "SplitRow: " & lngRow & vbCrLf & _
            "SplitCol: " & lngCol)
    
End Sub

Private Sub テキストボックスの255文字問題検証用()

    Dim shpList As Shapes
    Set shpList = ActiveSheet.Shapes
    
    'テキスト
    Dim shpTextBox As Shape
    Set shpTextBox = shpList.AddTextbox(msoTextOrientationHorizontal, _
                                10, 10, _
                                65, 400)
    shpTextBox.Line.Visible = msoTrue
    'shpTextBox.TextFrame.AutoSize = True
    
    Dim strTest As String
    Dim i As Long
    shpTextBox.TextFrame.Characters.Text = "test"
    strTest = ""
    For i = 1 To 25
        strTest = strTest & Format(i, "A000000000")
    Next
    strTest = strTest & "1234５"     '←255文字だと大丈夫
    'strTest = strTest & "123456"   '←256文字にすると、2003だとエラーとなる。2007は不明。

    shpTextBox.TextFrame.Characters.Text = strTest
    
    Call MsgBox(Right(shpTextBox.TextFrame.Characters.Text, 5))
    
    Call shpTextBox.Delete
    
End Sub



'オートシェイプの情報の一覧をイミディエイトウィンドウにデバッグ出力
Private Sub PutAutoShapeList()
    
    Dim shpList As Shapes
    Dim shpCur As Shape
    
    Set shpList = ActiveSheet.Shapes
    
    Debug.Print "【Shapes: " & shpList.Count & "】"
    For Each shpCur In shpList
        Debug.Print "ID:" & shpCur.ID & _
                    "／Name:" & shpCur.Name & _
                    "／ZOrder:" & Right("000" & shpCur.ZOrderPosition, 3) & _
                    "／Pos:(" & shpCur.Left & ", " & shpCur.Top & ")-(" & (shpCur.Left + shpCur.Width) & ", " & (shpCur.Top + shpCur.Height) & ")" & _
                    "／CellRange:" & shpCur.TopLeftCell.Address(False, False) & ":" & shpCur.BottomRightCell.Address(False, False) & _
                    "／" & GetShapeTypeName(shpCur.AutoShapeType) & _
                    ""
    Next
    
End Sub

'オートシェイプの種別を文字列で返す
Private Function GetShapeTypeName(ByVal shapeType As MsoAutoShapeType) As String
    Dim strTypeName As String
    
    Select Case shapeType
        Case msoShape24pointStar: strTypeName = "msoShape24pointStar"
        Case msoShape4pointStar: strTypeName = "msoShape4pointStar"
        Case msoShape8pointStar: strTypeName = "msoShape8pointStar"
        Case msoShapeActionButtonBeginning: strTypeName = "msoShapeActionButtonBeginning"
        Case msoShapeActionButtonDocument: strTypeName = "msoShapeActionButtonDocument"
        Case msoShapeActionButtonForwardorNext: strTypeName = "msoShapeActionButtonForwardorNext"
        Case msoShapeActionButtonHome: strTypeName = "msoShapeActionButtonHome"
        Case msoShapeActionButtonMovie: strTypeName = "msoShapeActionButtonMovie"
        Case msoShapeActionButtonSound: strTypeName = "msoShapeActionButtonSound"
        Case msoShapeBalloon: strTypeName = "msoShapeBalloon"
        Case msoShapeBentUpArrow: strTypeName = "msoShapeBentUpArrow"
        Case msoShapeBlockArc: strTypeName = "msoShapeBlockArc"
        Case msoShapeChevron: strTypeName = "msoShapeChevron"
        Case msoShapeCloudCallout: strTypeName = "msoShapeCloudCallout"
        Case msoShapeCube: strTypeName = "msoShapeCube"
        Case msoShapeCurvedDownRibbon: strTypeName = "msoShapeCurvedDownRibbon"
        Case msoShapeCurvedRightArrow: strTypeName = "msoShapeCurvedRightArrow"
        Case msoShapeCurvedUpRibbon: strTypeName = "msoShapeCurvedUpRibbon"
        Case msoShapeDonut: strTypeName = "msoShapeDonut"
        Case msoShapeDoubleBracket: strTypeName = "msoShapeDoubleBracket"
        Case msoShapeDownArrow: strTypeName = "msoShapeDownArrow"
        Case msoShapeDownRibbon: strTypeName = "msoShapeDownRibbon"
        Case msoShapeExplosion2: strTypeName = "msoShapeExplosion2"
        Case msoShapeFlowchartCard: strTypeName = "msoShapeFlowchartCard"
        Case msoShapeFlowchartConnector: strTypeName = "msoShapeFlowchartConnector"
        Case msoShapeFlowchartDecision: strTypeName = "msoShapeFlowchartDecision"
        Case msoShapeFlowchartDirectAccessStorage: strTypeName = "msoShapeFlowchartDirectAccessStorage"
        Case msoShapeFlowchartDisplay: strTypeName = "msoShapeFlowchartDisplay"
        Case msoShapeFlowchartDocument: strTypeName = "msoShapeFlowchartDocument"
        Case msoShapeFlowchartExtract: strTypeName = "msoShapeFlowchartExtract"
        Case msoShapeFlowchartInternalStorage: strTypeName = "msoShapeFlowchartInternalStorage"
        Case msoShapeFlowchartMagneticDisk: strTypeName = "msoShapeFlowchartMagneticDisk"
        Case msoShapeFlowchartManualInput: strTypeName = "msoShapeFlowchartManualInput"
        Case msoShapeFlowchartManualOperation: strTypeName = "msoShapeFlowchartManualOperation"
        Case msoShapeFlowchartMerge: strTypeName = "msoShapeFlowchartMerge"
        Case msoShapeFlowchartMultidocument: strTypeName = "msoShapeFlowchartMultidocument"
        Case msoShapeFlowchartOffpageConnector: strTypeName = "msoShapeFlowchartOffpageConnector"
        Case msoShapeFlowchartOr: strTypeName = "msoShapeFlowchartOr"
        Case msoShapeFlowchartPredefinedProcess: strTypeName = "msoShapeFlowchartPredefinedProcess"
        Case msoShapeFlowchartPreparation: strTypeName = "msoShapeFlowchartPreparation"
        Case msoShapeFlowchartProcess: strTypeName = "msoShapeFlowchartProcess"
        Case msoShapeFlowchartPunchedTape: strTypeName = "msoShapeFlowchartPunchedTape"
        Case msoShapeFlowchartSequentialAccessStorage: strTypeName = "msoShapeFlowchartSequentialAccessStorage"
        Case msoShapeFlowchartSort: strTypeName = "msoShapeFlowchartSort"
        Case msoShapeFlowchartStoredData: strTypeName = "msoShapeFlowchartStoredData"
        Case msoShapeFlowchartSummingJunction: strTypeName = "msoShapeFlowchartSummingJunction"
        Case msoShapeFlowchartTerminator: strTypeName = "msoShapeFlowchartTerminator"
        Case msoShapeFoldedCorner: strTypeName = "msoShapeFoldedCorner"
        Case msoShapeHeart: strTypeName = "msoShapeHeart"
        Case msoShapeHexagon: strTypeName = "msoShapeHexagon"
        Case msoShapeHorizontalScroll: strTypeName = "msoShapeHorizontalScroll"
        Case msoShapeIsoscelesTriangle: strTypeName = "msoShapeIsoscelesTriangle"
        Case msoShapeLeftArrow: strTypeName = "msoShapeLeftArrow"
        Case msoShapeLeftArrowCallout: strTypeName = "msoShapeLeftArrowCallout"
        Case msoShapeLeftBrace: strTypeName = "msoShapeLeftBrace"
        Case msoShapeLeftBracket: strTypeName = "msoShapeLeftBracket"
        Case msoShapeLeftRightArrow: strTypeName = "msoShapeLeftRightArrow"
        Case msoShapeLeftRightArrowCallout: strTypeName = "msoShapeLeftRightArrowCallout"
        Case msoShapeLeftRightUpArrow: strTypeName = "msoShapeLeftRightUpArrow"
        Case msoShapeLeftUpArrow: strTypeName = "msoShapeLeftUpArrow"
        Case msoShapeLightningBolt: strTypeName = "msoShapeLightningBolt"
        Case msoShapeLineCallout1: strTypeName = "msoShapeLineCallout1"
        Case msoShapeLineCallout1AccentBar: strTypeName = "msoShapeLineCallout1AccentBar"
        Case msoShapeLineCallout1BorderandAccentBar: strTypeName = "msoShapeLineCallout1BorderandAccentBar"
        Case msoShapeLineCallout1NoBorder: strTypeName = "msoShapeLineCallout1NoBorder"
        Case msoShapeLineCallout2: strTypeName = "msoShapeLineCallout2"
        Case msoShapeLineCallout2AccentBar: strTypeName = "msoShapeLineCallout2AccentBar"
        Case msoShapeLineCallout2BorderandAccentBar: strTypeName = "msoShapeLineCallout2BorderandAccentBar"
        Case msoShapeLineCallout2NoBorder: strTypeName = "msoShapeLineCallout2NoBorder"
        Case msoShapeLineCallout3: strTypeName = "msoShapeLineCallout3"
        Case msoShapeLineCallout3AccentBar: strTypeName = "msoShapeLineCallout3AccentBar"
        Case msoShapeLineCallout3BorderandAccentBar: strTypeName = "msoShapeLineCallout3BorderandAccentBar"
        Case msoShapeLineCallout3NoBorder: strTypeName = "msoShapeLineCallout3NoBorder"
        Case msoShapeLineCallout4: strTypeName = "msoShapeLineCallout4"
        Case msoShapeLineCallout4AccentBar: strTypeName = "msoShapeLineCallout4AccentBar"
        Case msoShapeLineCallout4BorderandAccentBar: strTypeName = "msoShapeLineCallout4BorderandAccentBar"
        Case msoShapeLineCallout4NoBorder: strTypeName = "msoShapeLineCallout4NoBorder"
        Case msoShapeMixed: strTypeName = "msoShapeMixed"
        Case msoShapeMoon: strTypeName = "msoShapeMoon"
        Case msoShapeNoSymbol: strTypeName = "msoShapeNoSymbol"
        Case msoShapeNotchedRightArrow: strTypeName = "msoShapeNotchedRightArrow"
        Case msoShapeNotPrimitive: strTypeName = "msoShapeNotPrimitive"
        Case msoShapeOctagon: strTypeName = "msoShapeOctagon"
        Case msoShapeOval: strTypeName = "msoShapeOval"
        Case msoShapeOvalCallout: strTypeName = "msoShapeOvalCallout"
        Case msoShapeParallelogram: strTypeName = "msoShapeParallelogram"
        Case msoShapePentagon: strTypeName = "msoShapePentagon"
        Case msoShapePlaque: strTypeName = "msoShapePlaque"
        Case msoShapeQuadArrowCallout: strTypeName = "msoShapeQuadArrowCallout"
        Case msoShapeRectangularCallout: strTypeName = "msoShapeRectangularCallout"
        Case msoShapeRightArrow: strTypeName = "msoShapeRightArrow"
        Case msoShapeRightBrace: strTypeName = "msoShapeRightBrace"
        Case msoShapeRightTriangle: strTypeName = "msoShapeRightTriangle"
        Case msoShapeRoundedRectangularCallout: strTypeName = "msoShapeRoundedRectangularCallout"
        Case msoShapeStripedRightArrow: strTypeName = "msoShapeStripedRightArrow"
        Case msoShapeTrapezoid: strTypeName = "msoShapeTrapezoid"
        Case msoShapeUpArrowCallout: strTypeName = "msoShapeUpArrowCallout"
        Case msoShapeUpDownArrowCallout: strTypeName = "msoShapeUpDownArrowCallout"
        Case msoShapeUTurnArrow: strTypeName = "msoShapeUTurnArrow"
        Case msoShapeWave: strTypeName = "msoShapeWave"
        Case msoShape16pointStar: strTypeName = "msoShape16pointStar"
        Case msoShape32pointStar: strTypeName = "msoShape32pointStar"
        Case msoShape5pointStar: strTypeName = "msoShape5pointStar"
        Case msoShapeActionButtonBackorPrevious: strTypeName = "msoShapeActionButtonBackorPrevious"
        Case msoShapeActionButtonCustom: strTypeName = "msoShapeActionButtonCustom"
        Case msoShapeActionButtonEnd: strTypeName = "msoShapeActionButtonEnd"
        Case msoShapeActionButtonHelp: strTypeName = "msoShapeActionButtonHelp"
        Case msoShapeActionButtonInformation: strTypeName = "msoShapeActionButtonInformation"
        Case msoShapeActionButtonReturn: strTypeName = "msoShapeActionButtonReturn"
        Case msoShapeArc: strTypeName = "msoShapeArc"
        Case msoShapeBentArrow: strTypeName = "msoShapeBentArrow"
        Case msoShapeBevel: strTypeName = "msoShapeBevel"
        Case msoShapeCan: strTypeName = "msoShapeCan"
        Case msoShapeCircularArrow: strTypeName = "msoShapeCircularArrow"
        Case msoShapeCross: strTypeName = "msoShapeCross"
        Case msoShapeCurvedDownArrow: strTypeName = "msoShapeCurvedDownArrow"
        Case msoShapeCurvedLeftArrow: strTypeName = "msoShapeCurvedLeftArrow"
        Case msoShapeCurvedUpArrow: strTypeName = "msoShapeCurvedUpArrow"
        Case msoShapeDiamond: strTypeName = "msoShapeDiamond"
        Case msoShapeDoubleBrace: strTypeName = "msoShapeDoubleBrace"
        Case msoShapeDoubleWave: strTypeName = "msoShapeDoubleWave"
        Case msoShapeDownArrowCallout: strTypeName = "msoShapeDownArrowCallout"
        Case msoShapeExplosion1: strTypeName = "msoShapeExplosion1"
        Case msoShapeFlowchartAlternateProcess: strTypeName = "msoShapeFlowchartAlternateProcess"
        Case msoShapeFlowchartCollate: strTypeName = "msoShapeFlowchartCollate"
        Case msoShapeFlowchartData: strTypeName = "msoShapeFlowchartData"
        Case msoShapeFlowchartDelay: strTypeName = "msoShapeFlowchartDelay"
        Case msoShapeQuadArrow: strTypeName = "msoShapeQuadArrow"
        Case msoShapeRectangle: strTypeName = "msoShapeRectangle"
        Case msoShapeRegularPentagon: strTypeName = "msoShapeRegularPentagon"
        Case msoShapeRightArrowCallout: strTypeName = "msoShapeRightArrowCallout"
        Case msoShapeRightBracket: strTypeName = "msoShapeRightBracket"
        Case msoShapeRoundedRectangle: strTypeName = "msoShapeRoundedRectangle"
        Case msoShapeSmileyFace: strTypeName = "msoShapeSmileyFace"
        Case msoShapeSun: strTypeName = "msoShapeSun"
        Case msoShapeUpArrow: strTypeName = "msoShapeUpArrow"
        Case msoShapeUpDownArrow: strTypeName = "msoShapeUpDownArrow"
        Case msoShapeUpRibbon: strTypeName = "msoShapeUpRibbon"
        Case msoShapeVerticalScroll: strTypeName = "msoShapeVerticalScroll"
        Case Default: strTypeName = "Error!"
    End Select
    
    GetShapeTypeName = strTypeName
End Function

