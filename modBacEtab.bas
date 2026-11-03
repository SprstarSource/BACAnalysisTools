Attribute VB_Name = "modBuildBacPanel"

Option Explicit

' =====================================================================
'  BAC Panel Builder - Macro: BuildBacPanel
'  Creates a professional Arabic dashboard with 3 independent buttons
'  (extract / fetch repeat status / analyze). All are always active:
'  they can be run in any order.
' =====================================================================

Private Const PANEL_SHEET As String = "ControlPanel"
Private Const UI_FONT As String = "Hacen Saudi Arabia XL"   ' change the font in ONE place
Private Const CARD_WIDTH As Double = 780
Private Const CARD_HEIGHT As Double = 250
Private Const CARD_MARGIN As Double = 40     ' inner left/right padding of the card
Private Const BTN_GAP As Double = 24         ' space between the buttons
Private Const BTN_FONT As Single = 17
Private Const BTN_HEIGHT As Double = 65

Private m_Busy As Boolean

' ---------------------------------------------------------------------
' 1) Build the UI
' ---------------------------------------------------------------------
Public Sub BuildBacPanel()
    Dim ws As Worksheet, shp As Shape
    Dim panelW As Double, headerH As Double
    Dim cardW As Double, cardLeft As Double, cardTop As Double
    Dim btnW As Double, btnTop As Double, x1 As Double, x2 As Double, x3 As Double
    Dim i As Long
    
    On Error GoTo BuildError
    Application.ScreenUpdating = False
    
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 801, , "The tool workbook is read-only."
    
    ' Setup worksheet
    Set ws = PanelSheet()
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = PANEL_SHEET
    End If
    
    If ws.ProtectContents Or ws.ProtectDrawingObjects Then
        Err.Raise vbObjectError + 803, , "Unprotect the control panel sheet before rebuilding it."
    End If
    
    ws.Activate
    ws.Cells.Clear
    ws.Cells.UnMerge
    
    ' Delete any existing shapes
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    
    ' Worksheet base settings
    With ws
        .DisplayRightToLeft = True
        .Cells.Interior.Color = RGB(245, 247, 250) ' Light gray background
        .Columns("A:N").ColumnWidth = 8.5
        .Rows("1:20").RowHeight = 20
    End With
    
    panelW = ws.Range("A1:N1").width
    If panelW < CARD_WIDTH + 40 Then            ' widen the columns so the card fits
        ws.Columns("A:N").ColumnWidth = 8.5 * (CARD_WIDTH + 40) / panelW
        panelW = ws.Range("A1:N1").width
    End If
    headerH = ws.Range("A1:N7").Height
    
    ' ---------------------------------------------------------
    ' 1. Header (Navy) + title + subtitle + gold accent line
    ' ---------------------------------------------------------
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, panelW, headerH)
    With shp
        .Name = "HeaderBG"
        .Fill.ForeColor.RGB = RGB(18, 33, 58) ' Dark Navy
        .Line.Visible = msoFalse
    End With
    
    AddLabel ws, "HeaderTitle", Ar("main_title"), 0, 24, panelW, 50, 30, True, RGB(255, 255, 255)
    AddLabel ws, "HeaderSubtitle", Ar("sub_title"), 0, 76, panelW, 32, 15, False, RGB(185, 200, 220)
    
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, headerH - 4, panelW, 4)
    With shp
        .Name = "HeaderAccent"
        .Fill.ForeColor.RGB = RGB(212, 175, 55) ' Gold accent
        .Line.Visible = msoFalse
    End With
    
    ' ---------------------------------------------------------
    ' 2. Main Card (White with soft shadow)
    ' ---------------------------------------------------------
    cardW = CARD_WIDTH
    If cardW > panelW - 40 Then cardW = panelW - 40   ' never wider than the sheet
    cardLeft = (panelW - cardW) / 2
    cardTop = ws.Range("A8").Top + 20
    
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLeft, cardTop, cardW, CARD_HEIGHT)
    With shp
        .Name = "MainCard"
        .Adjustments(1) = 0.05
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoFalse
        .Shadow.Type = msoShadow21
        .Shadow.Transparency = 0.8
        .Shadow.Size = 110
        .Shadow.Blur = 15
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 4
    End With
    
    ' Card title + thin divider
    AddLabel ws, "CardTitle", Ar("card_title"), cardLeft, cardTop + 14, cardW, 36, 17, True, RGB(18, 33, 58)
    
    Set shp = ws.Shapes.AddLine(cardLeft + CARD_MARGIN, cardTop + 58, cardLeft + cardW - CARD_MARGIN, cardTop + 58)
    With shp
        .Name = "CardDivider"
        .Line.ForeColor.RGB = RGB(225, 230, 238)
        .Line.Weight = 1
    End With
    
    ' ---------------------------------------------------------
    ' 3. Buttons (independent - no required order)
    ' ---------------------------------------------------------
    btnW = (cardW - 2 * CARD_MARGIN - 2 * BTN_GAP) / 3
    btnTop = cardTop + 80
    ' IMPORTANT: On a DisplayRightToLeft worksheet, Excel visually mirrors
    ' shape placement. The smallest left coordinate appears on the RIGHT.
    ' Visual order (right to left) = workflow order: extract, repeat, analyze.
    x1 = cardLeft + CARD_MARGIN
    x2 = x1 + btnW + BTN_GAP
    x3 = x2 + btnW + BTN_GAP
    
    ' Extract Button (Green) - appears on the right
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x1, btnTop, btnW, BTN_HEIGHT)
    StyleButtonShape shp, "btnBacExtract", Ar("btn_extract"), RGB(16, 124, 65)
    
    ' Fetch repeat status Button (Amber) - appears in the middle
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x2, btnTop, btnW, BTN_HEIGHT)
    StyleButtonShape shp, "btnBacRepeat", Ar("btn_repeat"), RGB(168, 96, 10)
    
    ' Analyze Button (Blue) - appears on the left
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x3, btnTop, btnW, BTN_HEIGHT)
    StyleButtonShape shp, "btnBacAnalyze", Ar("btn_analyze"), RGB(28, 73, 169)
    
    ' ---------------------------------------------------------
    ' 4. Footer status bar (text is set by RefreshBacPanelState)
    ' ---------------------------------------------------------
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLeft + CARD_MARGIN, cardTop + CARD_HEIGHT - 70, cardW - 2 * CARD_MARGIN, 45)
    With shp
        .Name = "InfoBar"
        .Adjustments(1) = 0.2
        .Fill.ForeColor.RGB = RGB(240, 243, 248)
        .Line.Visible = msoFalse
    End With
    ApplyText shp, Ar("info_hint"), 14, False, RGB(80, 90, 105)
    
    ' Keep every shape fixed in place (don't move/resize with cells)
    For Each shp In ws.Shapes
        shp.Placement = xlFreeFloating
    Next shp
    
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ws.Range("A1").Select
    End If
    
    RefreshBacPanelState
    Application.ScreenUpdating = True
    Exit Sub
    
BuildError:
    Application.ScreenUpdating = True
    MsgBox "Error building panel: " & Err.Description, vbExclamation, "Error"
End Sub

' ---------------------------------------------------------------------
' 2) Refresh panel: wire the buttons + update the status bar
' ---------------------------------------------------------------------
Public Sub RefreshBacPanelState()
    Dim ws As Worksheet, btnExt As Shape, btnAna As Shape, btnRep As Shape, bar As Shape
    Dim hasFile As Boolean, fileName As String
    
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set btnExt = ws.Shapes("btnBacExtract")
    Set btnAna = ws.Shapes("btnBacAnalyze")
    Set btnRep = ws.Shapes("btnBacRepeat")
    Set bar = ws.Shapes("InfoBar")
    On Error GoTo 0
    If btnExt Is Nothing Or btnAna Is Nothing Then Exit Sub
    
    ' All buttons are ALWAYS active - they can run in any order
    btnExt.OnAction = "'" & ThisWorkbook.Name & "'!RunBacStep1"
    btnAna.OnAction = "'" & ThisWorkbook.Name & "'!RunBacStep2"
    If Not btnRep Is Nothing Then btnRep.OnAction = "'" & ThisWorkbook.Name & "'!RunBacRepeat"
    
    ' Status bar: show the last extracted file if there is one
    If bar Is Nothing Then Exit Sub
    
    On Error Resume Next
    If Len(modBacEtab.gLastBacEtabPath) > 0 Then
        If Len(Dir(modBacEtab.gLastBacEtabPath)) > 0 Then
            hasFile = True
            fileName = Mid$(modBacEtab.gLastBacEtabPath, InStrRev(modBacEtab.gLastBacEtabPath, "\") + 1)
        End If
    End If
    On Error GoTo 0
    
    If hasFile Then
        ApplyText bar, Ar("status_last") & fileName, 14, True, RGB(16, 124, 65)
    Else
        ApplyText bar, Ar("info_hint"), 14, False, RGB(80, 90, 105)
    End If
End Sub

' ---------------------------------------------------------------------
' 3) Macros
' ---------------------------------------------------------------------
Public Sub RunBacStep1()
    If m_Busy Then Exit Sub
    m_Busy = True
    On Error GoTo Failed
    Application.Run "'" & ThisWorkbook.Name & "'!modBacEtab.Run_BacEtab"
    RefreshBacPanelState
    m_Busy = False
    Exit Sub
Failed:
    m_Busy = False
End Sub

Public Sub RunBacStep2()
    If m_Busy Then Exit Sub
    m_Busy = True
    On Error GoTo Failed
    Application.Run "'" & ThisWorkbook.Name & "'!modBacAnalysis.Run_BacAnalysis"
    m_Busy = False
    Exit Sub
Failed:
    m_Busy = False
End Sub

Public Sub RunBacRepeat()
    If m_Busy Then Exit Sub
    m_Busy = True
    On Error GoTo Failed
    ' Run_RepeatStatus reports its own errors; anything reaching Failed means
    ' the macro could not be started (e.g. modRepeatStatus was not imported).
    Application.Run "'" & ThisWorkbook.Name & "'!modRepeatStatus.Run_RepeatStatus"
    m_Busy = False
    Exit Sub
Failed:
    m_Busy = False
    MsgBox Ar("msg_missing") & vbLf & Err.Description, _
           vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading
End Sub

' ---------------------------------------------------------------------
' 4) Helpers
' ---------------------------------------------------------------------
Private Function PanelSheet() As Worksheet
    On Error Resume Next
    Set PanelSheet = ThisWorkbook.Worksheets(PANEL_SHEET)
    On Error GoTo 0
End Function

' Transparent text-only shape (titles / subtitles)
Private Function AddLabel(ByVal ws As Worksheet, ByVal shapeName As String, ByVal txt As String, _
        ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
        ByVal fontSize As Single, ByVal isBold As Boolean, ByVal textColor As Long) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
    With shp
        .Name = shapeName
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
    ApplyText shp, txt, fontSize, isBold, textColor
    Set AddLabel = shp
End Function

' One place that formats the text of any shape
Private Sub ApplyText(ByVal shp As Shape, ByVal txt As String, ByVal fontSize As Single, _
        ByVal isBold As Boolean, ByVal textColor As Long)
    With shp.TextFrame2
        .WordWrap = msoTrue
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 10
        .MarginRight = 10
        .MarginTop = 3
        .MarginBottom = 3
        With .TextRange
            .Text = txt
            .ParagraphFormat.Alignment = msoAlignCenter
            On Error Resume Next
            .ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
            On Error GoTo 0
            With .Font
                .Name = UI_FONT
                .NameComplexScript = UI_FONT
                .Size = fontSize
                .Bold = IIf(isBold, msoTrue, msoFalse)
                .Fill.ForeColor.RGB = textColor
            End With
        End With
    End With
End Sub

' Flat modern button: solid fill, no outline, soft shadow
Private Sub StyleButtonShape(ByVal btn As Shape, ByVal shapeName As String, _
        ByVal captionText As String, ByVal fillColor As Long)
    With btn
        .Name = shapeName
        .Placement = xlFreeFloating
        .Fill.Solid
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
        
        ' Adjustments only work for rounded rectangles
        On Error Resume Next
        .Adjustments(1) = 0.22
        On Error GoTo 0
        
        .Shadow.Type = msoShadow21
        .Shadow.Transparency = 0.6
        .Shadow.Size = 100
        .Shadow.Blur = 10
        .Shadow.OffsetY = 2
    End With
    ApplyText btn, captionText, BTN_FONT, True, RGB(255, 255, 255)
End Sub

' ---------------------------------------------------------------------
' Arabic Text Encoder
' ---------------------------------------------------------------------
Private Function ArHex(ByVal codes As String) As String
    Dim token As Variant
    For Each token In Split(codes, " ")
        If Len(CStr(token)) > 0 Then ArHex = ArHex & ChrW(CLng("&H" & CStr(token)))
    Next token
End Function

Private Function Ar(ByVal key As String) As String
    Select Case key
    Case "main_title": Ar = ArHex("644 648 62D 629 20 625 62F 627 631 629 20 646 62A 627 626 62C 20 627 644 628 643 627 644 648 631 64A 627")
    Case "sub_title": Ar = ArHex("646 638 627 645 20 625 62F 627 631 629 20 648 62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 645 624 633 633 629 20 627 644 62A 639 644 64A 645 64A 629")
    Case "card_title": Ar = ArHex("623 632 631 627 631 20 627 644 62A 62D 643 645")
    Case "btn_extract": Ar = ArHex("627 633 62A 62E 631 627 62C 20 646 62A 627 626 62C 20 627 644 645 624 633 633 629")
    Case "btn_analyze": Ar = ArHex("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 645 624 633 633 629")
    Case "btn_repeat": Ar = ArHex("62C 644 628 20 62D 627 644 627 62A 20 627 644 625 639 627 62F 629")
    Case "msg_missing": Ar = ArHex("62A 639 630 631 20 62A 634 63A 64A 644 20 627 644 645 627 643 631 648 2E 20 62A 623 643 62F 20 645 646 20 627 633 62A 64A 631 627 62F 20 627 644 648 62D 62F 629 20 6D 6F 64 52 65 70 65 61 74 53 74 61 74 75 73 20 641 64A 20 627 644 645 634 631 648 639 2E")
    Case "info_hint": Ar = ArHex("627 62E 62A 631 20 627 644 639 645 644 64A 629 20 627 644 645 637 644 648 628 629 3A 20 64A 645 643 646 643 20 62A 634 63A 64A 644 20 623 64A 20 632 631 20 641 64A 20 623 64A 20 648 642 62A 2E")
    Case "status_last": Ar = ArHex("622 62E 631 20 645 644 641 20 645 633 62A 62E 631 62C 3A 20")
    End Select
End Function
