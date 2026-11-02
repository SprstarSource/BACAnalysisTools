Attribute VB_Name = "modBuildBacPanel"
Option Explicit

' =====================================================================
  BAC Panel Builder - Macro: BuildBacPanel
  Creates a highly professional Arabic dashboard with 2 buttons.
  Enforces workflow: Analysis button is locked until Extraction succeeds.
' =====================================================================

Private Const PANEL_SHEET As String = "ControlPanel"
Private Const CARD_WIDTH As Double = 720
Private Const CARD_HEIGHT As Double = 320
Private Const BTN_WIDTH As Double = 300
Private Const BTN_HEIGHT As Double = 65

Private m_Busy As Boolean

' ---------------------------------------------------------------------
' 1) Build the UI
' ---------------------------------------------------------------------
Public Sub BuildBacPanel()
    Dim ws As Worksheet, shp As Shape
    Dim cardLeft As Double, cardTop As Double, btnTop As Double
    Dim x1 As Double, x2 As Double
    
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
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    
    ' Worksheet Base Settings
    With ws
        .DisplayRightToLeft = True
        .Cells.Interior.Color = RGB(245, 247, 250) ' Light gray background
        .Columns("A:N").ColumnWidth = 8.5
        .Rows("1:20").RowHeight = 20
    End With
    
    ' ---------------------------------------------------------
    ' 1. Header Area (Navy Blue)
    ' ---------------------------------------------------------
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, ws.Range("A1:N7").Width, ws.Range("A1:N7").Height)
    With shp
        .Name = "HeaderBG"
        .Fill.ForeColor.RGB = RGB(18, 33, 58) ' Dark Navy
        .Line.Visible = msoFalse
    End With
    
    ' Header Title
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, ws.Range("A2").Top, ws.Range("A1:N7").Width, 50)
    With shp
        .Name = "HeaderTitle"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = Ar("main_title")
            .TextRange.Font.Name = "Tahoma"
            .TextRange.Font.Size = 25
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    
    ' Header Subtitle
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, ws.Range("A3").Top + 15, ws.Range("A1:N7").Width, 40)
    With shp
        .Name = "HeaderSubtitle"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = Ar("sub_title")
            .TextRange.Font.Name = "Tahoma"
            .TextRange.Font.Size = 15
            .TextRange.Font.Bold = msoFalse
            .TextRange.Font.Fill.ForeColor.RGB = RGB(200, 210, 225)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    
    ' ---------------------------------------------------------
    ' 2. Main Card (White with Shadow)
    ' ---------------------------------------------------------
    cardLeft = (ws.Range("A1:N10").Width - CARD_WIDTH) / 2
    cardTop = ws.Range("A8").Top + 20
    
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLeft, cardTop, CARD_WIDTH, CARD_HEIGHT)
    With shp
        .Name = "MainCard"
        .Adjustments(1) = 0.05
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoFalse
        ' Apply soft shadow
        .Shadow.Type = msoShadow21
        .Shadow.Transparency = 0.8
        .Shadow.Size = 110
        .Shadow.Blur = 15
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 4
    End With
    
    ' Card Title ("أزرار التحكم")
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, cardLeft + CARD_WIDTH - 250, cardTop + 20, 230, 40)
    With shp
        .Name = "CardTitle"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = Ar("card_title")
            .TextRange.Font.Name = "Tahoma"
            .TextRange.Font.Size = 18
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(18, 33, 58)
            .TextRange.ParagraphFormat.Alignment = msoAlignRight
        End With
    End With
    
    ' ---------------------------------------------------------
    ' 3. Buttons (Using StyleButtonShape)
    ' ---------------------------------------------------------
    btnTop = cardTop + 100
    ' IMPORTANT: On a DisplayRightToLeft worksheet, Excel visually mirrors
    ' shape placement. Use the left coordinate for the button that must
    ' appear on the RIGHT, and the right coordinate for the button on the LEFT.
    x1 = cardLeft + 40
    x2 = cardLeft + CARD_WIDTH - BTN_WIDTH - 40
    
    ' Extract Button (Green)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x1, btnTop, BTN_WIDTH, BTN_HEIGHT)
    StyleButtonShape shp, "btnBacExtract", Ar("btn_extract")
    shp.Fill.ForeColor.RGB = RGB(16, 124, 65) ' Green color for extract
    
    ' Analyze Button (Blue)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x2, btnTop, BTN_WIDTH, BTN_HEIGHT)
    StyleButtonShape shp, "btnBacAnalyze", Ar("btn_analyze")
    shp.Fill.ForeColor.RGB = RGB(28, 73, 169) ' Blue color for analyze
    
    ' ---------------------------------------------------------
    ' 4. Footer Info Bar
    ' ---------------------------------------------------------
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLeft + 30, cardTop + CARD_HEIGHT - 60, CARD_WIDTH - 60, 45)
    With shp
        .Name = "InfoBar"
        .Adjustments(1) = 0.2
        .Fill.ForeColor.RGB = RGB(240, 243, 248)
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = Ar("info_steps")
            .TextRange.Font.Name = "Tahoma"
            .TextRange.Font.Size = 12
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(80, 90, 105)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
    End With
    
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If
    
    RefreshBacPanelState
    Application.ScreenUpdating = True
    Exit Sub
    
BuildError:
    Application.ScreenUpdating = True
    MsgBox "Error building panel: " & Err.Description, vbExclamation, "Error"
End Sub

' ---------------------------------------------------------------------
' 2) Update Button States (Lock / Unlock)
' ---------------------------------------------------------------------
Public Sub RefreshBacPanelState()
    Dim ws As Worksheet, btnExt As Shape, btnAna As Shape
    Dim isReady As Boolean
    
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set btnExt = ws.Shapes("btnBacExtract")
    Set btnAna = ws.Shapes("btnBacAnalyze")
    On Error GoTo 0
    If btnExt Is Nothing Or btnAna Is Nothing Then Exit Sub
    
    ' Check if extraction succeeded
    isReady = False
    On Error Resume Next
    If Len(modBacEtab.gLastBacEtabPath) > 0 Then
        If Dir(modBacEtab.gLastBacEtabPath) <> "" Then isReady = True
    End If
    On Error GoTo 0
    
    ' Extract Button is always active (Green)
    FormatButton btnExt, RGB(16, 124, 65), RGB(255, 255, 255), True
    btnExt.OnAction = "'" & ThisWorkbook.Name & "'!RunBacStep1"
    
    ' Analyze Button (Blue if ready, Gray if locked)
    If isReady Then
        FormatButton btnAna, RGB(28, 73, 169), RGB(255, 255, 255), True
        btnAna.OnAction = "'" & ThisWorkbook.Name & "'!RunBacStep2"
    Else
        FormatButton btnAna, RGB(220, 225, 230), RGB(150, 160, 170), False
        btnAna.OnAction = "'" & ThisWorkbook.Name & "'!BacAnalysisGuard"
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

Public Sub BacAnalysisGuard()
    MsgBox Ar("msg_guard"), vbExclamation, Ar("msg_guard_title")
End Sub

' ---------------------------------------------------------------------
' 4) Helpers
' ---------------------------------------------------------------------
Private Function PanelSheet() As Worksheet
    On Error Resume Next
    Set PanelSheet = ThisWorkbook.Worksheets(PANEL_SHEET)
    On Error GoTo 0
End Function

Private Sub FormatButton(ByVal btn As Shape, ByVal bgColor As Long, ByVal textColor As Long, ByVal isEnabled As Boolean)
    With btn
        .Fill.Solid
        .Fill.ForeColor.RGB = bgColor
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = textColor
        If isEnabled Then
            ' Add subtle shadow to active buttons
            .Shadow.Type = msoShadow21
            .Shadow.Transparency = 0.6
            .Shadow.Size = 100
            .Shadow.Blur = 10
            .Shadow.OffsetY = 2
        Else
            ' Remove shadow from disabled buttons
            .Shadow.Visible = msoFalse
        End If
    End With
End Sub

' --- YOUR CUSTOM SUBROUTINE ---
Private Sub StyleButtonShape(ByVal btn As Shape, ByVal shapeName As String, ByVal captionText As String)
    ' Fallback in case UI_FONT is not declared globally
    Dim fontName As String
    fontName = "Tahoma"
    
    With btn
        .Name = shapeName
        .Placement = xlFreeFloating
        .Line.Weight = 1.25
        
        ' Safely apply adjustments (only works for rounded rectangles)
        On Error Resume Next
        .Adjustments(1) = 0.22
        On Error GoTo 0
        
        With .TextFrame2
            .TextRange.Text = captionText
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 10
            .MarginRight = 10
            .MarginTop = 3
            .MarginBottom = 3
            With .TextRange.Font
                .Name = fontName
                .NameComplexScript = fontName
                .NameAscii = fontName
                .Size = 13
                .Bold = msoTrue
                .Fill.ForeColor.RGB = RGB(255, 255, 255) ' White text
            End With
        End With
    End With
End Sub
' ------------------------------

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
    Case "info_steps": Ar = ArHex("62E 637 648 627 62A 20 627 644 639 645 644 3A 20 64A 631 62C 649 20 627 633 62A 62E 631 627 62C 20 627 644 628 64A 627 646 627 62A 20 623 648 644 627 64B 20 62B 645 20 627 644 636 63A 637 20 639 644 649 20 62A 62D 644 64A 644 20 627 644 646 62A 627 626 62C 2E")
    Case "msg_guard": Ar = ArHex("64A 62C 628 20 62A 634 63A 64A 644 20 27 627 633 62A 62E 631 627 62C 20 646 62A 627 626 62C 20 627 644 645 624 633 633 629 27 20 623 648 644 627 64B 20 648 628 646 62C 627 62D 20 644 625 646 634 627 621 20 627 644 645 644 641 20 627 644 645 637 644 648 628 20 644 644 62A 62D 644 64A 644 2E")
    Case "msg_guard_title": Ar = ArHex("62E 637 648 629 20 645 637 644 648 628 629")
    End Select
End Function