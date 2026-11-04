Attribute VB_Name = "modBuildBacPanel"

Option Explicit

' =====================================================================
'  BAC Panel Builder - Macro: BuildBacPanel
'  Creates a professional Arabic dashboard:
'   - one MAIN button (process the whole BAC results in one click:
'     extract -> fetch repeat status -> analyze, in this order)
'   - 3 independent step buttons (extract / fetch repeat status / analyze),
'     always active, they can be run in any order.
'   - a status chip under each step button + a status bar: progress and
'     result are shown on the panel itself (no confirmation needed).
'
'  MAIN button rules:
'   * every stage module reports its OWN status (success / cancelled / failed
'     + detail); the main button never assumes that "the macro returned" means
'     "the macro succeeded".
'   * the sequence stops at the FIRST stage that does not succeed; one final
'     report lists: stages that succeeded, the stage that failed (and why),
'     and confirms that the following stages were NOT executed.
'   * stage success messages are suppressed in the main sequence: progress is
'     shown on the panel (chips + status bar), then ONE final report is shown.
' =====================================================================

Private Const PANEL_SHEET As String = "ControlPanel"
Private Const UI_FONT As String = "Hacen Saudi Arabia XL"   ' change the font in ONE place
Private Const CARD_WIDTH As Double = 780
Private Const CARD_HEIGHT As Double = 372
Private Const CARD_MARGIN As Double = 40     ' inner left/right padding of the card
Private Const BTN_GAP As Double = 24         ' space between the buttons
Private Const BTN_FONT As Single = 17
Private Const BTN_HEIGHT As Double = 65
Private Const MAIN_BTN_HEIGHT As Double = 66
Private Const MAIN_BTN_FONT As Single = 21

Private m_Busy As Boolean

' status reported by the stage modules
Private Const ST_NONE As Long = 0
Private Const ST_OK As Long = 1
Private Const ST_CANCEL As Long = 2
Private Const ST_FAIL As Long = 3

' visual state of a stage chip
Private Const CHIP_IDLE As Long = 0
Private Const CHIP_WAIT As Long = 1
Private Const CHIP_RUN As Long = 2
Private Const CHIP_OK As Long = 3
Private Const CHIP_FAIL As Long = 4
Private Const CHIP_CANCEL As Long = 5
Private Const CHIP_SKIP As Long = 6

' result of each stage of the main sequence (1 = extract, 2 = repeat status, 3 = analyze)
Private m_StSt(1 To 3) As Long
Private m_StDet(1 To 3) As String

' ---------------------------------------------------------------------
' 1) Build the UI
' ---------------------------------------------------------------------
Public Sub BuildBacPanel()
    Dim ws As Worksheet, shp As Shape
    Dim panelW As Double, headerH As Double
    Dim cardW As Double, cardLeft As Double, cardTop As Double
    Dim btnW As Double, btnTop As Double, x1 As Double, x2 As Double, x3 As Double
    Dim chipTop As Double
    Dim mainTop As Double
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
    ' 3. Main button (own row, full width) + 3 independent step buttons
    ' ---------------------------------------------------------
    mainTop = cardTop + 74
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLeft + CARD_MARGIN, mainTop, _
                                 cardW - 2 * CARD_MARGIN, MAIN_BTN_HEIGHT)
    StyleMainButton shp, "btnBacRunAll", Ar("btn_all")
    
    AddLabel ws, "StepsCaption", Ar("lbl_or"), cardLeft, mainTop + MAIN_BTN_HEIGHT + 8, cardW, 26, 12, False, RGB(120, 130, 145)
    
    btnW = (cardW - 2 * CARD_MARGIN - 2 * BTN_GAP) / 3
    btnTop = mainTop + MAIN_BTN_HEIGHT + 42
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
    
    ' Stage status chips: one under each step button (same x positions)
    chipTop = btnTop + BTN_HEIGHT + 6
    AddChip ws, "chip1", x1, chipTop, btnW
    AddChip ws, "chip2", x2, chipTop, btnW
    AddChip ws, "chip3", x3, chipTop, btnW
    
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
    Dim btnAll As Shape
    Dim hasFile As Boolean, fileName As String
    
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set btnExt = ws.Shapes("btnBacExtract")
    Set btnAna = ws.Shapes("btnBacAnalyze")
    Set btnRep = ws.Shapes("btnBacRepeat")
    Set bar = ws.Shapes("InfoBar")
    Set btnAll = ws.Shapes("btnBacRunAll")
    On Error GoTo 0
    If btnExt Is Nothing Or btnAna Is Nothing Then Exit Sub
    
    ' All buttons are ALWAYS active - they can run in any order
    btnExt.OnAction = "'" & ThisWorkbook.Name & "'!RunBacStep1"
    btnAna.OnAction = "'" & ThisWorkbook.Name & "'!RunBacStep2"
    If Not btnRep Is Nothing Then btnRep.OnAction = "'" & ThisWorkbook.Name & "'!RunBacRepeat"
    If Not btnAll Is Nothing Then btnAll.OnAction = "'" & ThisWorkbook.Name & "'!RunBacAll"
    
    ' Status bar: show the last extracted file if there is one
    If bar Is Nothing Then Exit Sub
    bar.Fill.ForeColor.RGB = RGB(240, 243, 248)     ' neutral (a previous result may have coloured it)
    
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
    Dim errText As String
    If m_Busy Then Exit Sub
    m_Busy = True
    On Error GoTo Failed
    StageAuto "modBacEtab", "BacEtab", False
    SetChip 1, CHIP_RUN
    Application.Run RunRef("modBacEtab.Run_BacEtab")
    ChipFromModule 1, "modBacEtab", "BacEtab"
    RefreshBacPanelState
    m_Busy = False
    Exit Sub
Failed:
    errText = Err.Description
    m_Busy = False
    SetChip 1, CHIP_FAIL
    MsgBox Ar("run_err") & vbLf & errText, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("btn_extract")
End Sub

Public Sub RunBacStep2()
    Dim errText As String
    If m_Busy Then Exit Sub
    m_Busy = True
    On Error GoTo Failed
    StageAuto "modBacAnalysis", "BacAnalysis", False
    SetChip 3, CHIP_RUN
    Application.Run RunRef("modBacAnalysis.Run_BacAnalysis")
    ChipFromModule 3, "modBacAnalysis", "BacAnalysis"
    m_Busy = False
    Exit Sub
Failed:
    errText = Err.Description
    m_Busy = False
    SetChip 3, CHIP_FAIL
    MsgBox Ar("run_err") & vbLf & errText, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("btn_analyze")
End Sub

Public Sub RunBacRepeat()
    Dim repMod As String, errText As String
    If m_Busy Then Exit Sub
    m_Busy = True
    On Error GoTo Failed
    ' Run_RepeatStatus reports its own errors; anything reaching Failed means
    ' the macro could not be started (module modStudentRepetitionStatus not imported).
    repMod = ResolveRepeatModule()
    If Len(repMod) = 0 Then repMod = "modStudentRepetitionStatus"
    StageAuto repMod, "RepeatStatus", False
    SetChip 2, CHIP_RUN
    Application.Run RunRef(repMod & ".Run_RepeatStatus")
    ChipFromModule 2, repMod, "RepeatStatus"
    m_Busy = False
    Exit Sub
Failed:
    errText = Err.Description
    m_Busy = False
    SetChip 2, CHIP_FAIL
    MsgBox Ar("msg_missing") & vbLf & errText, _
           vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("btn_repeat")
End Sub

' ---------------------------------------------------------------------
' MAIN macro: the three stages in one click, in this order:
'   1) extract  ->  2) fetch repeat status  ->  3) analyze
'
' * a stage counts as successful ONLY if its module reports success itself;
' * the sequence stops at the FIRST stage that fails or is cancelled;
' * no stage success message interrupts the sequence: progress is shown on the
'   panel (chips + status bar) and ONE final report is shown at the end.
' ---------------------------------------------------------------------
Public Sub RunBacAll()
    Dim repMod As String, errText As String
    Dim stepNo As Long, i As Long, failedAt As Long, icon As Long
    Dim ans As VbMsgBoxResult

    If m_Busy Then Exit Sub

    ans = MsgBox(Ar("all_confirm_head") & vbLf & vbLf & _
                 "1) " & Ar("btn_extract") & vbLf & _
                 "2) " & Ar("btn_repeat") & vbLf & _
                 "3) " & Ar("btn_analyze") & vbLf & vbLf & Ar("all_confirm_q"), _
                 vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("btn_all"))
    If ans <> vbYes Then Exit Sub

    m_Busy = True
    For i = 1 To 3
        m_StSt(i) = ST_NONE
        m_StDet(i) = ""
        SetChip i, CHIP_WAIT
    Next i

    On Error GoTo Failed
    For stepNo = 1 To 3
        SetChip stepNo, CHIP_RUN
        SetBar Ar("bar_stage") & stepNo & " : " & StageName(stepNo), CHIP_RUN
        Repaint

        Select Case stepNo
            Case 1
                m_StSt(1) = RunStage("modBacEtab", "BacEtab", "Run_BacEtab", m_StDet(1))
            Case 2
                repMod = ResolveRepeatModule()
                If Len(repMod) = 0 Then
                    m_StSt(2) = ST_FAIL
                    m_StDet(2) = Ar("msg_missing")
                Else
                    m_StSt(2) = RunStage(repMod, "RepeatStatus", "Run_RepeatStatus", m_StDet(2))
                End If
            Case 3
                m_StSt(3) = RunStage("modBacAnalysis", "BacAnalysis", "Run_BacAnalysis", m_StDet(3))
        End Select

        SetChip stepNo, ChipFor(m_StSt(stepNo))
        If m_StSt(stepNo) <> ST_OK Then         ' STOP: later stages are not executed
            failedAt = stepNo
            Exit For
        End If
    Next stepNo
    GoTo Reported

Failed:
    ' unexpected error raised while starting / reading a stage
    errText = Err.Description
    If Len(errText) = 0 Then errText = "#" & CStr(Err.Number)
    failedAt = stepNo
    m_StSt(stepNo) = ST_FAIL
    m_StDet(stepNo) = errText
    Resume Reported

Reported:
    On Error Resume Next
    StageAuto "modBacEtab", "BacEtab", False
    StageAuto "modBacAnalysis", "BacAnalysis", False
    If Len(repMod) > 0 Then StageAuto repMod, "RepeatStatus", False
    Application.StatusBar = False
    Application.ScreenUpdating = True
    On Error GoTo 0
    m_Busy = False

    If failedAt = 0 Then
        SetBar Ar("bar_ok"), CHIP_OK
        icon = vbInformation
    Else
        For i = failedAt + 1 To 3
            SetChip i, CHIP_SKIP
        Next i
        ShowPanel                               ' bring the panel back so the result is visible
        If m_StSt(failedAt) = ST_CANCEL Then
            SetBar Ar("bar_cancel") & failedAt & " : " & StageName(failedAt), CHIP_CANCEL
            icon = vbExclamation
        Else
            SetBar Ar("bar_fail") & failedAt & " : " & StageName(failedAt), CHIP_FAIL
            icon = vbCritical
        End If
    End If

    MsgBox BuildFinalReport(failedAt), icon + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("btn_all")
End Sub

' Runs one stage in automatic mode and returns the status the stage reported itself.
Private Function RunStage(ByVal modName As String, ByVal prefix As String, _
                          ByVal runName As String, ByRef detail As String) As Long
    Dim st As Long, errText As String
    On Error GoTo Bad
    Application.Run RunRef(modName & "." & prefix & "_SetAuto"), True
    Application.Run RunRef(modName & "." & runName)
    st = CLng(Application.Run(RunRef(modName & "." & prefix & "_LastStatus")))
    detail = CStr(Application.Run(RunRef(modName & "." & prefix & "_LastDetail")))
    If st <> ST_OK And st <> ST_CANCEL And st <> ST_FAIL Then
        st = ST_FAIL                            ' no report = never a success
        detail = Ar("no_report")
    End If
    RunStage = st
    Exit Function
Bad:
    errText = Err.Description
    If Len(errText) = 0 Then errText = "#" & CStr(Err.Number)
    RunStage = ST_FAIL
    detail = errText
End Function

' ---------------------------------------------------------------------
' Final report (kept under the MsgBox size limit)
' ---------------------------------------------------------------------
Private Function BuildFinalReport(ByVal failedAt As Long) As String
    Dim s As String, i As Long, lbl As String

    If failedAt = 0 Then
        s = Ar("rep_ok")
        For i = 1 To 3
            s = s & vbLf & vbLf & ChrW(&H2714) & " " & i & ") " & StageName(i)
            s = s & vbLf & ShortDetail(m_StDet(i), 2, 170)
        Next i
        BuildFinalReport = s
        Exit Function
    End If

    If m_StSt(failedAt) = ST_CANCEL Then s = Ar("rep_cancel") Else s = Ar("rep_fail")

    ' stages that succeeded
    s = s & vbLf & vbLf & Ar("rep_done")
    If failedAt = 1 Then
        s = s & " " & Ar("rep_none")
    Else
        For i = 1 To failedAt - 1
            s = s & vbLf & ChrW(&H2714) & " " & i & ") " & StageName(i)
            s = s & vbLf & "    " & ShortDetail(m_StDet(i), 1, 120)
        Next i
    End If

    ' the stage that failed (or was cancelled) + the reason
    If m_StSt(failedAt) = ST_CANCEL Then lbl = Ar("rep_canceled") Else lbl = Ar("rep_failed")
    s = s & vbLf & vbLf & lbl & vbLf & ChrW(&H2718) & " " & failedAt & ") " & StageName(failedAt)
    If m_StSt(failedAt) = ST_FAIL Then
        If Len(Trim$(m_StDet(failedAt))) = 0 Then
            s = s & vbLf & Ar("rep_reason") & " " & Ar("rep_noreason")
        Else
            s = s & vbLf & Ar("rep_reason") & vbLf & ShortDetail(m_StDet(failedAt), 6, 380)
        End If
    End If

    ' confirmation: the following stages were NOT executed
    s = s & vbLf & vbLf & Ar("rep_notrun")
    If failedAt = 3 Then
        s = s & " " & Ar("rep_nonext")
    Else
        For i = failedAt + 1 To 3
            s = s & vbLf & "- " & i & ") " & StageName(i)
        Next i
    End If

    If Len(s) > 1000 Then s = Left$(s, 1000)
    BuildFinalReport = s
End Function

' First lines of a text, each trimmed, the whole clipped
Private Function ShortDetail(ByVal txt As String, ByVal maxLines As Long, ByVal maxChars As Long) As String
    Dim parts() As String, i As Long, n As Long, s As String
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    parts = Split(Trim$(txt), vbLf)
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            n = n + 1
            If n > maxLines Then Exit For
            If Len(s) > 0 Then s = s & vbLf
            s = s & Trim$(parts(i))
        End If
    Next i
    If Len(s) > maxChars Then s = Left$(s, maxChars) & ChrW(&H2026)
    ShortDetail = s
End Function

Private Function StageName(ByVal stage As Long) As String
    Select Case stage
        Case 1: StageName = Ar("btn_extract")
        Case 2: StageName = Ar("btn_repeat")
        Case 3: StageName = Ar("btn_analyze")
    End Select
End Function

' ---------------------------------------------------------------------
' Panel feedback (chips + status bar): informative only, never asks anything
' ---------------------------------------------------------------------
Private Function ChipFor(ByVal st As Long) As Long
    Select Case st
        Case ST_OK: ChipFor = CHIP_OK
        Case ST_CANCEL: ChipFor = CHIP_CANCEL
        Case ST_FAIL: ChipFor = CHIP_FAIL
        Case Else: ChipFor = CHIP_IDLE
    End Select
End Function

Private Sub SetChip(ByVal stage As Long, ByVal kind As Long)
    Dim ws As Worksheet, shp As Shape, txt As String, fillC As Long, txtC As Long
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Set shp = ws.Shapes("chip" & stage)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub             ' panel built by an older version: rebuild it once

    Select Case kind
        Case CHIP_WAIT:   txt = Ar("c_wait"):   fillC = RGB(240, 243, 248): txtC = RGB(120, 130, 145)
        Case CHIP_RUN:    txt = Ar("c_run"):    fillC = RGB(255, 243, 205): txtC = RGB(150, 100, 0)
        Case CHIP_OK:     txt = ChrW(&H2714) & " " & Ar("c_ok"):   fillC = RGB(220, 242, 228): txtC = RGB(16, 124, 65)
        Case CHIP_FAIL:   txt = ChrW(&H2718) & " " & Ar("c_fail"): fillC = RGB(253, 226, 226): txtC = RGB(176, 32, 32)
        Case CHIP_CANCEL: txt = Ar("c_cancel"): fillC = RGB(255, 243, 205): txtC = RGB(150, 100, 0)
        Case CHIP_SKIP:   txt = Ar("c_skip"):   fillC = RGB(230, 233, 238): txtC = RGB(120, 130, 145)
        Case Else:        txt = Ar("c_idle"):   fillC = RGB(240, 243, 248): txtC = RGB(120, 130, 145)
    End Select
    shp.Fill.ForeColor.RGB = fillC
    ApplyText shp, txt, 12, True, txtC
End Sub

' Status after a manual (single-step) run: read what the module itself reported
Private Sub ChipFromModule(ByVal stage As Long, ByVal modName As String, ByVal prefix As String)
    Dim st As Long
    On Error GoTo Bad
    st = CLng(Application.Run(RunRef(modName & "." & prefix & "_LastStatus")))
    SetChip stage, ChipFor(st)
    Exit Sub
Bad:
    SetChip stage, CHIP_IDLE
End Sub

Private Sub SetBar(ByVal txt As String, ByVal kind As Long)
    Dim ws As Worksheet, bar As Shape, fillC As Long, txtC As Long
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Set bar = ws.Shapes("InfoBar")
    On Error GoTo 0
    If bar Is Nothing Then Exit Sub
    Select Case kind
        Case CHIP_OK:     fillC = RGB(220, 242, 228): txtC = RGB(16, 124, 65)
        Case CHIP_FAIL:   fillC = RGB(253, 226, 226): txtC = RGB(176, 32, 32)
        Case CHIP_CANCEL: fillC = RGB(255, 243, 205): txtC = RGB(150, 100, 0)
        Case Else:        fillC = RGB(240, 243, 248): txtC = RGB(18, 33, 58)
    End Select
    bar.Fill.ForeColor.RGB = fillC
    ApplyText bar, txt, 14, True, txtC
End Sub

Private Sub Repaint()
    Application.ScreenUpdating = True
    DoEvents
End Sub

Private Sub ShowPanel()
    Dim ws As Worksheet
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Activate
    ws.Activate
End Sub

Private Sub StageAuto(ByVal modName As String, ByVal prefix As String, ByVal v As Boolean)
    On Error Resume Next
    Application.Run RunRef(modName & "." & prefix & "_SetAuto"), v
End Sub

Private Sub AddChip(ByVal ws As Worksheet, ByVal shapeName As String, _
        ByVal x As Double, ByVal y As Double, ByVal w As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, 24)
    With shp
        .Name = shapeName
        .Placement = xlFreeFloating
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(240, 243, 248)
        .Line.Visible = msoFalse
        On Error Resume Next
        .Adjustments(1) = 0.3
        On Error GoTo 0
    End With
    ApplyText shp, Ar("c_idle"), 12, True, RGB(120, 130, 145)
End Sub

' ---------------------------------------------------------------------
' 4) Helpers
' ---------------------------------------------------------------------
Private Function PanelSheet() As Worksheet
    On Error Resume Next
    Set PanelSheet = ThisWorkbook.Worksheets(PANEL_SHEET)
    On Error GoTo 0
End Function

' "'Workbook.xlsm'!ModuleName.Macro"
Private Function RunRef(ByVal procName As String) As String
    RunRef = "'" & ThisWorkbook.Name & "'!" & procName
End Function

' Finds the module that holds the repeat-status macro (current name first, old name second)
Private Function ResolveRepeatModule() As String
    Dim nm As Variant, ok As Variant
    For Each nm In Array("modStudentRepetitionStatus", "modRepeatStatus")
        ok = Empty
        On Error Resume Next
        ok = Application.Run(RunRef(CStr(nm) & ".RepeatStatus_Probe"))
        If Err.Number = 0 Then
            If Not IsEmpty(ok) Then
                If CBool(ok) Then
                    ' an old module without the status API is NOT usable: success could not be verified
                    Application.Run RunRef(CStr(nm) & ".RepeatStatus_LastStatus")
                    If Err.Number = 0 Then ResolveRepeatModule = CStr(nm)
                End If
            End If
        End If
        Err.Clear
        On Error GoTo 0
        If Len(ResolveRepeatModule) > 0 Then Exit Function
    Next nm
End Function

' Main button: navy fill + gold outline (matches the header), larger text
Private Sub StyleMainButton(ByVal btn As Shape, ByVal shapeName As String, ByVal captionText As String)
    StyleButtonShape btn, shapeName, captionText, RGB(18, 33, 58)
    With btn
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(212, 175, 55)
        .Line.Weight = 2.5
        .Shadow.Transparency = 0.5
    End With
    ApplyText btn, captionText, MAIN_BTN_FONT, True, RGB(255, 255, 255)
End Sub

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
    Case "msg_missing": Ar = ArHex("62A 639 630 631 20 62A 634 63A 64A 644 20 627 644 645 627 643 631 648 2E 20 62A 623 643 62F 20 645 646 20 627 633 62A 64A")
        Ar = Ar & ArHex("631 627 62F 20 627 644 648 62D 62F 629 20 6D 6F 64 53 74 75 64 65 6E 74 52 65 70 65 74 69 74 69 6F 6E 53")
        Ar = Ar & ArHex("74 61 74 75 73 20 641 64A 20 627 644 645 634 631 648 639 2E")
    Case "info_hint": Ar = ArHex("627 62E 62A 631 20 627 644 639 645 644 64A 629 20 627 644 645 637 644 648 628 629 3A 20 64A 645 643 646 643 20 62A 634 63A 64A 644 20 623 64A 20 632 631 20 641 64A 20 623 64A 20 648 642 62A 2E")
    Case "btn_all": Ar = ArHex("645 639 627 644 62C 629 20 646 62A 627 626 62C 20 627 644 628 643 627 644 648 631 64A 627")
    Case "lbl_or": Ar = ArHex("623 648 20 634 63A 651 644 20 643 644 20 62E 637 648 629 20 639 644 649 20 62D 62F 629")
    Case "all_confirm_head": Ar = ArHex("633 64A 62A 645 20 62A 634 63A 64A 644 20 627 644 62E 637 648 627 62A 20 627 644 62A 627 644 64A 629 20 628 627 644 62A 631")
        Ar = Ar & ArHex("62A 64A 628 3A")
    Case "all_confirm_q": Ar = ArHex("647 644 20 62A 631 64A 62F 20 627 644 645 62A 627 628 639 629 61F")
    Case "all_err": Ar = ArHex("62A 648 642 641 62A 20 627 644 645 639 627 644 62C 629 20 628 633 628 628 20 62E 637 623 20 641 64A 20 627 644 62E 637 648")
        Ar = Ar & ArHex("629 20 631 642 645 20")
    Case "c_idle": Ar = ArHex("62C 627 647 632")
    Case "c_wait": Ar = ArHex("641 64A 20 627 644 627 646 62A 638 627 631")
    Case "c_run": Ar = ArHex("642 64A 62F 20 627 644 62A 646 641 64A 630 2026")
    Case "c_ok": Ar = ArHex("62A 645 62A 20 628 646 62C 627 62D")
    Case "c_fail": Ar = ArHex("641 634 644 62A")
    Case "c_cancel": Ar = ArHex("623 64F 644 63A 64A 62A")
    Case "c_skip": Ar = ArHex("644 645 20 62A 64F 646 641 64E 651 630")
    Case "bar_stage": Ar = ArHex("62C 627 631 64D 20 62A 646 641 64A 630 20 627 644 645 631 62D 644 629 20")
    Case "bar_ok": Ar = ArHex("627 643 62A 645 644 62A 20 62C 645 64A 639 20 627 644 645 631 627 62D 644 20 628 646 62C 627 62D 2E")
    Case "bar_fail": Ar = ArHex("62A 648 642 641 62A 20 627 644 645 639 627 644 62C 629 20 639 646 62F 20 627 644 645 631 62D 644 629 20")
    Case "bar_cancel": Ar = ArHex("623 64F 648 642 641 62A 20 627 644 645 639 627 644 62C 629 20 639 646 62F 20 627 644 645 631 62D 644 629 20")
    Case "rep_ok": Ar = ArHex("627 643 62A 645 644 62A 20 62C 645 64A 639 20 627 644 645 631 627 62D 644 20 628 646 62C 627 62D 2E")
    Case "rep_fail": Ar = ArHex("62A 648 642 641 62A 20 627 644 645 639 627 644 62C 629 20 641 648 631 20 641 634 644 20 625 62D 62F 649 20 627 644 645 631 627 62D 644 2E")
    Case "rep_cancel": Ar = ArHex("623 64F 648 642 641 62A 20 627 644 645 639 627 644 62C 629 20 628 637 644 628 20 645 646 20 627 644 645 633 62A 62E 62F 645 2E")
    Case "rep_done": Ar = ArHex("627 644 645 631 627 62D 644 20 627 644 62A 64A 20 646 62C 62D 62A 3A")
    Case "rep_none": Ar = ArHex("644 627 20 634 64A 621 2E")
    Case "rep_failed": Ar = ArHex("627 644 645 631 62D 644 629 20 627 644 62A 64A 20 641 634 644 62A 3A")
    Case "rep_canceled": Ar = ArHex("627 644 645 631 62D 644 629 20 627 644 645 644 63A 627 629 3A")
    Case "rep_reason": Ar = ArHex("627 644 633 628 628 3A")
    Case "rep_noreason": Ar = ArHex("644 645 20 62A 64F 62D 62F 650 651 62F 20 627 644 645 631 62D 644 629 20 633 628 628 64B 627 20 644 644 641 634 644 2E")
    Case "rep_notrun": Ar = ArHex("644 645 20 62A 64F 646 641 64E 651 630 20 627 644 645 631 627 62D 644 20 627 644 644 627 62D 642 629 3A")
    Case "rep_nonext": Ar = ArHex("644 627 20 62A 648 62C 62F 20 645 631 627 62D 644 20 644 627 62D 642 629 2E")
    Case "no_report": Ar = ArHex("627 646 62A 647 62A 20 627 644 645 631 62D 644 629 20 62F 648 646 20 623 646 20 62A 64F 628 644 63A 20 639 646 20 646 62A 64A 62C 62A 647")
        Ar = Ar & ArHex("627 60C 20 644 630 644 643 20 644 627 20 62A 64F 639 62F 651 20 646 627 62C 62D 629 2E")
    Case "run_err": Ar = ArHex("62A 639 630 651 631 20 62A 634 63A 64A 644 20 627 644 645 627 643 631 648 3A")
    Case "status_last": Ar = ArHex("622 62E 631 20 645 644 641 20 645 633 62A 62E 631 62C 3A 20")
    End Select
End Function
