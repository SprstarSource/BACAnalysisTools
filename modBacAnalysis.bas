Attribute VB_Name = "modBacAnalysis"
Option Explicit

' =====================================================================
'  BAC analysis builder  -  macro: Run_BacAnalysis
'  Runs AFTER Run_BacEtab (path comes from modBacEtab.gLastBacEtabPath,
'  or a file dialog if it is empty).
'   1) saves a copy "<prefix> <original name>.xlsx" next to the original
'   2) adds the Gen sheet (first sheet)
'   3) under each series sheet: counters strip, then 3 tables
'      (subject averages / statistical indicators / score bands)
'  Only rows with Moyenne > 0 are analysed (0.00 = absent pupil).
'  Sex is read as 1 / 2.  Repeater = "yes" word in column E.
' =====================================================================

Private Const HEADER_ROW As Long = 2
Private Const FIRST_SUBJ_COL As Long = 6
Private Const GAP_BELOW_DATA As Long = 2

Private Const HDR_ROW_HEIGHT As Double = 100
Private Const ROW_HEIGHT As Double = 24
Private Const TITLE_HEIGHT As Double = 36
Private Const GEN_HEAD_HEIGHT As Double = 40

Private Const FMT_COUNT As String = "00"
Private Const FMT_DEC As String = "00.00"
Private Const FMT_PCT As String = "[=1]%0;%00.00"

Private Const GT0 As String = """>0"""
Private Const GE0 As String = """>=0"""
Private Const GE10 As String = """>=10"""
Private Const EMPTYQ As String = """"""

Private Type SeriesInfo
    SheetName As String
    BranchName As String
    FirstRow As Long
    LastRow As Long
    avgCol As Long
    StripRow As Long
    RowPassF As Long
    RowPassM As Long
    RowPass As Long
    RowRepPass As Long
End Type

Private mFontBase As String
Private mFontName As String
Private mFontNum As String
Private mYes As String

' ---- run status (read by the control panel / main button) ----
Private Const ST_NONE As Long = 0       ' no report (the macro did not finish properly)
Private Const ST_OK As Long = 1         ' finished successfully
Private Const ST_CANCEL As Long = 2     ' cancelled by the user
Private Const ST_FAIL As Long = 3       ' failed (reason in mDetail)
Private mAuto As Boolean                ' automatic mode: no success message, no file dialog
Private mStatus As Long
Private mDetail As String

' ---------------------------------------------------------------------
Public Sub Run_BacAnalysis()
    Dim wb As Workbook, ws As Worksheet, info() As SeriesInfo
    Dim srcPath As String, dstPath As String, names As Variant, nm As Variant
    Dim n As Long, k As Long, yearVal As Long
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean, oldStatus As Variant
    Dim stateSaved As Boolean, Failed As Boolean, cancelled As Boolean, savedOK As Boolean
    Dim errText As String, summary As String

    mStatus = ST_NONE
    mDetail = ""
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldStatus = Application.StatusBar
    stateSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    mYes = Ar("yes")
    srcPath = modBacEtab.gLastBacEtabPath
    If Len(srcPath) > 0 Then
        If Not FileExists(srcPath) Then srcPath = ""
    End If
    If Len(srcPath) = 0 Then
        ' automatic mode never asks: no extraction file means the previous stage did not deliver
        If mAuto Then Err.Raise vbObjectError + 912, "modBacAnalysis", Ar("no_input")
        srcPath = PickFile()
    End If
    If Len(srcPath) = 0 Then
        cancelled = True
        GoTo Finished
    End If

    Application.StatusBar = Ar("working")
    dstPath = NewAnalysisPath(srcPath)
    MakeCopy srcPath, dstPath
    Set wb = Workbooks.Open(fileName:=dstPath, UpdateLinks:=0)

    names = Array("Philo", "LaEtr", "Math", "SE", "GeEc", "TGM", "TGE", "TGC", "TGP")
    For Each nm In names
        Set ws = FindSheet(wb, CStr(nm))
        If Not ws Is Nothing Then
            If ws.ListObjects.Count > 0 Then
                n = n + 1
                If n = 1 Then
                    ReDim info(1 To 1)
                Else
                    ReDim Preserve info(1 To n)
                End If
                info(n).SheetName = ws.Name
                info(n).BranchName = BranchOf(CStr(nm))
            End If
        End If
    Next nm
    If n = 0 Then Err.Raise vbObjectError + 911, "modBacAnalysis", Ar("no_series")

    Set ws = wb.Worksheets(info(1).SheetName)
    mFontBase = ws.Cells(HEADER_ROW, 1).Font.Name
    mFontName = ws.Cells(HEADER_ROW + 1, 2).Font.Name
    mFontNum = ws.Cells(HEADER_ROW + 1, FIRST_SUBJ_COL).Font.Name
    yearVal = YearFromSheet(ws)

    For k = 1 To n
        Application.StatusBar = Ar("working") & " " & info(k).SheetName
        Set ws = wb.Worksheets(info(k).SheetName)
        BuildSeries ws, info(k)
    Next k

    Application.StatusBar = Ar("working") & " Gen"
    BuildGen wb, info, n, yearVal

    ' Calculate all generated formulas, then freeze results as values.
    FreezeAnalysisFormulas wb, info, n

    wb.Worksheets("Gen").Activate
    wb.Worksheets("Gen").Range("A1").Select
    wb.Save
    If Not FileExists(dstPath) Then Err.Raise vbObjectError + 913, "modBacAnalysis", Ar("save_chk") & vbLf & dstPath
    savedOK = True

Finished:
    On Error Resume Next
    If Failed And Not savedOK And Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Kill dstPath
    End If
    If stateSaved Then
        Application.StatusBar = oldStatus
        Application.DisplayAlerts = oldAlerts
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    On Error GoTo 0
    If cancelled Then
        mStatus = ST_CANCEL
        mDetail = Ar("st_cancel")
        Exit Sub
    End If
    If Failed Then
        mStatus = ST_FAIL
        mDetail = errText
        If Not mAuto Then
            MsgBox Ar("failed") & vbLf & errText, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("app_title")
        End If
    Else
        mStatus = ST_OK
        mDetail = dstPath & vbLf & Ar("sheets") & ": " & (n + 1)
        If Not mAuto Then       ' automatic mode: the main button shows ONE final report
            summary = Ar("done") & vbLf & dstPath & vbLf & Ar("sheets") & ": " & (n + 1)
            MsgBox summary, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("app_title")
        End If
    End If
    Exit Sub
Failed:
    Failed = True
    errText = Err.Description
    If Len(errText) = 0 Then errText = "#" & CStr(Err.Number)
    Resume Finished
End Sub

' ---------------------------------------------------------------------
'  Status API (used by modBuildBacPanel)
' ---------------------------------------------------------------------
Public Sub BacAnalysis_SetAuto(ByVal v As Boolean)
    mAuto = v
End Sub

Public Function BacAnalysis_LastStatus() As Long
    BacAnalysis_LastStatus = mStatus
End Function

Public Function BacAnalysis_LastDetail() As String
    BacAnalysis_LastDetail = mDetail
End Function

' ---------------------------------------------------------------------
'  Files
' ---------------------------------------------------------------------
Private Function PickFile() As String
    On Error Resume Next
    With Application.FileDialog(3)
        .Title = Ar("pick_title")
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx"
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    FileExists = (Len(Dir$(fullPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0)
End Function

Private Function PathInUse(ByVal fullPath As String) As Boolean
    Dim book As Workbook
    If FileExists(fullPath) Then
        PathInUse = True
        Exit Function
    End If
    For Each book In Application.Workbooks
        If StrComp(book.FullName, fullPath, vbTextCompare) = 0 Then
            PathInUse = True
            Exit Function
        End If
    Next book
End Function

Private Function NewAnalysisPath(ByVal srcPath As String) As String
    Dim p As Long, folder As String, base As String, stem As String
    Dim fullPath As String, n As Long
    p = InStrRev(srcPath, Application.PathSeparator)
    folder = Left$(srcPath, p)
    base = Mid$(srcPath, p + 1)
    p = InStrRev(base, ".")
    If p > 1 Then base = Left$(base, p - 1)
    stem = Ar("prefix") & " " & base
    fullPath = folder & stem & ".xlsx"
    n = 1
    Do While PathInUse(fullPath)
        n = n + 1
        fullPath = folder & stem & "_" & Format$(n, "00") & ".xlsx"
    Loop
    NewAnalysisPath = fullPath
End Function

Private Sub MakeCopy(ByVal src As String, ByVal dst As String)
    Dim book As Workbook
    For Each book In Application.Workbooks
        If StrComp(book.FullName, src, vbTextCompare) = 0 Then
            book.SaveCopyAs dst
            Exit Sub
        End If
    Next book
    FileCopy src, dst
End Sub

Private Function FindSheet(ByVal wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function YearFromSheet(ByVal ws As Worksheet) As Long
    Dim t As String
    t = Trim$(CStr(ws.Range("B1").Value2))
    If Len(t) >= 4 Then
        If Right$(t, 4) Like "####" Then YearFromSheet = CLng(Right$(t, 4))
    End If
    If YearFromSheet = 0 Then YearFromSheet = Year(Date)
End Function

Private Function BranchOf(ByVal nm As String) As String
    Select Case LCase$(nm)
        Case "philo": BranchOf = Ar("b30")
        Case "laetr": BranchOf = Ar("b31")
        Case "geec":  BranchOf = Ar("b32")
        Case "se":    BranchOf = Ar("b33")
        Case "math":  BranchOf = Ar("b34")
        Case "tgm":   BranchOf = Ar("b35")
        Case "tge":   BranchOf = Ar("b36")
        Case "tgc":   BranchOf = Ar("b37")
        Case "tgp":   BranchOf = Ar("b38")
    End Select
End Function

' ---------------------------------------------------------------------
'  Small helpers
' ---------------------------------------------------------------------
Private Function ColL(ByVal c As Long) As String
    Dim s As String
    Do While c > 0
        s = Chr$(65 + ((c - 1) Mod 26)) & s
        c = (c - 1) \ 26
    Loop
    ColL = s
End Function

Private Function Rg(ByVal colLetter As String, ByVal r1 As Long, ByVal r2 As Long) As String
    Rg = "$" & colLetter & "$" & r1 & ":$" & colLetter & "$" & r2
End Function

Private Function IsAmazigh(ByVal v As Variant) As Boolean
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    IsAmazigh = (InStr(1, CStr(v), ArHex("645 627 632 64A 63A")) > 0)
End Function

Private Function ArrF(ByVal fn As String, ByVal cond As String, ByVal xr As String) As String
    ArrF = "=IFERROR(" & fn & "(IF(" & cond & "," & xr & "))," & EMPTYQ & ")"
End Function

Private Function SumStrip(ByRef info() As SeriesInfo, ByVal n As Long, ByVal colLetter As String) As String
    Dim i As Long, f As String
    For i = 1 To n
        If i > 1 Then f = f & "+"
        f = f & "'" & info(i).SheetName & "'!$" & colLetter & "$" & info(i).StripRow
    Next i
    SumStrip = "=" & f
End Function

' ---------------------------------------------------------------------
'  Styles
' ---------------------------------------------------------------------
Private Sub StyleGrid(ByVal rng As Range)
    With rng
        .Font.Name = mFontNum: .Font.Size = 15: .Font.Bold = False
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = False
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        .Rows.RowHeight = ROW_HEIGHT
    End With
End Sub

Private Sub StyleHead(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = mFontBase: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = True
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Private Sub StyleLabel(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(221, 235, 247)
        .Font.Name = mFontName: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = False
        .ShrinkToFit = True
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Private Sub MergeText(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal txt As String)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        If c2 > c1 Then .Merge
        .Value2 = txt
    End With
End Sub

' header row: A:E merged label + subject headers linked to row 2
Private Sub PutHeaderRow(ByVal ws As Worksheet, ByVal r As Long, ByVal avgCol As Long, ByVal lbl As String)
    Dim c As Long
    For c = FIRST_SUBJ_COL To avgCol
        ws.Cells(r, c).Formula = "=" & ColL(c) & "$" & HEADER_ROW
    Next c
    StyleHead ws.Range(ws.Cells(r, 1), ws.Cells(r, avgCol))
    ws.Range(ws.Cells(r, FIRST_SUBJ_COL), ws.Cells(r, avgCol)).Orientation = 90
    MergeText ws, r, 1, 5, lbl
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Font.Size = 16
    ws.Rows(r).RowHeight = HDR_ROW_HEIGHT
End Sub

Private Sub PutLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    MergeText ws, r, 1, 5, txt
    StyleLabel ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
End Sub

' ---------------------------------------------------------------------
'  Counters strip (row = last data row + 3)
' ---------------------------------------------------------------------
Private Sub PutStripCell(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, _
                         ByVal lbl As String, ByVal cv As Long, ByVal f As String)
    MergeText ws, r, c1, c2, lbl
    StyleHead ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
    With ws.Cells(r, cv)
        .Formula = f
        .NumberFormat = FMT_COUNT
        .Font.Name = mFontNum: .Font.Size = 15: .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Private Sub BuildStrip(ByVal ws As Worksheet, ByVal r As Long, ByVal avgRng As String, _
                       ByVal sexRng As String, ByVal repRng As String)
    Dim yesCrit As String
    yesCrit = Chr$(34) & mYes & Chr$(34)
    PutStripCell ws, r, 2, 2, Ar("s_count"), 3, "=COUNTIF(" & avgRng & "," & GT0 & ")"
    PutStripCell ws, r, 4, 6, Ar("s_male"), 7, "=COUNTIFS(" & sexRng & ",1," & avgRng & "," & GT0 & ")"
    PutStripCell ws, r, 8, 10, Ar("s_female"), 11, "=COUNTIFS(" & sexRng & ",2," & avgRng & "," & GT0 & ")"
    PutStripCell ws, r, 12, 14, Ar("s_rep"), 15, "=COUNTIFS(" & repRng & "," & yesCrit & "," & avgRng & "," & GT0 & ")"
    ws.Rows(r).RowHeight = ROW_HEIGHT + 4
End Sub

' ---------------------------------------------------------------------
'  One series sheet
' ---------------------------------------------------------------------
Private Sub BuildSeries(ByVal ws As Worksheet, ByRef s As SeriesInfo)
    Dim lo As ListObject
    Dim nCols As Long, avgCol As Long, avgL As String
    Dim t1 As Long, t2 As Long, t3 As Long, hdr3 As Long, r As Long, c As Long, i As Long
    Dim avgRng As String, sexRng As String, repRng As String, xr As String, cond As String
    Dim yesCrit As String, f As String, parts As String
    Dim lows As Variant, highs As Variant, bands As Variant
    Dim lastUsed As Long

    Set lo = ws.ListObjects(1)
    s.FirstRow = lo.DataBodyRange.Row
    s.LastRow = s.FirstRow + lo.DataBodyRange.Rows.Count - 1
    nCols = lo.Range.Columns.Count
    avgCol = nCols - 1
    s.avgCol = avgCol
    avgL = ColL(avgCol)
    avgRng = Rg(avgL, s.FirstRow, s.LastRow)
    sexRng = Rg("D", s.FirstRow, s.LastRow)
    repRng = Rg("E", s.FirstRow, s.LastRow)
    yesCrit = Chr$(34) & mYes & Chr$(34)

    ' ---- counters strip: 2 empty rows under the data, strip on the 3rd
    s.StripRow = s.LastRow + GAP_BELOW_DATA + 1
    BuildStrip ws, s.StripRow, avgRng, sexRng, repRng

    ' ================= Table 1 =================
    t1 = s.StripRow + 2
    s.RowPassF = t1 + 2
    s.RowPassM = t1 + 3
    s.RowPass = t1 + 4
    s.RowRepPass = t1 + 8
    StyleGrid ws.Range(ws.Cells(t1 + 1, 1), ws.Cells(t1 + 8, avgCol))
    PutHeaderRow ws, t1, avgCol, ""
    PutLabel ws, t1 + 1, Ar("l_avg")
    PutLabel ws, t1 + 2, Ar("l_passf")
    PutLabel ws, t1 + 3, Ar("l_passm")
    PutLabel ws, t1 + 4, Ar("l_pass")
    PutLabel ws, t1 + 5, Ar("l_rate")
    PutLabel ws, t1 + 6, Ar("l_max")
    PutLabel ws, t1 + 7, Ar("l_min")
    PutLabel ws, t1 + 8, Ar("l_reppass")

    For c = FIRST_SUBJ_COL To avgCol
        If Not IsAmazigh(ws.Cells(HEADER_ROW, c).Value) Then
            xr = Rg(ColL(c), s.FirstRow, s.LastRow)
            cond = "(" & avgRng & ">0)*(LEN(" & xr & ")>0)"
            ws.Cells(t1 + 1, c).Formula = "=IFERROR(AVERAGEIFS(" & xr & "," & avgRng & "," & GT0 & ")," & EMPTYQ & ")"
            ws.Cells(t1 + 2, c).Formula = "=COUNTIFS(" & xr & "," & GE10 & "," & sexRng & ",2," & avgRng & "," & GT0 & ")"
            ws.Cells(t1 + 3, c).Formula = "=COUNTIFS(" & xr & "," & GE10 & "," & sexRng & ",1," & avgRng & "," & GT0 & ")"
            ws.Cells(t1 + 4, c).Formula = "=COUNTIFS(" & xr & "," & GE10 & "," & avgRng & "," & GT0 & ")"
            ws.Cells(t1 + 5, c).Formula = "=IFERROR(COUNTIFS(" & xr & "," & GE10 & "," & avgRng & "," & GT0 & ")/COUNTIFS(" & _
                                          xr & "," & GE0 & "," & avgRng & "," & GT0 & ")," & EMPTYQ & ")"
            ws.Cells(t1 + 6, c).FormulaArray = ArrF("MAX", cond, xr)
            ws.Cells(t1 + 7, c).FormulaArray = ArrF("MIN", cond, xr)
            ws.Cells(t1 + 8, c).Formula = "=COUNTIFS(" & xr & "," & GE10 & "," & repRng & "," & yesCrit & "," & avgRng & "," & GT0 & ")"
        End If
    Next c
    With ws
        .Range(.Cells(t1 + 1, FIRST_SUBJ_COL), .Cells(t1 + 1, avgCol)).NumberFormat = FMT_DEC
        .Range(.Cells(t1 + 2, FIRST_SUBJ_COL), .Cells(t1 + 4, avgCol)).NumberFormat = FMT_COUNT
        .Range(.Cells(t1 + 5, FIRST_SUBJ_COL), .Cells(t1 + 5, avgCol)).NumberFormat = FMT_PCT
        .Range(.Cells(t1 + 6, FIRST_SUBJ_COL), .Cells(t1 + 7, avgCol)).NumberFormat = FMT_DEC
        .Range(.Cells(t1 + 8, FIRST_SUBJ_COL), .Cells(t1 + 8, avgCol)).NumberFormat = FMT_COUNT
    End With

    ' ================= Table 2 : statistical indicators =================
    t2 = t1 + 8 + 2
    StyleGrid ws.Range(ws.Cells(t2 + 1, 1), ws.Cells(t2 + 8, avgCol))
    PutHeaderRow ws, t2, avgCol, Ar("t2_head")
    PutLabel ws, t2 + 1, Ar("i_rate")
    PutLabel ws, t2 + 2, Ar("i_mean")
    PutLabel ws, t2 + 3, Ar("i_median")
    PutLabel ws, t2 + 4, Ar("i_mode")
    PutLabel ws, t2 + 5, Ar("i_range")
    PutLabel ws, t2 + 6, Ar("i_sd")
    PutLabel ws, t2 + 7, Ar("i_var")
    PutLabel ws, t2 + 8, Ar("i_cv")

    For c = FIRST_SUBJ_COL To avgCol
        If Not IsAmazigh(ws.Cells(HEADER_ROW, c).Value) Then
            xr = Rg(ColL(c), s.FirstRow, s.LastRow)
            cond = "(" & avgRng & ">0)*(LEN(" & xr & ")>0)"
            ws.Cells(t2 + 1, c).Formula = "=IFERROR(COUNTIFS(" & xr & "," & GE10 & "," & avgRng & "," & GT0 & ")/COUNTIFS(" & _
                                          xr & "," & GE0 & "," & avgRng & "," & GT0 & ")," & EMPTYQ & ")"
            ws.Cells(t2 + 2, c).Formula = "=IFERROR(AVERAGEIFS(" & xr & "," & avgRng & "," & GT0 & ")," & EMPTYQ & ")"
            ws.Cells(t2 + 3, c).FormulaArray = ArrF("MEDIAN", cond, xr)
            ws.Cells(t2 + 4, c).FormulaArray = ArrF("MODE", cond, xr)
            ws.Cells(t2 + 5, c).FormulaArray = "=IFERROR(MAX(IF(" & cond & "," & xr & "))-MIN(IF(" & cond & "," & xr & "))," & EMPTYQ & ")"
            ws.Cells(t2 + 6, c).FormulaArray = ArrF("STDEVP", cond, xr)
            ws.Cells(t2 + 7, c).FormulaArray = ArrF("VARP", cond, xr)
            ws.Cells(t2 + 8, c).FormulaArray = "=IFERROR(STDEVP(IF(" & cond & "," & xr & "))/AVERAGEIFS(" & xr & "," & avgRng & "," & GT0 & ")," & EMPTYQ & ")"
        End If
    Next c
    With ws
        .Range(.Cells(t2 + 1, FIRST_SUBJ_COL), .Cells(t2 + 1, avgCol)).NumberFormat = FMT_PCT
        .Range(.Cells(t2 + 2, FIRST_SUBJ_COL), .Cells(t2 + 7, avgCol)).NumberFormat = FMT_DEC
        .Range(.Cells(t2 + 8, FIRST_SUBJ_COL), .Cells(t2 + 8, avgCol)).NumberFormat = FMT_PCT
    End With

    ' ================= Table 3 : score bands =================
    t3 = t2 + 8 + 2
    hdr3 = t3 + 1
    MergeText ws, t3, 1, avgCol, Ar("t3_title")
    With ws.Range(ws.Cells(t3, 1), ws.Cells(t3, avgCol))
        .Font.Name = mFontName: .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
    ws.Rows(t3).RowHeight = TITLE_HEIGHT
    StyleGrid ws.Range(ws.Cells(hdr3 + 1, 1), ws.Cells(hdr3 + 8, avgCol))
    PutHeaderRow ws, hdr3, avgCol, Ar("t3_head")
    bands = Array("07 >", "08.99 - 07", "09.99 - 09", "11.99 - 10", "13.99 - 12", "15.99 - 14", "17.99 - 16", "18 =<")
    lows = Array("", ">=7", ">=9", ">=10", ">=12", ">=14", ">=16", ">=18")
    highs = Array("<7", "<9", "<10", "<12", "<14", "<16", "<18", "")
    For i = 0 To 7
        ws.Cells(hdr3 + 1 + i, 1).NumberFormat = "@"
        PutLabel ws, hdr3 + 1 + i, CStr(bands(i))
    Next i
    For c = FIRST_SUBJ_COL To avgCol
        If Not IsAmazigh(ws.Cells(HEADER_ROW, c).Value) Then
            xr = Rg(ColL(c), s.FirstRow, s.LastRow)
            For i = 0 To 7
                parts = ""
                If Len(lows(i)) > 0 Then parts = parts & xr & "," & Chr$(34) & lows(i) & Chr$(34) & ","
                If Len(highs(i)) > 0 Then parts = parts & xr & "," & Chr$(34) & highs(i) & Chr$(34) & ","
                ws.Cells(hdr3 + 1 + i, c).Formula = "=COUNTIFS(" & parts & avgRng & "," & GT0 & ")"
            Next i
        End If
    Next c
    ws.Range(ws.Cells(hdr3 + 1, FIRST_SUBJ_COL), ws.Cells(hdr3 + 8, avgCol)).NumberFormat = FMT_COUNT

    lastUsed = hdr3 + 8
    On Error Resume Next
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastUsed, lo.Range.Columns.Count)).Address
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------
'  Gen sheet
' ---------------------------------------------------------------------
Private Sub SectionTitle(ByVal g As Worksheet, ByVal r As Long, ByVal lastCol As Long, ByVal txt As String)
    MergeText g, r, 1, lastCol, txt
    With g.Range(g.Cells(r, 1), g.Cells(r, lastCol))
        .Font.Name = mFontName: .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    g.Rows(r).RowHeight = TITLE_HEIGHT
End Sub

Private Sub PutHeads(ByVal g As Worksheet, ByVal r As Long, ByVal heads As Variant)
    Dim i As Long
    For i = 0 To UBound(heads)
        g.Cells(r, 1 + i).Value2 = heads(i)
    Next i
    StyleHead g.Range(g.Cells(r, 1), g.Cells(r, 1 + UBound(heads)))
    g.Rows(r).RowHeight = GEN_HEAD_HEIGHT
End Sub

Private Sub StyleBlock(ByVal g As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal lastCol As Long)
    StyleGrid g.Range(g.Cells(r1, 1), g.Cells(r2, lastCol))
    With g.Range(g.Cells(r1, 1), g.Cells(r2, 1))
        .Font.Name = mFontName: .Font.Size = 14: .Font.Bold = True
        .ShrinkToFit = True
    End With
End Sub

Private Sub StyleTotalRow(ByVal g As Worksheet, ByVal r As Long, ByVal lastCol As Long)
    With g.Range(g.Cells(r, 1), g.Cells(r, lastCol))
        .Interior.Color = RGB(255, 255, 0)
        .Font.Bold = True
    End With
End Sub

Private Sub TotalSums(ByVal g As Worksheet, ByVal r As Long, ByVal r1 As Long, ByVal r2 As Long, _
                      ByVal c1 As Long, ByVal c2 As Long)
    Dim c As Long
    g.Cells(r, 1).Value2 = Ar("total")
    For c = c1 To c2
        g.Cells(r, c).Formula = "=SUM(" & ColL(c) & r1 & ":" & ColL(c) & r2 & ")"
    Next c
End Sub

Private Sub BuildGen(ByVal wb As Workbook, ByRef info() As SeriesInfo, ByVal n As Long, ByVal yearVal As Long)
    Dim g As Worksheet, r As Long, i As Long, j As Long, y As Long
    Dim tit(1 To 7) As Long, hd(1 To 7) As Long, fr(1 To 7) As Long, lr(1 To 7) As Long, tt(1 To 7) As Long
    Dim sh As String, avgL As String, yesCrit As String
    Dim dRng As String, eRng As String, aRng As String

    yesCrit = Chr$(34) & mYes & Chr$(34)
    Set g = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    g.Name = "Gen"
    g.DisplayRightToLeft = True
    g.Cells.Font.Name = mFontBase
    g.Cells.Font.Size = 14
    g.Cells.VerticalAlignment = xlCenter
    g.Columns("A").ColumnWidth = 30
    g.Columns("B:J").ColumnWidth = 15

    SectionTitle g, 1, 6, Ar("gen_title")

    ' ---------- 1 : registered ----------
    tit(1) = 3: hd(1) = 4: fr(1) = 5: lr(1) = fr(1) + n - 1: tt(1) = lr(1) + 1
    SectionTitle g, tit(1), 4, Ar("g1")
    PutHeads g, hd(1), Array(Ar("h_branch"), Ar("h_reg"), Ar("h_pres"), Ar("h_male"), Ar("h_female"), Ar("h_abs"))
    StyleBlock g, fr(1), tt(1), 6
    For i = 1 To n
        r = fr(1) + i - 1
        sh = "'" & info(i).SheetName & "'!"
        dRng = sh & Rg("D", info(i).FirstRow, info(i).LastRow)
        g.Cells(r, 1).Value2 = info(i).BranchName
        g.Cells(r, 2).Formula = "=COUNTA(" & sh & Rg("B", info(i).FirstRow, info(i).LastRow) & ")"
        g.Cells(r, 3).Formula = "=" & sh & "$C$" & info(i).StripRow
        g.Cells(r, 4).Formula = "=COUNTIF(" & dRng & ",1)"
        g.Cells(r, 5).Formula = "=COUNTIF(" & dRng & ",2)"
        g.Cells(r, 6).Formula = "=B" & r & "-C" & r
    Next i
    TotalSums g, tt(1), fr(1), lr(1), 2, 6
    StyleTotalRow g, tt(1), 6
    g.Range(g.Cells(fr(1), 2), g.Cells(tt(1), 6)).NumberFormat = FMT_COUNT

    ' ---------- 2 : general results ----------
    tit(2) = tt(1) + 2: hd(2) = tit(2) + 1: fr(2) = tit(2) + 2: lr(2) = fr(2) + n - 1: tt(2) = lr(2) + 1
    SectionTitle g, tit(2), 4, Ar("g2")
    PutHeads g, hd(2), Array(Ar("h_branch"), Ar("h_pres"), Ar("h_passed"), Ar("h_male"), Ar("h_female"), Ar("h_rate"))
    StyleBlock g, fr(2), tt(2), 6
    For i = 1 To n
        r = fr(2) + i - 1
        sh = "'" & info(i).SheetName & "'!"
        avgL = ColL(info(i).avgCol)
        g.Cells(r, 1).Value2 = info(i).BranchName
        g.Cells(r, 2).Formula = "=C" & (fr(1) + i - 1)
        g.Cells(r, 3).Formula = "=" & sh & "$" & avgL & "$" & info(i).RowPass
        g.Cells(r, 4).Formula = "=" & sh & "$" & avgL & "$" & info(i).RowPassM
        g.Cells(r, 5).Formula = "=" & sh & "$" & avgL & "$" & info(i).RowPassF
        g.Cells(r, 6).Formula = "=IFERROR(C" & r & "/B" & r & "," & EMPTYQ & ")"
    Next i
    TotalSums g, tt(2), fr(2), lr(2), 2, 5
    g.Cells(tt(2), 6).Formula = "=IFERROR(C" & tt(2) & "/B" & tt(2) & "," & EMPTYQ & ")"
    StyleTotalRow g, tt(2), 6
    g.Range(g.Cells(fr(2), 2), g.Cells(tt(2), 5)).NumberFormat = FMT_COUNT
    g.Range(g.Cells(fr(2), 6), g.Cells(tt(2), 6)).NumberFormat = FMT_PCT

    ' ---------- 3 : by sex ----------
    tit(3) = tt(2) + 2: hd(3) = tit(3) + 1: fr(3) = tit(3) + 2: lr(3) = fr(3) + 1: tt(3) = lr(3) + 1
    SectionTitle g, tit(3), 4, Ar("g3")
    PutHeads g, hd(3), Array(Ar("h_type"), Ar("h_presm"), Ar("h_passm"), Ar("h_rate"))
    StyleBlock g, fr(3), tt(3), 4
    g.Cells(fr(3), 1).Value2 = Ar("male")
    g.Cells(fr(3), 2).Formula = SumStrip(info, n, "G")
    g.Cells(fr(3), 3).Formula = "=D" & tt(2)
    g.Cells(lr(3), 1).Value2 = Ar("female")
    g.Cells(lr(3), 2).Formula = SumStrip(info, n, "K")
    g.Cells(lr(3), 3).Formula = "=E" & tt(2)
    For r = fr(3) To lr(3)
        g.Cells(r, 4).Formula = "=IFERROR(C" & r & "/B" & r & "," & EMPTYQ & ")"
    Next r
    TotalSums g, tt(3), fr(3), lr(3), 2, 3
    g.Cells(tt(3), 4).Formula = "=IFERROR(C" & tt(3) & "/B" & tt(3) & "," & EMPTYQ & ")"
    StyleTotalRow g, tt(3), 4
    g.Cells(tt(3), 1).Value2 = Ar("total2")
    g.Range(g.Cells(fr(3), 2), g.Cells(tt(3), 3)).NumberFormat = FMT_COUNT
    g.Range(g.Cells(fr(3), 4), g.Cells(tt(3), 4)).NumberFormat = FMT_PCT

    ' ---------- 4 : repeaters ----------
    tit(4) = tt(3) + 2: hd(4) = tit(4) + 1: fr(4) = tit(4) + 2: lr(4) = fr(4) + n - 1: tt(4) = lr(4) + 1
    SectionTitle g, tit(4), 4, Ar("g4")
    PutHeads g, hd(4), Array(Ar("h_branch"), Ar("h_repcnt"), Ar("h_passof"), Ar("h_rate"))
    StyleBlock g, fr(4), tt(4), 4
    For i = 1 To n
        r = fr(4) + i - 1
        sh = "'" & info(i).SheetName & "'!"
        avgL = ColL(info(i).avgCol)
        g.Cells(r, 1).Value2 = info(i).BranchName
        g.Cells(r, 2).Formula = "=" & sh & "$O$" & info(i).StripRow
        g.Cells(r, 3).Formula = "=" & sh & "$" & avgL & "$" & info(i).RowRepPass
        g.Cells(r, 4).Formula = "=IFERROR(C" & r & "/B" & r & "," & EMPTYQ & ")"
    Next i
    TotalSums g, tt(4), fr(4), lr(4), 2, 3
    g.Cells(tt(4), 4).Formula = "=IFERROR(C" & tt(4) & "/B" & tt(4) & "," & EMPTYQ & ")"
    StyleTotalRow g, tt(4), 4
    g.Range(g.Cells(fr(4), 2), g.Cells(tt(4), 3)).NumberFormat = FMT_COUNT
    g.Range(g.Cells(fr(4), 4), g.Cells(tt(4), 4)).NumberFormat = FMT_PCT

    ' ---------- 5 : provincial / national (manual cells left empty) ----------
    tit(5) = tt(4) + 2: hd(5) = tit(5) + 1: fr(5) = tit(5) + 2: lr(5) = fr(5) + n - 1: tt(5) = lr(5) + 1
    SectionTitle g, tit(5), 5, Ar("g5")
    PutHeads g, hd(5), Array(Ar("h_branch"), Ar("h_school"), Ar("h_prov"), Ar("h_nat"), Ar("h_rank"))
    StyleBlock g, fr(5), tt(5), 5
    For i = 1 To n
        r = fr(5) + i - 1
        g.Cells(r, 1).Value2 = info(i).BranchName
        g.Cells(r, 2).Formula = "=F" & (fr(2) + i - 1)
    Next i
    g.Cells(tt(5), 1).Value2 = Ar("total")
    g.Cells(tt(5), 2).Formula = "=F" & tt(2)
    StyleTotalRow g, tt(5), 5
    g.Range(g.Cells(fr(5), 2), g.Cells(tt(5), 4)).NumberFormat = FMT_PCT
    g.Range(g.Cells(fr(5), 5), g.Cells(tt(5), 5)).NumberFormat = FMT_COUNT

    ' ---------- 6 : previous sessions (manual cells left empty) ----------
    tit(6) = tt(5) + 2: hd(6) = tit(6) + 1: fr(6) = tit(6) + 2: lr(6) = fr(6) + n - 1: tt(6) = lr(6) + 1
    SectionTitle g, tit(6), 6, Ar("g6")
    g.Cells(hd(6), 1).Value2 = Ar("h_branch")
    For j = 0 To 4
        y = yearVal - 4 + j
        g.Cells(hd(6), 2 + j).Value2 = Ar("session") & " " & CStr(y)
    Next j
    StyleHead g.Range(g.Cells(hd(6), 1), g.Cells(hd(6), 6))
    g.Rows(hd(6)).RowHeight = GEN_HEAD_HEIGHT
    StyleBlock g, fr(6), tt(6), 6
    For i = 1 To n
        r = fr(6) + i - 1
        g.Cells(r, 1).Value2 = info(i).BranchName
        g.Cells(r, 6).Formula = "=F" & (fr(2) + i - 1)
    Next i
    g.Cells(tt(6), 1).Value2 = Ar("total")
    g.Cells(tt(6), 6).Formula = "=F" & tt(2)
    StyleTotalRow g, tt(6), 6
    g.Range(g.Cells(fr(6), 2), g.Cells(tt(6), 6)).NumberFormat = FMT_PCT

    ' ---------- 7 : repeaters by sex ----------
    tit(7) = tt(6) + 2: hd(7) = tit(7) + 1: fr(7) = tit(7) + 3: lr(7) = fr(7) + n - 1: tt(7) = lr(7) + 1
    SectionTitle g, tit(7), 10, Ar("g7")
    MergeText g, hd(7), 1, 1, Ar("h_branch")
    MergeText g, hd(7), 2, 4, Ar("h_rep")
    MergeText g, hd(7), 5, 7, Ar("h_passof")
    MergeText g, hd(7), 8, 10, Ar("h_rate")
    For j = 0 To 2
        g.Cells(hd(7) + 1, 2 + 3 * j).Value2 = Ar("male")
        g.Cells(hd(7) + 1, 3 + 3 * j).Value2 = Ar("female")
        g.Cells(hd(7) + 1, 4 + 3 * j).Value2 = Ar("total2")
    Next j
    g.Range(g.Cells(hd(7), 1), g.Cells(hd(7), 1)).Resize(2, 1).Merge
    StyleHead g.Range(g.Cells(hd(7), 1), g.Cells(hd(7) + 1, 10))
    g.Rows(hd(7)).RowHeight = GEN_HEAD_HEIGHT
    g.Rows(hd(7) + 1).RowHeight = GEN_HEAD_HEIGHT
    StyleBlock g, fr(7), tt(7), 10
    For i = 1 To n
        r = fr(7) + i - 1
        sh = "'" & info(i).SheetName & "'!"
        dRng = sh & Rg("D", info(i).FirstRow, info(i).LastRow)
        eRng = sh & Rg("E", info(i).FirstRow, info(i).LastRow)
        aRng = sh & Rg(ColL(info(i).avgCol), info(i).FirstRow, info(i).LastRow)
        g.Cells(r, 1).Value2 = info(i).BranchName
        g.Cells(r, 2).Formula = "=COUNTIFS(" & dRng & ",1," & eRng & "," & yesCrit & "," & aRng & "," & GT0 & ")"
        g.Cells(r, 3).Formula = "=COUNTIFS(" & dRng & ",2," & eRng & "," & yesCrit & "," & aRng & "," & GT0 & ")"
        g.Cells(r, 4).Formula = "=SUM(B" & r & ":C" & r & ")"
        g.Cells(r, 5).Formula = "=COUNTIFS(" & dRng & ",1," & eRng & "," & yesCrit & "," & aRng & "," & GE10 & ")"
        g.Cells(r, 6).Formula = "=COUNTIFS(" & dRng & ",2," & eRng & "," & yesCrit & "," & aRng & "," & GE10 & ")"
        g.Cells(r, 7).Formula = "=SUM(E" & r & ":F" & r & ")"
    Next i
    TotalSums g, tt(7), fr(7), lr(7), 2, 7
    For i = fr(7) To tt(7)
        g.Cells(i, 8).Formula = "=IFERROR(E" & i & "/B" & i & "," & EMPTYQ & ")"
        g.Cells(i, 9).Formula = "=IFERROR(F" & i & "/C" & i & "," & EMPTYQ & ")"
        g.Cells(i, 10).Formula = "=IFERROR(G" & i & "/D" & i & "," & EMPTYQ & ")"
    Next i
    StyleTotalRow g, tt(7), 10
    g.Range(g.Cells(fr(7), 2), g.Cells(tt(7), 7)).NumberFormat = FMT_COUNT
    g.Range(g.Cells(fr(7), 8), g.Cells(tt(7), 10)).NumberFormat = FMT_PCT

    On Error Resume Next
    g.Activate
    ActiveWindow.DisplayGridlines = False
    With g.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------
'  Arabic text
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
' Freeze generated analysis formulas as fixed values.
' Original source data rows are deliberately left untouched.
' ---------------------------------------------------------------------
Private Sub FreezeAnalysisFormulas(ByVal wb As Workbook, ByRef info() As SeriesInfo, ByVal n As Long)
    Dim ws As Worksheet, rng As Range
    Dim i As Long, lastRow As Long, lastCol As Long

    ' Calculate every formula before replacing it with its displayed result.
    Application.CalculateFull

    ' Freeze only the generated analysis area below each source data table.
    For i = 1 To n
        Set ws = wb.Worksheets(info(i).SheetName)
        lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

        If lastRow > info(i).LastRow Then
            Set rng = ws.Range(ws.Cells(info(i).LastRow + 1, 1), ws.Cells(lastRow, lastCol))
            rng.Value2 = rng.Value2
        End If
    Next i

    ' Freeze all calculated results on the generated general summary sheet.
    Set ws = wb.Worksheets("Gen")
    Set rng = ws.UsedRange
    If Not rng Is Nothing Then rng.Value2 = rng.Value2
End Sub

Private Function ArHex(ByVal codes As String) As String
    Dim token As Variant
    For Each token In Split(codes, " ")
        If Len(CStr(token)) > 0 Then ArHex = ArHex & ChrW(CLng("&H" & CStr(token)))
    Next token
End Function

Private Function Ar(ByVal key As String) As String
    Select Case key
    Case "prefix": Ar = ArHex("62A 62D 644 64A 644 20 646 62A 627 626 62C")
    Case "app_title": Ar = ArHex("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 628 643 627 644 648 631 64A 627")
    Case "working": Ar = ArHex("62C 627 631 64D 20 627 644 62A 62D 636 64A 631 3A")
    Case "pick_title": Ar = ArHex("627 62E 62A 631 20 645 644 641 20 646 62A 627 626 62C 20 627 644 62B 627 646 648 64A 629")
    Case "no_series": Ar = ArHex("644 645 20 64A 62A 645 20 627 644 639 62B 648 631 20 639 644 649 20 623 648 631 627 642 20 627 644 634 639 628 20 641 64A 20 627 644 645 644 641")
    Case "save_chk": Ar = ArHex("62A 639 630 651 631 20 627 644 62A 62D 642 642 20 645 646 20 648 62C 648 62F 20 645 644 641 20 627 644 62A 62D 644 64A 644 20 628 639 62F")
        Ar = Ar & ArHex("20 627 644 62D 641 638 3A")
    Case "st_cancel": Ar = ArHex("623 644 63A 649 20 627 644 645 633 62A 62E 62F 645 20 627 644 639 645 644 64A 629 2E")
    Case "no_input": Ar = ArHex("644 627 20 64A 648 62C 62F 20 645 644 641 20 646 62A 627 626 62C 20 645 64F 633 62A 62E 631 64E 62C 20 644 62A 62D 644 64A 644 647 2E 20")
        Ar = Ar & ArHex("646 641 650 651 630 20 645 631 62D 644 629 20 627 633 62A 62E 631 627 62C 20 627 644 646 62A 627 626 62C 20 623 648 644 627 64B 2E")
    Case "failed": Ar = ArHex("62A 639 630 631 20 625 646 634 627 621 20 645 644 641 20 627 644 62A 62D 644 64A 644 3A")
    Case "done": Ar = ArHex("62A 645 20 625 646 634 627 621 20 645 644 641 20 627 644 62A 62D 644 64A 644 3A")
    Case "sheets": Ar = ArHex("639 62F 62F 20 627 644 623 648 631 627 642")
    Case "yes": Ar = ArHex("646 639 645")
    Case "b30": Ar = ArHex("622 62F 627 628 20 648 641 644 633 641 629")
    Case "b31": Ar = ArHex("644 63A 627 62A 20 623 62C 646 628 64A 629")
    Case "b32": Ar = ArHex("62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F")
    Case "b33": Ar = ArHex("639 644 648 645 20 62A 62C 631 64A 628 64A 629")
    Case "b34": Ar = ArHex("631 64A 627 636 64A 627 62A")
    Case "b35": Ar = ArHex("62A 642 646 64A 20 631 64A 627 636 64A 20 2D 20 647 646 62F 633 629 20 645 64A 643 627 646 64A 643 64A 629")
    Case "b36": Ar = ArHex("62A 642 646 64A 20 631 64A 627 636 64A 20 2D 20 647 646 62F 633 629 20 643 647 631 628 627 626 64A 629")
    Case "b37": Ar = ArHex("62A 642 646 64A 20 631 64A 627 636 64A 20 2D 20 647 646 62F 633 629 20 645 62F 646 64A 629")
    Case "b38": Ar = ArHex("62A 642 646 64A 20 631 64A 627 636 64A 20 2D 20 647 646 62F 633 629 20 637 631 627 626 642")
    Case "s_count": Ar = ArHex("62A 639 62F 627 62F 20 627 644 642 633 645")
    Case "s_male": Ar = ArHex("639 20 627 644 630 643 648 631")
    Case "s_female": Ar = ArHex("639 20 627 644 625 646 627 62B")
    Case "s_rep": Ar = ArHex("627 644 645 639 64A 62F 648 646")
    Case "l_avg": Ar = ArHex("645 639 62F 644 627 62A 20 627 644 645 648 627 62F")
    Case "l_passf": Ar = ArHex("639 62F 62F 20 627 644 62A 645 64A 630 627 62A 20 627 644 646 627 62C 62D 627 62A")
    Case "l_passm": Ar = ArHex("639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 646 627 62C 62D 64A 646")
    Case "l_pass": Ar = ArHex("645 62C 645 648 639 20 627 644 646 627 62C 62D 64A 646")
    Case "l_rate": Ar = ArHex("646 633 628 629 20 627 644 646 62C 627 62D 20 641 64A 20 643 644 20 645 627 62F 629")
    Case "l_max": Ar = ArHex("623 639 644 649 20 645 639 62F 644")
    Case "l_min": Ar = ArHex("623 62F 646 649 20 645 639 62F 644")
    Case "l_reppass": Ar = ArHex("639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 645 639 64A 62F 64A 646 20 627 644 646 627 62C 62D 64A 646")
    Case "t2_head": Ar = ArHex("627 644 645 624 634 631 20 627 644 625 62D 635 627 626 64A")
    Case "i_rate": Ar = ArHex("646 633 628 629 20 627 644 646 62C 627 62D")
    Case "i_mean": Ar = ArHex("627 644 645 62A 648 633 637 20 627 644 62D 633 627 628 64A")
    Case "i_median": Ar = ArHex("627 644 648 633 64A 637")
    Case "i_mode": Ar = ArHex("627 644 645 646 648 627 644")
    Case "i_range": Ar = ArHex("627 644 645 62F 649")
    Case "i_sd": Ar = ArHex("627 644 627 646 62D 631 627 641 20 627 644 645 639 64A 627 631 64A")
    Case "i_var": Ar = ArHex("627 644 62A 628 627 64A 646")
    Case "i_cv": Ar = ArHex("645 639 627 645 644 20 627 644 627 62E 62A 644 627 641")
    Case "t3_title": Ar = ArHex("62A 648 632 64A 639 20 627 644 645 639 62F 644 627 62A 20 62D 633 628 20 627 644 641 626 627 62A")
    Case "t3_head": Ar = ArHex("627 644 641 626 627 62A 20 2F 20 627 644 645 648 627 62F")
    Case "gen_title": Ar = ArHex("625 62D 635 627 626 64A 627 62A 20 639 627 645 629")
    Case "g1": Ar = ArHex("31 2D 20 639 62F 62F 20 627 644 645 633 62C 644 64A 646")
    Case "g2": Ar = ArHex("32 2D 20 646 62A 627 626 62C 20 639 627 645 629")
    Case "g3": Ar = ArHex("33 2D 20 646 633 628 629 20 627 644 646 62C 627 62D 20 62D 633 628 20 627 644 646 648 639")
    Case "g4": Ar = ArHex("34 2D 20 646 633 628 629 20 627 644 646 62C 627 62D 20 641 64A 20 634 647 627 62F 629 20 627 644 628 643 627 644 648 631 64A 627 20 62D 633 628 20 627 644 645 639 64A 62F 64A 646")
    Case "g5": Ar = ArHex("35 2D 20 645 642 627 631 646 629 20 646 633 628 629 20 627 644 646 62C 627 62D 20 641 64A 20 634 647 627 62F 629 20 627 644 628 643 627 644 648 631 64A 627 20 645 639 20 627 644 646 62A 627 626 62C 20 627 644 648 644 627 626 64A 629 20 648 627 644 648 637 646 64A 629")
    Case "g6": Ar = ArHex("36 2D 20 645 642 627 631 646 629 20 646 633 628 629 20 627 644 646 62C 627 62D 20 641 64A 20 634 647 627 62F 629 20 627 644 628 643 627 644 648 631 64A 627 20 645 639 20 646 633 628 629 20 627 644 646 62C 627 62D 20 641 64A 20 627 644 62F 648 631 627 62A 20 627 644 645 627 636 64A 629")
    Case "g7": Ar = ArHex("37 2D 20 646 633 628 629 20 627 644 646 62C 627 62D 20 641 64A 20 634 647 627 62F 629 20 627 644 628 643 627 644 648 631 64A 627 20 62D 633 628 20 627 644 645 639 64A 62F 64A 646")
    Case "h_branch": Ar = ArHex("627 644 634 639 628 629")
    Case "h_reg": Ar = ArHex("627 644 645 633 62C 644 64A 646")
    Case "h_pres": Ar = ArHex("627 644 62D 627 636 631 64A 646")
    Case "h_male": Ar = ArHex("627 644 630 643 648 631")
    Case "h_female": Ar = ArHex("627 644 625 646 627 62B")
    Case "h_abs": Ar = ArHex("627 644 63A 627 626 628 64A 646")
    Case "h_passed": Ar = ArHex("627 644 646 627 62C 62D 64A 646")
    Case "h_type": Ar = ArHex("627 644 646 648 639")
    Case "h_presm": Ar = ArHex("627 644 62D 627 636 631 648 646")
    Case "h_passm": Ar = ArHex("627 644 646 627 62C 62D 648 646")
    Case "h_rate": Ar = ArHex("646 633 628 629 20 627 644 646 62C 627 62D")
    Case "h_repcnt": Ar = ArHex("639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 645 639 64A 62F 64A 646")
    Case "h_passof": Ar = ArHex("627 644 646 627 62C 62D 648 646 20 645 646 647 645")
    Case "h_school": Ar = ArHex("646 633 628 629 20 627 644 646 62C 627 62D 20 641 64A 20 627 644 62B 627 646 648 64A 629")
    Case "h_prov": Ar = ArHex("627 644 646 633 628 629 20 627 644 648 644 627 626 64A 629")
    Case "h_nat": Ar = ArHex("627 644 646 633 628 629 20 627 644 648 637 646 64A 629")
    Case "h_rank": Ar = ArHex("627 644 631 62A 628 629 20 648 644 627 626 64A 627")
    Case "h_rep": Ar = ArHex("627 644 62A 644 627 645 64A 630 20 627 644 645 639 64A 62F 648 646")
    Case "total": Ar = ArHex("627 644 645 62C 645 648 639 20 627 644 639 627 645")
    Case "total2": Ar = ArHex("627 644 645 62C 645 648 639")
    Case "male": Ar = ArHex("630 643 648 631")
    Case "female": Ar = ArHex("625 646 627 62B")
    Case "session": Ar = ArHex("62F 648 631 629")
    End Select
End Function
