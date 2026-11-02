Attribute VB_Name = "modRepeatStatus"
Option Explicit

' =====================================================================
'  Repeat-status matcher  -  macro: Run_RepeatStatus
'
'  One-step replacement for the old chain
'     modBuildRepeatStatus -> RepeatStatus.xlsx -> modMatchRepeatStatus
'  The intermediate RepeatStatus.xlsx file is NOT created any more: the
'  repeat data is built in memory and matched straight into the
'  extraction workbook.
'
'  QUESTION 1 (source): the user picks the 3rd-term results workbook. Its
'           name and location may differ, its structure must be identical.
'           - only sheets whose merged H1 starts with the 3rd-secondary prefix
'           - the last 3 characters of H1 (space + 2 digits) are removed to
'             get the branch group; sheets of the same group are pooled
'           - columns read: B = name, C = date of birth, E = repeat Yes/No
'           - duplicate pupil records are kept (as before)
'           - the source workbook is never modified
'  QUESTION 2 (paste file): the workbook that receives the values. If a
'           last extraction file exists it is offered as the default
'           (Yes = use it, No = choose another file); otherwise the user
'           is asked to pick it.
'  MATCH  : same sheet (branch code) + normalized name (col B) + date of
'           birth (col C). Name normalization handles alef forms, taa
'           marbuta, diacritics, tatweel and spaces. One-letter spelling
'           differences are accepted only when the date of birth matches
'           and there is exactly ONE candidate.
'  OUTPUT : values written to column E only (no formulas).
'           Everything is computed first; the target is touched (after a
'           timestamped backup) only when the whole computation succeeded.
' =====================================================================

Private Const FIRST_DATA_ROW As Long = 3
Private Const NAME_COL As Long = 2
Private Const DOB_COL As Long = 3
Private Const REPEAT_COL As Long = 5
Private Const SRC_LAST_COL As Long = 5          ' source columns A:E

Private Const MATCH_NONE As Long = 0
Private Const MATCH_EXACT As Long = 1
Private Const MATCH_FUZZY As Long = 2
Private Const MATCH_AMBIGUOUS As Long = 3

' Repeat records of ONE branch (one worksheet code), already normalized
Private Type RepeatSet
    n As Long
    nm() As String      ' normalized full name
    dob() As String     ' normalized date of birth (yyyymmdd)
    st() As String      ' raw repeat value (Yes / No word)
End Type

Private mSets() As RepeatSet
Private mLastSourcePath As String   ' remembered for the session (start folder)
Private mYes As String
Private mNo As String

' ---------------------------------------------------------------------
'  Main macro
' ---------------------------------------------------------------------
Public Sub Run_RepeatStatus()
    Dim sourcePath As String, targetPath As String
    Dim backupPath As String, backupExt As String, dotPos As Long
    Dim wbSource As Workbook, wbTarget As Workbook
    Dim wsSource As Worksheet, wsT As Worksheet
    Dim openedSource As Boolean, openedTarget As Boolean, saved As Boolean
    Dim groups As Object, setIndex As Object, pend As Object
    Dim titleText As String, groupKey As String, sheetCode As String
    Dim setCount As Long, idx As Long
    Dim includedSheets As Long, ignoredSheets As Long
    Dim lastT As Long, i As Long, totalWrites As Long
    Dim tData As Variant, writeVals As Variant, key As Variant
    Dim outVals() As Variant
    Dim tName As String, tDob As String, newVal As String
    Dim exactCount As Long, fuzzyCount As Long, notFound As Long, ambiguous As Long
    Dim skippedSheets As String, report As String, fatalText As String
    Dim prevScreen As Boolean, prevEvents As Boolean, prevAlerts As Boolean
    Dim prevSec As MsoAutomationSecurity, settingsCaptured As Boolean

    On Error GoTo FatalError

    mYes = Ar("yes")
    mNo = Ar("no")
    Erase mSets

    ' ---- Question 1 of 2: the SOURCE file (read only) ----
    sourcePath = AskSourceFile()
    If Len(sourcePath) = 0 Then Exit Sub            ' cancelled
    mLastSourcePath = sourcePath

    ' ---- Question 2 of 2: the file that RECEIVES the values ----
    targetPath = AskTargetFile()
    If Len(targetPath) = 0 Then Exit Sub            ' cancelled

    If StrComp(sourcePath, targetPath, vbTextCompare) = 0 Then
        ShowMsg Ar("same_file"), vbExclamation
        Exit Sub
    End If

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevSec = Application.AutomationSecurity
    settingsCaptured = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.StatusBar = Ar("working")

    ' ---- open the target first (fail early if it is read-only) ----
    Set wbTarget = GetOpenWorkbook(targetPath)
    If wbTarget Is Nothing Then
        Set wbTarget = Workbooks.Open(Filename:=targetPath, UpdateLinks:=False)
        openedTarget = True
    End If
    If wbTarget.ReadOnly Then Err.Raise vbObjectError + 861, "Run_RepeatStatus", Ar("target_ro")

    ' ---- open the source (read only, never modified) ----
    Set wbSource = GetOpenWorkbook(sourcePath)
    If wbSource Is Nothing Then
        Set wbSource = Workbooks.Open(Filename:=sourcePath, UpdateLinks:=False, _
                                      ReadOnly:=True, AddToMRU:=False)
        openedSource = True
    End If

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    Set setIndex = CreateObject("Scripting.Dictionary")
    setIndex.CompareMode = vbTextCompare
    Set pend = CreateObject("Scripting.Dictionary")
    pend.CompareMode = vbTextCompare

    ' =================================================================
    '  PHASE 1: build the repeat data in memory (replaces RepeatStatus.xlsx)
    ' =================================================================
    For Each wsSource In wbSource.Worksheets
        titleText = GetMergedCellText(wsSource.Range("H1"))

        If StartsWithThirdSecondary(titleText) Then
            includedSheets = includedSheets + 1
            groupKey = GetGroupKey(titleText)
            sheetCode = SheetCodeFromTitle(groupKey)

            If Len(sheetCode) = 0 Then
                Err.Raise vbObjectError + 852, "Run_RepeatStatus", _
                          Ar("no_code") & vbLf & wsSource.Name & " : " & titleText
            End If

            If Not groups.Exists(groupKey) Then
                If setIndex.Exists(sheetCode) Then
                    Err.Raise vbObjectError + 853, "Run_RepeatStatus", _
                              Ar("dup_code") & vbLf & sheetCode
                End If
                groups.Add groupKey, sheetCode
                setCount = setCount + 1
                If setCount = 1 Then
                    ReDim mSets(1 To 1)
                Else
                    ReDim Preserve mSets(1 To setCount)
                End If
                setIndex.Add sheetCode, setCount
            End If

            AppendSourceRows wsSource, CLng(setIndex(sheetCode))
        Else
            ignoredSheets = ignoredSheets + 1
        End If
    Next wsSource

    If groups.Count = 0 Then
        Err.Raise vbObjectError + 851, "Run_RepeatStatus", _
                  Ar("no_group") & vbLf & ThirdSecondaryPrefix()
    End If

    ' =================================================================
    '  PHASE 2: match every target pupil (nothing is written yet)
    ' =================================================================
    For Each wsT In wbTarget.Worksheets
        If StrComp(wsT.Name, "Gen", vbTextCompare) <> 0 Then
            If setIndex.Exists(wsT.Name) Then
                idx = CLng(setIndex(wsT.Name))
                lastT = LastUsedRow(wsT, NAME_COL)

                If lastT >= FIRST_DATA_ROW Then
                    tData = wsT.Range(wsT.Cells(FIRST_DATA_ROW, 1), _
                                      wsT.Cells(lastT, REPEAT_COL)).Value
                    ReDim outVals(1 To UBound(tData, 1))

                    For i = 1 To UBound(tData, 1)
                        tName = NormalizeArabicName(CStrSafe(tData(i, NAME_COL)))
                        tDob = NormalizeDate(tData(i, DOB_COL))

                        If Len(tName) > 0 And Len(tDob) > 0 Then
                            Select Case MatchPupil(idx, tName, tDob, newVal)
                                Case MATCH_EXACT
                                    outVals(i) = newVal
                                    exactCount = exactCount + 1
                                    totalWrites = totalWrites + 1
                                Case MATCH_FUZZY
                                    outVals(i) = newVal
                                    fuzzyCount = fuzzyCount + 1
                                    totalWrites = totalWrites + 1
                                Case MATCH_AMBIGUOUS
                                    ambiguous = ambiguous + 1
                                Case Else
                                    notFound = notFound + 1
                            End Select
                        Else
                            notFound = notFound + 1
                        End If
                    Next i

                    pend.Add wsT.Name, outVals
                End If
            Else
                If Len(skippedSheets) > 0 Then skippedSheets = skippedSheets & ", "
                skippedSheets = skippedSheets & wsT.Name
            End If
        End If
    Next wsT

    If pend.Count = 0 Then
        Err.Raise vbObjectError + 862, "Run_RepeatStatus", _
                  Ar("no_match_sheets") & vbLf & skippedSheets
    End If

    ' =================================================================
    '  PHASE 3: backup, write column E, save
    ' =================================================================
    If totalWrites > 0 Then
        dotPos = InStrRev(targetPath, ".")
        If dotPos > InStrRev(targetPath, Application.PathSeparator) Then
            backupExt = Mid$(targetPath, dotPos)
        Else
            backupExt = ".xlsx"
        End If
        backupPath = Left$(targetPath, InStrRev(targetPath, Application.PathSeparator)) & _
                     "Backup_RepeatStatus_" & Format(Now, "yyyymmdd_hhnnss") & backupExt
        wbTarget.SaveCopyAs backupPath

        For Each key In pend.Keys
            Set wsT = wbTarget.Worksheets(CStr(key))
            writeVals = pend(key)
            For i = LBound(writeVals) To UBound(writeVals)
                If Not IsEmpty(writeVals(i)) Then
                    wsT.Cells(FIRST_DATA_ROW + i - 1, REPEAT_COL).Value = writeVals(i)
                End If
            Next i
        Next key

        wbTarget.Save
        saved = True
    End If

    GoTo CleanExit

FatalError:
    fatalText = CStr(Err.Number) & " - " & Err.Description
    Resume CleanExit

CleanExit:
    On Error Resume Next
    If openedSource Then wbSource.Close SaveChanges:=False
    If openedTarget And Not saved Then wbTarget.Close SaveChanges:=False
    If settingsCaptured Then
        Application.AutomationSecurity = prevSec
        Application.DisplayAlerts = prevAlerts
        Application.EnableEvents = prevEvents
        Application.ScreenUpdating = prevScreen
        Application.StatusBar = False
    End If
    Erase mSets
    On Error GoTo 0

    If Len(fatalText) > 0 Then
        ShowMsg Ar("stopped") & vbLf & fatalText, vbExclamation
        Exit Sub
    End If

    If saved Then
        report = Ar("done") & vbLf & vbLf
    Else
        report = Ar("no_change") & vbLf & vbLf
    End If
    report = report & _
             Ar("r_source") & vbLf & sourcePath & vbLf & _
             Ar("r_target") & vbLf & targetPath & vbLf & vbLf & _
             Ar("src_sheets") & includedSheets & vbLf & _
             Ar("src_groups") & groups.Count & vbLf & vbLf & _
             Ar("r_exact") & exactCount & vbLf & _
             Ar("r_fuzzy") & fuzzyCount & vbLf & _
             Ar("r_none") & notFound & vbLf & _
             Ar("r_amb") & ambiguous
    If Len(skippedSheets) > 0 Then
        report = report & vbLf & vbLf & Ar("r_skipped") & skippedSheets
    End If
    If saved Then
        report = report & vbLf & vbLf & Ar("r_backup") & vbLf & backupPath
    End If
    ShowMsg report, vbInformation
End Sub

' ---------------------------------------------------------------------
'  Phase 1 helper: read B / C / E of one source sheet into its branch set
' ---------------------------------------------------------------------
Private Sub AppendSourceRows(ByVal ws As Worksheet, ByVal idx As Long)
    Dim lastRow As Long, data As Variant, r As Long, base As Long, cap As Long
    Dim nmN As String, dobN As String

    lastRow = LastDataRowAE(ws)
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    data = ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastRow, SRC_LAST_COL)).Value
    base = mSets(idx).n
    cap = base + UBound(data, 1)

    If base = 0 Then
        ReDim mSets(idx).nm(1 To cap)
        ReDim mSets(idx).dob(1 To cap)
        ReDim mSets(idx).st(1 To cap)
    Else
        ReDim Preserve mSets(idx).nm(1 To cap)
        ReDim Preserve mSets(idx).dob(1 To cap)
        ReDim Preserve mSets(idx).st(1 To cap)
    End If

    For r = 1 To UBound(data, 1)
        nmN = NormalizeArabicName(CStrSafe(data(r, NAME_COL)))
        ' Rows without a name or a date of birth can never match: skip them
        If Len(nmN) > 0 Then
            dobN = NormalizeDate(data(r, DOB_COL))
            If Len(dobN) > 0 Then
                base = base + 1
                mSets(idx).nm(base) = nmN
                mSets(idx).dob(base) = dobN
                mSets(idx).st(base) = CStrSafe(data(r, REPEAT_COL))
            End If
        End If
    Next r
    mSets(idx).n = base
End Sub

' ---------------------------------------------------------------------
'  Phase 2 helper: match ONE target pupil against its branch set
'  Returns MATCH_xxx; outVal receives the value to write (Yes / No word)
' ---------------------------------------------------------------------
Private Function MatchPupil(ByVal idx As Long, ByVal tName As String, _
                            ByVal tDob As String, ByRef outVal As String) As Long
    Dim rR As Long, cnt As Long, cand As Long, d As Long
    Dim firstStat As String, sameStat As Boolean

    MatchPupil = MATCH_NONE
    outVal = ""

    With mSets(idx)
        ' Pass 1: normalized exact name + exact date of birth
        For rR = 1 To .n
            If .dob(rR) = tDob Then
                If .nm(rR) = tName Then
                    cnt = cnt + 1
                    cand = rR
                End If
            End If
        Next rR

        If cnt = 1 Then
            If IsValidRepeatValue(.st(cand)) Then
                outVal = Trim$(.st(cand))
                MatchPupil = MATCH_EXACT
            End If

        ElseIf cnt > 1 Then
            ' Repeated identity is safe only if every matching row carries
            ' the same valid Yes/No value.
            firstStat = .st(cand)
            sameStat = IsValidRepeatValue(firstStat)
            If sameStat Then
                For rR = 1 To .n
                    If .dob(rR) = tDob Then
                        If .nm(rR) = tName Then
                            If .st(rR) <> firstStat Then
                                sameStat = False
                                Exit For
                            End If
                        End If
                    End If
                Next rR
            End If
            If sameStat Then
                outVal = Trim$(firstStat)
                MatchPupil = MATCH_EXACT
            Else
                MatchPupil = MATCH_AMBIGUOUS
            End If

        Else
            ' Pass 2: one spelling edit allowed, same date of birth, unique only
            cand = 0
            For rR = 1 To .n
                If .dob(rR) = tDob Then
                    d = LevenshteinDistance(tName, .nm(rR))
                    If d = 1 Then
                        cnt = cnt + 1
                        cand = rR
                    End If
                End If
            Next rR

            If cnt = 1 Then
                If IsValidRepeatValue(.st(cand)) Then
                    outVal = Trim$(.st(cand))
                    MatchPupil = MATCH_FUZZY
                End If
            ElseIf cnt > 1 Then
                MatchPupil = MATCH_AMBIGUOUS
            End If
        End If
    End With
End Function

' ---------------------------------------------------------------------
'  Source-title helpers (from the old modBuildRepeatStatus)
' ---------------------------------------------------------------------
Private Function GetMergedCellText(ByVal targetCell As Range) As String
    Dim v As Variant

    If targetCell.MergeCells Then
        v = targetCell.MergeArea.Cells(1, 1).Value2
    Else
        v = targetCell.Value2
    End If

    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    GetMergedCellText = Trim$(Replace(CStr(v), ChrW(160), " "))
End Function

Private Function StartsWithThirdSecondary(ByVal titleText As String) As Boolean
    StartsWithThirdSecondary = _
        (StrComp(Left$(Trim$(titleText), Len(ThirdSecondaryPrefix())), _
                 ThirdSecondaryPrefix(), vbTextCompare) = 0)
End Function

Private Function GetGroupKey(ByVal titleText As String) As String
    Dim s As String, lastTwo As String

    s = Trim$(Replace(titleText, ChrW(160), " "))

    ' Remove exactly the final three characters only when they are
    ' a separating space followed by two digits (e.g. " 01", " 02").
    If Len(s) >= 3 Then
        lastTwo = Right$(s, 2)
        If Mid$(s, Len(s) - 2, 1) = " " And IsNumeric(lastTwo) Then
            s = Left$(s, Len(s) - 3)
        End If
    End If

    GetGroupKey = Trim$(s)
End Function

Private Function SheetCodeFromTitle(ByVal groupTitle As String) As String
    ' Worksheet codes are identical to modBacEtab.SeriesSheetName:
    ' Philo, LaEtr, GeEc, SE, Math, TGM, TGE, TGC, TGP.
    If InStr(1, groupTitle, ArabicText(&H622, &H62F, &H627, &H628), vbTextCompare) > 0 Or _
       InStr(1, groupTitle, ArabicText(&H641, &H644, &H633, &H641, &H629), vbTextCompare) > 0 Then
        SheetCodeFromTitle = "Philo"
    ElseIf InStr(1, groupTitle, ArabicText(&H644, &H63A, &H627, &H62A, &H20, &H623, &H62C, &H646, &H628, &H64A, &H629), vbTextCompare) > 0 Or _
           InStr(1, groupTitle, ArabicText(&H644, &H63A, &H627, &H62A, &H20, &H627, &H62C, &H646, &H628, &H64A, &H629), vbTextCompare) > 0 Then
        SheetCodeFromTitle = "LaEtr"
    ElseIf InStr(1, groupTitle, ArabicText(&H62A, &H633, &H64A, &H64A, &H631), vbTextCompare) > 0 Or _
           InStr(1, groupTitle, ArabicText(&H627, &H642, &H62A, &H635, &H627, &H62F), vbTextCompare) > 0 Then
        SheetCodeFromTitle = "GeEc"
    ElseIf InStr(1, groupTitle, ArabicText(&H639, &H644, &H648, &H645, &H20, &H62A, &H62C, &H631, &H64A, &H628, &H64A, &H629), vbTextCompare) > 0 Then
        SheetCodeFromTitle = "SE"
    ElseIf InStr(1, groupTitle, ArabicText(&H631, &H64A, &H627, &H636, &H64A, &H627, &H62A), vbTextCompare) > 0 And _
           InStr(1, groupTitle, ArabicText(&H62A, &H642, &H646, &H64A), vbTextCompare) = 0 Then
        SheetCodeFromTitle = "Math"
    ElseIf InStr(1, groupTitle, ArabicText(&H645, &H64A, &H643, &H627, &H646, &H64A, &H643, &H64A, &H629), vbTextCompare) > 0 Then
        SheetCodeFromTitle = "TGM"
    ElseIf InStr(1, groupTitle, ArabicText(&H643, &H647, &H631, &H628, &H627, &H626, &H64A, &H629), vbTextCompare) > 0 Then
        SheetCodeFromTitle = "TGE"
    ElseIf InStr(1, groupTitle, ArabicText(&H645, &H62F, &H646, &H64A, &H629), vbTextCompare) > 0 Then
        SheetCodeFromTitle = "TGC"
    ElseIf InStr(1, groupTitle, ArabicText(&H637, &H631, &H627, &H626, &H642), vbTextCompare) > 0 Then
        SheetCodeFromTitle = "TGP"
    End If
End Function

Private Function ThirdSecondaryPrefix() As String
    ThirdSecondaryPrefix = ArabicText(&H62B, &H627, &H644, &H62B, &H629, &H20, _
                                      &H62B, &H627, &H646, &H648, &H64A)
End Function

Private Function LastDataRowAE(ByVal ws As Worksheet) As Long
    Dim searchArea As Range, foundCell As Range

    Set searchArea = ws.Range("A:E")
    Set foundCell = searchArea.Find(What:="*", After:=searchArea.Cells(1, 1), _
                                    LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                    MatchCase:=False)

    If foundCell Is Nothing Then
        LastDataRowAE = 0
    Else
        LastDataRowAE = foundCell.Row
    End If
End Function

' ---------------------------------------------------------------------
'  Matching helpers (from the old modMatchRepeatStatus)
' ---------------------------------------------------------------------
Private Function NormalizeArabicName(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    s = Trim$(s)
    s = Replace(s, ChrW(&H640), "") ' Tatweel

    ' Remove Arabic diacritics and normalize common Arabic letter variants.
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536

        If (code >= &H64B And code <= &H65F) Or code = &H670 Then
            ' Skip tashkeel.
        Else
            Select Case code
                Case &H623, &H625, &H622, &H671
                    ch = ChrW(&H627) ' alef with hamza / madda / wasla -> plain alef
                Case &H629
                    ch = ChrW(&H647) ' taa marbuta -> haa
            End Select

            If ch = vbTab Or ch = ChrW(160) Or ch = vbCr Or ch = vbLf Then ch = " "
            out = out & ch
        End If
    Next i

    out = LCase$(out)
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    NormalizeArabicName = Trim$(out)
End Function

Private Function NormalizeDate(ByVal v As Variant) As String
    Dim s As String
    On Error GoTo BadDate
    If IsEmpty(v) Or IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then Exit Function

    If IsDate(v) Then
        NormalizeDate = Format$(CDate(v), "yyyymmdd")
    Else
        s = Trim$(CStr(v))
        s = Replace(s, "/", "-")
        s = Replace(s, ".", "-")
        If IsDate(s) Then NormalizeDate = Format$(CDate(s), "yyyymmdd")
    End If
    Exit Function
BadDate:
    NormalizeDate = ""
End Function

Private Function LevenshteinDistance(ByVal s As String, ByVal t As String) As Long
    Dim d() As Long, i As Long, j As Long, cost As Long
    Dim n As Long, m As Long, v1 As Long, v2 As Long, v3 As Long

    n = Len(s): m = Len(t)
    If n = 0 Then LevenshteinDistance = m: Exit Function
    If m = 0 Then LevenshteinDistance = n: Exit Function
    If Abs(n - m) > 1 Then LevenshteinDistance = 2: Exit Function

    ReDim d(0 To n, 0 To m)
    For i = 0 To n: d(i, 0) = i: Next i
    For j = 0 To m: d(0, j) = j: Next j

    For i = 1 To n
        For j = 1 To m
            If Mid$(s, i, 1) = Mid$(t, j, 1) Then cost = 0 Else cost = 1
            v1 = d(i - 1, j) + 1
            v2 = d(i, j - 1) + 1
            v3 = d(i - 1, j - 1) + cost
            d(i, j) = Min3(v1, v2, v3)
        Next j
    Next i
    LevenshteinDistance = d(n, m)
End Function

Private Function Min3(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Min3 = a
    If b < Min3 Then Min3 = b
    If c < Min3 Then Min3 = c
End Function

Private Function IsValidRepeatValue(ByVal v As Variant) As Boolean
    Dim s As String
    s = Trim$(CStrSafe(v))
    IsValidRepeatValue = (s = mYes Or s = mNo)
End Function

' ---------------------------------------------------------------------
'  General helpers
' ---------------------------------------------------------------------
Private Function CStrSafe(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        CStrSafe = ""
    Else
        CStrSafe = CStr(v)
    End If
End Function

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal colNum As Long) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
End Function

Private Function GetOpenWorkbook(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set GetOpenWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    If Len(fullPath) = 0 Then Exit Function
    FileExists = (Len(Dir$(fullPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0)
End Function

Private Function PickFile(ByVal dlgTitle As String, ByVal startPath As String, _
                          ByVal pattern As String) As String
    On Error Resume Next
    With Application.FileDialog(3)
        .Title = dlgTitle
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel", pattern
        If Len(startPath) > 0 Then .InitialFileName = startPath
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

' Question 1 of 2 - explains first, then opens the file dialog
Private Function AskSourceFile() As String
    Dim startPath As String

    If MsgBox(Ar("q1_head") & vbLf & vbLf & Ar("q1_note") & vbLf & Ar("q1_safe"), _
              vbOKCancel + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, _
              Ar("q1_title")) <> vbOK Then Exit Function

    startPath = mLastSourcePath
    If Len(startPath) = 0 And Len(ThisWorkbook.Path) > 0 Then
        startPath = ThisWorkbook.Path & Application.PathSeparator
    End If
    AskSourceFile = PickFile(Ar("q1_dialog"), startPath, "*.xlsx;*.xlsm;*.xls")
End Function

' Question 2 of 2 - the last extraction file is offered as the default
Private Function AskTargetFile() As String
    Dim lastExtract As String, answer As VbMsgBoxResult, startPath As String

    On Error Resume Next
    lastExtract = modBacEtab.gLastBacEtabPath
    On Error GoTo 0

    If FileExists(lastExtract) Then
        answer = MsgBox(Ar("q2_head") & vbLf & vbLf & Ar("q2_last") & vbLf & lastExtract & vbLf & vbLf & _
                        Ar("q2_yes") & vbLf & Ar("q2_no") & vbLf & Ar("q2_cancel"), _
                        vbYesNoCancel + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, _
                        Ar("q2_title"))
        Select Case answer
            Case vbYes
                AskTargetFile = lastExtract
                Exit Function
            Case vbNo
                startPath = lastExtract
            Case Else
                Exit Function
        End Select
    Else
        If MsgBox(Ar("q2_head") & vbLf & vbLf & Ar("q2_pick"), _
                  vbOKCancel + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, _
                  Ar("q2_title")) <> vbOK Then Exit Function
        If Len(ThisWorkbook.Path) > 0 Then startPath = ThisWorkbook.Path & Application.PathSeparator
    End If

    AskTargetFile = PickFile(Ar("q2_dialog"), startPath, "*.xlsx;*.xlsm")
End Function

Private Sub ShowMsg(ByVal txt As String, ByVal icon As Long)
    MsgBox txt, icon + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("title")
End Sub

' ---------------------------------------------------------------------
'  Arabic text (ASCII-safe: no Arabic literals in the source file)
' ---------------------------------------------------------------------
Private Function ArabicText(ParamArray codePoints() As Variant) As String
    Dim i As Long
    For i = LBound(codePoints) To UBound(codePoints)
        ArabicText = ArabicText & ChrW(CLng(codePoints(i)))
    Next i
End Function

Private Function ArHex(ByVal codes As String) As String
    Dim token As Variant
    For Each token In Split(codes, " ")
        If Len(CStr(token)) > 0 Then ArHex = ArHex & ChrW(CLng("&H" & CStr(token)))
    Next token
End Function

Private Function Ar(ByVal key As String) As String
    Select Case key
    Case "title": Ar = ArHex("62C 644 628 20 62D 627 644 627 62A 20 627 644 625 639 627 62F 629")
    Case "working": Ar = ArHex("62C 627 631 64A 20 645 637 627 628 642 629 20 627 644 623 633 645 627 621 20 648 62A 648 627 631 64A 62E 20 627 644 645 64A 644 627 62F 2E 2E 2E")
    Case "q1_title": Ar = ArHex("627 644 633 624 627 644 20 31 20 645 646 20 32 3A 20 627 644 645 644 641 20 627 644 645 635 62F 631")
    Case "q1_head": Ar = ArHex("627 62E 62A 631 20 627 644 622 646 20 627 644 645 644 641 20 627 644 645 635 62F 631 3A 20 648 647 648 20 645 644 641 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 644 62B 20 627 644 630 64A 20 62A 64F 62C 644 628 20 645 646 647 20 62D 627 644 627 62A 20 627 644 625 639 627 62F 629 20 28 646 639 645 20 2F 20 644 627 29 2E")
    Case "q1_note": Ar = ArHex("64A 645 643 646 20 623 646 20 64A 62E 62A 644 641 20 627 633 645 647 20 648 645 643 627 646 647 60C 20 644 643 646 20 64A 62C 628 20 623 646 20 62A 643 648 646 20 628 646 64A 62A 647 20 645 637 627 628 642 629 20 62A 645 627 645 627 64B 20 644 644 645 644 641 20 627 644 645 639 62A 627 62F 2E")
    Case "q1_safe": Ar = ArHex("644 646 20 64A 62A 645 20 62A 639 62F 64A 644 20 647 630 627 20 627 644 645 644 641 2E")
    Case "q1_dialog": Ar = ArHex("31 2F 32 20 2D 20 627 62E 62A 631 20 627 644 645 644 641 20 627 644 645 635 62F 631 20 28 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 644 62B 29")
    Case "q2_title": Ar = ArHex("627 644 633 624 627 644 20 32 20 645 646 20 32 3A 20 645 644 641 20 627 644 644 635 642")
    Case "q2_head": Ar = ArHex("645 644 641 20 627 644 644 635 642 20 647 648 20 627 644 645 644 641 20 627 644 630 64A 20 633 62A 64F 643 62A 628 20 641 64A 647 20 62D 627 644 627 62A 20 627 644 625 639 627 62F 629 20 641 64A 20 627 644 639 645 648 62F 20 45 2E")
    Case "q2_last": Ar = ArHex("622 62E 631 20 645 644 641 20 627 633 62A 62E 631 627 62C 20 647 648 3A")
    Case "q2_yes": Ar = ArHex("646 639 645 3A 20 627 633 62A 62E 62F 627 645 20 647 630 627 20 627 644 645 644 641")
    Case "q2_no": Ar = ArHex("644 627 3A 20 627 62E 62A 64A 627 631 20 645 644 641 20 622 62E 631")
    Case "q2_cancel": Ar = ArHex("625 644 63A 627 621 3A 20 625 64A 642 627 641 20 627 644 639 645 644 64A 629")
    Case "q2_pick": Ar = ArHex("627 62E 62A 631 20 627 644 622 646 20 645 644 641 20 627 644 644 635 642 60C 20 648 647 648 20 639 627 62F 629 20 645 644 641 20 627 644 627 633 62A 62E 631 627 62C 20 627 644 646 627 62A 62C 20 639 646 20 632 631 20 AB 627 633 62A 62E 631 627 62C 20 646 62A 627 626 62C 20 627 644 645 624 633 633 629 BB 2E")
    Case "q2_dialog": Ar = ArHex("32 2F 32 20 2D 20 627 62E 62A 631 20 645 644 641 20 627 644 644 635 642 20 28 627 644 630 64A 20 62A 64F 643 62A 628 20 641 64A 647 20 62D 627 644 627 62A 20 627 644 625 639 627 62F 629 29")
    Case "same_file": Ar = ArHex("627 644 645 644 641 20 627 644 645 635 62F 631 20 648 645 644 641 20 627 644 644 635 642 20 647 645 627 20 627 644 645 644 641 20 646 641 633 647 2E 20 627 62E 62A 631 20 645 644 641 64A 646 20 645 62E 62A 644 641 64A 646 2E")
    Case "target_ro": Ar = ArHex("645 644 641 20 627 644 644 635 642 20 645 641 62A 648 62D 20 644 644 642 631 627 621 629 20 641 642 637 2E 20 623 63A 644 642 647 20 62B 645 20 623 639 62F 20 627 644 62A 634 63A 64A 644 2E")
    Case "no_group": Ar = ArHex("644 627 20 62A 648 62C 62F 20 648 631 642 629 20 62A 628 62F 623 20 642 64A 645 629 20 627 644 62E 644 64A 629 20 48 31 20 641 64A 647 627 20 628 639 628 627 631 629 3A")
    Case "no_code": Ar = ArHex("62A 639 630 631 20 62A 62D 62F 64A 62F 20 631 645 632 20 627 644 634 639 628 629 20 645 646 20 627 644 62E 644 64A 629 20 48 31 20 641 64A 20 627 644 648 631 642 629 3A")
    Case "dup_code": Ar = ArHex("623 643 62B 631 20 645 646 20 639 646 648 627 646 20 641 64A 20 627 644 645 635 62F 631 20 64A 642 627 628 644 20 627 644 648 631 642 629 20 646 641 633 647 627 60C 20 64A 631 62C 649 20 645 631 627 62C 639 629 20 639 646 627 648 64A 646 20 48 31 3A")
    Case "no_match_sheets": Ar = ArHex("645 644 641 20 627 644 644 635 642 20 644 627 20 64A 62D 62A 648 64A 20 623 64A 20 648 631 642 629 20 62A 642 627 628 644 20 627 644 634 639 628 20 627 644 645 648 62C 648 62F 629 20 641 64A 20 627 644 645 635 62F 631 2E 20 62A 623 643 62F 20 645 646 20 627 62E 62A 64A 627 631 20 627 644 645 644 641 20 627 644 635 62D 64A 62D 2E 20 623 648 631 627 642 20 645 644 641 20 627 644 644 635 642 3A")
    Case "stopped": Ar = ArHex("62A 648 642 641 62A 20 627 644 639 645 644 64A 629 20 628 633 628 628 20 62E 637 623 3A")
    Case "done": Ar = ArHex("627 643 62A 645 644 62A 20 627 644 645 637 627 628 642 629 20 648 62D 64F 641 638 62A 20 627 644 642 64A 645 20 641 64A 20 627 644 639 645 648 62F 20 45 2E")
    Case "no_change": Ar = ArHex("627 643 62A 645 644 62A 20 627 644 645 637 627 628 642 629 20 648 644 645 20 62A 64F 643 62A 628 20 623 64A 20 642 64A 645 629 60C 20 648 644 645 20 62A 64F 646 634 623 20 646 633 62E 629 20 627 62D 62A 64A 627 637 64A 629 2E")
    Case "r_source": Ar = ArHex("627 644 645 644 641 20 627 644 645 635 62F 631 3A")
    Case "r_target": Ar = ArHex("645 644 641 20 627 644 644 635 642 3A")
    Case "src_sheets": Ar = ArHex("623 648 631 627 642 20 627 644 645 635 62F 631 20 627 644 645 639 62A 645 62F 629 3A 20")
    Case "src_groups": Ar = ArHex("627 644 634 639 628 20 627 644 645 62F 645 62C 629 3A 20")
    Case "r_exact": Ar = ArHex("645 637 627 628 642 627 62A 20 62F 642 64A 642 629 20 628 639 62F 20 62A 648 62D 64A 62F 20 627 644 643 62A 627 628 629 3A 20")
    Case "r_fuzzy": Ar = ArHex("645 637 627 628 642 627 62A 20 628 627 62E 62A 644 627 641 20 625 645 644 627 626 64A 20 628 633 64A 637 3A 20")
    Case "r_none": Ar = ArHex("63A 64A 631 20 645 637 627 628 642 629 3A 20")
    Case "r_amb": Ar = ArHex("62D 627 644 627 62A 20 645 644 62A 628 633 629 20 644 645 20 62A 64F 639 62F 651 644 3A 20")
    Case "r_skipped": Ar = ArHex("623 648 631 627 642 20 641 64A 20 645 644 641 20 627 644 644 635 642 20 644 627 20 64A 642 627 628 644 647 627 20 645 635 62F 631 3A 20")
    Case "r_backup": Ar = ArHex("646 633 62E 629 20 627 62D 62A 64A 627 637 64A 629 20 642 628 644 20 627 644 62A 639 62F 64A 644 3A")
    Case "yes": Ar = ArHex("646 639 645")
    Case "no": Ar = ArHex("644 627")
    End Select
End Function
