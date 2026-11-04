Attribute VB_Name = "modStudentRepetitionStatus"
Option Explicit

' =====================================================================
' Repeat-status matcher - macro: Run_RepeatStatus
' (merged with RepeatingStudents01: reads the RAW files of \Source03\ directly)
'
' Reads the raw third-year files in the Source03 folder (next to this workbook),
' prepares them in memory with the same steps as Collector01 (conversion of the
' web-exported files to .xlsx, then collection) - no results file is created -
' and writes the repeat status (Yes / No) into column E of the extraction
' workbook (TARGET, values only; a temporary backup is deleted after a verified save).
'
' MATCH   : same sheet code + normalized name (col B) + date of birth
'           (col C). EXACT matching only (no fuzzy matching).
'           A pupil is searched only in the sheet of his own branch
'           (SE pools the two source sheets "... 01" and "... 02").
' ORDER   : 1) read Source03 and detect duplicate identities
'           2) per target sheet: detect duplicates, then match
'           3) ask for confirmation, temporary backup, write, save, verify, delete backup
'           4) create a separate report workbook (summary + details)
' REPORT  : locations are written as "SheetCode: number in column A",
'           e.g. SE: 41
'
' Second macro: Run_RepeatingStudents01 - builds RepeatingStudents.xlsx (one
' sheet per stream, A:E only) from the same Source03 data. It is optional and
' not needed by Run_RepeatStatus.
'
' SELF-CONTAINED except for one optional link: the last extraction file offered
' as the default target comes from modBacEtab. Set HAS_BACETAB to False below
' if this module is used in a workbook that does not contain modBacEtab.
' =====================================================================

#Const HAS_BACETAB = True

Private Const FIRST_DATA_ROW As Long = 3
Private Const ID_COL As Long = 1          ' A = sequence number
Private Const NAME_COL As Long = 2        ' B = full name
Private Const DOB_COL As Long = 3         ' C = date of birth
Private Const REPEAT_COL As Long = 5      ' E = repeat Yes / No
Private Const SRC_LAST_COL As Long = 5    ' source columns A:E
Private Const MIN_VALID_RATIO As Double = 0.5   ' source structure check

' Source03 / RepeatingStudents01 settings
Private Const SOURCE_FOLDER_NAME As String = "Source03"
Private Const OUTPUT_FILE_NAME   As String = "RepeatingStudents.xlsx"
Private Const KEEP_COLS          As Long = 5      ' identity columns A:E
Private Const TERM_HDR_ROW       As Long = 2      ' header row of a collected sheet
Private Const TERM_FIRST_ROW     As Long = 3      ' first data row of a collected sheet

' ---- Collector01 constants (embedded) ----
Private Const DATA_START_ROW       As Long = 6      ' header row in the source sheet
Private Const TRAILING_ROWS_SKIP   As Long = 1      ' footer rows after last used row (original logic: -1). Set 0 if none
Private Const IDENTITY_LAST_COL    As Long = 5      ' columns A:E = identity columns (AutoFit)
Private Const COL_WIDTH            As Double = 5.7  ' width for columns F:LastCol
Private Const FONT_SIZE            As Long = 14
Private Const FONT_ARABIC          As String = "Arabic UI Display Light"
Private Const FONT_NUMERIC         As String = "Bahnschrift SemiBold"
Private Const FONT_FALLBACK        As String = "Calibri"
Private Const ARCHIVE_FOLDER       As String = "_Original"

Private Const PREFERRED_FONT As String = "Dubai"
Private Const FALLBACK_FONT  As String = "Arial"
Private Const HEADER_FONT_SIZE As Long = 16
Private Const HEADER_ROW_HEIGHT As Double = 30

Private g_ActiveArabicFont  As String
Private g_ActiveNumericFont As String

Private Const KIND_OK As Byte = 0
Private Const KIND_INVALID As Byte = 1
Private Const KIND_CONFLICT As Byte = 2

Private Const CAT_NOTSRC As Long = 1
Private Const CAT_DUPSRC As Long = 2
Private Const CAT_DUPTGT As Long = 3
Private Const CAT_AMB As Long = 4
Private Const CAT_INV As Long = 5
Private Const CAT_INC As Long = 6
Private Const CAT_SKIP As Long = 7
Private Const CAT_COUNT As Long = 7

' Repeat records of ONE branch (one worksheet code)
Private Type RepeatSet
    code As String
    n As Long
    nm() As String        ' normalized full name
    dob() As String       ' normalized date of birth (yyyymmdd)
    st() As String        ' repeat value (Yes / No word)
    loc() As String       ' location text  "SE: 41"
    shName() As String    ' real source sheet name
    dupCnt() As Long      ' >0 only on the FIRST record of an identity
    kind() As Byte        ' KIND_OK / KIND_INVALID / KIND_CONFLICT
    hit() As Boolean      ' matched by at least one target pupil
    dict As Object        ' key (dob|name) -> index of first record
    targetFound As Boolean
    sIncomplete As Long
    sInv As Long
    sUnique As Long
    sDupIds As Long
    sNotTgt As Long
    tTotal As Long
    tWritten As Long
    tNotSrc As Long
    tDup As Long
    tAmb As Long
    tInv As Long
    tIncomplete As Long
End Type

Private mSets() As RepeatSet
Private mYes As String
Private mNo As String

' log of the report
Private mLogN As Long
Private mLogCat() As Long
Private mLogLoc() As String
Private mLogSide() As String
Private mLogSheet() As String
Private mLogDet() As String

' totals
Private mTotWrite As Long, mTotTgt As Long, mTotNotSrc As Long
Private mTotDupTgt As Long, mTotAmb As Long, mTotInv As Long, mTotIncT As Long
Private mTotSrcUnique As Long, mTotDupSrc As Long, mTotIncS As Long
Private mTotSrcNotTgt As Long, mTotInvSrc As Long

' ---------------------------------------------------------------------
' Main macro
' ---------------------------------------------------------------------
Public Sub Run_RepeatStatus()
    Dim sourcePath As String, targetPath As String
    Dim backupPath As String, backupExt As String, dotPos As Long
    Dim reportPath As String, stamp As String, folder As String
    Dim wbSource As Workbook, wbTarget As Workbook
    Dim wsSource As Worksheet, wsT As Worksheet
    Dim openedSource As Boolean, openedTarget As Boolean
    Dim saved As Boolean, doWrite As Boolean, hasData As Boolean
    Dim groups As Object, setIndex As Object, pend As Object
    Dim titleText As String, groupKey As String, sheetCode As String, reason As String
    Dim setCount As Long, idx As Long, i As Long
    Dim includedSheets As Long, ignoredSheets As Long
    Dim skippedSrc As Long, skippedTgt As Long
    Dim lastRow As Long, srcData As Variant, writeVals As Variant, key As Variant
    Dim outVals() As Variant
    Dim skipText As String, skipShown As Long
    Dim report As String, fatalText As String, answer As VbMsgBoxResult
    Dim prevScreen As Boolean, prevEvents As Boolean, prevAlerts As Boolean
    Dim prevSec As MsoAutomationSecurity, settingsCaptured As Boolean
    Dim srcFolder As String, srcList As String, lineTxt As String, srcLines() As String
    Dim srcCopied As Long, srcSkipped As Long
    Dim convTotal As Long, convOK As Long, convFailed As Long
    Dim backupNote As String, tBeforeSave As Date

    On Error GoTo FatalError

    mYes = Ar("yes")
    mNo = Ar("no")
    Erase mSets
    mLogN = 0
    Erase mLogCat: Erase mLogLoc: Erase mLogSide: Erase mLogSheet: Erase mLogDet

    ' ---- pre-flight: this workbook must be saved locally and Source03 must exist ----
    If Len(ThisWorkbook.Path) = 0 Then
        ShowMsg Ar("save_first"), vbExclamation
        Exit Sub
    End If
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        ShowMsg Ar("url_path"), vbExclamation
        Exit Sub
    End If
    srcFolder = ThisWorkbook.Path & Application.PathSeparator & SOURCE_FOLDER_NAME & Application.PathSeparator
    If Not FolderExistsPath(Left$(srcFolder, Len(srcFolder) - 1)) Then
        ShowMsg Ar("no_src_folder") & vbLf & srcFolder, vbExclamation
        Exit Sub
    End If
    sourcePath = srcFolder

    ' ---- the only question: the file that RECEIVES the values ----
    targetPath = AskTargetFile()
    If Len(targetPath) = 0 Then Exit Sub

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevSec = Application.AutomationSecurity
    settingsCaptured = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.StatusBar = Ar("sb_read")

    ' ---- open the target first (fail early if it is read-only) ----
    Set wbTarget = GetOpenWorkbook(targetPath)
    If wbTarget Is Nothing Then
        Set wbTarget = Workbooks.Open(Filename:=targetPath, UpdateLinks:=False)
        openedTarget = True
    End If
    If wbTarget.ReadOnly Then Err.Raise vbObjectError + 861, "Run_RepeatStatus", Ar("target_ro")

    ' ---- Source03 raw files -> results prepared in memory (Collector steps) ----
    If Not SrcBuildWorkbook(srcFolder, wbSource, srcCopied, srcSkipped, srcList, _
                            convTotal, convOK, convFailed) Then
        Err.Raise vbObjectError + 852, "Run_RepeatStatus", Ar("src_failed") & vbLf & Left$(srcList, 500)
    End If
    openedSource = True

    ' raw files that could not be used are listed in the report details
    If Len(srcList) > 0 Then
        srcLines = Split(srcList, vbCrLf)
        For i = LBound(srcLines) To UBound(srcLines)
            lineTxt = Trim$(srcLines(i))
            If Len(lineTxt) > 0 Then AddLog CAT_SKIP, "", Ar("side_src"), "", lineTxt
        Next i
    End If

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    Set setIndex = CreateObject("Scripting.Dictionary")
    setIndex.CompareMode = vbTextCompare
    Set pend = CreateObject("Scripting.Dictionary")
    pend.CompareMode = vbTextCompare

    ' =================================================================
    ' PHASE 1: read the source, check its structure, detect duplicates
    ' =================================================================
    For Each wsSource In wbSource.Worksheets
        titleText = GetMergedCellText(wsSource.Range("H1"))
        If StartsWithThirdSecondary(titleText) Then
            Application.StatusBar = Ar("sb_read") & wsSource.Name
            groupKey = GetGroupKey(titleText)
            sheetCode = SheetCodeFromTitle(groupKey)
            reason = ""

            If Len(sheetCode) = 0 Then
                reason = Ar("sk_nocode") & titleText
            ElseIf Not groups.Exists(groupKey) And setIndex.Exists(sheetCode) Then
                reason = Ar("sk_dupcode") & sheetCode
            Else
                lastRow = LastDataRowAE(wsSource)
                If lastRow < FIRST_DATA_ROW Then
                    reason = Ar("sk_nodata")
                Else
                    srcData = wsSource.Range(wsSource.Cells(FIRST_DATA_ROW, 1), _
                                             wsSource.Cells(lastRow, SRC_LAST_COL)).Value
                    SourceStructureOK srcData, reason
                End If
            End If

            If Len(reason) > 0 Then
                skippedSrc = skippedSrc + 1
                AddLog CAT_SKIP, wsSource.Name, Ar("side_src"), wsSource.Name, reason
                If skipShown < 8 Then
                    skipText = skipText & vbLf & wsSource.Name & " : " & reason
                    skipShown = skipShown + 1
                End If
            Else
                includedSheets = includedSheets + 1
                If Not groups.Exists(groupKey) Then
                    groups.Add groupKey, sheetCode
                    setCount = setCount + 1
                    If setCount = 1 Then
                        ReDim mSets(1 To 1)
                    Else
                        ReDim Preserve mSets(1 To setCount)
                    End If
                    mSets(setCount).code = sheetCode
                    Set mSets(setCount).dict = CreateObject("Scripting.Dictionary")
                    setIndex.Add sheetCode, setCount
                End If
                AppendSourceRows wsSource.Name, CLng(setIndex(sheetCode)), srcData
            End If
        Else
            ignoredSheets = ignoredSheets + 1
        End If
    Next wsSource

    If groups.Count = 0 Then
        Err.Raise vbObjectError + 851, "Run_RepeatStatus", _
            Ar("no_group") & ThirdSecondaryPrefix() & Ar("no_group2") & Left$(skipText, 600)
    End If

    ' =================================================================
    ' PHASE 2: per target sheet - duplicates first, then matching
    ' (nothing is written yet)
    ' =================================================================
    For Each wsT In wbTarget.Worksheets
        If StrComp(wsT.Name, "Gen", vbTextCompare) <> 0 Then
            If setIndex.Exists(wsT.Name) Then
                idx = CLng(setIndex(wsT.Name))
                Application.StatusBar = Ar("sb_match") & wsT.Name
                MatchTargetSheet wsT, idx, outVals, hasData
                If hasData Then
                    mSets(idx).targetFound = True
                    pend.Add wsT.Name, outVals
                Else
                    skippedTgt = skippedTgt + 1
                    AddLog CAT_SKIP, wsT.Name, Ar("side_tgt"), "", Ar("sk_tgtempty")
                End If
            Else
                skippedTgt = skippedTgt + 1
                AddLog CAT_SKIP, wsT.Name, Ar("side_tgt"), "", Ar("sk_tgtnone")
            End If
        End If
    Next wsT

    If pend.Count = 0 Then
        Err.Raise vbObjectError + 862, "Run_RepeatStatus", Ar("no_match_sheets") & vbLf & wbTarget.Name
    End If

    ComputeTotals setCount

    ' =================================================================
    ' PHASE 3: confirmation, backup, write column E, save
    ' =================================================================
    stamp = Format(Now, "yyyymmdd_hhnnss")
    folder = Left$(targetPath, InStrRev(targetPath, Application.PathSeparator))

    If mTotWrite > 0 Then
        Application.StatusBar = False
        answer = MsgBox(Ar("c_write") & mTotWrite & vbLf & vbLf & Ar("c_yes") & vbLf & Ar("c_no"), _
                        vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("c_title"))
        doWrite = (answer = vbYes)
    End If

    If doWrite Then
        dotPos = InStrRev(targetPath, ".")
        If dotPos > InStrRev(targetPath, Application.PathSeparator) Then
            backupExt = Mid$(targetPath, dotPos)
        Else
            backupExt = ".xlsx"
        End If
        backupPath = folder & "Backup_RepeatStatus_" & stamp & backupExt
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
        tBeforeSave = Now
        wbTarget.Save
        saved = TargetSaveVerified(wbTarget, tBeforeSave)

        ' the backup is removed ONLY after a verified save
        If saved Then
            If DeleteBackupFile(backupPath) Then
                backupNote = Ar("bk_deleted")
            Else
                backupNote = Ar("bk_delfail") & " " & backupPath
            End If
        Else
            backupNote = Ar("bk_keep") & " " & backupPath
        End If
    End If

    ' =================================================================
    ' PHASE 4: report workbook (separate file, target structure untouched)
    ' =================================================================
    Application.StatusBar = Ar("sb_report")
    reportPath = BuildReport(folder, stamp, sourcePath, targetPath, backupNote, doWrite, _
                             includedSheets, skippedSrc, ignoredSheets, skippedTgt, groups.Count, setCount)
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
        If Len(backupPath) > 0 Then
            If FileExistsPath(backupPath) Then
                fatalText = fatalText & vbLf & vbLf & Ar("bk_keep") & vbLf & backupPath
            End If
        End If
        ShowMsg Ar("stopped") & vbLf & fatalText, vbExclamation
        Exit Sub
    End If

    If saved Then
        report = Ar("done")
    ElseIf doWrite Then
        report = Ar("save_unver")
    Else
        report = Ar("no_change")
    End If
    report = report & vbLf & vbLf & _
        Ar("r_written") & mTotWrite & vbLf & _
        Ar("r_notsrc") & mTotNotSrc & vbLf & _
        Ar("r_dupsrc") & mTotDupSrc & vbLf & _
        Ar("r_duptgt") & mTotDupTgt & vbLf & _
        Ar("r_amb") & mTotAmb & vbLf & _
        Ar("r_inv") & mTotInv & vbLf & _
        Ar("r_inct") & mTotIncT & vbLf & _
        Ar("r_incs") & mTotIncS & vbLf & _
        Ar("r_srcnot") & mTotSrcNotTgt
    report = report & vbLf & Ar("src_files") & srcCopied & "  |  " & Ar("r_conv") & convOK & "/" & convTotal
    If srcSkipped > 0 Then report = report & "  |  " & Ar("r_srcbad") & srcSkipped
    If skippedSrc > 0 Then report = report & vbLf & Ar("src_skipped") & skippedSrc
    If doWrite Then report = report & vbLf & vbLf & Ar("r_backup") & vbLf & backupNote
    If Len(reportPath) > 0 Then
        report = report & vbLf & vbLf & Ar("r_report") & vbLf & reportPath
    Else
        report = report & vbLf & vbLf & Ar("r_reportfail")
    End If
    ShowMsg report, vbInformation
End Sub

' ---------------------------------------------------------------------
' Phase 1 helpers
' ---------------------------------------------------------------------

' Structure check of a source sheet; reason stays empty when it is fine
Private Sub SourceStructureOK(ByRef data As Variant, ByRef reason As String)
    Dim r As Long, named As Long, okDob As Long, okRep As Long
    For r = 1 To UBound(data, 1)
        If Len(Trim$(CStrSafe(data(r, NAME_COL)))) > 0 Then
            named = named + 1
            If Len(NormalizeDate(data(r, DOB_COL))) > 0 Then okDob = okDob + 1
            If IsValidRepeatValue(data(r, REPEAT_COL)) Then okRep = okRep + 1
        End If
    Next r
    If named = 0 Then
        reason = Ar("sk_noname")
    ElseIf okDob < named * MIN_VALID_RATIO Then
        reason = Ar("sk_baddob")
    ElseIf okRep < named * MIN_VALID_RATIO Then
        reason = Ar("sk_badrep")
    End If
End Sub

' Read B / C / E of one source sheet into its branch set.
' Duplicate identities and invalid values are detected right here.
Private Sub AppendSourceRows(ByVal shName As String, ByVal idx As Long, ByRef data As Variant)
    Dim r As Long, base As Long, cap As Long, f As Long, rowNo As Long
    Dim nameRaw As String, nmN As String, dobN As String, stv As String
    Dim k As String, aV As String, locTxt As String

    base = mSets(idx).n
    cap = base + UBound(data, 1)

    With mSets(idx)
        If base = 0 Then
            ReDim .nm(1 To cap)
            ReDim .dob(1 To cap)
            ReDim .st(1 To cap)
            ReDim .loc(1 To cap)
            ReDim .shName(1 To cap)
            ReDim .dupCnt(1 To cap)
            ReDim .kind(1 To cap)
            ReDim .hit(1 To cap)
        Else
            ReDim Preserve .nm(1 To cap)
            ReDim Preserve .dob(1 To cap)
            ReDim Preserve .st(1 To cap)
            ReDim Preserve .loc(1 To cap)
            ReDim Preserve .shName(1 To cap)
            ReDim Preserve .dupCnt(1 To cap)
            ReDim Preserve .kind(1 To cap)
            ReDim Preserve .hit(1 To cap)
        End If

        For r = 1 To UBound(data, 1)
            rowNo = FIRST_DATA_ROW + r - 1
            nameRaw = Trim$(CStrSafe(data(r, NAME_COL)))
            nmN = ""
            If Len(nameRaw) > 0 Then nmN = NormalizeArabicName(nameRaw)
            dobN = NormalizeDate(data(r, DOB_COL))
            aV = CStrSafe(data(r, ID_COL))

            If Len(nmN) = 0 Or Len(dobN) = 0 Then
                ' a completely empty row is ignored, a partly filled one is reported
                If Len(nameRaw) > 0 Or Len(Trim$(CStrSafe(data(r, DOB_COL)))) > 0 Then
                    .sIncomplete = .sIncomplete + 1
                    AddLog CAT_INC, LocText(.code, aV, rowNo), Ar("side_src"), shName, Ar("d_inc")
                End If
            Else
                stv = CleanStatus(CStrSafe(data(r, REPEAT_COL)))
                locTxt = LocText(.code, aV, rowNo)
                base = base + 1
                .nm(base) = nmN
                .dob(base) = dobN
                .st(base) = stv
                .loc(base) = locTxt
                .shName(base) = shName
                k = dobN & "|" & nmN

                If Not .dict.Exists(k) Then
                    .dict.Add k, base
                    .dupCnt(base) = 1
                    If Not IsValidRepeatValue(stv) Then
                        .kind(base) = KIND_INVALID
                        .sInv = .sInv + 1
                        AddLog CAT_INV, locTxt, Ar("side_src"), shName, Ar("d_inv") & stv
                    End If
                Else
                    f = CLng(.dict.Item(k))
                    .dupCnt(f) = .dupCnt(f) + 1
                    If .dupCnt(f) = 2 Then
                        AddLog CAT_DUPSRC, .loc(f), Ar("side_src"), .shName(f), Ar("d_dup")
                    End If
                    AddLog CAT_DUPSRC, locTxt, Ar("side_src"), shName, Ar("d_dup")
                    If Not IsValidRepeatValue(stv) Then
                        .sInv = .sInv + 1
                        AddLog CAT_INV, locTxt, Ar("side_src"), shName, Ar("d_inv") & stv
                    End If
                    If .kind(f) <> KIND_CONFLICT Then
                        If Not IsValidRepeatValue(stv) Or Not IsValidRepeatValue(.st(f)) Then
                            .kind(f) = KIND_CONFLICT
                        ElseIf StrComp(stv, .st(f), vbBinaryCompare) <> 0 Then
                            .kind(f) = KIND_CONFLICT
                        End If
                    End If
                End If
            End If
        Next r
        .n = base
    End With
End Sub

' ---------------------------------------------------------------------
' Phase 2: one target sheet (duplicates first, then exact matching)
' ---------------------------------------------------------------------
Private Sub MatchTargetSheet(ByVal wsT As Worksheet, ByVal idx As Long, _
                             ByRef outVals() As Variant, ByRef hasData As Boolean)
    Dim lastT As Long, tData As Variant, i As Long, cnt As Long, rowNo As Long
    Dim rawName As String, tName As String, tDob As String, k As String
    Dim keys() As String, locs() As String, f As Long, first As Long
    Dim seen As Object, logged As Object

    hasData = False
    lastT = LastUsedRow(wsT, NAME_COL)
    If lastT < FIRST_DATA_ROW Then Exit Sub
    hasData = True

    tData = wsT.Range(wsT.Cells(FIRST_DATA_ROW, 1), wsT.Cells(lastT, REPEAT_COL)).Value
    cnt = UBound(tData, 1)
    ReDim outVals(1 To cnt)
    ReDim keys(1 To cnt)
    ReDim locs(1 To cnt)
    Set seen = CreateObject("Scripting.Dictionary")
    Set logged = CreateObject("Scripting.Dictionary")

    ' ---- Pass A: normalize, detect duplicates inside the target sheet ----
    For i = 1 To cnt
        rowNo = FIRST_DATA_ROW + i - 1
        rawName = Trim$(CStrSafe(tData(i, NAME_COL)))
        tName = ""
        If Len(rawName) > 0 Then tName = NormalizeArabicName(rawName)
        tDob = NormalizeDate(tData(i, DOB_COL))
        locs(i) = LocText(mSets(idx).code, CStrSafe(tData(i, ID_COL)), rowNo)

        If Len(tName) = 0 Or Len(tDob) = 0 Then
            If Len(rawName) > 0 Or Len(Trim$(CStrSafe(tData(i, DOB_COL)))) > 0 Then
                mSets(idx).tIncomplete = mSets(idx).tIncomplete + 1
                AddLog CAT_INC, locs(i), Ar("side_tgt"), "", Ar("d_inc")
            End If
        Else
            mSets(idx).tTotal = mSets(idx).tTotal + 1
            k = tDob & "|" & tName
            keys(i) = k
            If seen.Exists(k) Then
                first = CLng(seen.Item(k))
                If Not logged.Exists(k) Then
                    logged.Add k, True
                    mSets(idx).tDup = mSets(idx).tDup + 1
                    AddLog CAT_DUPTGT, locs(first), Ar("side_tgt"), "", Ar("d_dup")
                End If
                AddLog CAT_DUPTGT, locs(i), Ar("side_tgt"), "", Ar("d_dup")
            Else
                seen.Add k, i
            End If
        End If
    Next i

    ' ---- Pass B: exact matching against the branch set ----
    With mSets(idx)
        For i = 1 To cnt
            If Len(keys(i)) > 0 Then
                If .dict.Exists(keys(i)) Then
                    f = CLng(.dict.Item(keys(i)))
                    .hit(f) = True
                    Select Case .kind(f)
                        Case KIND_OK
                            outVals(i) = .st(f)
                            .tWritten = .tWritten + 1
                        Case KIND_INVALID
                            .tInv = .tInv + 1
                            AddLog CAT_INV, locs(i), Ar("side_tgt"), "", Ar("d_invt")
                        Case Else
                            .tAmb = .tAmb + 1
                            AddLog CAT_AMB, locs(i), Ar("side_tgt"), "", Ar("d_amb")
                    End Select
                Else
                    .tNotSrc = .tNotSrc + 1
                    AddLog CAT_NOTSRC, locs(i), Ar("side_tgt"), "", Ar("d_notsrc")
                End If
            End If
        Next i
    End With
End Sub

Private Sub ComputeTotals(ByVal setCount As Long)
    Dim s As Long, r As Long
    mTotWrite = 0: mTotTgt = 0: mTotNotSrc = 0: mTotDupTgt = 0: mTotAmb = 0
    mTotInv = 0: mTotIncT = 0: mTotSrcUnique = 0: mTotDupSrc = 0: mTotIncS = 0
    mTotSrcNotTgt = 0: mTotInvSrc = 0
    For s = 1 To setCount
        With mSets(s)
            .sUnique = 0: .sDupIds = 0: .sNotTgt = 0
            For r = 1 To .n
                If .dupCnt(r) > 0 Then
                    .sUnique = .sUnique + 1
                    If .dupCnt(r) > 1 Then .sDupIds = .sDupIds + 1
                    If .targetFound And Not .hit(r) Then .sNotTgt = .sNotTgt + 1
                End If
            Next r
            mTotWrite = mTotWrite + .tWritten
            mTotTgt = mTotTgt + .tTotal
            mTotNotSrc = mTotNotSrc + .tNotSrc
            mTotDupTgt = mTotDupTgt + .tDup
            mTotAmb = mTotAmb + .tAmb
            mTotInv = mTotInv + .tInv
            mTotIncT = mTotIncT + .tIncomplete
            mTotSrcUnique = mTotSrcUnique + .sUnique
            mTotDupSrc = mTotDupSrc + .sDupIds
            mTotIncS = mTotIncS + .sIncomplete
            mTotSrcNotTgt = mTotSrcNotTgt + .sNotTgt
            mTotInvSrc = mTotInvSrc + .sInv
        End With
    Next s
End Sub

' ---------------------------------------------------------------------
' Log + report
' ---------------------------------------------------------------------
Private Sub AddLog(ByVal cat As Long, ByVal loc As String, ByVal side As String, _
                   ByVal sh As String, ByVal det As String)
    If mLogN = 0 Then
        ReDim mLogCat(1 To 256)
        ReDim mLogLoc(1 To 256)
        ReDim mLogSide(1 To 256)
        ReDim mLogSheet(1 To 256)
        ReDim mLogDet(1 To 256)
    ElseIf mLogN >= UBound(mLogCat) Then
        ReDim Preserve mLogCat(1 To mLogN * 2)
        ReDim Preserve mLogLoc(1 To mLogN * 2)
        ReDim Preserve mLogSide(1 To mLogN * 2)
        ReDim Preserve mLogSheet(1 To mLogN * 2)
        ReDim Preserve mLogDet(1 To mLogN * 2)
    End If
    mLogN = mLogN + 1
    mLogCat(mLogN) = cat
    mLogLoc(mLogN) = loc
    mLogSide(mLogN) = side
    mLogSheet(mLogN) = sh
    mLogDet(mLogN) = det
End Sub

Private Function CatName(ByVal cat As Long) As String
    Select Case cat
        Case CAT_NOTSRC: CatName = Ar("c_notsrc")
        Case CAT_DUPSRC: CatName = Ar("c_dupsrc")
        Case CAT_DUPTGT: CatName = Ar("c_duptgt")
        Case CAT_AMB:    CatName = Ar("c_amb")
        Case CAT_INV:    CatName = Ar("c_inv")
        Case CAT_INC:    CatName = Ar("c_inc")
        Case CAT_SKIP:   CatName = Ar("c_skip")
    End Select
End Function

Private Sub KV(ByRef arr() As Variant, ByRef k As Long, ByVal a As String, ByVal b As Variant)
    k = k + 1
    arr(k, 1) = a
    arr(k, 2) = b
End Sub

' Creates a separate report workbook next to the target file.
' Returns its path, or "" when it could not be created.
Private Function BuildReport(ByVal folder As String, ByVal stamp As String, _
        ByVal sourcePath As String, ByVal targetPath As String, ByVal backupNote As String, _
        ByVal wrote As Boolean, ByVal incSheets As Long, ByVal skipSheets As Long, _
        ByVal ignSheets As Long, ByVal skipTgt As Long, ByVal groupCount As Long, _
        ByVal setCount As Long) As String
    Dim wb As Workbook, wsS As Worksheet, wsL As Worksheet
    Dim sm() As Variant, br() As Variant, lg() As Variant
    Dim k As Long, s As Long, c As Long, rw As Long, nRows As Long
    Dim path As String

    On Error GoTo Fail
    Set wb = Workbooks.Add
    Do While wb.Worksheets.Count > 2
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    Do While wb.Worksheets.Count < 2
        wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
    Loop
    Set wsS = wb.Worksheets(1)
    Set wsL = wb.Worksheets(2)
    wsS.Name = Ar("sh_sum")
    wsL.Name = Ar("sh_log")

    ' ---------- summary ----------
    ReDim sm(1 To 30, 1 To 2)
    k = 0
    KV sm, k, Ar("r_source"), sourcePath
    KV sm, k, Ar("r_target"), targetPath
    If wrote Then
        KV sm, k, Ar("r_backup"), backupNote
    Else
        KV sm, k, Ar("no_change"), ""
    End If
    k = k + 1
    KV sm, k, Ar("src_sheets"), incSheets
    KV sm, k, Ar("src_skipped"), skipSheets
    KV sm, k, Ar("src_ignored"), ignSheets
    KV sm, k, Ar("src_groups"), groupCount
    KV sm, k, Ar("tgt_skipped"), skipTgt
    k = k + 1
    KV sm, k, Ar("r_tgt"), mTotTgt
    KV sm, k, Ar("r_written"), mTotWrite
    KV sm, k, Ar("r_notsrc"), mTotNotSrc
    KV sm, k, Ar("r_duptgt"), mTotDupTgt
    KV sm, k, Ar("r_amb"), mTotAmb
    KV sm, k, Ar("r_inv"), mTotInv
    KV sm, k, Ar("r_inct"), mTotIncT
    k = k + 1
    KV sm, k, Ar("r_srcuniq"), mTotSrcUnique
    KV sm, k, Ar("r_dupsrc"), mTotDupSrc
    KV sm, k, Ar("r_invsrc"), mTotInvSrc
    KV sm, k, Ar("r_incs"), mTotIncS
    KV sm, k, Ar("r_srcnot"), mTotSrcNotTgt
    wsS.Range("A1").Resize(k, 2).Value = sm

    ' per-branch table
    wsS.Cells(k + 2, 1).Resize(1, 12).Value = Array(Ar("h_code"), Ar("h_tgt"), Ar("h_wr"), _
        Ar("h_ns"), Ar("h_dt"), Ar("h_am"), Ar("h_inv"), Ar("h_it"), Ar("h_su"), _
        Ar("h_ds"), Ar("h_is"), Ar("h_sn"))
    wsS.Cells(k + 2, 1).Resize(1, 12).Font.Bold = True
    If setCount > 0 Then
        ReDim br(1 To setCount, 1 To 12)
        For s = 1 To setCount
            With mSets(s)
                br(s, 1) = .code
                br(s, 2) = .tTotal
                br(s, 3) = .tWritten
                br(s, 4) = .tNotSrc
                br(s, 5) = .tDup
                br(s, 6) = .tAmb
                br(s, 7) = .tInv
                br(s, 8) = .tIncomplete
                br(s, 9) = .sUnique
                br(s, 10) = .sDupIds
                br(s, 11) = .sIncomplete
                br(s, 12) = .sNotTgt
            End With
        Next s
        wsS.Cells(k + 3, 1).Resize(setCount, 12).Value = br
    End If
    wsS.Range("A1:A" & k).Font.Bold = True
    FormatSummarySheet wsS, k + 1, k + 2     ' k + 2 = header row of the branch table (row 25)

    ' ---------- details ----------
    wsL.Columns(3).NumberFormat = "@"
    wsL.Range("A1:F1").Value = Array(Ar("h_n"), Ar("h_cat"), Ar("h_loc"), Ar("h_side"), Ar("h_sheet"), Ar("h_det"))
    wsL.Range("A1:F1").Font.Bold = True
    If mLogN = 0 Then
        wsL.Range("B2").Value = Ar("d_none")
    Else
        ReDim lg(1 To mLogN, 1 To 6)
        rw = 0
        For c = 1 To CAT_COUNT
            For s = 1 To mLogN
                If mLogCat(s) = c Then
                    rw = rw + 1
                    lg(rw, 1) = rw
                    lg(rw, 2) = CatName(c)
                    lg(rw, 3) = mLogLoc(s)
                    lg(rw, 4) = mLogSide(s)
                    lg(rw, 5) = mLogSheet(s)
                    lg(rw, 6) = mLogDet(s)
                End If
            Next s
        Next c
        wsL.Range("A2").Resize(mLogN, 6).Value = lg
        wsL.Range("A1:F1").AutoFilter
    End If
    wsL.Columns(1).ColumnWidth = 7
    wsL.Columns(2).ColumnWidth = 34
    wsL.Columns(3).ColumnWidth = 14
    wsL.Columns(4).ColumnWidth = 12
    wsL.Columns(5).ColumnWidth = 24
    wsL.Columns(6).ColumnWidth = 70

    On Error Resume Next
    wsL.Activate
    wsL.DisplayRightToLeft = True
    wsS.Activate
    wsS.DisplayRightToLeft = True
    On Error GoTo Fail

    path = folder & "RepeatStatus_Report_" & stamp & ".xlsx"
    wb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    BuildReport = path
    Exit Function

Fail:
    BuildReport = ""
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

' ---------------------------------------------------------------------
' Summary sheet formatting (applied to this sheet only)
' ---------------------------------------------------------------------
Private Sub FormatSummarySheet(ByVal ws As Worksheet, ByVal mergeLastRow As Long, _
                               ByVal rotateRow As Long)
    Const SCRATCH_COL As Long = 200
    Const SUM_LAST_COL As Long = 12          ' summary sheet uses A:L
    Dim r As Long, c As Long, w As Double
    Dim needL As Double, needR As Double, hdrW As Double
    Dim prevAlerts As Boolean

    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    ' 1) move B -> E BEFORE merging, so nothing is lost
    For r = 1 To mergeLastRow
        If Not IsEmpty(ws.Cells(r, 2).Value) Then
            ws.Cells(r, 5).Value = ws.Cells(r, 2).Value
            ws.Cells(r, 2).ClearContents
        End If
    Next r

    ' 2) merge A:D and E:H
    For r = 1 To mergeLastRow
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Merge
        ws.Range(ws.Cells(r, 5), ws.Cells(r, 8)).Merge
    Next r

    ' 3) general alignment
    With ws.Cells
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 4) wrap text + rotate up on the header row
    With ws.Rows(rotateRow)
        .WrapText = True
        .Orientation = xlUpward
    End With

    ' 5) measure the text of the merged areas and of the rotated headers
    For r = 1 To mergeLastRow
        w = MeasureColWidth(ws, ws.Cells(r, 1).Value, True, SCRATCH_COL)
        If w > needL Then needL = w
        w = MeasureColWidth(ws, ws.Cells(r, 5).Value, False, SCRATCH_COL)
        If w > needR Then needR = w
    Next r
    For c = 1 To SUM_LAST_COL
        w = MeasureColWidth(ws, ws.Cells(rotateRow, c).Value, True, SCRATCH_COL)
        If w > hdrW Then hdrW = w
    Next c
    ws.Columns(SCRATCH_COL).Delete

    ' 6) AutoFit: height of the rotated row first, then the columns
    '    (AutoFit ignores merged cells, so they are widened afterwards)
    If hdrW > 0 Then
        w = hdrW * 5.4 + 8
        If w > 409 Then w = 409
        ws.Rows(rotateRow).RowHeight = w
    End If
    ws.Range(ws.Columns(1), ws.Columns(SUM_LAST_COL)).AutoFit

    GrowRangeWidth ws, 1, 4, needL + 2
    GrowRangeWidth ws, 5, 8, needR + 2

    Application.DisplayAlerts = prevAlerts
End Sub

' Width (in column units) needed by a text on one line; uses a scratch column
Private Function MeasureColWidth(ByVal ws As Worksheet, ByVal v As Variant, _
                                 ByVal isBold As Boolean, ByVal scratchCol As Long) As Double
    Dim s As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    s = CStr(v)
    If Len(s) = 0 Then Exit Function
    With ws.Cells(1, scratchCol)
        .NumberFormat = "@"
        .WrapText = False
        .Orientation = xlHorizontal
        .Font.Bold = isBold
        .Value = s
    End With
    ws.Columns(scratchCol).AutoFit
    MeasureColWidth = ws.Columns(scratchCol).ColumnWidth
    ws.Cells(1, scratchCol).ClearContents
End Function

' Widens columns c1..c2 equally when their total width is below needW
Private Sub GrowRangeWidth(ByVal ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long, _
                           ByVal needW As Double)
    Dim c As Long, have As Double, add As Double
    For c = c1 To c2
        have = have + ws.Columns(c).ColumnWidth
    Next c
    If needW <= have Then Exit Sub
    add = (needW - have) / (c2 - c1 + 1)
    For c = c1 To c2
        ws.Columns(c).ColumnWidth = ws.Columns(c).ColumnWidth + add
    Next c
End Sub

' "SE: 41" (number in column A); falls back to the row number
Private Function LocText(ByVal code As String, ByVal aVal As String, ByVal rowNo As Long) As String
    aVal = Trim$(aVal)
    If Len(aVal) = 0 Then
        LocText = code & ": [" & Ar("row") & " " & rowNo & "]"
    Else
        LocText = code & ": " & aVal
    End If
End Function

' ---------------------------------------------------------------------
' Source-title helpers
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
    Dim p As String
    p = NormAr(ThirdSecondaryPrefix())
    StartsWithThirdSecondary = (StrComp(Left$(NormAr(titleText), Len(p)), p, vbTextCompare) = 0)
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
' Normalization helpers
' ---------------------------------------------------------------------
Private Function NormalizeArabicName(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String

    s = Trim$(Replace(s, ChrW(&H640), ""))        ' tatweel

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536

        If (code >= &H64B And code <= &H65F) Or code = &H670 Then
            ' tashkeel: skip
        ElseIf (code >= &H200B And code <= &H200F) Or (code >= &H202A And code <= &H202E) _
               Or code = &HFEFF& Or code = &H61C Then
            ' invisible direction / zero-width marks: skip
        Else
            Select Case code
                Case &H623, &H625, &H622, &H671
                    ch = ChrW(&H627)              ' alef variants -> alef
                Case &H629
                    ch = ChrW(&H647)              ' taa marbuta -> haa
                Case &H649, &H626, &H6CC
                    ch = ChrW(&H64A)              ' alef maqsura / yaa-hamza / farsi yeh -> yaa
                Case &H624
                    ch = ChrW(&H648)              ' waw-hamza -> waw
                Case &H6A9
                    ch = ChrW(&H643)              ' farsi kaf -> kaf
                Case 9, 10, 13, 160
                    ch = " "
            End Select
            out = out & ch
        End If
    Next i

    out = LCase$(out)
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    NormalizeArabicName = Trim$(out)
End Function

Private Function ToWesternDigits(ByVal s As String) As String
    Dim i As Long, code As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        If code >= &H660 And code <= &H669 Then
            ch = Chr$(48 + code - &H660)
        ElseIf code >= &H6F0 And code <= &H6F9 Then
            ch = Chr$(48 + code - &H6F0)
        End If
        out = out & ch
    Next i
    ToWesternDigits = out
End Function

' Returns yyyymmdd or "" (real dates, serial numbers and d/m/yyyy text)
Private Function NormalizeDate(ByVal v As Variant) As String
    Dim s As String, parts() As String, p As Long
    Dim d As Long, m As Long, y As Long

    On Error GoTo BadDate
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    Select Case VarType(v)
        Case vbDate
            NormalizeDate = Format$(v, "yyyymmdd")
            Exit Function
        Case vbDouble, vbSingle, vbCurrency, vbLong, vbInteger
            If v >= 1 And v < 100000 Then NormalizeDate = Format$(CDate(CDbl(v)), "yyyymmdd")
            Exit Function
    End Select

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    s = ToWesternDigits(s)
    p = InStr(s, " ")
    If p > 0 Then s = Left$(s, p - 1)           ' drop a time part
    s = Replace(s, "/", "-")
    s = Replace(s, ".", "-")
    s = Replace(s, "\", "-")

    parts = Split(s, "-")
    If UBound(parts) = 2 Then
        If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
            If Len(parts(0)) = 4 Then
                y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
            Else
                d = CLng(parts(0)): m = CLng(parts(1)): y = CLng(parts(2))
                If y < 100 Then
                    If y > 30 Then y = 1900 + y Else y = 2000 + y
                End If
            End If
            If y >= 1900 And y <= 2100 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
                If Day(DateSerial(y, m, d)) = d Then
                    NormalizeDate = Format$(DateSerial(y, m, d), "yyyymmdd")
                End If
            End If
            Exit Function
        End If
    End If
    If IsDate(s) Then NormalizeDate = Format$(CDate(s), "yyyymmdd")
    Exit Function
BadDate:
    NormalizeDate = ""
End Function

Private Function CleanStatus(ByVal s As String) As String
    CleanStatus = Trim$(Replace(s, ChrW(160), " "))
End Function

Private Function IsValidRepeatValue(ByVal v As Variant) As Boolean
    Dim s As String
    s = CleanStatus(CStrSafe(v))
    IsValidRepeatValue = (s = mYes Or s = mNo)
End Function

' ---------------------------------------------------------------------
' General helpers
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

' The only question - the last extraction file is offered as the default
Private Function AskTargetFile() As String
    Dim lastExtract As String, answer As VbMsgBoxResult, startPath As String

#If HAS_BACETAB Then
    On Error Resume Next
    lastExtract = modBacEtab.gLastBacEtabPath
    On Error GoTo 0
#End If

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
    If Len(txt) > 1000 Then txt = Left$(txt, 1000)
    MsgBox txt, icon + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("title")
End Sub

' ---------------------------------------------------------------------
' Arabic text (ASCII-safe: no Arabic literals in the source file)
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
    Dim s As String
    Select Case key
        Case "title"
            s = ArHex("62C 644 628 20 62D 627 644 627 62A 20 627 644 625 639 627 62F 629")
        Case "sb_read"
            s = ArHex("62C 627 631 64D 20 642 631 627 621 629 20 648 631 642 629 20 627 644 645 635 62F 631 3A 20")
        Case "sb_match"
            s = ArHex("62C 627 631 64D 20 645 637 627 628 642 629 20 648 631 642 629 3A 20")
        Case "sb_report"
            s = ArHex("62C 627 631 64D 20 625 646 634 627 621 20 627 644 62A 642 631 64A 631 2E 2E 2E")
        Case "q2_title"
            s = ArHex("645 644 641 20 627 644 644 635 642")
        Case "q2_head"
            s = ArHex("645 644 641 20 627 644 644 635 642 20 647 648 20 627 644 645 644 641 20 627 644 630 64A 20 633 62A 64F 643 62A 628 20 641 64A 647 20 62D 627 644 627 62A 20 627 644 625 639")
            s = s & ArHex("627 62F 629 20 641 64A 20 627 644 639 645 648 62F 20 45 2E")
        Case "q2_last"
            s = ArHex("622 62E 631 20 645 644 641 20 627 633 62A 62E 631 627 62C 20 647 648 3A")
        Case "q2_yes"
            s = ArHex("646 639 645 3A 20 627 633 62A 62E 62F 627 645 20 647 630 627 20 627 644 645 644 641")
        Case "q2_no"
            s = ArHex("644 627 3A 20 627 62E 62A 64A 627 631 20 645 644 641 20 622 62E 631")
        Case "q2_cancel"
            s = ArHex("625 644 63A 627 621 3A 20 625 64A 642 627 641 20 627 644 639 645 644 64A 629")
        Case "q2_pick"
            s = ArHex("627 62E 62A 631 20 627 644 622 646 20 645 644 641 20 627 644 644 635 642 60C 20 648 647 648 20 639 627 62F 629 20 645 644 641 20 627 644 627 633 62A 62E 631 627 62C 20 627")
            s = s & ArHex("644 646 627 62A 62C 20 639 646 20 632 631 20 AB 627 633 62A 62E 631 627 62C 20 646 62A 627 626 62C 20 627 644 645 624 633 633 629 BB 2E")
        Case "q2_dialog"
            s = ArHex("627 62E 62A 631 20 645 644 641 20 627 644 644 635 642 20 28 627 644 630 64A 20 633 62A 643 62A 628 20 641 64A 647 20 62D 627 644 627 62A 20 627 644 625 639 627 62F 629 29")
        Case "target_ro"
            s = ArHex("645 644 641 20 627 644 644 635 642 20 645 641 62A 648 62D 20 644 644 642 631 627 621 629 20 641 642 637 2E 20 623 63A 644 642 647 20 62B 645 20 623 639 62F 20 627 644 62A")
            s = s & ArHex("634 63A 64A 644 2E")
        Case "no_group"
            s = ArHex("644 645 20 64A 64F 639 62B 631 20 641 64A 20 627 644 645 635 62F 631 20 639 644 649 20 623 64A 20 648 631 642 629 20 635 627 644 62D 629 20 28 62A 628 62F 623 20 627 644")
            s = s & ArHex("62E 644 64A 629 20 48 31 20 641 64A 647 627 20 628 639 628 627 631 629 3A")
        Case "no_group2"
            s = ArHex("29 2E 20 623 633 628 627 628 20 62A 62C 627 648 632 20 627 644 623 648 631 627 642 3A")
        Case "no_match_sheets"
            s = ArHex("645 644 641 20 627 644 644 635 642 20 644 627 20 64A 62D 62A 648 64A 20 623 64A 20 648 631 642 629 20 62A 642 627 628 644 20 627 644 634 639 628 20 627 644 645 648 62C 648")
            s = s & ArHex("62F 629 20 641 64A 20 627 644 645 635 62F 631 2E 20 62A 623 643 62F 20 645 646 20 627 62E 62A 64A 627 631 20 627 644 645 644 641 20 627 644 635 62D 64A 62D 2E 20 623 648")
            s = s & ArHex("631 627 642 20 645 644 641 20 627 644 644 635 642 3A")
        Case "stopped"
            s = ArHex("62A 648 642 641 62A 20 627 644 639 645 644 64A 629 20 628 633 628 628 20 62E 637 623 3A")
        Case "c_title"
            s = ArHex("62A 623 643 64A 62F 20 627 644 643 62A 627 628 629")
        Case "c_write"
            s = ArHex("627 643 62A 645 644 62A 20 627 644 645 637 627 628 642 629 20 648 644 645 20 64A 64F 639 62F 64E 651 644 20 623 64A 20 634 64A 621 20 628 639 62F 2E 20 639 62F 62F 20 627")
            s = s & ArHex("644 642 64A 645 20 627 644 62A 64A 20 633 62A 64F 643 62A 628 20 641 64A 20 627 644 639 645 648 62F 20 45 3A 20")
        Case "c_yes"
            s = ArHex("646 639 645 3A 20 643 62A 627 628 629 20 627 644 642 64A 645 20 648 62D 641 638 20 627 644 645 644 641 20")
            s = s & ArHex("28 645 639 20 646 633 62E 629 20 627 62D 62A 64A 627 637 64A 629 20 645 624 642 62A 629 20 62A 64F 62D 630")
            s = s & ArHex("641 20 628 639 62F 20 646 62C 627 62D 20 627 644 62D 641 638 29")
        Case "bk_deleted"
            s = ArHex("62D 64F 630 641 62A 20 646 647 627 626 64A 627 64B 20 628 639 62F 20 627 644 62A 623 643 62F 20 645 646 20")
            s = s & ArHex("646 62C 627 62D 20 627 644 62D 641 638 2E")
        Case "bk_delfail"
            s = ArHex("62A 639 630 631 20 62D 630 641 20 627 644 646 633 62E 629 20 627 644 627 62D 62A 64A 627 637 64A 629 60C 20")
            s = s & ArHex("627 62D 630 641 647 627 20 64A 62F 648 64A 627 64B 3A")
        Case "bk_keep"
            s = ArHex("627 644 646 633 62E 629 20 627 644 627 62D 62A 64A 627 637 64A 629 20 645 62D 641 648 638 629 20 644 627 633")
            s = s & ArHex("62A 639 627 62F 629 20 627 644 628 64A 627 646 627 62A 3A")
        Case "save_unver"
            s = ArHex("644 645 20 64A 62A 623 643 62F 20 646 62C 627 62D 20 62D 641 638 20 645 644 641 20 627 644 644 635 642 60C")
            s = s & ArHex("20 644 630 644 643 20 623 64F 628 642 64A 62A 20 627 644 646 633 62E 629 20 627 644 627 62D 62A 64A 627 637")
            s = s & ArHex("64A 629 2E")
        Case "c_no"
            s = ArHex("644 627 3A 20 639 62F 645 20 62A 639 62F 64A 644 20 627 644 645 644 641 60C 20 645 639 20 625 646 634 627 621 20 627 644 62A 642 631 64A 631 20 641 642 637")
        Case "done"
            s = ArHex("627 643 62A 645 644 62A 20 627 644 645 637 627 628 642 629 20 648 62D 64F 641 638 62A 20 627 644 642 64A 645 20 641 64A 20 627 644 639 645 648 62F 20 45 2E")
        Case "no_change"
            s = ArHex("644 645 20 64A 62A 645 20 62A 639 62F 64A 644 20 645 644 641 20 627 644 644 635 642 20 28 644 645 20 62A 64F 643 62A 628 20 623 64A 20 642 64A 645 629 20 648 644 645 20")
            s = s & ArHex("62A 64F 646 634 623 20 646 633 62E 629 20 627 62D 62A 64A 627 637 64A 629 29 2E")
        Case "r_written"
            s = ArHex("642 64A 645 20 627 644 625 639 627 62F 629 20 627 644 645 637 627 628 64E 642 629 3A 20")
        Case "r_notsrc"
            s = ArHex("62A 644 627 645 64A 630 20 63A 64A 631 20 645 648 62C 648 62F 64A 646 20 641 64A 20 627 644 645 635 62F 631 3A 20")
        Case "r_dupsrc"
            s = ArHex("62A 643 631 627 631 627 62A 20 641 64A 20 627 644 645 635 62F 631 20 28 639 62F 62F 20 627 644 62A 644 627 645 64A 630 29 3A 20")
        Case "r_duptgt"
            s = ArHex("62A 643 631 627 631 627 62A 20 641 64A 20 645 644 641 20 627 644 644 635 642 20 28 639 62F 62F 20 627 644 62A 644 627 645 64A 630 29 3A 20")
        Case "r_amb"
            s = ArHex("645 644 62A 628 633 629 20 28 642 64A 645 20 645 62A 636 627 631 628 629 60C 20 644 645 20 62A 64F 643 62A 628 29 3A 20")
        Case "r_inv"
            s = ArHex("644 645 20 62A 64F 643 62A 628 20 644 639 62F 645 20 635 644 627 62D 64A 629 20 627 644 642 64A 645 629 20 641 64A 20 627 644 645 635 62F 631 3A 20")
        Case "r_inct"
            s = ArHex("628 64A 627 646 627 62A 20 646 627 642 635 629 20 641 64A 20 645 644 641 20 627 644 644 635 642 3A 20")
        Case "r_incs"
            s = ArHex("628 64A 627 646 627 62A 20 646 627 642 635 629 20 641 64A 20 627 644 645 635 62F 631 3A 20")
        Case "r_invsrc"
            s = ArHex("633 62C 644 627 62A 20 628 642 64A 645 629 20 625 639 627 62F 629 20 63A 64A 631 20 635 627 644 62D 629 20 641 64A 20 627 644 645 635 62F 631 3A 20")
        Case "r_srcuniq"
            s = ArHex("62A 644 627 645 64A 630 20 627 644 645 635 62F 631 20 28 647 648 64A 627 62A 20 641 631 64A 62F 629 29 3A 20")
        Case "r_srcnot"
            s = ArHex("641 64A 20 627 644 645 635 62F 631 20 648 63A 64A 631 20 645 648 62C 648 62F 64A 646 20 641 64A 20 645 644 641 20 627 644 644 635 642 20 28 644 644 639 644 645 29 3A 20")
        Case "r_tgt"
            s = ArHex("62A 644 627 645 64A 630 20 645 644 641 20 627 644 644 635 642 20 28 628 64A 627 646 627 62A 20 643 627 645 644 629 29 3A 20")
        Case "src_sheets"
            s = ArHex("623 648 631 627 642 20 627 644 645 635 62F 631 20 627 644 645 639 62A 645 62F 629 3A 20")
        Case "src_skipped"
            s = ArHex("623 648 631 627 642 20 627 644 645 635 62F 631 20 627 644 645 62A 62C 627 648 632 629 3A 20")
        Case "src_ignored"
            s = ArHex("623 648 631 627 642 20 627 644 645 635 62F 631 20 63A 64A 631 20 627 644 645 639 646 64A 629 3A 20")
        Case "src_groups"
            s = ArHex("627 644 634 639 628 20 627 644 645 62F 645 62C 629 3A 20")
        Case "tgt_skipped"
            s = ArHex("623 648 631 627 642 20 645 644 641 20 627 644 644 635 642 20 627 644 645 62A 62C 627 648 632 629 3A 20")
        Case "r_source"
            s = ArHex("645 62C 644 62F 20 627 644 645 635 62F 631 20 28 627 644 645 644 641 627 62A 20 627 644 62E 627 645 29 3A 20")
        Case "r_target"
            s = ArHex("645 644 641 20 627 644 644 635 642 3A 20")
        Case "r_backup"
            s = ArHex("646 633 62E 629 20 627 62D 62A 64A 627 637 64A 629 20 642 628 644 20 627 644 62A 639 62F 64A 644 3A 20")
        Case "r_report"
            s = ArHex("645 644 641 20 627 644 62A 642 631 64A 631 3A 20")
        Case "r_reportfail"
            s = ArHex("62A 639 630 631 20 625 646 634 627 621 20 645 644 641 20 627 644 62A 642 631 64A 631 2E")
        Case "see_report"
            s = ArHex("627 644 62A 641 627 635 64A 644 20 648 627 644 645 648 627 642 639 20 641 64A 20 645 644 641 20 627 644 62A 642 631 64A 631 2E")
        Case "sh_sum"
            s = ArHex("645 644 62E 635")
        Case "sh_log"
            s = ArHex("627 644 62A 641 627 635 64A 644")
        Case "h_n"
            s = ArHex("631 642 645")
        Case "h_cat"
            s = ArHex("627 644 641 626 629")
        Case "h_loc"
            s = ArHex("627 644 645 648 642 639")
        Case "h_side"
            s = ArHex("627 644 62C 647 629")
        Case "h_sheet"
            s = ArHex("648 631 642 629 20 627 644 645 635 62F 631 20 627 644 641 639 644 64A 629")
        Case "h_det"
            s = ArHex("627 644 62A 641 627 635 64A 644")
        Case "h_code"
            s = ArHex("627 644 634 639 628 629 20 28 648 631 642 629 29")
        Case "h_tgt"
            s = ArHex("62A 644 627 645 64A 630 20 627 644 644 635 642")
        Case "h_wr"
            s = ArHex("642 64A 645 20 645 637 627 628 64E 642 629")
        Case "h_ns"
            s = ArHex("63A 64A 631 20 645 648 62C 648 62F 20 641 64A 20 627 644 645 635 62F 631")
        Case "h_dt"
            s = ArHex("62A 643 631 627 631 20 641 64A 20 627 644 644 635 642")
        Case "h_am"
            s = ArHex("645 644 62A 628 633")
        Case "h_inv"
            s = ArHex("642 64A 645 629 20 63A 64A 631 20 635 627 644 62D 629")
        Case "h_it"
            s = ArHex("646 627 642 635 20 28 627 644 644 635 642 29")
        Case "h_su"
            s = ArHex("62A 644 627 645 64A 630 20 627 644 645 635 62F 631")
        Case "h_ds"
            s = ArHex("62A 643 631 627 631 20 641 64A 20 627 644 645 635 62F 631")
        Case "h_is"
            s = ArHex("646 627 642 635 20 28 627 644 645 635 62F 631 29")
        Case "h_sn"
            s = ArHex("641 64A 20 627 644 645 635 62F 631 20 648 644 64A 633 20 641 64A 20 627 644 644 635 642")
        Case "side_src"
            s = ArHex("627 644 645 635 62F 631")
        Case "side_tgt"
            s = ArHex("645 644 641 20 627 644 644 635 642")
        Case "row"
            s = ArHex("635 641")
        Case "c_notsrc"
            s = ArHex("63A 64A 631 20 645 648 62C 648 62F 20 641 64A 20 627 644 645 635 62F 631")
        Case "c_dupsrc"
            s = ArHex("62A 643 631 627 631 20 641 64A 20 627 644 645 635 62F 631")
        Case "c_duptgt"
            s = ArHex("62A 643 631 627 631 20 641 64A 20 645 644 641 20 627 644 644 635 642")
        Case "c_amb"
            s = ArHex("645 644 62A 628 633 3A 20 642 64A 645 20 645 62A 636 627 631 628 629 20 28 644 645 20 64A 64F 643 62A 628 29")
        Case "c_inv"
            s = ArHex("642 64A 645 629 20 625 639 627 62F 629 20 63A 64A 631 20 635 627 644 62D 629")
        Case "c_inc"
            s = ArHex("628 64A 627 646 627 62A 20 646 627 642 635 629")
        Case "c_skip"
            s = ArHex("648 631 642 629 20 645 62A 62C 627 648 632 629")
        Case "d_notsrc"
            s = ArHex("644 627 20 64A 648 62C 62F 20 641 64A 20 648 631 642 629 20 627 644 645 635 62F 631 20 62A 644 645 64A 630 20 628 646 641 633 20 627 644 627 633 645 20 648 62A 627 631 64A")
            s = s & ArHex("62E 20 627 644 645 64A 644 627 62F")
        Case "d_dup"
            s = ArHex("62A 644 645 64A 630 20 645 643 631 631 20 628 646 641 633 20 627 644 627 633 645 20 648 627 644 644 642 628 20 648 62A 627 631 64A 62E 20 627 644 645 64A 644 627 62F")
        Case "d_amb"
            s = ArHex("644 645 20 62A 64F 643 62A 628 20 627 644 642 64A 645 629 3A 20 633 62C 644 627 62A 20 645 643 631 631 629 20 641 64A 20 627 644 645 635 62F 631 20 628 642 64A 645 20 625")
            s = s & ArHex("639 627 62F 629 20 645 62A 636 627 631 628 629 20 623 648 20 63A 64A 631 20 635 627 644 62D 629")
        Case "d_inv"
            s = ArHex("627 644 639 645 648 62F 20 45 20 644 64A 633 20 AB 646 639 645 BB 20 648 644 627 20 AB 644 627 BB 2E 20 627 644 642 64A 645 629 3A 20")
        Case "d_invt"
            s = ArHex("644 645 20 62A 64F 643 62A 628 20 627 644 642 64A 645 629 20 644 623 646 20 642 64A 645 629 20 627 644 625 639 627 62F 629 20 641 64A 20 627 644 645 635 62F 631 20 63A 64A")
            s = s & ArHex("631 20 635 627 644 62D 629")
        Case "d_inc"
            s = ArHex("627 633 645 20 623 648 20 62A 627 631 64A 62E 20 645 64A 644 627 62F 20 646 627 642 635 20 623 648 20 63A 64A 631 20 645 642 631 648 621")
        Case "d_none"
            s = ArHex("644 627 20 62A 648 62C 62F 20 62D 627 644 627 62A 20 644 644 625 628 644 627 63A")
        Case "sk_nocode"
            s = ArHex("62A 639 630 631 20 62A 62D 62F 64A 62F 20 631 645 632 20 627 644 634 639 628 629 20 645 646 20 627 644 639 646 648 627 646 3A 20")
        Case "sk_dupcode"
            s = ArHex("631 645 632 20 627 644 634 639 628 629 20 645 648 62C 648 62F 20 645 633 628 642 627 64B 20 628 639 646 648 627 646 20 645 62E 62A 644 641 3A 20")
        Case "sk_nodata"
            s = ArHex("644 627 20 62A 648 62C 62F 20 628 64A 627 646 627 62A 20 627 628 62A 62F 627 621 64B 20 645 646 20 627 644 635 641 20 33")
        Case "sk_noname"
            s = ArHex("644 627 20 62A 648 62C 62F 20 623 633 645 627 621 20 641 64A 20 627 644 639 645 648 62F 20 42")
        Case "sk_baddob"
            s = ArHex("627 644 639 645 648 62F 20 43 20 644 627 20 64A 62D 62A 648 64A 20 62A 648 627 631 64A 62E 20 645 64A 644 627 62F 20 645 642 631 648 621 629 20 641 64A 20 623 63A 644 628")
            s = s & ArHex("20 627 644 635 641 648 641")
        Case "sk_badrep"
            s = ArHex("627 644 639 645 648 62F 20 45 20 644 627 20 64A 62D 62A 648 64A 20 AB 646 639 645 BB 20 623 648 20 AB 644 627 BB 20 641 64A 20 623 63A 644 628 20 627 644 635 641 648")
            s = s & ArHex("641")
        Case "sk_tgtnone"
            s = ArHex("648 631 642 629 20 644 627 20 64A 642 627 628 644 647 627 20 645 635 62F 631")
        Case "sk_tgtempty"
            s = ArHex("648 631 642 629 20 645 644 641 20 627 644 644 635 642 20 644 627 20 62A 62D 62A 648 64A 20 628 64A 627 646 627 62A")
        Case "src_files"
            s = ArHex("645 644 641 627 62A 20 627 644 645 635 62F 631 20 627 644 62E 627 645 20 627 644 645 642 631 648 621 629 3A 20")
        Case "r_conv"
            s = ArHex("627 644 645 62D 648 644 629 20 625 644 649 20 78 6C 73 78 3A 20")
        Case "r_srcbad"
            s = ArHex("627 644 62A 64A 20 62A 639 630 631 62A 20 642 631 627 621 62A 647 627 3A 20")
        Case "no_src_folder"
            s = ArHex("645 62C 644 62F 20 627 644 645 635 62F 631 20 53 6F 75 72 63 65 30 33 20 63A 64A 631 20 645 648 62C 648 62F 20 628 62C 627 646 628 20 647 630 627 20 627 644 645 635 646 641 3A")
        Case "src_failed"
            s = ArHex("644 645 20 64A 639 62B 631 20 641 64A 20 645 62C 644 62F 20 53 6F 75 72 63 65 30 33 20 639 644 649 20 623 64A 20 645 644 641 20 635 627 644 62D 20 644 644 642 631 627 621 629 2E")
        Case "save_first"
            s = ArHex("627 62D 641 638 20 647 630 627 20 627 644 645 635 646 641 20 623 648 644 627 20 62B 645 20 623 639 62F 20 627 644 62A 634 63A 64A 644 2E")
        Case "url_path"
            s = ArHex("647 630 627 20 627 644 645 635 646 641 20 639 644 649 20 645 633 627 631 20 625 644 643 62A 631 648 646 64A 20 28 4F 6E 65 44 72 69 76 65 2F 53 68 61 72 65 50 6F 69 6E 74 29")
            s = s & ArHex("2E 20 627 646 642 644 647 20 625 644 649 20 645 62C 644 62F 20 645 62D 644 64A 20 62B 645 20 623 639 62F 20 627 644 62A 634 63A 64A 644 2E")
        Case "yes"
            s = ArHex("646 639 645")
        Case "no"
            s = ArHex("644 627")
    End Select
    Ar = s
End Function

' ==============================================================================
' SOURCE03 BUILD - Collector01 logic embedded, so this module needs no other module
' STEP 1: silent conversion of the \Source03\ files to .xlsx (web-exported .xls /
'         .xlsb / .xlsm; originals are moved to \Source03\_Original, never deleted)
' STEP 2: collection into ONE new workbook returned OPEN and UNSAVED (nothing is
'         written to Term01Results.xlsx). The caller closes it with SaveChanges:=False.
' Returns False when nothing usable was collected (reasons in skippedList).
' ==============================================================================
Private Function SrcBuildWorkbook(ByVal srcFolder As String, _
                                  ByRef wbOut As Workbook, _
                                  ByRef copied As Long, _
                                  ByRef skipped As Long, _
                                  ByRef skippedList As String, _
                                  ByRef convTotal As Long, _
                                  ByRef convSuccess As Long, _
                                  ByRef convFailed As Long) As Boolean
    Dim wb As Workbook, oldSec As Long, oldAlerts As Boolean, oldCalc As Long

    SrcBuildWorkbook = False
    Set wbOut = Nothing
    copied = 0: skipped = 0: skippedList = ""
    convTotal = 0: convSuccess = 0: convFailed = 0

    If LCase(Left(srcFolder, 4)) = "http" Then
        skippedList = "  - The workbook is on OneDrive/SharePoint (URL path). Move it to a local folder first." & vbCrLf
        Exit Function
    End If
    If Right(srcFolder, 1) <> "\" Then srcFolder = srcFolder & "\"
    If Dir(srcFolder, vbDirectory) = "" Then
        skippedList = "  - Source03 folder not found: " & srcFolder & vbCrLf
        Exit Function
    End If

    oldSec = Application.AutomationSecurity
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable   ' block Auto_Open in .xlsm/.xls
    On Error GoTo EH

    ' 1) Silent auto-conversion inside \Source03\ first
    ConvertSourceFilesToXLSX srcFolder, convTotal, convSuccess, convFailed

    ' 2) Collect all files into a new (unsaved) workbook
    Set wb = Workbooks.Add
    Do While wb.Sheets.Count > 1: wb.Sheets(wb.Sheets.Count).Delete: Loop

    SrcCollectAll wb, srcFolder, copied, skipped, skippedList

    If copied > 0 Then
        If wb.Sheets.Count > 1 Then
            If Left(wb.Sheets(1).Name, 5) = "Sheet" And wb.Sheets(1).UsedRange.Cells.Count = 1 Then wb.Sheets(1).Delete
        End If
        Set wbOut = wb
        SrcBuildWorkbook = True
    Else
        wb.Close SaveChanges:=False
    End If

Done:
    Application.AutomationSecurity = oldSec
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    Exit Function

EH:
    skippedList = skippedList & "  - Build failed: " & Err.Number & " : " & Err.Description & vbCrLf
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wbOut = Nothing
    SrcBuildWorkbook = False
    Resume Done
End Function

' ==============================================================================
' HELPERS
' ==============================================================================
' True only when the save really happened: Saved flag, file present, not empty, fresh timestamp
Private Function TargetSaveVerified(ByVal wb As Workbook, ByVal savedAfter As Date) As Boolean
    Dim fp As String
    On Error GoTo Bad
    If Not wb.Saved Then Exit Function
    fp = wb.FullName
    If LCase$(Left$(fp, 4)) = "http" Then Exit Function
    If Not FileExistsPath(fp) Then Exit Function
    If FileLen(fp) = 0 Then Exit Function
    If FileDateTime(fp) < DateAdd("s", -5, savedAfter) Then Exit Function
    TargetSaveVerified = True
    Exit Function
Bad:
    TargetSaveVerified = False
End Function

Private Function DeleteBackupFile(ByVal p As String) As Boolean
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    SetAttr p, vbNormal
    Kill p
    Err.Clear
    DeleteBackupFile = Not FileExistsPath(p)
    On Error GoTo 0
End Function

Private Function FileExistsPath(ByVal p As String) As Boolean
    On Error Resume Next
    FileExistsPath = ((GetAttr(p) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExistsPath = False: Err.Clear
End Function

Private Function SortedNames(ByVal col As Collection) As Variant
    Dim arr() As String, i As Long, j As Long, tmp As String
    If col.Count = 0 Then SortedNames = Array(): Exit Function
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count: arr(i) = CStr(col(i)): Next i
    For i = 2 To UBound(arr)
        tmp = arr(i): j = i - 1
        Do While j >= 1
            If StrComp(arr(j), tmp, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
    SortedNames = arr
End Function

' ==============================================================================
' STEP 1: SILENT CONVERSION INSIDE \Source03\ FOLDER
' ==============================================================================
Private Sub ConvertSourceFilesToXLSX(ByVal targetFolderPath As String, _
                                     ByRef totalFiles As Long, _
                                     ByRef convertedFiles As Long, _
                                     ByRef failedFiles As Long)
    Dim fso As Object, folderObj As Object, fileObj As Object
    Dim paths As Collection, p As Variant
    Dim wb As Workbook
    Dim newFilePath As String, baseName As String, ext As String
    Dim cleanFolder As String, archPath As String, archFile As String
    Dim saveErr As Long

    totalFiles = 0: convertedFiles = 0: failedFiles = 0

    If Right(targetFolderPath, 1) = "\" Then
        cleanFolder = Left(targetFolderPath, Len(targetFolderPath) - 1)
    Else
        cleanFolder = targetFolderPath
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(cleanFolder) Then Exit Sub
    Set folderObj = fso.GetFolder(cleanFolder)
    archPath = cleanFolder & "\" & ARCHIVE_FOLDER

    ' Snapshot the file list first (do not modify the folder while enumerating it)
    Set paths = New Collection
    For Each fileObj In folderObj.Files
        If Left(fileObj.Name, 2) <> "~$" Then
            ext = LCase(fso.GetExtensionName(fileObj.Name))
            If ext = "xls" Or ext = "xlsb" Or ext = "xlsm" Then
                If StrComp(fileObj.Path, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
                    paths.Add fileObj.Path
                End If
            End If
        End If
    Next fileObj

    For Each p In paths
        totalFiles = totalFiles + 1
        baseName = fso.GetBaseName(CStr(p))
        newFilePath = cleanFolder & "\" & baseName & ".xlsx"

        Set wb = Nothing
        On Error Resume Next
        Set wb = Workbooks.Open(fileName:=CStr(p), UpdateLinks:=False, ReadOnly:=True, _
                                AddToMRU:=False, Password:="")
        Err.Clear
        On Error GoTo 0

        If wb Is Nothing Then
            failedFiles = failedFiles + 1
        ElseIf fso.FileExists(newFilePath) Then
            ' Target already exists: do not overwrite (duplicate is skipped later in collection)
            wb.Close SaveChanges:=False
            failedFiles = failedFiles + 1
        Else
            On Error Resume Next
            wb.SaveAs fileName:=newFilePath, FileFormat:=xlOpenXMLWorkbook
            saveErr = Err.Number: Err.Clear
            wb.Close SaveChanges:=False
            Err.Clear
            On Error GoTo 0

            If saveErr = 0 Then
                ' Archive the original instead of deleting it
                On Error Resume Next
                If Not fso.FolderExists(archPath) Then fso.CreateFolder archPath
                archFile = archPath & "\" & fso.GetFileName(CStr(p))
                If fso.FileExists(archFile) Then
                    archFile = archPath & "\" & Format(Now, "yyyymmdd_hhnnss") & "_" & fso.GetFileName(CStr(p))
                End If
                fso.MoveFile CStr(p), archFile
                Err.Clear
                On Error GoTo 0
                convertedFiles = convertedFiles + 1
            Else
                failedFiles = failedFiles + 1
            End If
        End If
        Set wb = Nothing
    Next p
End Sub

' ==============================================================================
' STEP 2: FONT CHECK & COLLECTION INTO Term01Results.xlsx
' ==============================================================================
' Checks the Windows font registry (HKLM + HKCU). If the check itself cannot run,
' returns True (keeps the requested font instead of forcing the fallback).
Private Function FontExists(ByVal FontName As String) As Boolean
    Const HKCU As Long = &H80000001
    Const HKLM As Long = &H80000002
    Dim reg As Object, hv As Variant
    Dim names As Variant, types As Variant
    Dim i As Long, p As Long, nm As String, enumOK As Boolean

    On Error GoTo Fallback
    Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

    For Each hv In Array(HKLM, HKCU)
        names = Empty
        If reg.EnumValues(CLng(hv), "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", names, types) = 0 Then
            enumOK = True
            If IsArray(names) Then
                For i = LBound(names) To UBound(names)
                    nm = CStr(names(i))
                    p = InStr(nm, " (")
                    If p > 0 Then nm = Left(nm, p - 1)          ' strip "(TrueType)" / "(OpenType)"
                    ' exact name, or family name of a variable font (e.g. "Bahnschrift" for "Bahnschrift SemiBold")
                    If StrComp(nm, FontName, vbTextCompare) = 0 Then FontExists = True: Exit Function
                    If StrComp(Left(FontName, Len(nm) + 1), nm & " ", vbTextCompare) = 0 Then FontExists = True: Exit Function
                Next i
            End If
        End If
    Next hv

    FontExists = Not enumOK     ' enumeration worked and nothing matched -> False; enumeration failed -> True
    Exit Function
Fallback:
    FontExists = True
End Function

Private Sub SetupFonts()
    If FontExists(FONT_ARABIC) Then g_ActiveArabicFont = FONT_ARABIC Else g_ActiveArabicFont = FONT_FALLBACK
    If FontExists(FONT_NUMERIC) Then g_ActiveNumericFont = FONT_NUMERIC Else g_ActiveNumericFont = FONT_FALLBACK
End Sub

Private Function SrcProcessFile(ByRef wbResult As Workbook, _
                                        ByVal SourcePath As String, _
                                        ByVal fileName As String, _
                                        ByVal SheetNumber As Long, _
                                        ByRef ReasonOut As String) As Boolean
    Dim wbSource As Workbook, wsSource As Worksheet, wsResult As Worksheet
    Dim LastRow As Long, LastCol As Long, DataArr As Variant
    Dim errNum As Long, errDesc As String, lastDataRow As Long, wrapEndRow As Long
    Dim idLastCol As Long, lastFind As Range
    Dim tbl As ListObject, tblRange As Range, seenHeaders As Object
    Dim h As Long, headerText As String, genCounter As Long

    SrcProcessFile = False
    ReasonOut = ""
    On Error GoTo EH

    On Error Resume Next
    Set wbSource = Workbooks.Open(fileName:=SourcePath & fileName, UpdateLinks:=False, _
                                  ReadOnly:=True, AddToMRU:=False, Password:="")
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If wbSource Is Nothing Or errNum <> 0 Then
        ReasonOut = "OPEN_FAILED | " & errDesc: Exit Function
    End If

    If wbSource.Sheets.Count = 0 Then
        ReasonOut = "EMPTY_WORKBOOK (no sheets)": wbSource.Close SaveChanges:=False: Exit Function
    End If

    Set wsSource = wbSource.Sheets(1)

    On Error Resume Next
    LastCol = wsSource.Cells(DATA_START_ROW, wsSource.Columns.Count).End(xlToLeft).Column
    Set lastFind = wsSource.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, _
                                       SearchDirection:=xlPrevious)
    If lastFind Is Nothing Then LastRow = 0 Else LastRow = lastFind.Row - TRAILING_ROWS_SKIP
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If errNum <> 0 Then
        ReasonOut = "DETECTION_FAILED | " & errDesc: wbSource.Close SaveChanges:=False: Exit Function
    End If

    If LastRow < DATA_START_ROW Or LastCol < 1 Then
        ReasonOut = "EMPTY_WORKBOOK (no data at row " & DATA_START_ROW & ")"
        wbSource.Close SaveChanges:=False: Exit Function
    End If

    On Error Resume Next
    Set wsResult = wbResult.Sheets.Add(After:=wbResult.Sheets(wbResult.Sheets.Count))
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If wsResult Is Nothing Then
        ReasonOut = "ADD_SHEET_FAILED | " & errDesc: wbSource.Close SaveChanges:=False: Exit Function
    End If

    On Error Resume Next
    wsResult.Name = Format(SheetNumber, "00")
    Err.Clear
    On Error GoTo EH

    wsResult.DisplayRightToLeft = True

    ' ---- Read source data
    On Error Resume Next
    DataArr = wsSource.Range(wsSource.Cells(DATA_START_ROW, 1), wsSource.Cells(LastRow, LastCol)).Value
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If errNum <> 0 Or Not IsArray(DataArr) Then
        ReasonOut = "DATA_READ_FAILED | " & errDesc
        wbSource.Close SaveChanges:=False
        wsResult.Delete
        Exit Function
    End If

    lastDataRow = 1 + UBound(DataArr, 1)     ' header at row 2 + data rows

    ' ---- Number formats BEFORE writing data (so "@" columns store text)
    With wsResult
        .Columns("B").NumberFormat = "@"
        .Columns("C").NumberFormat = "yyyy-mm-dd"
        If LastCol >= 4 Then .Columns("D").NumberFormat = "@"
        If LastCol >= 5 Then .Columns("E").NumberFormat = "@"
        If LastCol >= 6 Then .Range(.Columns("F"), .Columns(LastCol)).NumberFormat = "0.00"
    End With

    wsResult.Range("A2").Resize(UBound(DataArr, 1), UBound(DataArr, 2)).Value = DataArr

    ' ---- Unique, non-empty headers (case-insensitive, like Excel tables)
    Set seenHeaders = CreateObject("Scripting.Dictionary")
    seenHeaders.CompareMode = vbTextCompare
    genCounter = 0

    For h = 1 To LastCol
        headerText = Trim(CStr(wsResult.Cells(2, h).Value))
        If headerText = "" Or seenHeaders.Exists(headerText) Then
            genCounter = genCounter + 1
            Do While seenHeaders.Exists("Column" & genCounter): genCounter = genCounter + 1: Loop
            wsResult.Cells(2, h).Value = "Column" & genCounter
            seenHeaders.Add "Column" & genCounter, True
        Else
            seenHeaders.Add headerText, True
        End If
    Next h

    ' ---- Dynamic table (range ends exactly at the last data row)
    Set tblRange = wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(lastDataRow, LastCol))
    Set tbl = Nothing

    On Error Resume Next
    Set tbl = wsResult.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    If Err.Number = 0 Then
        tbl.Name = "Data" & Format(SheetNumber, "00")
        tbl.TableStyle = ""
        tbl.ShowTableStyleColumnStripes = False
        tbl.ShowTableStyleRowStripes = False
        tbl.ShowTableStyleFirstColumn = False
        tbl.ShowTableStyleLastColumn = False
        tbl.ShowAutoFilter = True
    End If
    Err.Clear
    On Error GoTo EH

    ' ---- Formatting
    With wsResult
        .Columns("B").Font.Name = g_ActiveArabicFont: .Columns("B").Font.Size = FONT_SIZE
        .Columns("C").Font.Name = g_ActiveNumericFont: .Columns("C").Font.Size = FONT_SIZE
        If LastCol >= 4 Then .Columns("D").Font.Name = g_ActiveArabicFont: .Columns("D").Font.Size = FONT_SIZE
        If LastCol >= 5 Then .Columns("E").Font.Name = g_ActiveArabicFont: .Columns("E").Font.Size = FONT_SIZE
        If LastCol >= 6 Then
            With .Range(.Columns("F"), .Columns(LastCol))
                .Font.Name = g_ActiveNumericFont
                .Font.Size = FONT_SIZE
            End With
        End If

        wrapEndRow = lastDataRow
        If wrapEndRow < 2 Then wrapEndRow = 2
        .Range(.Cells(2, 1), .Cells(wrapEndRow, LastCol)).WrapText = False

        With .Range(.Cells(1, 1), .Cells(lastDataRow, LastCol))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ' Identity columns: AutoFit. Score columns: fixed width, rotated bold headers.
        idLastCol = IDENTITY_LAST_COL
        If LastCol < idLastCol Then idLastCol = LastCol
        .Range(.Columns(1), .Columns(idLastCol)).EntireColumn.AutoFit

        If LastCol >= 6 Then .Range(.Columns(IDENTITY_LAST_COL + 1), .Columns(LastCol)).ColumnWidth = COL_WIDTH

        If LastCol >= 4 Then
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Orientation = xlUpward
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Font.Bold = True
        End If
    End With

    ' Apply title formatting last so table column fonts do not override it.
    PlaceHeaderTitle wsSource, wsResult

    On Error Resume Next
    wbSource.Close SaveChanges:=False
    Set wbSource = Nothing
    Err.Clear
    On Error GoTo 0

    SrcProcessFile = True
    Exit Function

EH:
    ReasonOut = "UNEXPECTED | " & Err.Number & " : " & Err.Description
    Err.Clear
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    If Not wsResult Is Nothing Then wsResult.Delete
    Err.Clear
    SrcProcessFile = False
End Function

Private Sub SrcCollectAll(ByRef wbResult As Workbook, _
                             ByVal SourcePath As String, _
                             ByRef CopiedCount As Long, _
                             ByRef SkippedCount As Long, _
                             ByRef SkippedList As String)
    Dim fileName As String, SheetNumber As Long, Reason As String
    Dim names As Collection, sorted As Variant, i As Long
    Dim f As String, ext As String, baseName As String, dotPos As Long

    CopiedCount = 0: SkippedCount = 0: SheetNumber = 1: SkippedList = ""

    ' Collect names first (Dir cannot be nested with other Dir calls)
    Set names = New Collection
    fileName = Dir(SourcePath & "*.xls*")
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            dotPos = InStrRev(fileName, ".")
            ext = LCase(Mid(fileName, dotPos + 1))
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb" Then names.Add fileName
        End If
        fileName = Dir
    Loop

    If names.Count = 0 Then
        MsgBox "Source03 folder is empty or contains no Excel files:" & vbCrLf & SourcePath, vbExclamation, "Empty Source03 Folder"
        Exit Sub
    End If

    sorted = SortedNames(names)     ' deterministic sheet order (01, 02, ...)
    SetupFonts

    For i = LBound(sorted) To UBound(sorted)
        f = CStr(sorted(i))
        dotPos = InStrRev(f, ".")
        ext = LCase(Mid(f, dotPos + 1))
        baseName = Left(f, dotPos - 1)

        If StrComp(SourcePath & f, ThisWorkbook.FullName, vbTextCompare) = 0 Then
            ' skip the macro workbook itself
        ElseIf ext <> "xlsx" And FileExistsPath(SourcePath & baseName & ".xlsx") Then
            SkippedCount = SkippedCount + 1
            SkippedList = SkippedList & "  - " & f & " : DUPLICATE (.xlsx version exists)" & vbCrLf
        ElseIf SrcProcessFile(wbResult, SourcePath, f, SheetNumber, Reason) Then
            CopiedCount = CopiedCount + 1: SheetNumber = SheetNumber + 1
        Else
            SkippedCount = SkippedCount + 1
            SkippedList = SkippedList & "  - " & f & " : " & Reason & vbCrLf
        End If
    Next i

    If CopiedCount = 0 Then
        MsgBox "None of the source files could be processed.", vbCritical, "Collection Failed"
    End If
End Sub


Private Sub PlaceHeaderTitle(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet)

    Dim rawText As String
    Dim partTitle As String
    Dim partGroup As String

    ' 1. Capture and clean the text from A5
    rawText = CleanText(CStr(wsSource.Range("A5").Value))

    ' 2. Split at the end of the school year
    SplitAtSchoolYear rawText, partTitle, partGroup


    ' 3. Write and format both merged ranges
    WriteMergedCell wsTarget.Range("B1:F1"), partTitle
    WriteMergedCell wsTarget.Range("H1:Q1"), partGroup

    ' Merged cells do not auto-fit, so set the row height explicitly
    wsTarget.Rows(1).RowHeight = HEADER_ROW_HEIGHT

End Sub

'----------------------------------------------------------
' Cleans text: line breaks, non-breaking spaces, control
' characters and repeated spaces.
'----------------------------------------------------------
Private Function CleanText(ByVal s As String) As String

    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, Chr(160), " ")                       ' non-breaking space
    s = Application.WorksheetFunction.Clean(s)          ' control characters
    CleanText = Application.WorksheetFunction.Trim(s)   ' extra spaces

End Function

'----------------------------------------------------------
' Splits text into:
'   part1 = everything up to the end of the school year
'   part2 = the remainder (class / group name)
' If no year is found, part1 = full text and part2 = "".
'----------------------------------------------------------
Private Sub SplitAtSchoolYear(ByVal txt As String, _
                              ByRef part1 As String, _
                              ByRef part2 As String)

    Dim rx As Object
    Dim m As Object
    Dim splitPos As Long
    Dim digit As String
    Dim seps As String

    part1 = txt
    part2 = ""
    If Len(txt) = 0 Then Exit Sub

    ' Western (0-9) and Arabic-Indic (U+0660-U+0669) digits
    digit = "[0-9\u0660-\u0669]"
    ' Separators: slash, backslash, hyphen, en dash, em dash
    seps = "[/\\\-\u2013\u2014]"

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True

    ' First choice: full school-year pattern (e.g. 2025-2026)
    rx.Pattern = digit & "{4}\s*" & seps & "\s*" & digit & "{4}"

    If Not rx.Test(txt) Then
        ' Fallback: a single 4-digit year
        rx.Pattern = digit & "{4}"
        If Not rx.Test(txt) Then Exit Sub
    End If

    Set m = rx.Execute(txt)
    splitPos = m(0).FirstIndex + m(0).Length

    part1 = Trim(Left(txt, splitPos))

    If splitPos < Len(txt) Then
        part2 = Trim(Mid(txt, splitPos + 1))

        ' Remove any separator left at the start of the group name
        Do While Len(part2) > 0
            If InStr(1, "-:/" & ChrW(&H2013) & ChrW(&H2014), Left(part2, 1)) = 0 Then Exit Do
            part2 = Trim(Mid(part2, 2))
        Loop
    End If

End Sub

'----------------------------------------------------------
' Merges the range, writes the text and applies formatting.
'----------------------------------------------------------
Private Sub WriteMergedCell(ByVal rng As Range, ByVal txt As String)

    With rng
        .UnMerge                                ' remove any previous merge
        .ClearContents                          ' clear leftovers
        .Merge                                  ' merge the range
        .Value = txt
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL                   ' right-to-left reading order
        .Font.Name = GetAvailableFont(PREFERRED_FONT, FALLBACK_FONT)
        .Font.Size = HEADER_FONT_SIZE
        .Font.Bold = True
    End With

End Sub

'----------------------------------------------------------
' Returns the preferred font if installed, otherwise the fallback.
'----------------------------------------------------------
Private Function GetAvailableFont(ByVal wanted As String, _
                                  ByVal fallback As String) As String

    If IsFontInstalled(wanted) Then
        GetAvailableFont = wanted
    Else
        GetAvailableFont = fallback
    End If

End Function

'----------------------------------------------------------
' Checks the system and per-user font folders for a file whose
' name starts with the font name (e.g. Cairo-Regular.ttf).
'----------------------------------------------------------
Private Function IsFontInstalled(ByVal fontName As String) As Boolean

    Dim folders(1 To 2) As String
    Dim i As Long

    folders(1) = Environ("WINDIR") & "\Fonts\"
    folders(2) = Environ("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\"

    On Error Resume Next
    For i = 1 To 2
        If Len(Dir(folders(i) & fontName & "*.*")) > 0 Then
            IsFontInstalled = True
            Exit Function
        End If
    Next i
    On Error GoTo 0

End Function

' ==============================================================================
' HELPERS
' ==============================================================================

' Arabic-tolerant comparison text: unified alef / yeh / teh marbuta, digits to 0-9
Private Function NormAr(ByVal v As Variant) As String
    Dim t As String, i As Long
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    t = CStr(v)
    t = Replace(t, ChrW(&H623), ChrW(&H627))
    t = Replace(t, ChrW(&H625), ChrW(&H627))
    t = Replace(t, ChrW(&H622), ChrW(&H627))
    t = Replace(t, ChrW(&H649), ChrW(&H64A))
    t = Replace(t, ChrW(&H629), ChrW(&H647))
    t = Replace(t, ChrW(&H640), "")
    t = Replace(t, ChrW(&HA0), " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, ChrW(&H200E), "")
    t = Replace(t, ChrW(&H200F), "")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbCr, " ")
    For i = 0 To 9
        t = Replace(t, ChrW(&H660 + i), CStr(i))
        t = Replace(t, ChrW(&H6F0 + i), CStr(i))
    Next i
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormAr = Trim$(t)
End Function

' Same cleaning as the key of ResultFlow01 (group title)
Private Function KeyText(ByVal s As String) As String
    Dim c As Variant
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    For Each c In Array(&H200B, &H200E, &H200F, &H202A, &H202B, &H202C, &H640)
        s = Replace(s, ChrW(CLng(c)), "")
    Next c
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    KeyText = Trim(s)
End Function

Private Function CellText(ByVal r As Range) As String
    If IsError(r.Value) Then CellText = "" Else CellText = CStr(r.Value)
End Function

Private Function FolderExistsPath(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExistsPath = ((GetAttr(p) And vbDirectory) <> 0)
    If Err.Number <> 0 Then FolderExistsPath = False: Err.Clear
End Function

Private Function IsWorkbookOpen(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function

' Last row containing anything, in any column (0 if the sheet is empty)
Private Function LastUsedRowAny(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If f Is Nothing Then LastUsedRowAny = 0 Else LastUsedRowAny = f.Row
End Function

' "ThirdYear Stream 02" -> "ThirdYear Stream" (the last section number is removed)
Private Function StripSectionNumber(ByVal s As String) As String
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "\s*[0-9\u0660-\u0669]+\s*$"
    rx.Global = False
    StripSectionNumber = Trim(rx.Replace(s, ""))
End Function

' Last number of the text as a Long (0 when none)
Private Function SectionNumberOf(ByVal s As String) As Long
    Dim rx As Object, m As Object, d As String, i As Long
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "([0-9\u0660-\u0669]+)\s*$"
    If Not rx.Test(s) Then Exit Function
    Set m = rx.Execute(s)
    d = m(0).SubMatches(0)
    For i = 0 To 9
        d = Replace(d, ChrW(&H660 + i), CStr(i))
    Next i
    If IsNumeric(d) Then SectionNumberOf = CLng(d)
End Function

' The word for "section" (with its usual one-letter / "al" prefixes) becomes the
' word for "stream"; only whole words are replaced.
Private Function ReplaceSectionWord(ByVal s As String) As String
    Dim rx As Object
    If Len(s) = 0 Then Exit Function
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "(^|[\s:\-\u2013\u2014/])((?:\u0627\u0644|\u0644\u0644|\u0628\u0627\u0644|\u0648\u0627\u0644|\u0644|\u0628|\u0648)?)" & _
                 "\u0642\u0633\u0645(?=$|[\s:\-\u2013\u2014/.,\u060C\u061B])"
    ReplaceSectionWord = rx.Replace(s, "$1$2" & ChrW(&H634) & ChrW(&H639) & ChrW(&H628) & ChrW(&H629))
End Function

' Valid, unique worksheet name (max 31 chars, no forbidden characters)
Private Function SafeSheetName(ByVal base As String, ByVal wb As Workbook) As String
    Dim s As String, ch As Variant, cand As String, n As Long
    Dim ws As Worksheet, clash As Boolean
    s = base
    For Each ch In Array("\", "/", "?", "*", "[", "]", ":")
        s = Replace(s, CStr(ch), " ")
    Next ch
    s = Trim(s)
    If Len(s) = 0 Then s = "Group"
    If Len(s) > 31 Then s = Trim(Left(s, 31))
    cand = s: n = 1
    Do
        clash = False
        For Each ws In wb.Worksheets
            If StrComp(ws.Name, cand, vbTextCompare) = 0 Then clash = True: Exit For
        Next ws
        If Not clash Then Exit Do
        n = n + 1
        cand = Left(s, 31 - Len(CStr(n)) - 1) & "_" & CStr(n)
    Loop
    SafeSheetName = cand
End Function

' Column of the term-3 average in a collected sheet (header holds "average" + "term",
' preferably with the digit 3 or the word for "third"); 0 when not found
Private Function FindAverageColumn(ByVal ws As Worksheet, ByVal lastCol As Long) As Long
    Dim c As Long, h As String, wAvg As String, wTerm As String, wThird As String
    Dim lastAny As Long, preferred As Long
    wAvg = NormAr(ArHex("0645 0639 062F 0644"))
    wTerm = NormAr(ArHex("0641 0635 0644"))
    wThird = NormAr(ArHex("062B 0627 0644 062B"))
    For c = 1 To lastCol
        h = NormAr(CellText(ws.Cells(TERM_HDR_ROW, c)))
        If InStr(h, wAvg) > 0 And InStr(h, wTerm) > 0 Then
            lastAny = c
            If preferred = 0 Then
                If InStr(h, "3") > 0 Or InStr(h, wThird) > 0 Then preferred = c
            End If
        End If
    Next c
    If preferred > 0 Then FindAverageColumn = preferred Else FindAverageColumn = lastAny
End Function

Private Function ColumnLetterOf(ByVal n As Long) As String
    ColumnLetterOf = Split(Cells(1, n).Address(True, False), "$")(0)
End Function

' One column as a 2-D array (n x 1), also for a single row
Private Function ReadColumn(ByVal ws As Worksheet, ByVal col As Long, _
                            ByVal r1 As Long, ByVal r2 As Long) As Variant
    Dim v As Variant, one(1 To 1, 1 To 1) As Variant
    v = ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)).Value2
    If IsArray(v) Then
        ReadColumn = v
    Else
        one(1, 1) = v
        ReadColumn = one
    End If
End Function

' Numeric average of row r (very low when blank / not a number so it sorts last)
Private Function AvgValue(ByVal arr As Variant, ByVal r As Long) As Double
    Dim v As Variant
    AvgValue = -1E+30
    If IsEmpty(arr) Then Exit Function
    v = arr(r, 1)
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then AvgValue = CDbl(v)
End Function

' Stable sort of the index array by value, highest first
Private Sub SortIndexDesc(ByRef idx() As Long, ByRef v() As Double, ByVal n As Long)
    Dim i As Long, j As Long, k As Long
    For i = 2 To n
        k = idx(i): j = i - 1
        Do While j >= 1
            If v(idx(j)) >= v(k) Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = k
    Next i
End Sub

' ==============================================================================
' ONE STREAM -> ONE SHEET
' ==============================================================================
Private Function WriteGroupSheet(ByVal wbTerm As Workbook, _
                                 ByVal wbOut As Workbook, _
                                 ByRef secNames() As String, _
                                 ByVal nSec As Long, _
                                 ByVal sheetName As String, _
                                 ByVal baseText As String, _
                                 ByVal gIdx As Long, _
                                 ByRef students As Long, _
                                 ByRef info As String) As Boolean

    Dim i As Long, r As Long, c As Long, p As Long
    Dim wsT As Worksheet, wsOut As Worksheet, lo As ListObject
    Dim secRows() As Long, secLast() As Long, secAvgCol() As Long
    Dim total As Long, lastRow As Long, lastCol As Long
    Dim allData() As Variant, allAvg() As Double, idx() As Long
    Dim outData As Variant, tmp As Variant, avgArr As Variant
    Dim titleText As String, partTitle As String
    Dim noAvg As Boolean

    WriteGroupSheet = False
    students = 0: info = ""

    ReDim secRows(1 To nSec)
    ReDim secLast(1 To nSec)
    ReDim secAvgCol(1 To nSec)

    For i = 1 To nSec
        Set wsT = wbTerm.Worksheets(secNames(i))
        lastRow = LastUsedRowAny(wsT)
        secLast(i) = lastRow
        secRows(i) = lastRow - TERM_HDR_ROW
        If secRows(i) < 0 Then secRows(i) = 0
        lastCol = wsT.Cells(TERM_HDR_ROW, wsT.Columns.Count).End(xlToLeft).Column
        secAvgCol(i) = FindAverageColumn(wsT, lastCol)
        total = total + secRows(i)
    Next i

    If total = 0 Then Exit Function

    ' ---- data (A:E only) ----
    If nSec = 1 Then
        Set wsT = wbTerm.Worksheets(secNames(1))
        outData = wsT.Range(wsT.Cells(TERM_FIRST_ROW, 1), wsT.Cells(secLast(1), KEEP_COLS)).Value
        info = " (as in the source)"
    Else
        ReDim allData(1 To total, 1 To KEEP_COLS)
        ReDim allAvg(1 To total)
        p = 0
        For i = 1 To nSec
            If secRows(i) > 0 Then
                Set wsT = wbTerm.Worksheets(secNames(i))
                tmp = wsT.Range(wsT.Cells(TERM_FIRST_ROW, 1), wsT.Cells(secLast(i), KEEP_COLS)).Value
                If secAvgCol(i) > 0 Then
                    avgArr = ReadColumn(wsT, secAvgCol(i), TERM_FIRST_ROW, secLast(i))
                Else
                    avgArr = Empty
                    noAvg = True
                End If
                For r = 1 To secRows(i)
                    p = p + 1
                    For c = 1 To KEEP_COLS
                        allData(p, c) = tmp(r, c)
                    Next c
                    allAvg(p) = AvgValue(avgArr, r)
                Next r
            End If
        Next i

        ReDim idx(1 To total)
        For p = 1 To total: idx(p) = p: Next p
        If Not noAvg Then SortIndexDesc idx, allAvg, total

        ReDim outData(1 To total, 1 To KEEP_COLS)
        For p = 1 To total
            outData(p, 1) = p
            For c = 2 To KEEP_COLS
                outData(p, c) = allData(idx(p), c)
            Next c
        Next p

        If noAvg Then
            info = " (WARNING: average column not found - sections kept in order)"
        Else
            info = " (sorted by average, column " & ColumnLetterOf(secAvgCol(1)) & ")"
        End If
    End If

    ' ---- sheet: copy of the first section sheet, so the Collector01 formatting is kept ----
    Set wsT = wbTerm.Worksheets(secNames(1))
    partTitle = ReplaceSectionWord(CellText(wsT.Range("B1")))
    titleText = Trim(partTitle & " " & baseText)

    wsT.Copy After:=wbOut.Sheets(wbOut.Sheets.Count)
    Set wsOut = wbOut.Sheets(wbOut.Sheets.Count)

    On Error Resume Next
    wsOut.Name = sheetName
    Err.Clear
    On Error GoTo 0

    wsOut.Rows(1).UnMerge
    wsOut.Rows(1).ClearContents
    wsOut.Range(wsOut.Columns(KEEP_COLS + 1), wsOut.Columns(wsOut.Columns.Count)).Delete
    WriteMergedCell wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, KEEP_COLS)), titleText

    ' ---- dynamic table over A2:E(last) ----
    Set lo = Nothing
    If wsOut.ListObjects.Count > 0 Then Set lo = wsOut.ListObjects(1)
    If lo Is Nothing Then
        Set lo = wsOut.ListObjects.Add(xlSrcRange, _
                 wsOut.Range(wsOut.Cells(TERM_HDR_ROW, 1), wsOut.Cells(TERM_HDR_ROW + total, KEEP_COLS)), , xlYes)
        lo.TableStyle = ""
        lo.ShowTableStyleColumnStripes = False
        lo.ShowTableStyleRowStripes = False
        lo.ShowTableStyleFirstColumn = False
        lo.ShowTableStyleLastColumn = False
        lo.ShowAutoFilter = True
    Else
        lo.Resize wsOut.Range(wsOut.Cells(TERM_HDR_ROW, 1), wsOut.Cells(TERM_HDR_ROW + total, KEEP_COLS))
    End If
    On Error Resume Next
    lo.Name = "Data" & Format(gIdx, "00")
    Err.Clear
    On Error GoTo 0

    wsOut.Range(wsOut.Cells(TERM_FIRST_ROW, 1), wsOut.Cells(TERM_HDR_ROW + total, KEEP_COLS)).Value = outData

    students = total
    WriteGroupSheet = True
End Function

' ==============================================================================
' SECOND MACRO: RepeatingStudents.xlsx (optional)
' One sheet per stream (code of the extraction workbook), A:E only, built from
' the same Source03 data. Sections of a stream are stacked under one table header.
' ==============================================================================
Public Sub Run_RepeatingStudents01()

    Dim prevScreen As Boolean, prevAlerts As Boolean, prevEvents As Boolean
    Dim wbTerm As Workbook, wbOut As Workbook, ws As Worksheet
    Dim srcFolder As String, outPath As String
    Dim srcCopied As Long, srcSkipped As Long, srcList As String
    Dim convTotal As Long, convOK As Long, convFailed As Long
    Dim codeOrder As Variant
    Dim Groups As Object, GroupSecs As Object, GroupText As Object, GroupName As Object
    Dim OrderKeys As Collection
    Dim hText As String, baseText As String, key As String, code As String, grpName As String
    Dim secNo As Long
    Dim i As Long, j As Long, g As Long, nSec As Long
    Dim kv As Variant
    Dim secNames() As String, secNos() As Long, tmpName As String, tmpNo As Long
    Dim sheetName As String, students As Long, info As String
    Dim defaultCount As Long, made As Long, totalStudents As Long
    Dim notThird As Long
    Dim lines As String, Report As String, errMsg As String
    Dim saveErr As Long, saveDesc As String

    ' ==================== Pre-flight checks ====================
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Please save this workbook before running the macro.", vbExclamation, "Repeating Students"
        Exit Sub
    End If
    If LCase(Left(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "This workbook is on OneDrive/SharePoint (URL path)." & vbCrLf & _
               "Move it to a local folder, or sync the folder, then run again.", vbExclamation, "Repeating Students"
        Exit Sub
    End If

    prevScreen = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    prevEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo UnexpectedError

    srcFolder = ThisWorkbook.Path & "\" & SOURCE_FOLDER_NAME & "\"
    outPath = ThisWorkbook.Path & "\" & OUTPUT_FILE_NAME

    If Not FolderExistsPath(ThisWorkbook.Path & "\" & SOURCE_FOLDER_NAME) Then
        MsgBox SOURCE_FOLDER_NAME & " folder was not found next to this workbook." & vbCrLf & vbCrLf & _
               "Expected path:" & vbCrLf & srcFolder, vbCritical, "Repeating Students"
        GoTo CleanExit
    End If

    If IsWorkbookOpen(outPath) Then
        MsgBox OUTPUT_FILE_NAME & " is open in this Excel session. Close it first.", vbExclamation, "Repeating Students"
        GoTo CleanExit
    End If

    ' ==================== Collector01 steps (in memory) ====================
    Set wbTerm = Nothing
    If Not SrcBuildWorkbook(srcFolder, wbTerm, srcCopied, srcSkipped, srcList, _
                            convTotal, convOK, convFailed) Then
        MsgBox "No usable source file was found in " & SOURCE_FOLDER_NAME & "." & vbCrLf & vbCrLf & srcList, _
               vbCritical, "Repeating Students"
        GoTo CleanExit
    End If

    Set Groups = CreateObject("Scripting.Dictionary"):    Groups.CompareMode = vbTextCompare
    Set GroupSecs = CreateObject("Scripting.Dictionary"): GroupSecs.CompareMode = vbTextCompare
    Set GroupText = CreateObject("Scripting.Dictionary"): GroupText.CompareMode = vbTextCompare
    Set GroupName = CreateObject("Scripting.Dictionary"): GroupName.CompareMode = vbTextCompare

    ' ==================== Keep only the third-year sheets, grouped by stream code ====================
    For Each ws In wbTerm.Worksheets
        hText = KeyText(CellText(ws.Range("H1")))

        If Not StartsWithThirdSecondary(hText) Then
            notThird = notThird + 1
        Else
            baseText = StripSectionNumber(hText)
            secNo = SectionNumberOf(hText)
            code = SheetCodeFromTitle(baseText)
            If Len(code) > 0 Then
                key = "#" & code
                grpName = code
            Else
                key = NormAr(baseText)
                grpName = baseText
            End If
            If Not Groups.Exists(key) Then
                Groups.Add key, New Collection
                GroupSecs.Add key, New Collection
                GroupText.Add key, baseText
                GroupName.Add key, grpName
            End If
            Groups.Item(key).Add ws.Name
            GroupSecs.Item(key).Add secNo
        End If
    Next ws

    If Groups.Count = 0 Then
        MsgBox "No third-year sheet was found in " & SOURCE_FOLDER_NAME & "." & vbCrLf & _
               "Nothing was created.", vbExclamation, "Repeating Students"
        GoTo CleanExit
    End If

    ' Output order: Philo, GeEc, Math, SE, LaEtr, then the technical streams, then any new stream
    codeOrder = Array("Philo", "GeEc", "Math", "SE", "LaEtr", "TGM", "TGE", "TGC", "TGP")
    Set OrderKeys = New Collection
    For i = 0 To UBound(codeOrder)
        key = "#" & codeOrder(i)
        If Groups.Exists(key) Then OrderKeys.Add key
    Next i
    For Each kv In Groups.Keys
        If Left$(CStr(kv), 1) <> "#" Then OrderKeys.Add CStr(kv)
    Next kv

    ' ==================== Build the output workbook ====================
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    defaultCount = wbOut.Sheets.Count
    g = 0

    For Each kv In OrderKeys
        key = CStr(kv)
        nSec = Groups.Item(key).Count

        ReDim secNames(1 To nSec)
        ReDim secNos(1 To nSec)
        For i = 1 To nSec
            secNames(i) = CStr(Groups.Item(key).Item(i))
            secNos(i) = CLng(GroupSecs.Item(key).Item(i))
        Next i

        ' sections in ascending order of their number (stable)
        For i = 2 To nSec
            tmpName = secNames(i): tmpNo = secNos(i): j = i - 1
            Do While j >= 1
                If secNos(j) <= tmpNo Then Exit Do
                secNames(j + 1) = secNames(j): secNos(j + 1) = secNos(j): j = j - 1
            Loop
            secNames(j + 1) = tmpName: secNos(j + 1) = tmpNo
        Next i

        sheetName = SafeSheetName(CStr(GroupName.Item(key)), wbOut)

        g = g + 1
        If WriteGroupSheet(wbTerm, wbOut, secNames, nSec, sheetName, CStr(GroupText.Item(key)), g, students, info) Then
            made = made + 1
            totalStudents = totalStudents + students
            lines = lines & "  - " & sheetName & " : " & nSec & " section(s), " & students & " student(s)" & info & vbCrLf
        Else
            g = g - 1
            lines = lines & "  - " & CStr(GroupText.Item(key)) & " : skipped (no student rows)" & vbCrLf
        End If
    Next kv

    If made = 0 Then
        MsgBox "The third-year sheets contain no student rows." & vbCrLf & "Nothing was created.", _
               vbExclamation, "Repeating Students"
        GoTo CleanExit
    End If

    ' remove the default empty sheet(s) that Excel created with the workbook
    For i = defaultCount To 1 Step -1
        wbOut.Sheets(i).Delete
    Next i

    ' ==================== Save (replace directly, no backup) ====================
    On Error Resume Next
    wbOut.SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook
    saveErr = Err.Number: saveDesc = Err.Description
    Err.Clear
    On Error GoTo UnexpectedError

    If saveErr <> 0 Then
        MsgBox "The output file could not be saved:" & vbCrLf & saveDesc & vbCrLf & vbCrLf & outPath, _
               vbCritical, "Repeating Students"
        GoTo CleanExit
    End If

    wbOut.Close SaveChanges:=False
    Set wbOut = Nothing

    ' ==================== Report ====================
    Report = "Repeating students file created." & vbCrLf & vbCrLf & _
             "Output : " & OUTPUT_FILE_NAME & vbCrLf & vbCrLf & _
             "--- " & SOURCE_FOLDER_NAME & " ---" & vbCrLf & _
             "Source files read : " & srcCopied & vbCrLf & _
             "Converted to .xlsx : " & convOK & " of " & convTotal & " (failed " & convFailed & ")" & vbCrLf & _
             "Unreadable source files : " & srcSkipped & vbCrLf & _
             "Sheets left out (not third year) : " & notThird & vbCrLf & vbCrLf & _
             "--- Sheets (" & made & ") - " & totalStudents & " student(s) ---" & vbCrLf & lines

    If srcSkipped > 0 Then
        Report = Report & vbCrLf & "Unreadable / skipped source files:" & vbCrLf & srcList
    End If

    MsgBox Report, vbInformation, "Repeating Students"

CleanExit:
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    If Not wbTerm Is Nothing Then wbTerm.Close SaveChanges:=False
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Exit Sub

UnexpectedError:
    errMsg = "An unexpected error occurred:" & vbCrLf & "Error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    If Not wbTerm Is Nothing Then wbTerm.Close SaveChanges:=False
    Set wbOut = Nothing: Set wbTerm = Nothing
    On Error GoTo 0
    MsgBox errMsg, vbCritical, "Repeating Students"
    Resume CleanExit
End Sub
