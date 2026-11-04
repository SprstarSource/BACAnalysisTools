Attribute VB_Name = "modRepeatStatus"
Option Explicit

' =====================================================================
' Repeat-status matcher - macro: Run_RepeatStatus
'
' Reads the 3rd-term results workbook (SOURCE, never modified) and writes
' the repeat status (Yes / No) into column E of the extraction workbook
' (TARGET, values only, after a timestamped backup).
'
' MATCH   : same sheet code + normalized name (col B) + date of birth
'           (col C). EXACT matching only (no fuzzy matching).
'           A pupil is searched only in the sheet of his own branch
'           (SE pools the two source sheets "... 01" and "... 02").
' ORDER   : 1) read the source and detect duplicate identities
'           2) per target sheet: detect duplicates, then match
'           3) ask for confirmation, backup, write, save
'           4) create a separate report workbook (summary + details)
' REPORT  : locations are written as "SheetCode: number in column A",
'           e.g. SE: 41
' =====================================================================

Private Const FIRST_DATA_ROW As Long = 3
Private Const ID_COL As Long = 1          ' A = sequence number
Private Const NAME_COL As Long = 2        ' B = full name
Private Const DOB_COL As Long = 3         ' C = date of birth
Private Const REPEAT_COL As Long = 5      ' E = repeat Yes / No
Private Const SRC_LAST_COL As Long = 5    ' source columns A:E
Private Const MIN_VALID_RATIO As Double = 0.5   ' source structure check

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
Private mLastSourcePath As String
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

    On Error GoTo FatalError

    mYes = Ar("yes")
    mNo = Ar("no")
    Erase mSets
    mLogN = 0
    Erase mLogCat: Erase mLogLoc: Erase mLogSide: Erase mLogSheet: Erase mLogDet

    ' ---- Question 1 of 2: the SOURCE file (read only) ----
    sourcePath = AskSourceFile()
    If Len(sourcePath) = 0 Then Exit Sub
    mLastSourcePath = sourcePath

    ' ---- Question 2 of 2: the file that RECEIVES the values ----
    targetPath = AskTargetFile()
    If Len(targetPath) = 0 Then Exit Sub

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
    Application.StatusBar = Ar("sb_read")

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
        wbTarget.Save
        saved = True
    End If

    ' =================================================================
    ' PHASE 4: report workbook (separate file, target structure untouched)
    ' =================================================================
    Application.StatusBar = Ar("sb_report")
    reportPath = BuildReport(folder, stamp, sourcePath, targetPath, backupPath, saved, _
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
        ShowMsg Ar("stopped") & vbLf & fatalText, vbExclamation
        Exit Sub
    End If

    If saved Then
        report = Ar("done")
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
    If skippedSrc > 0 Then report = report & vbLf & Ar("src_skipped") & skippedSrc
    If saved Then report = report & vbLf & vbLf & Ar("r_backup") & vbLf & backupPath
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
        ByVal sourcePath As String, ByVal targetPath As String, ByVal backupPath As String, _
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
        KV sm, k, Ar("r_backup"), backupPath
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
    wsS.Columns(1).ColumnWidth = 48
    wsS.Columns("B:L").ColumnWidth = 18
    wsS.Range("A1:A" & k).Font.Bold = True

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
        Case "q1_title"
            s = ArHex("627 644 633 624 627 644 20 31 20 645 646 20 32 3A 20 627 644 645 644 641 20 627 644 645 635 62F 631")
        Case "q1_head"
            s = ArHex("627 62E 62A 631 20 627 644 622 646 20 627 644 645 644 641 20 627 644 645 635 62F 631 3A 20 648 647 648 20 645 644 641 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627")
            s = s & ArHex("644 62B 627 644 62B 20 627 644 630 64A 20 62A 64F 62C 644 628 20 645 646 647 20 62D 627 644 627 62A 20 627 644 625 639 627 62F 629 20 28 646 639 645 20 2F 20 644 627 29")
            s = s & ArHex("2E")
        Case "q1_note"
            s = ArHex("64A 645 643 646 20 623 646 20 64A 62E 62A 644 641 20 627 633 645 647 20 648 645 643 627 646 647 60C 20 644 643 646 20 64A 62C 628 20 623 646 20 62A 643 648 646 20 628 646")
            s = s & ArHex("64A 62A 647 20 645 637 627 628 642 629 20 62A 645 627 645 627 64B 20 644 644 645 644 641 20 627 644 645 639 62A 627 62F 2E")
        Case "q1_safe"
            s = ArHex("644 646 20 64A 62A 645 20 62A 639 62F 64A 644 20 647 630 627 20 627 644 645 644 641 2E")
        Case "q1_dialog"
            s = ArHex("31 2F 32 20 2D 20 627 62E 62A 631 20 627 644 645 644 641 20 627 644 645 635 62F 631 20 28 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 644 62B 29")
        Case "q2_title"
            s = ArHex("627 644 633 624 627 644 20 32 20 645 646 20 32 3A 20 645 644 641 20 627 644 644 635 642")
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
            s = ArHex("32 2F 32 20 2D 20 627 62E 62A 631 20 645 644 641 20 627 644 644 635 642 20 28 627 644 630 64A 20 633 62A 64F 643 62A 628 20 641 64A 647 20 62D 627 644 627 62A 20 627")
            s = s & ArHex("644 625 639 627 62F 629 29")
        Case "same_file"
            s = ArHex("627 644 645 644 641 20 627 644 645 635 62F 631 20 648 645 644 641 20 627 644 644 635 642 20 647 645 627 20 627 644 645 644 641 20 646 641 633 647 2E 20 627 62E 62A 631 20")
            s = s & ArHex("645 644 641 64A 646 20 645 62E 62A 644 641 64A 646 2E")
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
            s = ArHex("646 639 645 3A 20 643 62A 627 628 629 20 627 644 642 64A 645 20 648 62D 641 638 20 627 644 645 644 641 20 28 645 639 20 646 633 62E 629 20 627 62D 62A 64A 627 637 64A 629")
            s = s & ArHex("29")
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
            s = ArHex("627 644 645 644 641 20 627 644 645 635 62F 631 3A 20")
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
        Case "yes"
            s = ArHex("646 639 645")
        Case "no"
            s = ArHex("644 627")
    End Select
    Ar = s
End Function
