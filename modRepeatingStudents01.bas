Attribute VB_Name = "modRepeatingStudents01"
Option Explicit

'==============================================================================
' Module  : RepeatingStudents01
' Purpose : Special copy of Collector01 for the THIRD-YEAR repeating students.
'           Reads \Source03\ with the same steps as Collector01 (conversion of
'           the web-exported files to .xlsx, then collection) but the collected
'           workbook (Term03Results) stays in memory only - it is never saved.
'           From there:
'             - only the sheets whose A5 text holds the third-year phrase are kept
'             - sections of the same stream (names differing only by the last
'               section number) are stacked in ONE sheet, one table header
'             - when a stream has several sections the students are sorted by
'               the term average (descending) and numbered 1..N; a stream with a
'               single section is kept exactly as it is in the source
'             - every column from F onward is removed: the sheets hold A:E only
'             - row 1 title: the word for "section" becomes "stream" and the
'               stream name is written without its last section number
'           Sheet names: Philo / GeEc / Math / SE / LaEtr (a new stream gets a
'           sheet named after its own title). Output: RepeatingStudents.xlsx next
'           to this workbook, replaced directly if it already exists. No backups.
' Entry   : Run_RepeatingStudents01()
' Needs   : the Source03 folder next to the workbook that holds this module.
'           SELF-CONTAINED: no other module is required.
'==============================================================================

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

' ==============================================================================
' REPEATING STUDENTS LOGIC
' ==============================================================================

' ==============================================================================
' HELPERS
' ==============================================================================

' Builds a string from space-separated hex code points, e.g. HexText("0627 0644")
Private Function HexText(ByVal codes As String) As String
    Dim parts() As String, i As Long, result As String
    parts = Split(Trim(codes), " ")
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then result = result & ChrW(CLng("&H" & parts(i)))
    Next i
    HexText = result
End Function

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
Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If f Is Nothing Then LastUsedRow = 0 Else LastUsedRow = f.Row
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
    wAvg = NormAr(HexText("0645 0639 062F 0644"))
    wTerm = NormAr(HexText("0641 0635 0644"))
    wThird = NormAr(HexText("062B 0627 0644 062B"))
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
        lastRow = LastUsedRow(wsT)
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
' ENTRY POINT
' ==============================================================================
Public Sub Run_RepeatingStudents01()

    Dim prevScreen As Boolean, prevAlerts As Boolean, prevEvents As Boolean
    Dim wbTerm As Workbook, wbOut As Workbook, ws As Worksheet
    Dim srcFolder As String, outPath As String
    Dim srcCopied As Long, srcSkipped As Long, srcList As String
    Dim convTotal As Long, convOK As Long, convFailed As Long
    Dim knownNames As Variant, knownCodes As Variant
    Dim NameByKey As Object, Groups As Object, GroupSecs As Object, GroupText As Object
    Dim OrderKeys As Collection
    Dim phraseNorm As String
    Dim hText As String, tText As String, baseText As String, key As String
    Dim secNo As Long
    Dim i As Long, j As Long, g As Long, nSec As Long
    Dim kv As Variant
    Dim secNames() As String, secNos() As Long, tmpName As String, tmpNo As Long
    Dim sheetName As String, students As Long, info As String
    Dim defaultCount As Long, made As Long, totalStudents As Long
    Dim notThird As Long, noGroup As String, noGroupCount As Long
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

    ' ==================== Known streams (third year, section number removed) ====================
    knownNames = Array("Philo", "GeEc", "Math", "SE", "LaEtr")
    knownCodes = Array( _
        "062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A 0020 0622 062F 0627 0628 0020 0648 0641 0644 0633 0641 0629", _
        "062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A 0020 062A 0633 064A 064A 0631 0020 0648 0627 0642 062A 0635 0627 062F", _
        "062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A 0020 0631 064A 0627 0636 064A 0627 062A", _
        "062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A 0020 0639 0644 0648 0645 0020 062A 062C 0631 064A 0628 064A 0629", _
        "062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A 0020 0644 063A 0627 062A 0020 0623 062C 0646 0628 064A 0629 0020 0623 0644 0645 0627 0646 064A 0629")

    Set NameByKey = CreateObject("Scripting.Dictionary"): NameByKey.CompareMode = vbTextCompare
    Set Groups = CreateObject("Scripting.Dictionary"):    Groups.CompareMode = vbTextCompare
    Set GroupSecs = CreateObject("Scripting.Dictionary"): GroupSecs.CompareMode = vbTextCompare
    Set GroupText = CreateObject("Scripting.Dictionary"): GroupText.CompareMode = vbTextCompare

    For i = 0 To UBound(knownNames)
        NameByKey.Add NormAr(HexText(CStr(knownCodes(i)))), CStr(knownNames(i))
    Next i

    phraseNorm = NormAr(HexText("062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A"))

    ' ==================== Keep only the third-year sheets, grouped by stream ====================
    For Each ws In wbTerm.Worksheets
        hText = KeyText(CellText(ws.Range("H1")))
        tText = CellText(ws.Range("B1"))

        If InStr(NormAr(tText & " " & hText), phraseNorm) = 0 Then
            notThird = notThird + 1
        ElseIf Len(hText) = 0 Then
            noGroupCount = noGroupCount + 1
            noGroup = noGroup & "  - sheet " & ws.Name & " (no group title after the school year)" & vbCrLf
        Else
            baseText = StripSectionNumber(hText)
            secNo = SectionNumberOf(hText)
            key = NormAr(baseText)
            If Not Groups.Exists(key) Then
                Groups.Add key, New Collection
                GroupSecs.Add key, New Collection
                GroupText.Add key, baseText
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

    ' Output order: the known streams first (as listed), then any new stream
    Set OrderKeys = New Collection
    For i = 0 To UBound(knownNames)
        key = NormAr(HexText(CStr(knownCodes(i))))
        If Groups.Exists(key) Then OrderKeys.Add key
    Next i
    For Each kv In Groups.Keys
        If Not NameByKey.Exists(CStr(kv)) Then OrderKeys.Add CStr(kv)
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

        If NameByKey.Exists(key) Then sheetName = NameByKey.Item(key) Else sheetName = CStr(GroupText.Item(key))
        sheetName = SafeSheetName(sheetName, wbOut)

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

    If noGroupCount > 0 Then
        Report = Report & vbCrLf & "Sheets without a group title (not used):" & vbCrLf & noGroup
    End If
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

' ==============================================================================
' SOURCE01 BUILD - Collector01 logic embedded, so this module needs no other module
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
