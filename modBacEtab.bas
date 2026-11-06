Attribute VB_Name = "modBacEtab"
Option Explicit

' =====================================================================
'  BAC results splitter  -  macro: Run_BacEtab
'  Updates:
'   - Reads source files exclusively from subfolder "BacSource" (next to this tool).
'   - Dynamic source file discovery: accepts any Excel filename/extension,
'     selecting the file whose columns match LooksLikeData and LoadData.
'   - Saves the output workbook inside folder "معالجة نتائج البكالوريا" (created if missing).
' =====================================================================

Private Const SOURCE_FOLDER_NAME As String = "BacSource"
Private Const OUT_FOLDER_NAME As String = "معالجة نتائج البكالوريا"
Private Const OUT_FOLDER_ALT As String = "معالجه نتائج البكالوريا"
Private Const SOURCE_SHEET As String = "Data"

Private Const TITLE_ROW As Long = 1
Private Const HEADER_ROW As Long = 2
Private Const FIRST_DATA_ROW As Long = 3
Private Const NUM_SRC_COLS As Long = 14
Private Const FIRST_NUM_COL As Long = 6
Private Const MAX_CODE_DIGITS As Long = 15

Private Const TITLE_ROW_HEIGHT As Double = 40
Private Const HEADER_ROW_HEIGHT As Double = 100
Private Const DATA_ROW_HEIGHT As Double = 24
Private Const ID_COL_WIDTH As Double = 4.3
Private Const NUM_COL_WIDTH As Double = 6.3

Private Const CHECK_FONTS As Boolean = True
Private Const FONT_BASE As String = "Dubai"
Private Const FONT_NAME As String = "Dubai Medium"
Private Const FONT_NUM As String = "Aptos Slab"
Private Const FONT_SAFE_TEXT As String = "Arial"
Private Const FONT_SAFE_NUM As String = "Calibri"

Public gLastBacEtabPath As String
Public gExtractedYear As String   ' Stores 4-digit year extracted from source filename if present

Private Const ST_NONE As Long = 0
Private Const ST_OK As Long = 1
Private Const ST_CANCEL As Long = 2
Private Const ST_FAIL As Long = 3
Private mAuto As Boolean
Private mStatus As Long
Private mDetail As String

Private mCEtab As Long, mCSerie As Long, mCNom As Long, mCPre As Long
Private mCDay As Long, mCMon As Long, mCYear As Long, mCSex As Long
Private mCDesi As Long
Private mNumCols(1 To 14) As Long
Private mFontBase As String, mFontName As String, mFontNum As String, mFontNotes As String

' ---------------------------------------------------------------------
Public Sub Run_BacEtab()
    Dim source As Workbook, target As Workbook, dataSheet As Worksheet, ws As Worksheet
    Dim openedSource As Boolean, savedTarget As Boolean, Failed As Boolean
    Dim cancelled As Boolean, stateSaved As Boolean, loaded As Boolean
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean, oldStatus As Variant
    Dim errText As String, baseFolder As String, sourceFolder As String, sourceName As String, sourcePath As String
    Dim yearText As String, outFolder As String, outputPath As String, codeText As String
    Dim prefill As String, errMsg As String, problems As String, summary As String
    Dim vals As Variant, codeNum As Double, total As Long, otherSeries As Long
    Dim cnt(1 To 9) As Long, grp() As Byte, data() As Variant, hdr(1 To 19) As Variant
    Dim k As Long, sheetsMade As Long, nCols As Long
    Dim desiName As String, layoutIssues As String

    mStatus = ST_NONE
    mDetail = ""
    gExtractedYear = ""
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldStatus = Application.StatusBar
    stateSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    baseFolder = ThisWorkbook.Path
    If Len(baseFolder) = 0 Then Err.Raise vbObjectError + 901, "modBacEtab", Ar("save_tool")
    
    sourceFolder = baseFolder & Application.PathSeparator & SOURCE_FOLDER_NAME & Application.PathSeparator
    If Not FolderExists(sourceFolder) Then
        Err.Raise vbObjectError + 902, "modBacEtab", Ar("no_src_folder") & vbLf & sourceFolder
    End If

    If Not FindDynamicSourceFile(sourceFolder, sourceName, yearText) Then
        Err.Raise vbObjectError + 902, "modBacEtab", Ar("no_source") & vbLf & sourceFolder
    End If
    gExtractedYear = yearText
    sourcePath = sourceFolder & sourceName
    prefill = "47000000"

    Do
        If Not AskCode(prefill, errMsg, sourceName, codeText) Then
            cancelled = True
            GoTo Finished
        End If
        prefill = codeText
        codeNum = CDbl(codeText)

        If Not loaded Then
            Application.StatusBar = Ar("working") & " " & sourceName
            Set source = OpenSource(sourcePath, openedSource)
            Set dataSheet = FindDataSheet(source)
            If dataSheet Is Nothing Then
                Err.Raise vbObjectError + 903, "modBacEtab", Ar("no_sheet") & vbLf & "Cd_etab , Cd_serie"
            End If
            If Not LoadData(dataSheet, vals, problems) Then
                Err.Raise vbObjectError + 904, "modBacEtab", problems
            End If
            loaded = True
            If openedSource Then
                source.Close SaveChanges:=False
                openedSource = False
            End If
            Set source = Nothing
        End If

        total = CountMatches(vals, codeNum, cnt, grp, otherSeries)
        If total > 0 Then Exit Do
        If otherSeries > 0 Then errMsg = Ar("other_series") Else errMsg = Ar("not_found")
    Loop

    Set target = Application.Workbooks.Add(xlWBATWorksheet)
    ResolveFonts target.Worksheets(1)

    For k = 1 To 9
        If cnt(k) > 0 Then
            nCols = 8 + SubjectCount(k)
            ReDim data(1 To cnt(k), 1 To nCols)
            BuildHeaders k, hdr
            FillData data, vals, grp, k
            sheetsMade = sheetsMade + 1
            Application.StatusBar = Ar("working") & " " & SeriesSheetName(k)
            If sheetsMade = 1 Then
                Set ws = target.Worksheets(1)
            Else
                Set ws = target.Worksheets.Add(After:=target.Worksheets(target.Worksheets.Count))
            End If
            BuildSheet ws, k, data, cnt(k), hdr, yearText
        End If
    Next k

    For Each ws in target.Worksheets
        nCols = LastDataCol(ws)
        ApplyLayout ws, nCols
        layoutIssues = layoutIssues & CheckLayout(ws, nCols)
    Next ws

    target.Worksheets(1).Activate

    outFolder = EnsureOutputFolder(baseFolder)
    If Len(outFolder) = 0 Then Err.Raise vbObjectError + 908, "modBacEtab", Ar("no_out_folder")

    desiName = GetDesiValue(vals, codeNum)
    If Len(desiName) > 0 Then
        outputPath = NewOutputPath(outFolder, SafeFileName(desiName))
    Else
        outputPath = NewOutputPath(outFolder, "BAC" & IIf(Len(yearText) > 0, yearText, "Results") & "_" & Format$(codeNum, "0"))
    End If

    target.CheckCompatibility = False
    target.SaveAs fileName:=outputPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    If Not FileExists(outputPath) Then Err.Raise vbObjectError + 905, "modBacEtab", Ar("save_chk") & vbLf & outputPath
    gLastBacEtabPath = outputPath
    savedTarget = True

Finished:
    On Error Resume Next
    If openedSource And Not source Is Nothing Then source.Close SaveChanges:=False
    If Failed And Not savedTarget And Not target Is Nothing Then target.Close SaveChanges:=False
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
        mDetail = outputPath & vbLf & Ar("sheets") & ": " & sheetsMade & "  |  " & Ar("pupils") & ": " & total
        If Len(layoutIssues) > 0 Then mDetail = mDetail & vbLf & Ar("layout_warn") & layoutIssues
        If Not mAuto Then
            summary = Ar("done") & vbLf & outputPath & vbLf & _
                      Ar("sheets") & ": " & sheetsMade & vbLf & Ar("pupils") & ": " & total & vbLf
            For k = 1 To 9
                If cnt(k) > 0 Then summary = summary & vbLf & SeriesSheetName(k) & " : " & cnt(k)
            Next k
            If Len(mFontNotes) > 0 Then summary = summary & vbLf & vbLf & Ar("fonts") & mFontNotes
            If Len(layoutIssues) > 0 Then summary = summary & vbLf & vbLf & Ar("layout_warn") & layoutIssues
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

Public Sub BacEtab_SetAuto(ByVal v As Boolean)
    mAuto = v
End Sub

Public Function BacEtab_LastStatus() As Long
    BacEtab_LastStatus = mStatus
End Function

Public Function BacEtab_LastDetail() As String
    BacEtab_LastDetail = mDetail
End Function

Private Function AskCode(ByVal prefill As String, ByVal errMsg As String, ByVal infoText As String, ByRef code As String) As Boolean
    Dim f As Object
    On Error Resume Next
    Set f = VBA.UserForms.Add("frmEtabCode")
    On Error GoTo 0
    If f Is Nothing Then
        AskCode = AskCodeFallback(prefill, errMsg, code)
        Exit Function
    End If
    f.Setup prefill, errMsg, infoText
    f.Show
    If f.Accepted Then
        code = NormalizeDigits(CStr(f.codeText))
        AskCode = (Len(code) > 0)
    End If
    Unload f
    Set f = Nothing
End Function

Private Function AskCodeFallback(ByVal prefill As String, ByVal errMsg As String, ByRef code As String) As Boolean
    Dim v As Variant, s As String, prompt As String
    prompt = Ar("ask_prompt")
    If Len(errMsg) > 0 Then prompt = errMsg & vbLf & vbLf & prompt
    Do
        v = Application.InputBox(prompt:=prompt, Title:=Ar("app_title"), Default:=prefill, Type:=2)
        If VarType(v) = vbBoolean Then Exit Function
        s = NormalizeDigits(CStr(v))
        If Len(s) > 0 And Len(s) <= MAX_CODE_DIGITS Then
            If Not (s Like "*[!0-9]*") Then
                code = s
                AskCodeFallback = True
                Exit Function
            End If
        End If
        prompt = Ar("err_digits") & vbLf & vbLf & Ar("ask_prompt")
        prefill = CStr(v)
    Loop
End Function

' ---------------------------------------------------------------------
' Dynamic Source Discovery inside "BacSource"
' ---------------------------------------------------------------------
Private Function FolderExists(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Right$(p, 1) = Application.PathSeparator Then p = Left$(p, Len(p) - 1)
    FolderExists = fso.FolderExists(p)
    If Err.Number <> 0 Then FolderExists = False: Err.Clear
End Function

Private Function EnsureOutputFolder(ByVal baseFolder As String) As String
    Dim fso As Object, p1 As String, p2 As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    p1 = baseFolder & Application.PathSeparator & OUT_FOLDER_NAME
    p2 = baseFolder & Application.PathSeparator & OUT_FOLDER_ALT
    If fso.FolderExists(p1) Then
        EnsureOutputFolder = p1
        Exit Function
    ElseIf fso.FolderExists(p2) Then
        EnsureOutputFolder = p2
        Exit Function
    End If
    fso.CreateFolder p1
    If fso.FolderExists(p1) Then EnsureOutputFolder = p1
    Err.Clear
End Function

Private Function ExtractYearFromName(ByVal fileName As String) As String
    Dim rx As Object, matches As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "(19|20)\d{2}"
    rx.Global = False
    Set matches = rx.Execute(fileName)
    If matches.Count > 0 Then ExtractYearFromName = matches(0).Value
End Function

' Scans all Excel files in BacSource folder; validates columns against LooksLikeData/LoadData
Private Function FindDynamicSourceFile(ByVal folderPath As String, ByRef foundName As String, ByRef yearText As String) As Boolean
    Dim fileName As String, fullPath As String, ext As String
    Dim wbTest As Workbook, wsTest As Worksheet
    Dim problems As String, dummyVals As Variant

    fileName = Dir$(folderPath & "*.xls*")
    Do While Len(fileName) > 0
        If Left$(fileName, 2) <> "~$" Then
            ext = LCase$(Mid$(fileName, InStrRev(fileName, ".") + 1))
            If ext = "xlsx" Or ext = "xls" Or ext = "xlsm" Or ext = "xlsb" Then
                fullPath = folderPath & fileName
                On Error Resume Next
                Set wbTest = Workbooks.Open(Filename:=fullPath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
                If Not wbTest Is Nothing Then
                    Set wsTest = FindDataSheet(wbTest)
                    If Not wsTest Is Nothing Then
                        If LoadData(wsTest, dummyVals, problems) Then
                            foundName = fileName
                            yearText = ExtractYearFromName(fileName)
                            wbTest.Close SaveChanges:=False
                            FindDynamicSourceFile = True
                            Exit Function
                        End If
                    End If
                    wbTest.Close SaveChanges:=False
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
        fileName = Dir$
    Loop
End Function

Private Function OpenSource(ByVal fullPath As String, ByRef opened As Boolean) As Workbook
    Dim book As Workbook, nm As String
    nm = Mid$(fullPath, InStrRev(fullPath, Application.PathSeparator) + 1)
    For Each book In Application.Workbooks
        If StrComp(book.FullName, fullPath, vbTextCompare) = 0 Then
            Set OpenSource = book
            Exit Function
        End If
        If StrComp(book.Name, nm, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 906, "modBacEtab", Ar("same_name_open") & vbLf & book.FullName
        End If
    Next book
    If Not FileExists(fullPath) Then Err.Raise vbObjectError + 907, "modBacEtab", Ar("missing_source") & vbLf & fullPath
    Set OpenSource = Application.Workbooks.Open(fileName:=fullPath, UpdateLinks:=0, ReadOnly:=True, _
                                              IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
    opened = True
End Function

Private Function LooksLikeData(ByVal ws As Worksheet) As Boolean
    Dim lastCol As Long, hdr As Variant, c As Long, a As Boolean, b As Boolean
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then Exit Function
    hdr = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value2
    For c = 1 To lastCol
        Select Case LCase$(Trim$(SafeText(hdr(1, c))))
            Case "cd_etab": a = True
            Case "cd_serie": b = True
        End Select
    Next c
    LooksLikeData = (a And b)
End Function

Private Function FindDataSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SOURCE_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        If LooksLikeData(ws) Then
            Set FindDataSheet = ws
            Exit Function
        End If
    End If
    For Each ws In wb.Worksheets
        If LooksLikeData(ws) Then
            Set FindDataSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function ReqCol(ByRef hdr As Variant, ByVal lastCol As Long, ByVal nm As String, ByRef problems As String) As Long
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(SafeText(hdr(1, c))), nm, vbTextCompare) = 0 Then
            ReqCol = c
            Exit Function
        End If
    Next c
    problems = problems & vbLf & nm
End Function

Private Function FindCol(ByRef hdr As Variant, ByVal lastCol As Long, ByVal nm As String) As Long
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(SafeText(hdr(1, c))), nm, vbTextCompare) = 0 Then
            FindCol = c
            Exit Function
        End If
    Next c
End Function

Private Function LoadData(ByVal ws As Worksheet, ByRef vals As Variant, ByRef problems As String) As Boolean
    Dim lastCol As Long, LastRow As Long, hdr As Variant, j As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    hdr = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value2

    problems = ""
    mCEtab = ReqCol(hdr, lastCol, "Cd_etab", problems)
    mCSerie = ReqCol(hdr, lastCol, "Cd_serie", problems)
    mCNom = ReqCol(hdr, lastCol, "Nom", problems)
    mCPre = ReqCol(hdr, lastCol, "Prenom", problems)
    mCDay = ReqCol(hdr, lastCol, "Date_jour", problems)
    mCMon = ReqCol(hdr, lastCol, "Date_mois", problems)
    mCYear = ReqCol(hdr, lastCol, "Date_annee", problems)
    mCSex = ReqCol(hdr, lastCol, "Sexe", problems)
    mCDesi = FindCol(hdr, lastCol, "Desi")
    mNumCols(1) = ReqCol(hdr, lastCol, "Eps", problems)
    For j = 1 To 11
        mNumCols(1 + j) = ReqCol(hdr, lastCol, "Notef" & CStr(j), problems)
    Next j
    mNumCols(13) = ReqCol(hdr, lastCol, "Moyenne", problems)
    mNumCols(14) = ReqCol(hdr, lastCol, "MoyenneMe", problems)
    If Len(problems) > 0 Then
        problems = Ar("bad_headers") & problems
        Exit Function
    End If

    LastRow = ws.Cells(ws.Rows.Count, mCEtab).End(xlUp).Row
    If LastRow < 2 Then
        problems = Ar("no_rows")
        Exit Function
    End If
    vals = ws.Range(ws.Cells(1, 1), ws.Cells(LastRow, lastCol)).Value2
    LoadData = True
End Function

Private Function CountMatches(ByRef vals As Variant, ByVal codeNum As Double, ByRef cnt() As Long, _
                              ByRef grp() As Byte, ByRef other As Long) As Long
    Dim r As Long, n As Long, k As Long, c As Double, s As Double
    n = UBound(vals, 1)
    ReDim grp(1 To n)
    For k = 1 To 9
        cnt(k) = 0
    Next k
    other = 0
    For r = 2 To n
        If ToCode(vals(r, mCEtab), c) Then
            If c = codeNum Then
                If ToCode(vals(r, mCSerie), s) Then
                    If s >= 30 And s <= 38 Then
                        k = CLng(s) - 29
                        cnt(k) = cnt(k) + 1
                        grp(r) = CByte(k)
                        CountMatches = CountMatches + 1
                    Else
                        other = other + 1
                    End If
                Else
                    other = other + 1
                End If
            End If
        End If
    Next r
End Function

Private Function GetDesiValue(ByRef vals As Variant, ByVal codeNum As Double) As String
    Dim r As Long, n As Long, c As Double, s As Double, t As String
    If mCDesi <= 0 Then Exit Function
    n = UBound(vals, 1)
    For r = 2 To n
        If ToCode(vals(r, mCEtab), c) Then
            If c = codeNum Then
                If ToCode(vals(r, mCSerie), s) Then
                    If s >= 30 And s <= 38 Then
                        t = Trim$(SafeText(vals(r, mCDesi)))
                        If Len(t) > 0 Then
                            GetDesiValue = t
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, CStr(bad), "_")
    Next bad
    s = Trim$(s)
    Do While Len(s) > 0
        If Right$(s, 1) = "." Or Right$(s, 1) = " " Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    If Len(s) > 150 Then s = Left$(s, 150)
    SafeFileName = s
End Function

Private Function SubjectCount(ByVal k As Long) As Long
    Select Case k
        Case 1: SubjectCount = 8
        Case 2: SubjectCount = 9
        Case 3: SubjectCount = 11
        Case Else: SubjectCount = 10
    End Select
End Function

Private Function Subj(ByVal id As Long) As String
    Select Case id
        Case 1:  Subj = ArHex("627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627")
        Case 2:  Subj = ArHex("627 644 641 644 633 641 629")
        Case 3:  Subj = ArHex("627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627")
        Case 4:  Subj = ArHex("627 644 644 63A 629 20 627 644 641 631 646 633 64A 629")
        Case 5:  Subj = ArHex("627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629")
        Case 6:  Subj = ArHex("627 644 631 64A 627 636 64A 627 62A")
        Case 7:  Subj = ArHex("627 644 644 63A 629 20 627 644 623 645 627 632 64A 63A 64A 629")
        Case 8:  Subj = ArHex("627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629")
        Case 9:  Subj = ArHex("627 644 644 63A 629 20 627 644 623 62C 646 628 64A 629 20 627 644 62B 627 644 62B 629")
        Case 10: Subj = ArHex("62A 2E 20 627 644 645 62D 627 633 628 64A 20 648 627 644 645 627 644 64A")
        Case 11: Subj = ArHex("627 644 627 642 62A 635 627 62F 20 648 627 644 645 646 627 62C 645 646 62A")
        Case 12: Subj = ArHex("627 644 642 627 646 648 646")
        Case 13: Subj = ArHex("639 644 648 645 20 637 628 64A 639 64A 629 20 648 62D 64A 627 629")
        Case 14: Subj = ArHex("627 644 639 644 648 645 20 627 644 641 64A 632 64A 627 626 64A 629")
        Case 15: Subj = ArHex("647 646 62F 633 629 20 645 64A 643 627 646 64A 643 64A 629")
        Case 16: Subj = ArHex("647 646 62F 633 629 20 643 647 631 628 627 626 64A 629")
        Case 17: Subj = ArHex("647 646 62F 633 629 20 645 62F 646 64A 629")
        Case 18: Subj = ArHex("647 646 62F 633 629 20 637 631 627 626 642")
    End Select
End Function

Private Function NotefHeaderAr(ByVal k As Long, ByVal j As Long) As String
    Select Case k
        Case 1
            Select Case j
                Case 1: NotefHeaderAr = Subj(1)
                Case 2: NotefHeaderAr = Subj(2)
                Case 3: NotefHeaderAr = Subj(3)
                Case 4: NotefHeaderAr = Subj(4)
                Case 5: NotefHeaderAr = Subj(5)
                Case 6: NotefHeaderAr = Subj(6)
                Case 7: NotefHeaderAr = Subj(7)
                Case 8: NotefHeaderAr = Subj(8)
            End Select
        Case 2
            Select Case j
                Case 1: NotefHeaderAr = Subj(1)
                Case 2: NotefHeaderAr = Subj(4)
                Case 3: NotefHeaderAr = Subj(5)
                Case 4: NotefHeaderAr = Subj(9)
                Case 5: NotefHeaderAr = Subj(2)
                Case 6: NotefHeaderAr = Subj(3)
                Case 7: NotefHeaderAr = Subj(6)
                Case 8: NotefHeaderAr = Subj(7)
                Case 9: NotefHeaderAr = Subj(8)
            End Select
        Case 3
            Select Case j
                Case 1:  NotefHeaderAr = Subj(10)
                Case 2:  NotefHeaderAr = Subj(11)
                Case 3:  NotefHeaderAr = Subj(6)
                Case 4:  NotefHeaderAr = Subj(3)
                Case 5:  NotefHeaderAr = Subj(1)
                Case 6:  NotefHeaderAr = Subj(12)
                Case 7:  NotefHeaderAr = Subj(4)
                Case 8:  NotefHeaderAr = Subj(5)
                Case 9:  NotefHeaderAr = Subj(2)
                Case 10: NotefHeaderAr = Subj(7)
                Case 11: NotefHeaderAr = Subj(8)
            End Select
        Case 4
            Select Case j
                Case 1:  NotefHeaderAr = Subj(13)
                Case 2:  NotefHeaderAr = Subj(14)
                Case 3:  NotefHeaderAr = Subj(6)
                Case 4:  NotefHeaderAr = Subj(1)
                Case 5:  NotefHeaderAr = Subj(4)
                Case 6:  NotefHeaderAr = Subj(5)
                Case 7:  NotefHeaderAr = Subj(2)
                Case 8:  NotefHeaderAr = Subj(3)
                Case 9:  NotefHeaderAr = Subj(7)
                Case 10: NotefHeaderAr = Subj(8)
            End Select
        Case 5
            Select Case j
                Case 1:  NotefHeaderAr = Subj(6)
                Case 2:  NotefHeaderAr = Subj(14)
                Case 3:  NotefHeaderAr = Subj(1)
                Case 4:  NotefHeaderAr = Subj(13)
                Case 5:  NotefHeaderAr = Subj(3)
                Case 6:  NotefHeaderAr = Subj(4)
                Case 7:  NotefHeaderAr = Subj(5)
                Case 8:  NotefHeaderAr = Subj(2)
                Case 9:  NotefHeaderAr = Subj(7)
                Case 10: NotefHeaderAr = Subj(8)
            End Select
        Case 6
            Select Case j
                Case 1:  NotefHeaderAr = Subj(15)
                Case 2:  NotefHeaderAr = Subj(6)
                Case 3:  NotefHeaderAr = Subj(14)
                Case 4:  NotefHeaderAr = Subj(1)
                Case 5:  NotefHeaderAr = Subj(4)
                Case 6:  NotefHeaderAr = Subj(5)
                Case 7:  NotefHeaderAr = Subj(2)
                Case 8:  NotefHeaderAr = Subj(3)
                Case 9:  NotefHeaderAr = Subj(7)
                Case 10: NotefHeaderAr = Subj(8)
            End Select
        Case 7
            Select Case j
                Case 1:  NotefHeaderAr = Subj(16)
                Case 2:  NotefHeaderAr = Subj(6)
                Case 3:  NotefHeaderAr = Subj(14)
                Case 4:  NotefHeaderAr = Subj(1)
                Case 5:  NotefHeaderAr = Subj(4)
                Case 6:  NotefHeaderAr = Subj(5)
                Case 7:  NotefHeaderAr = Subj(2)
                Case 8:  NotefHeaderAr = Subj(3)
                Case 9:  NotefHeaderAr = Subj(7)
                Case 10: NotefHeaderAr = Subj(8)
            End Select
        Case 8
            Select Case j
                Case 1:  NotefHeaderAr = Subj(17)
                Case 2:  NotefHeaderAr = Subj(6)
                Case 3:  NotefHeaderAr = Subj(14)
                Case 4:  NotefHeaderAr = Subj(1)
                Case 5:  NotefHeaderAr = Subj(4)
                Case 6:  NotefHeaderAr = Subj(5)
                Case 7:  NotefHeaderAr = Subj(2)
                Case 8:  NotefHeaderAr = Subj(3)
                Case 9:  NotefHeaderAr = Subj(7)
                Case 10: NotefHeaderAr = Subj(8)
            End Select
        Case 9
            Select Case j
                Case 1:  NotefHeaderAr = Subj(18)
                Case 2:  NotefHeaderAr = Subj(6)
                Case 3:  NotefHeaderAr = Subj(14)
                Case 4:  NotefHeaderAr = Subj(1)
                Case 5:  NotefHeaderAr = Subj(4)
                Case 6:  NotefHeaderAr = Subj(5)
                Case 7:  NotefHeaderAr = Subj(2)
                Case 8:  NotefHeaderAr = Subj(3)
                Case 9:  NotefHeaderAr = Subj(7)
                Case 10: NotefHeaderAr = Subj(8)
            End Select
    End Select
End Function

Private Sub BuildHeaders(ByVal k As Long, ByRef hdr() As Variant)
    Dim j As Long, m As Long
    m = SubjectCount(k)
    hdr(1) = Ar("h_num")
    hdr(2) = Ar("h_name")
    hdr(3) = Ar("h_birth")
    hdr(4) = Ar("h_sex")
    hdr(5) = Ar("h_repeat")
    hdr(6) = Ar("h_eps")
    For j = 1 To m
        hdr(6 + j) = NotefHeaderAr(k, j)
    Next j
    hdr(7 + m) = Ar("h_avg")
    hdr(8 + m) = Ar("h_avgme")
End Sub

Private Sub FillData(ByRef data() As Variant, ByRef vals As Variant, ByRef grp() As Byte, ByVal k As Long)
    Dim r As Long, p As Long, j As Long, n As Long, m As Long
    n = UBound(vals, 1)
    m = SubjectCount(k)
    For r = 2 To n
        If grp(r) = k Then
            p = p + 1
            data(p, 2) = Trim$(SafeText(vals(r, mCNom)) & " " & SafeText(vals(r, mCPre)))
            data(p, 3) = BirthText(vals(r, mCDay), vals(r, mCMon), vals(r, mCYear))
            data(p, 4) = vals(r, mCSex)
            data(p, 6) = vals(r, mNumCols(1))
            For j = 1 To m
                data(p, 6 + j) = vals(r, mNumCols(1 + j))
            Next j
            data(p, 7 + m) = vals(r, mNumCols(13))
            data(p, 8 + m) = vals(r, mNumCols(14))
        End If
    Next r
    SortByAvgDesc data, p, 7 + m
End Sub

Private Function AvgKey(ByVal v As Variant) As Double
    AvgKey = -1
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    AvgKey = CDbl(v)
End Function

Private Sub SortByAvgDesc(ByRef data() As Variant, ByVal n As Long, ByVal avgCol As Long)
    Dim key() As Double, idx() As Long, tmp() As Variant
    Dim i As Long, j As Long, c As Long, t As Long, nc As Long
    If n < 1 Then Exit Sub
    nc = UBound(data, 2)
    ReDim key(1 To n)
    ReDim idx(1 To n)
    For i = 1 To n
        key(i) = AvgKey(data(i, avgCol))
        idx(i) = i
    Next i
    For i = 2 To n
        t = idx(i)
        j = i - 1
        Do While j >= 1
            If key(idx(j)) >= key(t) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = t
    Next i
    ReDim tmp(1 To n, 1 To nc)
    For i = 1 To n
        For c = 1 To nc
            tmp(i, c) = data(idx(i), c)
        Next c
    Next i
    For i = 1 To n
        For c = 1 To nc
            data(i, c) = tmp(i, c)
        Next c
        data(i, 1) = i
    Next i
End Sub

Private Function BirthText(ByVal d As Variant, ByVal m As Variant, ByVal y As Variant) As String
    Dim dn As Long, mn As Long, yn As Long
    dn = ToLongSafe(d)
    mn = ToLongSafe(m)
    yn = ToLongSafe(y)
    If dn = 0 And mn = 0 And yn = 0 Then Exit Function
    BirthText = PadNum(yn, 4) & "-" & PadNum(mn, 2) & "-" & PadNum(dn, 2)
End Function

Private Function PadNum(ByVal n As Long, ByVal width As Long) As String
    Dim s As String
    s = CStr(n)
    If Len(s) < width Then s = String$(width - Len(s), "0") & s
    PadNum = s
End Function

Private Function ToLongSafe(ByVal v As Variant) As Long
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    On Error Resume Next
    ToLongSafe = CLng(v)
    If Err.Number <> 0 Then ToLongSafe = 0
    On Error GoTo 0
    If ToLongSafe < 0 Then ToLongSafe = 0
End Function

Private Function ToCode(ByVal v As Variant, ByRef n As Double) As Boolean
    Dim s As String, d As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbLong, vbInteger, vbByte, vbDecimal
            d = CDbl(v)
            If d < 0 Or d <> Int(d) Or d > 999999999999999# Then Exit Function
            n = d
            ToCode = True
        Case vbString
            s = NormalizeDigits(CStr(v))
            If Len(s) = 0 Or Len(s) > MAX_CODE_DIGITS Then Exit Function
            If s Like "*[!0-9]*" Then Exit Function
            n = CDbl(s)
            ToCode = True
    End Select
End Function

Private Function NormalizeDigits(ByVal s As String) As String
    Dim i As Long
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(&HA0), "")
    s = Replace(s, ChrW(&H200E), "")
    s = Replace(s, ChrW(&H200F), "")
    For i = 0 To 9
        s = Replace(s, ChrW(&H660 + i), CStr(i))
        s = Replace(s, ChrW(&H6F0 + i), CStr(i))
    Next i
    NormalizeDigits = s
End Function

Private Function SafeText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    SafeText = CStr(v)
End Function

Private Function SeriesSheetName(ByVal k As Long) As String
    SeriesSheetName = CStr(Choose(k, "Philo", "LaEtr", "GeEc", "SE", "Math", "TGM", "TGE", "TGC", "TGP"))
End Function

Private Sub BuildSheet(ByVal ws As Worksheet, ByVal k As Long, ByRef data() As Variant, ByVal n As Long, _
                       ByRef hdr() As Variant, ByVal yearText As String)
    Dim LastRow As Long, area As Range, lo As ListObject, nCols As Long
    nCols = 8 + SubjectCount(k)

    LastRow = FIRST_DATA_ROW + n - 1
    If LastRow > ws.Rows.Count Then Err.Raise vbObjectError + 905, "modBacEtab", Ar("too_many_rows") & " " & SeriesSheetName(k)

    ws.Name = SeriesSheetName(k)
    ws.DisplayRightToLeft = True
    ws.Columns("A").ColumnWidth = ID_COL_WIDTH
    ws.Columns("B").AutoFit
    ws.Columns("C").ColumnWidth = 10.5
    ws.Range(ws.Columns("D"), ws.Columns(nCols)).ColumnWidth = NUM_COL_WIDTH

    Set area = ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(LastRow, nCols))

    ws.Range(ws.Cells(FIRST_DATA_ROW, 2), ws.Cells(LastRow, 3)).NumberFormat = "@"
    ws.Range(ws.Cells(FIRST_DATA_ROW, 5), ws.Cells(LastRow, 5)).NumberFormat = "@"
    ws.Range(ws.Cells(FIRST_DATA_ROW, FIRST_NUM_COL), ws.Cells(LastRow, nCols)).NumberFormat = "00.00"

    ws.Cells(HEADER_ROW, 1).Resize(1, nCols).Value2 = hdr
    ws.Cells(FIRST_DATA_ROW, 1).Resize(n, nCols).Value2 = data

    With area
        .Font.Name = mFontBase: .Font.Size = 14: .Font.Bold = False
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .Rows.RowHeight = DATA_ROW_HEIGHT
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
    With ws.Range(ws.Cells(FIRST_DATA_ROW, 2), ws.Cells(LastRow, 2))
        .Font.Name = mFontName: .Font.Size = 14
        .ShrinkToFit = True
    End With
    With ws.Range(ws.Cells(FIRST_DATA_ROW, FIRST_NUM_COL), ws.Cells(LastRow, nCols))
        .Font.Name = mFontNum: .Font.Size = 15
    End With

    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=area, XlListObjectHasHeaders:=xlYes)
    lo.Name = "T_" & ws.Name
    lo.TableStyle = "": lo.ShowTotals = False
    lo.ShowTableStyleRowStripes = False: lo.ShowTableStyleColumnStripes = False: lo.ShowAutoFilter = True
    With lo.HeaderRowRange
        .Interior.Color = RGB(28, 55, 103): .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True: .WrapText = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = HEADER_ROW_HEIGHT
    End With
    With ws.Range(ws.Cells(HEADER_ROW, FIRST_NUM_COL), ws.Cells(HEADER_ROW, nCols))
        .WrapText = True
        .Orientation = 90
    End With

    ws.Rows(TITLE_ROW).RowHeight = TITLE_ROW_HEIGHT
    With ws.Range("B1:F1")
        .Merge
        .Value2 = Ar("title_main") & IIf(Len(yearText) > 0, " " & yearText, "")
    End With
    With ws.Range("H1:L1")
        .Merge
        .Value2 = Ar("branch_word") & " " & Ar("s" & CStr(29 + k))
    End With
    StyleBanner ws.Range("B1:F1")
    StyleBanner ws.Range("H1:L1")

    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(LastRow, nCols)).Address
        .PrintTitleRows = "$1:$2"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    On Error GoTo 0
End Sub

Private Function LastDataCol(ByVal ws As Worksheet) As Long
    LastDataCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Sub ApplyLayout(ByVal ws As Worksheet, ByVal nCols As Long)
    If nCols < 4 Then Exit Sub
    ws.Columns(1).ColumnWidth = ID_COL_WIDTH
    ws.Columns(2).AutoFit
    ws.Range(ws.Columns(4), ws.Columns(nCols)).ColumnWidth = NUM_COL_WIDTH
    ws.Rows(TITLE_ROW).RowHeight = TITLE_ROW_HEIGHT
    ws.Rows(HEADER_ROW).RowHeight = HEADER_ROW_HEIGHT
End Sub

Private Function CheckLayout(ByVal ws As Worksheet, ByVal nCols As Long) As String
    Dim c As Long, s As String
    If Abs(ws.Columns(1).ColumnWidth - ID_COL_WIDTH) > 0.2 Then
        s = s & vbLf & ws.Name & " : A = " & Format$(ws.Columns(1).ColumnWidth, "0.00")
    End If
    For c = 4 To nCols
        If Abs(ws.Columns(c).ColumnWidth - NUM_COL_WIDTH) > 0.2 Then
            s = s & vbLf & ws.Name & " : " & Replace(ws.Cells(1, c).Address(True, False), "$1", "") & _
                " = " & Format$(ws.Columns(c).ColumnWidth, "0.00")
        End If
    Next c
    If Abs(ws.Rows(TITLE_ROW).RowHeight - TITLE_ROW_HEIGHT) > 1 Then
        s = s & vbLf & ws.Name & " : row 1 = " & Format$(ws.Rows(TITLE_ROW).RowHeight, "0.0")
    End If
    If Abs(ws.Rows(HEADER_ROW).RowHeight - HEADER_ROW_HEIGHT) > 1 Then
        s = s & vbLf & ws.Name & " : row 2 = " & Format$(ws.Rows(HEADER_ROW).RowHeight, "0.0")
    End If
    CheckLayout = s
End Function

Private Sub StyleBanner(ByVal rng As Range)
    With rng
        .Font.Name = mFontBase: .Font.Size = 16: .Font.Bold = True
        .Interior.Color = RGB(28, 55, 103): .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL
    End With
End Sub

Private Sub ResolveFonts(ByVal ws As Worksheet)
    mFontNotes = ""
    mFontBase = PickFont(ws, FONT_BASE, FONT_SAFE_TEXT)
    mFontName = PickFont(ws, FONT_NAME, FONT_BASE, FONT_SAFE_TEXT)
    mFontNum = PickFont(ws, FONT_NUM, FONT_SAFE_NUM, FONT_SAFE_TEXT)
    ws.Cells.Clear
End Sub

Private Function PickFont(ByVal ws As Worksheet, ParamArray names() As Variant) As String
    Dim i As Long, first As String, chosen As String
    first = CStr(names(LBound(names)))
    If Not CHECK_FONTS Then
        PickFont = first
        Exit Function
    End If
    chosen = CStr(names(UBound(names)))
    For i = LBound(names) To UBound(names) - 1
        If FontAvailable(ws, CStr(names(i))) Then
            chosen = CStr(names(i))
            Exit For
        End If
    Next i
    PickFont = chosen
    If StrComp(chosen, first, vbTextCompare) <> 0 Then
        If InStr(mFontNotes, first & " ->") = 0 Then mFontNotes = mFontNotes & vbLf & first & " -> " & chosen
    End If
End Function

Private Function FontAvailable(ByVal ws As Worksheet, ByVal fontName As String) As Boolean
    Dim samples(1 To 2) As String, i As Long, wBase As Double, wTest As Double
    samples(1) = "MWmwIl0123456789 The quick brown fox " & ArHex("627 644 639 631 628 64A 629 20 645 62F 631 633 629")
    samples(2) = "00.00 18.75 05.50 Notef11 MoyenneMe"
    On Error GoTo Unknown
    For i = 1 To 2
        wBase = MeasureWidth(ws, "Zz No Such Font Zz", samples(i))
        wTest = MeasureWidth(ws, fontName, samples(i))
        If Abs(wTest - wBase) > 0.3 Then
            FontAvailable = True
            Exit Function
        End If
    Next i
    Exit Function
Unknown:
    FontAvailable = True
End Function

Private Function MeasureWidth(ByVal ws As Worksheet, ByVal fontName As String, ByVal sample As String) As Double
    With ws.Range("A1")
        .Value2 = sample
        .Font.Name = fontName
        .Font.Size = 14
        .WrapText = False
    End With
    ws.Columns(1).AutoFit
    MeasureWidth = ws.Columns(1).ColumnWidth
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(fullPath)
    If Err.Number <> 0 Then FileExists = False: Err.Clear
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

Private Function NewOutputPath(ByVal folder As String, ByVal stem As String) As String
    Dim fullPath As String, n As Long
    If Right$(folder, 1) = Application.PathSeparator Then folder = Left$(folder, Len(folder) - 1)
    fullPath = folder & Application.PathSeparator & stem & ".xlsx"
    n = 0
    Do While PathInUse(fullPath)
        n = n + 1
        fullPath = folder & Application.PathSeparator & stem & "_" & Format$(n, "00") & ".xlsx"
    Loop
    NewOutputPath = fullPath
End Function

Private Function ArHex(ByVal codes As String) As String
    Dim token As Variant
    For Each token In Split(codes, " ")
        If Len(CStr(token)) > 0 Then ArHex = ArHex & ChrW(CLng("&H" & CStr(token)))
    Next token
End Function

Private Function Ar(ByVal key As String) As String
    Select Case key
    Case "prefix": Ar = ArHex("62A 62D 644 64A 644 20 646 62A 627 626 62C")
    Case "app_title": Ar = ArHex("646 62A 627 626 62C 20 627 644 628 643 627 644 648 631 64A 627")
    Case "ask_prompt": Ar = ArHex("64A 631 62C 649 20 625 62F 62E 627 644 20 643 648 62F 20 627 644 645 624 633 633 629")
    Case "err_digits": Ar = ArHex("627 644 643 648 62F 20 64A 62A 643 648 646 20 645 646 20 623 631 642 627 645 20 641 642 637")
    Case "working": Ar = ArHex("62C 627 631 64D 20 627 644 62A 62D 636 64A 631 3A")
    Case "save_tool": Ar = ArHex("627 62D 641 638 20 645 644 641 20 627 644 623 62F 627 629 20 623 648 644 627 64B 20 641 64A 20 645 62C 644 62F 60C 20 62B 645 20 636 639 20 645 62C 644 62F 20 627 644 645 635 62F 631 20 628 62C 627 646 628 647 2E")
    Case "no_src_folder": Ar = ArHex("645 62C 644 62F 20 627 644 645 635 62F 631 627 634 20 645 641 64A 64E 651 644 20 627 644 645 635 62F 631 3A")
    Case "no_source": Ar = ArHex("644 645 20 623 62C 62F 20 645 644 641 20 627 644 646 62A 627 626 62C 20 628 62C 627 646 628 20 645 644 641 20 627 644 623 62F 627 629 2E 20 627 644 627 633 645 627 644 645 637 644 648 628 3A")
    Case "same_name_open": Ar = ArHex("64A 648 62C 62F 20 645 644 641 20 628 627 644 627 633 645 20 646 641 633 647 20 645 641 62A 648 62D 20 645 646 20 645 62C 644 62F 20 622 62E 631 2E 20 623 63A 644 642 647 20 62B 645 20 623 639 62F 20 627 644 645 62D 627 648 644 629 3A")
    Case "missing_source": Ar = ArHex("645 644 641 20 627 644 646 62A 627 626 62C 20 63A 64A 631 20 645 648 62C 648 62F 3A")
    Case "no_sheet": Ar = ArHex("644 645 20 623 62C 62F 20 648 631 642 629 20 628 64A 627 646 627 62A 20 62A 62D 62A 648 64A 20 639 644 649 20 627 644 639 645 648 62F 64A 646 3A")
    Case "bad_headers": Ar = ArHex("62A 639 630 631 20 627 644 639 62B 648 631 639 644 649 647 630 647 627 644 623 639 645 62F 629 641 64A 648 631 642 629 627 644 628 64A 627 646 627 62A 3A")
    Case "no_rows": Ar = ArHex("648 631 642 629 20 627 644 628 64A 627 646 627 62A 20 644 627 20 62A 62D 62A 648 64A 639 644 649 23 623 64A 635 641 2E")
    Case "not_found": Ar = ArHex("644 645 20 62A 648 62C 62F 646 62A 627 626 62C 644 647 630 627 627 644 643 648 62F 641 64A 627 644 645 644 641")
    Case "other_series": Ar = ArHex("648 64F 62C 62F 20 627 644 643 648 62F 644 646 631 645 632 627 644 634 639 628 629 62E 627 631 62C 627 644 645 62C 627 644 627 644 645 633 645 648 62D")
    Case "too_many_rows": Ar = ArHex("639 62F 62F 627 644 635 641 648 641 64A 62A 62C 627 648 632 633 639 629 627 644 648 631 642 629 3A")
    Case "layout_warn": Ar = ArHex("62A 646 628 64A 647 3A 20 62A 639 630 631 636 628 637 627 644 623 628 639 627 62F 627 644 62A 627 644 64A 629 643 645 627 647 648 645 637 644 648 628 3A")
    Case "done": Ar = ArHex("62A 645 20 625 646 634 627 621 20 627 644 645 644 641 3A")
    Case "sheets": Ar = ArHex("639 62F 62F 20 627 644 623 648 631 627 642")
    Case "pupils": Ar = ArHex("639 62F 62F 20 627 644 62A 644 627 645 64A 630")
    Case "save_chk": Ar = ArHex("62A 639 630 651 631 20 627 644 62A 62D 642 642 20 645 646 20 648 62C 648 62F 20 645 644 641 20 627 644 646 62A 627 626 62C 20 628 639 62F 627 644 62D 641 638 3A")
    Case "st_cancel": Ar = ArHex("623 644 63A 649 20 627 644 645 633 62A 62E 62F 645 20 627 644 639 645 644 64A 629 2E")
    Case "failed": Ar = ArHex("62A 639 630 631 20 625 646 634 627 621 20 627 644 645 644 641 3A")
    Case "fonts": Ar = ArHex("62A 645 20 627 633 62A 628 62F 627 644 20 62E 637 648 637 20 63A 64A 631 20 645 62A 648 641 631 629 20 628 62E 637 648 637 20 628 62F 64A 644 629 3A")
    Case "title_main": Ar = ArHex("646 62A 627 626 62C 20 634 647 627 62F 629 20 627 644 628 643 627 644 648 631 64A 627")
    Case "branch_word": Ar = ArHex("634 639 628 629")
    Case "no_out_folder": Ar = ArHex("62A 639 630 631 20 625 646 634 627 621 20 645 62C 644 62F 627 629 20 627 644 625 635 62F 631 627 634 2E")
    Case "h_num": Ar = ArHex("627 644 631 642 645")
    Case "h_name": Ar = ArHex("627 644 644 642 628 20 648 627 644 627 633 645")
    Case "h_birth": Ar = ArHex("62A 627 631 64A 62E 20 627 644 645 64A 644 627 62F")
    Case "h_sex": Ar = ArHex("627 644 62C 646 633")
    Case "h_repeat": Ar = ArHex("627 644 625 639 627 62F 629")
    Case "h_eps": Ar = ArHex("62A 20 628 62F 646 64A 629 20 648 631 64A 627 636 64A 629")
    Case "h_avg": Ar = ArHex("627 644 645 639 62F 644")
    Case "h_avgme": Ar = ArHex("645 20 645 20 627 644 623 633 627 633 64A 629")
    Case "s30": Ar = ArHex("622 62F 627 628 20 648 641 644 633 641 629")
    Case "s31": Ar = ArHex("644 63A 627 62A 20 623 62C 646 628 64A 629")
    Case "s32": Ar = ArHex("62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F")
    Case "s33": Ar = ArHex("639 644 648 645 20 62A 62C 631 64A 628 64A 629")
    Case "s34": Ar = ArHex("631 64A 627 636 64A 627 62A")
    Case "s35": Ar = ArHex("62A 642 646 64A 20 631 64A 627 636 64A 20 2D 20 647 646 62F 633 629 20 645 64A 643 627 646 64A 643 64A 629")
    Case "s36": Ar = ArHex("62A 642 646 64A 20 631 64A 627 636 64A 20 2D 20 647 646 62F 633 629 20 643 647 631 628 627 626 64A 629")
    Case "s37": Ar = ArHex("62A 642 646 64A 20 631 64A 627 636 64A 20 2D 20 647 646 62F 633 629 20 645 62F 646 64A 629")
    Case "s38": Ar = ArHex("62A 642 646 64A 20 631 64A 627 636 64A 20 2D 20 647 646 62F 633 629 20 637 631 627 626 642")
    End Select
End Function