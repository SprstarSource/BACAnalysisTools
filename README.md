BACAnalysisTools
Excel VBA toolkit that extracts the complete Baccalaureate (BAC) results of a single school from the provincial results file, completes them with repeat status, and produces a statistical analysis workbook for every série and for the school as a whole.
---
1. Purpose and design challenge
The provincial BAC results file is a single flat table covering every candidate in the province: typically around 30,000 rows and columns extending to `BT`. Extracting one school's results from it by hand is slow and error-prone, because each candidate's marks sit in positional columns (`Notef1` … `Notef11`) whose meaning depends on the candidate's série.
This project automates the whole job in three steps:
Extraction – isolate every candidate of one school, group them by série, map each positional mark column to the correct subject, and rank candidates by their average. The result is a formatted school workbook named after the school.
Repeat status – complete the school workbook with each candidate's repeat status, taken from the school's third-term results.
Analysis – build statistical tables for each série and a consolidated summary for the whole school.
The original provincial file is only read. It is never modified, moved or deleted.
---
2. Pipeline
```
 Provincial file ─────────► [1] Extract ─────────► School workbook (.xlsx, named after the school)
 BAC<year>ResultatsOriginal.xlsx                    one sheet per série
 (read-only)                                                │
                                                            ▼
 Third-term results workbook ─► [2] Fetch repeat status ─► column E filled in the school workbook
 (read-only)                                                │
                                                            ▼
                                       [3] Analyse ─────► analysis copy: tables per série + "Gen" sheet
```
Step	Module	Macro	Responsibility
1	`modBacEtab`	`Run_BacEtab`	Extract the school's candidates, split by série, sort by average, issue the school workbook.
2	`modRepeatStatus`	`Run_RepeatStatus`	Write the repeat status of each candidate into column E of the school workbook.
3	`modBacAnalysis`	`Run_BacAnalysis`	Build the analysis tables and the `Gen` summary sheet in a copy of the school workbook.
Steps 2 and 3 offer the workbook produced by the previous step as their default input (session variable `gLastBacEtabPath`), so no file has to be located manually when the steps are run in order.
---
3. Step 1 – Extraction (`modBacEtab`, macro `Run_BacEtab`)
3.1 Input
Item	Requirement
File	`BAC????ResultatsOriginal.xlsx`, where `????` is the four-digit year. If several match, the latest year is used.
Location	The same folder as the tool workbook (which must therefore be saved).
Worksheet	`Data`. If absent, the first worksheet whose headers contain `Cd_etab` and `Cd_serie` is used.
Headers (row 1)	`Cd_etab`, `Cd_serie`, `Nom`, `Prenom`, `Date_jour`, `Date_mois`, `Date_annee`, `Sexe`, `Eps`, `Notef1` … `Notef11`, `Moyenne`, `MoyenneMe`. Optional: `Desi`.
Columns are resolved by header name, never by fixed position, so the width of the provincial file does not matter. The data range is read once into a memory array and processed there; the source workbook is opened read-only and closed without saving.
3.2 Processing
The user enters the school's national code (`Cd_etab`) in a dialog (`frmEtabCode`, or a standard `InputBox` if the form is not installed). The dialog validates the entry inline (not empty, digits only, maximum length); Arabic-Indic digits are normalised.
All rows whose `Cd_etab` matches are selected: these are the school's candidates. A code that exists only under series outside 30–38 produces a distinct message from a code that does not exist at all; in both cases the user is prompted again.
Each candidate is assigned to a série from `Cd_serie` (30 to 38).
Within each série, the marks (`Eps`, `Notef1` … `NotefN`) and the averages (`Moyenne`, `MoyenneMe`) are transferred, with `Notef1` … `NotefN` mapped to subject names through a hardcoded table keyed by `Cd_serie`. Unused `Notef` positions are skipped entirely: no column is created and no data is transferred.
Candidates are sorted by `Moyenne` in descending order (stable sort), and column A is numbered 1…n after sorting.
The data is placed in formatted tables, and the school workbook is issued with one sheet per série present for that school.
3.3 Conveniences applied during extraction
Surname (`Nom`) and first name (`Prenom`) are merged into a single name column.
Day, month and year columns are combined into one birth date, formatted `yyyy-mm-dd` and stored as text.
Sex is carried over from `Sexe`.
3.4 Séries
Code	Sheet	Série	Subject columns
30	`Philo`	Literature and Philosophy	8
31	`LaEtr`	Foreign Languages	9
32	`GeEc`	Management and Economics	11
33	`SE`	Experimental Sciences	10
34	`Math`	Mathematics	10
35	`TGM`	Technical Mathematics – Mechanical Engineering	10
36	`TGE`	Technical Mathematics – Electrical Engineering	10
37	`TGC`	Technical Mathematics – Civil Engineering	10
38	`TGP`	Technical Mathematics – Process Engineering	10
Subject names use unified Arabic spellings, so the same subject carries the same header in every série.
3.5 Output sheet layout
Column	Content
A	Rank (after descending sort by average)
B	Full name (surname + first name)
C	Birth date
D	Sex (`1` = male, `2` = female, as in the source)
E	Repeater flag (filled in Step 2)
F	Physical education (`Eps`)
G …	Subject columns for the série
next to last	Average (`Moyenne`)
last	Core-subjects average (`MoyenneMe`)
Each sheet is right-to-left, holds an Excel table named `T_<sheet>`, and is print-ready (landscape, fit to one page wide). Fixed layout: column A width 4.3; columns D to the last data column width 6.3; row 1 height 40; row 2 height 100 with rotated subject headers. The layout is applied last and verified by reading the values back; any deviation is reported in the final message.
Fonts: `Dubai` / `Dubai Medium` for text and `Aptos Slab` for numbers, with automatic fallback to `Arial` / `Calibri` when they are not installed (the substitution is reported).
3.6 Output file name
The workbook is saved as `.xlsx` in the tool folder. The name is the school's `Desi` value (sanitised, truncated to 150 characters) or, if unavailable, `BAC<year>_<code>`. An existing file is never overwritten: a sequential suffix (`_02`, `_03`, …) is added.
---
4. Step 2 – Repeat status (`modRepeatStatus`, macro `Run_RepeatStatus`)
The provincial file does not state whether a candidate is a repeater. This step reads it from the school's third-term results workbook, produced by the tools of the ResultsAnalysisTerm03 repository, and writes it into column E of the school workbook.
The repeat data is built in memory: no intermediate file is created. The source workbook is never modified.
4.1 Workflow
Source file. The user selects the third-term results workbook. Its name and location may vary; its structure must be identical to the one produced by the Term03 tools.
Target file. The user selects the workbook that receives the values. If a school workbook was issued in the current session, it is offered as the default (Yes = use it, No = choose another file, Cancel = stop).
The repeat data is built from the source, every candidate of the target is matched, and the results are written to column E.
4.2 Reading the source
Only sheets whose merged cell `H1` starts with the third-year-secondary prefix (`ثالثة ثانوي`) are read; other sheets are ignored.
The last three characters of `H1` (a space and two digits) are removed to obtain the branch group. Sheets of the same group are pooled, and each group is mapped to its série sheet code (`Philo`, `LaEtr`, …) from keywords in the title. An unrecognised title, or two groups mapping to the same code, stops the process with an explicit message.
Columns read, from row 3: `B` full name, `C` date of birth, `E` repeat status (`نعم` / `لا`).
Duplicate candidate records are kept.
4.3 Matching
A target candidate is matched within the sheet of the same série code, on normalised name and date of birth:
Name normalisation: alef variants unified, taa marbuta folded to haa, diacritics, tatweel and extra or non-breaking spaces removed.
Date normalisation: both sides are compared as `yyyymmdd`, whether stored as dates or as text with `/`, `-` or `.` separators.
Exact match: one matching record, or several records that all carry the same valid status.
One-letter tolerance: a single spelling difference is accepted only when the date of birth matches and there is exactly one such candidate.
Ambiguous: several candidates, or contradictory statuses, are never written.
A status other than `نعم` / `لا` is never written.
4.4 Writing
Values only, written to column `E` from row 3; unmatched candidates keep their existing cell content.
Everything is computed first. The target is modified only if the whole computation succeeded, and only after a timestamped backup (`Backup_RepeatStatus_<yyyymmdd_hhnnss>`, same extension) is saved next to it.
The target is saved, and a report gives the counts of exact matches, one-letter matches, unmatched candidates and ambiguous cases, plus any target sheets with no source counterpart and the backup path.
Source and target must be different files, and the target must not be read-only. Macros of the workbooks being opened are disabled during the operation, and Excel's settings are restored afterwards.
Step 3 counts a candidate as a repeater when column E contains the word `نعم` ("yes").
---
5. Step 3 – Analysis (`modBacAnalysis`, macro `Run_BacAnalysis`)
5.1 Workflow
Take the school workbook (from `gLastBacEtabPath`, or a file dialog if it is empty or the file is missing).
Save a copy named `تحليل نتائج <original name>.xlsx` next to the original; the school workbook is not altered. An existing name is never overwritten (sequential suffix).
In every série sheet that holds a table, build the analysis blocks below the data.
Insert the `Gen` sheet as the first sheet.
Calculate everything, then freeze all generated formulas as static values, which avoids problems when sheets are copied or values are transferred to other workbooks. The candidate data rows are left untouched.
If anything fails, the copy is closed and deleted, and Excel's settings are restored.
5.2 Conventions
Only candidates with `Moyenne > 0` are analysed; an average of 0.00 denotes an absent candidate.
Pass mark: 10 or above, on a 20-point scale.
Per-subject pass rates use as denominator the present candidates holding a mark in that subject.
Subjects identified as Tamazight are excluded from the analysis tables.
Number formats: counts `00`, decimals `00.00`, percentages `[=1]%0;%00.00`.
5.3 Blocks under each série sheet
The blocks start two blank rows below the last data row, with one column per subject (physical education through the average).
Block	Content
Counters strip	Candidates analysed, males, females, repeaters.
Table 1 – Results	Subject average; passing females; passing males; total passing; pass rate; highest mark; lowest mark; passing repeaters.
Table 2 – Statistical indicators	Pass rate; mean; median; mode; range; standard deviation; variance; coefficient of variation.
Table 3 – Score distribution	Candidate counts per band: `<7`, `7–8.99`, `9–9.99`, `10–11.99`, `12–13.99`, `14–15.99`, `16–17.99`, `≥18`.
Standard deviation and variance are population measures (`STDEVP`, `VARP`); the coefficient of variation is the population standard deviation divided by the mean.
5.4 `Gen` sheet – whole-school summary
#	Section
1	Registered, present, male, female and absent candidates per série, with totals.
2	General results per série: present, passed (total, male, female) and pass rate.
3	Pass rate by sex.
4	Repeaters per série: count, passed, pass rate.
5	School pass rate against provincial and national rates (provincial, national and rank cells are left for manual entry).
6	School pass rate against the previous five sessions (past sessions left for manual entry).
7	Repeaters by sex: count, passed and pass rate.
---
6. Control panel (`modBuildBacPanel`)
`BuildBacPanel` creates a `ControlPanel` sheet (right-to-left, Arabic interface) with three buttons. Any button can be run at any time.
Button (in running order)	Macro
Extract school results	`Run_BacEtab`
Fetch repeat cases	`Run_RepeatStatus`
Analyse school results	`Run_BacAnalysis`
A busy flag prevents re-entrant clicks.
---
7. Repository contents
The repository holds the exported VBA source only.
File	Role
`modBacEtab.bas`	Step 1 – extraction, grouping by série, ranking.
`modRepeatStatus.bas`	Step 2 – repeat status from the third-term results.
`modBacAnalysis.bas`	Step 3 – per-série analysis tables and `Gen` sheet.
`modBuildBacPanel.bas`	Control panel.
`frmEtabCode.frm`, `frmEtabCode_code.txt`	School-code input dialog (optional).
---
8. Installation and usage
Create a macro-enabled workbook and save it in the working folder.
In the VBA editor, import the `.bas` modules (File → Import File).
(Optional) Add the code dialog: either import `frmEtabCode.frm`, or insert a blank UserForm named `frmEtabCode` and paste the code from `frmEtabCode_code.txt` into its code window. The dialog creates its controls at runtime, so no designer layout is required. Without it, a standard `InputBox` is used.
Run `BuildBacPanel` once to create the control panel.
Place the provincial file `BAC<year>ResultatsOriginal.xlsx` in the same folder as the tool workbook.
Click Extract, enter the school code; then Fetch repeat cases, selecting the third-term results workbook; then Analyse.
---
9. Technical notes
Array-based processing. The provincial data is read once, filtered and transformed in memory, and written to each sheet in a single block assignment, which keeps a 30,000-row source practical.
Safe writes. Existing files are never overwritten. The repeat-status step computes everything first and backs up the target before touching it; the analysis step works on a copy and deletes it on failure.
Interruption handling. All steps restore `ScreenUpdating`, `EnableEvents`, `DisplayAlerts` and the status bar on completion, cancellation or failure.
Static analysis values. After generation, formulas are replaced by values, so results do not change if the data is edited afterwards.
Arabic text encoding. Arabic strings are built from Unicode code points (`ArHex`, `ArabicText`), because the VBA editor does not reliably store Arabic literals.
Statistical functions. The analysis uses `COUNTIFS`, `AVERAGEIFS` and legacy array formulas; no dynamic-array functions are required.
---
10. Requirements
Microsoft Excel for Windows with VBA enabled.
The provincial results file and, for repeat status, the school's third-term results workbook, with the structure described in sections 3.1 and 4.2.
Optional fonts: Dubai, Dubai Medium, Aptos Slab.
---
11. Planned improvements
Unmatched candidates in the repeat-status report. Each candidate that cannot be matched will be listed by sheet and by the number in column A, and the process will continue without interruption. Currently the report gives counts only.
---
12. Validation status
This repository contains source modules only. It includes no automated tests and no sample data. Results should be checked against a known school before any official use, and compatibility with Excel versions and source-file layouts other than those described here has not been formally established.
---
13. Data privacy and repository scope
This repository preserves and tracks the VBA source code. Do not commit candidate records, provincial or school result workbooks, or any file containing personal data, passwords or tokens.
---
14. Related repositories and roadmap
ResultsAnalysisTerm02 – second-term results analysis and orientation.
ResultsAnalysisTerm03 – third-term results tools; produces the third-term results workbook used for repeat status.
The long-term goal is to merge all of these repositories into a single, comprehensive application.
---
Platform: Microsoft Excel / VBA
Scope: Single-school extraction and analysis of Baccalaureate results
