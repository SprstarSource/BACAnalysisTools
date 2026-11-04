# BACAnalysisTools

**An Excel VBA toolkit for extracting, completing, and statistically analysing school-level Baccalaureate (BAC) results.**

BACAnalysisTools extracts the complete Baccalaureate results of a single school from the provincial results file, retrieves candidates' repeat status from the school's third-term results, and generates a statistical analysis workbook for every *série* and for the school as a whole.

The toolkit is designed to handle large provincial datasets efficiently while preserving the original source files and ensuring that each processing stage reports its execution status explicitly.

---

## 1. Purpose and Design

The provincial BAC results file is a single flat table covering all candidates in the province. It typically contains around 30,000 rows, with columns extending to `BT`.

Extracting one school's results manually is time-consuming and error-prone. Each candidate's marks are stored in positional columns (`Notef1` … `Notef11`), whose subject meanings depend on the candidate's *série*.

BACAnalysisTools automates this process through three stages:

1. **Extraction** — Isolate all candidates belonging to one school, group them by *série*, map their marks to the appropriate subjects, and rank them by average.
2. **Repeat Status** — Retrieve each candidate's repeat status from the school's third-term results and transfer it to the school workbook.
3. **Analysis** — Generate statistical tables for each *série* and a consolidated school-wide summary.

The original provincial results file is opened read-only. It is never modified, moved, or deleted.

### Design principles

- Independent modules that can be executed individually.
- A coordinated workflow for running all three stages in sequence.
- Explicit execution status for every stage.
- Immediate termination when a stage fails or is cancelled.
- No overwriting of existing output files.
- In-memory processing wherever practical.
- Safe handling of source and target workbooks.
- Static analysis outputs that remain unchanged after generation.

---

## 2. Processing Pipeline

```text
Provincial BAC Results File
          |
          v
   [1] Extract School Results
          |
          v
   School Workbook (.xlsx)
   One worksheet per série
          |
          v
   [2] Fetch Repeat Status
          |
          v
   Column E updated with repeat status
          |
          v
   [3] Analyse School Results
          |
          v
   Analysis Workbook (.xlsx)
   - Statistical tables per série
   - General school summary (Gen)
```

Each stage has its own execution status and detailed result information.

When the stages are executed through the master processing command, the next stage starts only after the previous stage has explicitly reported success.

A missing status report is treated as a failure. The workflow does not infer success from a macro returning normally or from the existence of a recently created file.

---

## 3. Modules and Responsibilities

| Stage | Module | Macro | Responsibility |
|---|---|---|---|
| 1 | `modBacEtab` | `Run_BacEtab` | Extract candidates, group them by série, rank them, and generate the school workbook. |
| 2 | `modStudentRepetitionStatus` | `Run_RepeatStatus` | Retrieve and transfer candidates' repeat status. |
| 3 | `modBacAnalysis` | `Run_BacAnalysis` | Generate per-série statistical tables and the `Gen` summary sheet. |
| Interface | `modBuildBacPanel` | `BuildBacPanel` | Build the BAC control interface and coordinate the processing workflow. |

The extraction module publishes the generated school workbook path through the session variable `gLastBacEtabPath`.

Stages 2 and 3 can use this path as their default target, avoiding unnecessary manual file selection when the workflow is executed in sequence.

---

## 4. Stage 1 — School Results Extraction

**Module:** `modBacEtab`  
**Macro:** `Run_BacEtab`

### 4.1 Input requirements

| Item | Requirement |
|---|---|
| File | `BAC????ResultatsOriginal.xlsx`, where `????` represents the four-digit year. |
| File selection | If several matching files exist, the latest year is selected. |
| Location | The same folder as the saved tool workbook. |
| Preferred worksheet | `Data`. |
| Fallback worksheet | The first worksheet containing the required headers `Cd_etab` and `Cd_serie`. |

Required headers in row 1:

`Cd_etab`, `Cd_serie`, `Nom`, `Prenom`, `Date_jour`, `Date_mois`, `Date_annee`, `Sexe`, `Eps`, `Notef1` … `Notef11`, `Moyenne`, `MoyenneMe`.

Optional header: `Desi`.

Columns are resolved by header name rather than fixed position. Consequently, the provincial file's column positions and overall width do not need to remain constant.

The source data is read into a memory array and processed there. The provincial workbook is opened read-only and closed without saving.

### 4.2 Processing

The user enters the school's national establishment code (`Cd_etab`) through the `frmEtabCode` dialog.

If the form is unavailable, a standard `InputBox` is used.

The input is validated for:

- Non-empty value.
- Numeric characters only.
- Maximum permitted length.
- Normalisation of Arabic-Indic digits.

All rows matching the establishment code are selected.

A distinction is made between:

- A code that does not exist in the provincial file.
- A code that exists but has candidates only in series outside the supported range (30–38).

In either case, the user receives a specific message and is prompted again.

Each candidate is assigned to a *série* according to `Cd_serie`.

The marks (`Eps`, `Notef1` … `NotefN`) and averages (`Moyenne`, `MoyenneMe`) are transferred. The positional mark columns are mapped to subject names through a predefined mapping table keyed by `Cd_serie`.

Unused `Notef` positions are excluded completely: no corresponding subject column is created and no data is transferred.

Candidates are sorted by `Moyenne` in descending order using a stable sort. Column A is then numbered sequentially from 1.

The output contains one worksheet for each *série* represented among the school's candidates.

### 4.3 Data transformations

- `Nom` and `Prenom` are combined into one full-name column.
- `Date_jour`, `Date_mois`, and `Date_annee` are combined into a birth-date value.
- Birth dates are formatted as `yyyy-mm-dd` and stored as text.
- Sex is transferred from `Sexe`.

### 4.4 Supported series

| Code | Sheet Série | Série | Subject columns |
|---|---|---|---:|
| 30 | Philo | Literature and Philosophy | 8 |
| 31 | LaEtr | Foreign Languages | 9 |
| 32 | GeEc | Management and Economics | 11 |
| 33 | SE | Experimental Sciences | 10 |
| 34 | Math | Mathematics | 10 |
| 35 | TGM | Technical Mathematics – Mechanical Engineering | 10 |
| 36 | TGE | Technical Mathematics – Electrical Engineering | 10 |
| 37 | TGC | Technical Mathematics – Civil Engineering | 10 |
| 38 | TGP | Technical Mathematics – Process Engineering | 10 |

Subject names use unified Arabic spellings to ensure consistent headers across the different series.

### 4.5 Output worksheet layout

| Column | Content |
|---|---|
| A | Rank after descending average sort |
| B | Full name |
| C | Birth date |
| D | Sex |
| E | Repeater flag, populated in Stage 2 |
| F | Physical education (`Eps`) |
| G onward | Subject marks for the relevant série |
| Penultimate | General average (`Moyenne`) |
| Last | Core-subject average (`MoyenneMe`) |

Formatting and layout:

- Right-to-left worksheet direction.
- Excel table named `T_<sheet>`.
- Print-ready landscape orientation, fitted to one page wide.
- Column A width: 4.3.
- Columns D through the last data column: width 6.3.
- Row 1 height: 40.
- Row 2 height: 100, with rotated subject headers.
- Dubai / Dubai Medium for text.
- Aptos Slab for numeric values.
- Automatic fallback to Arial / Calibri when the preferred fonts are unavailable.

The final layout is applied and verified by reading the values back. Any deviation is reported.

### 4.6 Output file naming

The workbook is saved as `.xlsx` in the tool folder.

The filename is based on the school's `Desi` value, sanitised and truncated to a maximum of 150 characters.

If `Desi` is unavailable, the fallback name is:

`BAC<year>_<code>`

Existing files are never overwritten. Sequential suffixes are added when necessary:

`_02`, `_03`, and so on.

---

## 5. Stage 2 — Repeat Status Processing

**Module:** `modStudentRepetitionStatus`  
**Macro:** `Run_RepeatStatus`

The provincial BAC file does not contain the candidates' repeat status.

This stage retrieves repeat-status information directly from the raw source files used by the `ResultsAnalysisTerm03` workflow. It reads the files placed in the `Source03` directory beside the BACAnalysisTools workbook, prepares and combines the required data internally, and writes the resulting statuses into column E of the school BAC workbook.

This design separates raw-data preparation from repeat-status matching: the stage does not depend on a previously generated, consolidated third-term results workbook. The required source data is collected and processed in memory; no intermediate consolidated workbook is required.

### 5.1 Workflow

**Raw source directory**

Place the raw third-term source files in a directory named `Source03`, located beside the saved BACAnalysisTools workbook. The module reads the required source files from this directory and performs the necessary preparation and consolidation internally. It does not ask the user to select a pre-generated, consolidated third-term results workbook.

When a source file requires format conversion, the converted working copy is used for processing and the original file is preserved in `Source03/_Original` rather than deleted.

**Target workbook**

The user selects the workbook to receive the repeat-status values.

If a school BAC workbook was generated during the current session, it is offered as the default:

- Yes: use the current session workbook.
- No: select another workbook.
- Cancel: stop the operation.

The source records are prepared internally from `Source03`, target candidates are matched, and the resulting statuses are written to column E.

### 5.2 Source data extraction

Only worksheets whose merged cell `H1` starts with the third-year secondary prefix:

`ثالثة ثانوي`

are processed.

Other worksheets are ignored.

The final three characters of `H1` (a space followed by two digits) are removed to obtain the branch group.

Worksheets belonging to the same group are pooled. Each group is mapped to its corresponding BAC série code (`Philo`, `LaEtr`, etc.) using keywords in the worksheet title.

The following conditions stop the operation with an explicit error:

- An unrecognised worksheet title.
- Two different groups mapping to the same série code.

Data is read from row 3:

| Column | Content |
|---|---|
| B | Full name |
| C | Date of birth |
| E | Repeat status (`نعم` / `لا`) |

Duplicate source records are retained for matching and ambiguity detection.

### 5.3 Candidate matching

Matching is performed within the target worksheet corresponding to the same série code.

The matching identity consists of:

- Normalised full name.
- Normalised date of birth.

**Name normalisation**

The normalisation procedure includes:

- Unifying Alef variants.
- Folding Taa Marbuta into Haa.
- Removing diacritics and Tatweel.
- Removing extra, leading, trailing, and non-breaking spaces.

**Date normalisation**

Dates are converted to a common `yyyymmdd` representation.

The procedure supports dates stored as Excel date values or as text using `/`, `-`, or `.` separators.

**Matching outcomes**

| Matching condition | Action |
|---|---|
| Exact identity match with a single valid status | Transfer the status. |
| Multiple exact source records with identical valid statuses | Transfer the common status. |
| One-character spelling discrepancy with matching date of birth and a unique candidate | Transfer the status and flag the discrepancy for review. |
| Multiple plausible candidates | Do not transfer. |
| Contradictory source statuses | Do not transfer. |
| Invalid status | Do not transfer. |
| Incomplete identity or date-of-birth mismatch | Do not transfer. |

A one-character discrepancy is not automatically rejected. It is accepted only when the date of birth matches and the candidate is uniquely identifiable.

Such transfers are reported as:

`Transferred with Name Discrepancy`

This allows the person responsible for the operation to review the spelling difference without unnecessarily losing a valid match.

Ambiguous cases are never resolved automatically.

### 5.4 Safe writing and backup handling

Only values are written to column E, starting at row 3.

Candidates that cannot be matched retain their existing cell contents.

The operation follows a transactional approach:

1. Collect and validate all source records.
2. Calculate all matching outcomes in memory.
3. Create a temporary, timestamped backup of the target workbook beside the original.
4. Apply the calculated values to the target.
5. Save the target workbook.
6. Verify that the updated workbook was saved successfully.
7. Delete the temporary backup only after successful saving and verification.

**Failure handling**

The temporary backup is retained if data transfer, saving, or verification fails.

The target is never modified until the complete matching computation has succeeded.

Additional safeguards:

- Source and target must be different files.
- The target must not be read-only.
- Macros in workbooks opened during the operation are disabled.
- Excel application settings are restored after completion, cancellation, or failure.

### 5.5 Repeat-status report

The operation generates a report containing the matching results, including:

- Exact matches.
- One-character spelling discrepancies transferred with a warning.
- Unmatched candidates.
- Ambiguous cases.
- Target worksheets with no corresponding source worksheet.
- Relevant execution details.

The report provides an audit trail for the transfer and identifies cases requiring human review.

### 5.6 Repeater identification in Stage 3

A candidate is counted as a repeater when column E contains:

`نعم`

---

## 6. Stage 3 — Statistical Analysis

**Module:** `modBacAnalysis`  
**Macro:** `Run_BacAnalysis`

### 6.1 Workflow

The module obtains the school workbook path from `gLastBacEtabPath`.

If the path is empty or the file no longer exists, the user must select the workbook manually when running the module independently.

A separate analysis copy is created next to the original:

`تحليل نتائج <original name>.xlsx`

The original school workbook is never modified.

Existing output files are never overwritten. Sequential suffixes are added when necessary.

For each série worksheet containing a valid data table, the analysis blocks are generated below the candidate data.

A `Gen` worksheet is inserted as the first worksheet.

All calculations are completed before generated formulas are converted into static values. The candidate data rows remain untouched.

If the operation fails, the analysis copy is closed and deleted, and Excel application settings are restored.

### 6.2 Statistical conventions

- Only candidates with `Moyenne > 0` are included in the analysis.
- An average of `0.00` represents an absent candidate.
- The passing threshold is 10/20.
- Subject pass rates use the number of present candidates with a valid mark in that subject as the denominator.
- Subjects identified as Tamazight are excluded from the analysis tables.

Number formats:

| Data type | Format |
|---|---|
| Counts | `00` |
| Decimal values | `00.00` |
| Percentages | `[=1]%0;%00.00` |

Population statistical measures are used:

- Standard deviation: `STDEVP` / `STDEV.P`.
- Variance: `VARP` / `VAR.P`.
- Coefficient of variation: population standard deviation divided by the mean.

### 6.3 Analysis blocks per série

The analysis blocks begin two blank rows below the last candidate data row.

Each block uses one column per subject, from physical education through the general average.

**Counters strip**

- Candidates analysed.
- Male candidates.
- Female candidates.
- Repeaters.

**Table 1 — Results**

For each subject:

- Subject average.
- Passing females.
- Passing males.
- Total passing candidates.
- Pass rate.
- Highest mark.
- Lowest mark.
- Passing repeaters.

**Table 2 — Statistical indicators**

- Pass rate.
- Mean.
- Median.
- Mode.
- Range.
- Standard deviation.
- Variance.
- Coefficient of variation.

**Table 3 — Score distribution**

Candidate counts across the following score bands:

| Score band |
|---|
| Below 7 |
| 7–8.99 |
| 9–9.99 |
| 10–11.99 |
| 12–13.99 |
| 14–15.99 |
| 16–17.99 |
| 18 and above |

### 6.4 General summary — `Gen`

The `Gen` worksheet consolidates the school's results.

It contains seven sections:

1. **Candidate population:** Registered, present, male, female, and absent candidates per série, with totals.
2. **General results:** Present candidates, passed candidates (total, male, and female), and pass rate per série.
3. **Pass rate by sex.**
4. **Repeaters:** Number of repeaters, repeaters who passed, and their pass rate per série.
5. **Provincial and national comparison:** School pass rate compared with provincial and national rates. Reference rates and ranking are left for manual entry.
6. **Historical comparison:** School pass rate compared with the previous five examination sessions. Historical values are left for manual entry.
7. **Repeaters by sex:** Counts, passed candidates, and pass rates.

---

## 7. Control Interface and Workflow Orchestration

**Module:** `modBuildBacPanel`

The BAC workflow is integrated into the Term03 data-processing interface.

The master command:

**Process Baccalaureate Results**  
`معالجة نتائج البكالوريا`

is placed in its own row at the top of the Term03 processing panel.

This placement follows the source-oriented organisation of the tools: the Term03 processing workflow and the BAC repeat-status workflow share the `Source03` raw-data directory.

The master command coordinates the three BAC stages:

1. Extract school results.
2. Fetch repeat status.
3. Analyse school results.

The individual stage commands remain available for independent execution. The master command does not replace them.

### 7.1 Explicit execution status

Each processing module exposes its own execution status and details.

| Module | Status interface |
|---|---|
| `modBacEtab` | `BacEtab_SetAuto`, `BacEtab_LastStatus`, `BacEtab_LastDetail` |
| `modStudentRepetitionStatus` | `RepeatStatus_SetAuto`, `RepeatStatus_LastStatus`, `RepeatStatus_LastDetail` |
| `modBacAnalysis` | `BacAnalysis_SetAuto`, `BacAnalysis_LastStatus`, `BacAnalysis_LastDetail` |

The status values are:

| Status | Meaning |
|---|---|
| `OK` | The stage completed successfully. |
| `Cancelled` | The user or operation cancelled the stage. |
| `Failed` | The stage encountered an error or could not complete its required operation. |

The orchestrator uses `RunStage()` to execute each stage in automatic mode and retrieve its reported status and details.

A missing status report is treated as a failure.

### 7.2 Stop-on-failure behaviour

The master workflow stops immediately when a stage returns anything other than `OK`.

It never proceeds to a later stage after a failure or cancellation.

The final report identifies:

- Successfully completed stages.
- The stage that failed or was cancelled.
- The reported reason.
- Subsequent stages that were skipped.

This prevents incomplete or invalid intermediate results from being treated as successful workflow outputs.

### 7.3 Interface feedback

The control interface provides non-blocking execution feedback:

- A status indicator beneath each stage button.
- A status bar showing the current workflow state.
- Updated indicators after independent manual execution.
- A consolidated final report after the master workflow completes.
- The panel is brought to the foreground when a failure or cancellation occurs.

Success dialogs are suppressed during automatic execution to avoid interrupting the coordinated workflow. The master command presents the final consolidated report instead.

A busy flag prevents re-entrant execution and overlapping processing operations.

---

## 8. Repository Contents

The repository contains exported VBA source code and project documentation. The raw `Source03` data directory and school datasets are operational inputs and are not part of the repository.

| File | Role |
|---|---|
| `modBacEtab.bas` | Stage 1 — School results extraction and ranking. |
| `modStudentRepetitionStatus.bas` | Stage 2 — Repeat-status matching and transfer. |
| `modBacAnalysis.bas` | Stage 3 — Statistical analysis and `Gen` summary. |
| `modBuildBacPanel.bas` | Control interface and workflow orchestration. |
| `frmEtabCode.frm` | Optional school-code input form. |
| `ARCHITECTURE.md` | Project architecture and implementation documentation. |

The repository does not contain school candidate datasets or provincial results workbooks.

---

## 9. Installation and Usage

### 9.1 Requirements

- Microsoft Excel for Windows.
- VBA macros enabled.
- The provincial BAC results file with the structure described in Section 4.
- The raw third-term source files placed in the `Source03` directory, as described in Section 5.
- Optional fonts: Dubai, Dubai Medium, and Aptos Slab.

### 9.2 Installation

1. Create a macro-enabled Excel workbook (`.xlsm`) and save it in the intended working folder.
2. Open the VBA editor.
3. Import the `.bas` modules using **File → Import File**.
4. Optionally install the school-code input form:
   - Import `frmEtabCode.frm`, or
   - Import the supplied `frmEtabCode.frm` form file.
5. The form creates its controls at runtime; no designer layout is required.
6. If the form is not installed, the extraction module uses a standard `InputBox`.
7. Run `BuildBacPanel` once to create or initialise the control interface.

### 9.3 Source files

Place the provincial file in the same folder as the tool workbook:

`BAC<year>ResultatsOriginal.xlsx`

The raw third-term source files must be placed in a directory named `Source03` beside the tool workbook. The module prepares the required data internally. When source formats require conversion, original files are preserved in `Source03/_Original`.

### 9.4 Running the complete workflow

Use the **Process Baccalaureate Results** command in the Term03 processing panel.

The workflow executes:

1. Extract school results and enter the establishment code.
2. Fetch repeat status from the third-term results.
3. Generate the statistical analysis workbook.

The user receives a consolidated report showing the outcome of each stage.

### 9.5 Independent execution

Each stage can also be executed independently through its individual command.

This allows users to run a specific operation without having to execute the complete workflow.

---

## 10. Technical Notes

### Array-based processing

The provincial data is read once into memory, filtered and transformed in memory, and written to the output worksheets in block assignments. This keeps processing practical for large provincial datasets.

### Safe file handling

- The provincial source is read-only.
- Raw source files are processed from `Source03`. When source formats require conversion, the original files are preserved in `Source03/_Original` rather than deleted.
- Existing output files are never overwritten.
- Repeat-status processing uses a temporary backup and deletes it only after successful saving and verification.
- The analysis stage works on a separate copy and removes that copy if processing fails.

### Explicit status reporting

Each stage reports its own execution outcome and error details.

The orchestrator relies on these explicit reports rather than indirect success indicators.

### Excel application-state restoration

The processing modules restore the relevant Excel application settings after completion, cancellation, or failure, including:

- `ScreenUpdating`
- `EnableEvents`
- `DisplayAlerts`
- Status bar state

### Static analysis values

Generated analysis formulas are converted into values after calculation. This ensures that the final analysis workbook retains its results independently of subsequent changes to formulas or external workbook references.

### Arabic text encoding

Arabic strings are constructed using Unicode code points through the `ArHex` and `ArabicText` helpers because the VBA editor does not reliably preserve all Arabic literals.

### Statistical compatibility

The analysis uses standard Excel statistical functions such as `COUNTIFS`, `AVERAGEIFS`, and legacy array formulas. Dynamic-array functions are not required.

---

## 11. Validation and Operational Status

**Status: Successfully tested and operational**

All BACAnalysisTools modules have been tested in Microsoft Excel. The individual stages, their integration, the control interface, and the stop-on-failure mechanism have been verified through practical execution.

### 11.1 Component validation

| Component | Responsibility | Status |
|---|---|---|
| `modBacEtab` | School results extraction and candidate grouping | Passed |
| `modStudentRepetitionStatus` | Repeat-status processing and transfer | Passed |
| `modBacAnalysis` | Statistical analysis and report generation | Passed |
| `modBuildBacPanel` | Control interface and workflow orchestration | Passed |
| `frmEtabCode` | School-code input and validation | Passed |

### 11.2 Workflow validation

The integrated workflow has been tested under successful and unsuccessful execution conditions.

| Test scenario | Expected behaviour | Result |
|---|---|---|
| All three stages complete successfully | Generate the school workbook, transfer repeat status, and produce the analysis workbook | Passed |
| A processing stage encounters an error | Stop the workflow immediately | Passed |
| `Source03` is missing | Stop at the repeat-status stage and skip the analysis stage | Passed |
| The source folder is restored with the correct name | Allow the workflow to run successfully | Passed |
| Automatic execution | Display stage indicators and one consolidated final report | Passed |
| Independent stage execution | Update the corresponding stage status indicator | Passed |

### 11.3 Successful end-to-end execution

A complete test run confirmed:

- School results extracted successfully.
- Five output worksheets generated.
- 147 candidates extracted.
- 145 repeat-status values matched.
- Repeat-status report generated successfully.
- Six worksheets generated in the analysis workbook.
- All three stages reported successful completion.

The difference between the extracted candidate count and the number of matched repeat-status values is reported for review; it is not automatically treated as a processing failure.

### 11.4 Failure and recovery validation

The workflow was also tested with the required `Source03` directory missing.

The system:

- Completed the extraction stage.
- Detected the missing source directory during the repeat-status stage.
- Reported the failing stage and the reason.
- Stopped the workflow immediately.
- Skipped the analysis stage.

After restoring the correct `Source03` directory name, the complete workflow executed successfully.

### 11.5 Operational conclusion

The tests confirm that the toolkit handles both successful execution and expected failure conditions correctly.

The coordinated workflow does not continue after a failed or cancelled stage, and it provides explicit execution feedback through the control interface and final report.

**Validation scope:** Practical manual testing in Microsoft Excel.

**Current status:** All tested repository components and the integrated BAC processing workflow are operational. No known blocking issues remain in the tested functionality.

---

## 12. Privacy and Repository Security

The repository is intended to distribute the VBA source code and technical documentation, not school or candidate data.

Before publishing changes, ensure that commits do not contain:

- Candidate names or personal information.
- Real student results or source workbooks.
- Confidential school records.
- Passwords, credentials, or other secrets.
- Unintended local paths or private operational information.

VBA source files and documentation should be reviewed before committing. The `.bas` extension alone does not guarantee that a file contains no sensitive information.

---

## 13. Planned Improvements

Potential future enhancements include:

- Detailed listing of unmatched candidates in the repeat-status report, including their worksheet and candidate rank.
- More detailed presentation of ambiguous matching cases and their possible source candidates.
- Further improvements to operational reporting and validation.

These enhancements are separate from the currently tested core workflow.

---

## About

BACAnalysisTools is an Excel VBA toolkit designed to automate the extraction, completion, and statistical analysis of school-level Baccalaureate results.

It focuses on reliable processing, modular execution, safe file handling, and transparent reporting throughout the workflow.

**Language:** VBA  
**Platform:** Microsoft Excel for Windows  
**Project:** BACAnalysisTools  
**Repository:** [SprstarSource/bac-analysis-tools](https://github.com/SprstarSource/bac-analysis-tools)
