# Architecture

This document records how **BACAnalysisTools** is organised, which design decisions shaped it, and what each decision costs. It is meant for maintainers and reviewers.

**What the tool does.** A school-level Excel/VBA toolkit for the Baccalaureate (BAC): it extracts one school's results from the national results export, optionally fills in each pupil's repeater status from the third-term results, and produces an analysis workbook (per-series statistics plus a school summary).

**Basis.** Written from the source on the `main` branch (one commit): `modBacEtab`, `modRepeatStatus`, `modBacAnalysis` and `modBuildBacPanel` were read in full (the end of `modBacEtab` is only its Arabic text table). `frmEtabCode` and the binary `BacTools.xlsm` were **not** inspected; what is said about the form comes from how `modBacEtab` calls it.

## Contents

1. [Overview](#1-overview)
2. [Components](#2-components)
3. [The extraction workbook: the contract between steps](#3-the-extraction-workbook-the-contract-between-steps)
4. [Decisions](#4-decisions)
5. [Known limitations and risks](#5-known-limitations-and-risks)
6. [Rules for maintainers](#6-rules-for-maintainers)
7. [Appendix: series table](#7-appendix-series-table)

---

## 1. Overview

```
                    ControlPanel (modBuildBacPanel)
        ┌──────────────────┬──────────────────┬──────────────────┐
        ▼                  ▼                  ▼
  [1] Extract         [2] Repeat status    [3] Analyse
  RunBacStep1         RunBacRepeat         RunBacStep2
  Run_BacEtab         Run_RepeatStatus     Run_BacAnalysis
        │                  │                  │
 BAC<year>Resultats        │ reads            │ reads a COPY of
 Original.xlsx             │ 3rd-term results │ the extraction file
 (national export)         │ workbook         │
        │                  │                  │
        ▼                  ▼                  ▼
  extraction workbook ◄── column E ──   "<prefix> <name>.xlsx"
  (one sheet per series)   (backup first)     (Gen sheet + 3 tables per series)
        └──── path shared through modBacEtab.gLastBacEtabPath ────┘
```

The three steps are separate macros. They are connected only by **a workbook**: the extraction file written by step 1, which steps 2 and 3 read or modify. The panel does not enforce an order (see D3).

## 2. Components

| File | Entry point | Responsibility |
| :--- | :--- | :--- |
| `modBuildBacPanel.bas` | `BuildBacPanel`, `RunBacStep1`, `RunBacRepeat`, `RunBacStep2` | Builds the `ControlPanel` sheet (header, card, three buttons, status bar) and routes each button to a macro with `Application.Run`. A module-level `m_Busy` flag blocks re-entry. |
| `modBacEtab.bas` | `Run_BacEtab` | Finds the newest `BAC????ResultatsOriginal.xlsx` next to the tool, asks for the establishment code, reads the `Data` sheet into memory, splits the school's pupils into one sheet per series, formats them and saves a new workbook. Publishes the saved path in `gLastBacEtabPath`. |
| `frmEtabCode.frm` (+ `frmEtabCode_code.txt`) | `Setup`, `Accepted`, `codeText` | Input form for the establishment code. `modBacEtab` falls back to `Application.InputBox` if the form cannot be created. |
| `modRepeatStatus.bas` | `Run_RepeatStatus` | Reads the third-term results workbook, matches pupils by sheet code, normalised name and date of birth, and writes the repeat flag into column E of the extraction workbook. |
| `modBacAnalysis.bas` | `Run_BacAnalysis` | Copies the extraction workbook, adds statistics under each series table and a `Gen` summary sheet, then freezes the generated formulas as values. |
| `BacTools.xlsm` | | The macro workbook that holds the modules and the panel. |

## 3. The extraction workbook: the contract between steps

Steps 2 and 3 do not call step 1's code. They rely on the **layout** it produces. Any change to this layout must be made in all three modules.

| Item | Value |
| :--- | :--- |
| Sheets | One per series present for the school: `Philo`, `LaEtr`, `GeEc`, `SE`, `Math`, `TGM`, `TGE`, `TGC`, `TGP` (see the [appendix](#7-appendix-series-table)). |
| Row 1 | Banner: title with the year, and the branch name. The year is read back from `B1` (last four characters) by the analysis. |
| Row 2 | Header row. |
| Rows 3 and below | One pupil per row, sorted by average, highest first; column A numbers them 1..n. |
| Columns A to E | Number, name, birth date (text `yyyy-mm-dd`), sex (1 = male, 2 = female), repeater flag (text; the "yes" word is `نعم`). |
| Column F | Physical education (EPS). |
| Columns G onwards | One column per subject of the series, in the series' own order. |
| Last two columns | `Moyenne` (the average used everywhere) and `MoyenneMe` (the second average, "الأساسية"). |
| Table | Each sheet holds a table named `T_<sheet name>` that starts at the header row. |

Who depends on what:

- `modBacAnalysis` assumes header row 2, first subject column F, sex in D, repeater in E, and the average at `nCols - 1`, where `nCols` is the width of the table.
- `modRepeatStatus` assumes first data row 3, name in B, birth date in C, repeater in E, and sheet names equal to the series codes.
- Both skip the sheet named `Gen`.

## 4. Decisions

### D1. Pupils are selected by establishment code, from a national export
- **Decision.** `Run_BacEtab` loads the whole national `Data` sheet into one array and keeps only the rows whose `Cd_etab` equals the code typed by the user and whose `Cd_serie` is between 30 and 38. If the code exists only under other series, the user is told so; if it is absent, the dialog is shown again with an error.
- **Why.** The national file holds every school; the tool is operated by one school at a time.
- **Consequences.** The file is opened once, read-only, and closed straight after loading. Codes are normalised (Arabic-Indic digits, spaces, bidi marks) and limited to 15 digits so they can be compared as numbers.

### D2. Source columns are located by header name; subject titles are fixed by series
- **Decision.** `LoadData` finds each required column by its header (`Cd_etab`, `Cd_serie`, `Nom`, `Prenom`, `Date_jour`, `Date_mois`, `Date_annee`, `Sexe`, `Eps`, `Notef1`..`Notef11`, `Moyenne`, `MoyenneMe`; `Desi` is optional). A missing header stops the run and lists the missing names. The meaning of `Notef1..n`, however, comes from a hard-coded table per series (`NotefHeaderAr`, `SubjectCount`).
- **Why.** Header lookup survives a reordering of columns in the export; the subject names are not in the export (only `Notef` numbers), so they have to be supplied by the tool.
- **Consequences.** If the Ministry changes which subject sits at which `Notef` position, the data are still read but labelled with the wrong subject, **without any error**. This is the most important silent-failure risk in the tool (see section 5).

### D3. Three independent macros joined by a file, not by calls
- **Decision.** Extraction, repeat status and analysis are separate macros. The panel states that all buttons are always active and can be run in any order.
- **Why.** The steps are run at different moments (the third-term file may only arrive later), and a failure in one must not block the others.
- **Consequences.** The coupling is the workbook layout (section 3) plus one public variable, `modBacEtab.gLastBacEtabPath`, used as a default by steps 2 and 3. If it is empty or the file is gone, the tool asks for a file (`PickFile`).
- **Order still matters for the data.** The analysis reads a **copy** of the extraction file made at the moment it runs and freezes its results as values (D6). Running the repeat step **after** the analysis does not update an analysis file that already exists; repeater counts in it come from column E as it was at that time.

### D4. Repeat status is matched in memory, with a conservative fuzzy rule
- **Decision.** `Run_RepeatStatus` works in three phases: (1) build the repeat data of every third-secondary sheet in memory, grouped by series code derived from the title in `H1`; (2) match every pupil of the extraction workbook without writing anything; (3) make a timestamped backup (`Backup_RepeatStatus_yyyymmdd_hhnnss`), write column E as **values**, save.
- **Matching rule.** Same sheet code, same normalised name (alef variants, taa marbuta, diacritics, tatweel and spacing are normalised), same date of birth. A one-letter spelling difference is accepted only if the date of birth matches and **exactly one** candidate exists; duplicates with conflicting values are counted as ambiguous and left unwritten.
- **Safeguards.** The target is opened first so a read-only file fails early; the source is opened read-only and never modified; `AutomationSecurity` is forced to disable macros while opening files and is restored afterwards; if nothing matched, nothing is written and no backup is made.
- **Consequences.** Unmatched and ambiguous pupils keep an empty flag, and the final report shows how many fell into each class so the user can check them by hand. The earlier two-file chain (`modBuildRepeatStatus` → `RepeatStatus.xlsx` → `modMatchRepeatStatus`) was replaced by this single module, which removes an intermediate file that could go stale.

### D5. Analysis on a copy, with fonts and styles taken from the extraction sheets
- **Decision.** `Run_BacAnalysis` saves a copy named with the prefix `تحليل نتائج` plus the original file name (sequential suffix `_02`, `_03`, ... if the name exists), opens the copy, and adds content below the data of every series sheet. The original is never modified. The fonts used for new cells are read from existing cells of the first series sheet.
- **Content per series sheet.** A counters strip (count, males, females, repeaters), then three tables: subject averages with pass counts (by sex, total, rate, highest, lowest, repeaters who passed); statistical indicators (pass rate, mean, median, mode, range, population standard deviation, variance, coefficient of variation); score bands (eight bands from below 7 to 18 and above).
- **Population rules.** Only pupils with `Moyenne > 0` are analysed (zero means absent). Population statistics are used (`STDEVP`, `VARP`). The Amazigh column is skipped by matching its header text. A subject cell counts only if it is not empty.
- **`Gen` sheet** (first sheet), seven sections: registered and present pupils per branch; general results; pass rate by sex; pass rate of repeaters; comparison with provincial and national rates (cells left empty for manual entry); previous sessions (years derived from the file's year; manual entry); repeaters by sex.
- **Consequences.** The analysis is repeatable and non-destructive, and the extraction file stays the single source of truth.

### D6. Generated formulas are frozen as values
- **Decision.** After building, the tool runs `Application.CalculateFull` and then replaces every formula below the data tables, and every formula on `Gen`, with its value. The data rows themselves are left untouched.
- **Why.** The analysis file is a report that is sent or printed; it should not recalculate, break on a missing function, or change when someone edits a source cell.
- **Consequences.** The analysis does not update when data change: regenerate it. The formulas are visible only while the run is in progress. The cells meant for manual entry on `Gen` are left empty on purpose.

### D7. In-memory processing and bulk writes
- **Decision.** The national sheet is read once into a `Variant` array; per-series data arrays are built, sorted by average (a stable insertion sort on a separate key array) and written to the sheet in one assignment. The repeat matcher keeps its data in arrays of a user-defined type, one set per series.
- **Why.** Per-cell access to a national-size sheet would be slow.

### D8. Fail safely, restore Excel's state, never overwrite
- **Decision.** Each macro saves screen updating, events, alerts and the status bar before starting and restores them on every exit path. A failed extraction closes its unsaved workbook; a failed analysis closes and deletes its copy. Output names are never reused: `NewOutputPath`, `NewAnalysisPath` and the backup naming add a sequence number or timestamp. An output name is also treated as taken while a workbook of that name is open.
- **Naming.** The extraction file is named after the school (`Desi` value, with illegal file-name characters replaced) or, when that field is missing, `BAC<year>_<code>`.
- **Consequences.** Re-running is always safe; old results stay on disk until the user removes them.

### D9. Pure-ASCII source, Arabic built from code points
- **Decision.** All Arabic text is constructed with `ArHex` / `ArabicText` from hexadecimal code points, so the `.bas` files contain only ASCII.
- **Why.** The source files do not state a reason. The likely one is that `.bas` files are imported in the system ANSI code page, which corrupts Arabic literals on some machines.
- **Consequences.** Texts are unreadable in the source. Each module carries its own copy of `ArHex`.

### D10. Self-contained modules
- **Decision.** Style helpers (`StyleGrid`, `StyleHead`, `StyleLabel`, `StyleBanner`), the font-detection code and the file helpers are repeated in `modBacEtab` and `modBacAnalysis`; `modRepeatStatus` has its own message and file helpers. The only shared names are `modBacEtab.gLastBacEtabPath` and the entry-point macros.
- **Why (inferred).** Each step can be imported and run on its own; the panel even tells the user if `modRepeatStatus` was not imported (`msg_missing`).
- **Consequences.** Repeated code must be kept in step by hand: constants such as colours, number formats (`00.00`, `[=1]%0;%00.00`) and layout sizes are copied between `modBacEtab` and `modBacAnalysis`.

## 5. Known limitations and risks

1. **Silent mislabelling if the `Notef` order changes (D2).** Add a sanity check, for example comparing the subject count with the number of non-empty `Notef` columns for each series.
2. **Hard-coded Arabic keywords.** Series codes are derived from the title in `H1` of the third-term workbook by keyword (for example, mathematics must contain "رياضيات" and not "تقني"). A change in wording stops the repeat step with a "code not found" message.
3. **Order dependence hidden by "any order" (D3).** The analysis is stale if the repeat step runs later.
4. **Panel error handling is uneven.** `RunBacRepeat` shows a message if the macro cannot start; `RunBacStep1` and `RunBacStep2` swallow such an error silently and rely on the called macro to report its own errors.
5. **Windows assumptions.** The panel computes the displayed file name with a hard-coded `\`; other modules use `Application.PathSeparator`.
6. **Repo hygiene.** The repository has a single commit, no README and no licence file. `BacTools.xlsm` is a binary copy of the modules, so it can drift from the `.bas` files and cannot be reviewed in a diff; check that it contains no pupil data before keeping it in a public repository. The national export `BAC????ResultatsOriginal.xlsx` and the extraction outputs contain personal data and must never be committed.
7. **Validation scope.** No automated tests were found in the repository.
8. **Modules not reviewed.** `frmEtabCode` was not read.

## 6. Rules for maintainers

1. Treat the layout in [section 3](#3-the-extraction-workbook-the-contract-between-steps) as an interface: change it in `modBacEtab`, `modBacAnalysis` and `modRepeatStatus` together.
2. A new series or a changed subject list must update `SubjectCount`, `NotefHeaderAr`, `SeriesSheetName`, the series names in `modBacAnalysis` (`names` array and `BranchOf`) and the keyword rules in `modRepeatStatus.SheetCodeFromTitle`.
3. Never write to the source workbooks, never overwrite an existing output; keep the backup-before-write step in `Run_RepeatStatus`.
4. Compute everything first, touch the target last (D4).
5. Restore Excel's state on every exit path (D8).
6. Keep source files ASCII-only (D9).
7. Keep real data out of the repository: no national exports, extraction or analysis workbooks, no pupil names.

## 7. Appendix: series table

`Cd_serie` in the export maps to a sheet and a subject count; sheet index `k = Cd_serie - 29`.

| `Cd_serie` | Sheet | Branch | Subject columns |
| :---: | :--- | :--- | :---: |
| 30 | `Philo` | Literature and Philosophy | 8 |
| 31 | `LaEtr` | Foreign Languages | 9 |
| 32 | `GeEc` | Management and Economics | 11 |
| 33 | `SE` | Experimental Sciences | 10 |
| 34 | `Math` | Mathematics | 10 |
| 35 | `TGM` | Technical Mathematics, Mechanical Engineering | 10 |
| 36 | `TGE` | Technical Mathematics, Electrical Engineering | 10 |
| 37 | `TGC` | Technical Mathematics, Civil Engineering | 10 |
| 38 | `TGP` | Technical Mathematics, Process Engineering | 10 |

Each sheet also has EPS (column F) before the subjects, and `Moyenne` / `MoyenneMe` after them, so a sheet has `8 + subject count` columns.
