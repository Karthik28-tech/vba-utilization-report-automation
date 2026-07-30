# Timesheet Validation & Utilization Report Generator (VBA)

An advanced VBA project that validates raw timesheet entries, flags data quality issues, and generates an automated utilization report with a programmatically-built chart — built to demonstrate VBA beyond recorded macros: Class Modules, array processing, Dictionary-based duplicate detection, robust error handling, custom worksheet functions, and event-driven automation.

## Business Problem
Manually checking timesheet entries for missing data, duplicate submissions, or invalid values (like impossible hour counts) doesn't scale and is easy to get wrong. This workbook automates that validation, flags issues clearly, and produces a ready-to-read utilization summary without any manual formula-building.

## How It Works
1. Raw timesheet entries go on the **Raw Data** sheet (or are entered via a UserForm)
2. `ProcessTimesheets()` reads the entire raw range into memory as an array, validates each row using a custom `TimesheetEntry` class, and writes all results to the **Processed** sheet in one operation — flagging each row Valid/Invalid
3. A UserForm (`frmTimesheetEntry`) allows manual entry with real-time validation and duplicate detection (via a `Scripting.Dictionary`) before a new row is ever written
4. `GenerateSummaryReport()` aggregates billable hours per employee and builds a bar chart entirely through code on the **Summary** sheet
5. `Workbook_Open()` automatically checks for unresolved invalid entries every time the file is opened, and alerts the user if any exist
6. Any runtime error is caught and logged with a timestamp to an auto-created **ErrorLog** sheet, rather than crashing with a raw VBA error

## Advanced VBA Techniques Demonstrated

| Technique | Where |
|---|---|
| Class Modules (custom objects) | `TimesheetEntry` class with properties + a `Validate()` method |
| Array processing | `ProcessTimesheets()` reads/writes the entire data range in single operations instead of looping cell-by-cell |
| Scripting.Dictionary | Duplicate-entry detection in the UserForm submit handler |
| UserForms with real-time validation | `frmTimesheetEntry` — rejects invalid/duplicate entries before they're written |
| Custom Functions (UDF) | `=UtilizationRate(billableHours, totalHours)` — usable directly as a worksheet formula |
| Robust error handling | `On Error GoTo` with a dedicated `LogError()` routine that writes to an ErrorLog sheet rather than silently failing |
| Events | `Workbook_Open()` — automatic validation check on file open |
| Programmatic charting | `GenerateSummaryReport()` builds a bar chart via `ChartObjects.Add` and `.SetSourceData`, no manual chart insertion |

Full code: see [`code/`](code) — each module exported as readable text (`.cls`, `.bas`, `.frm`/`.frx`).

## A Real Bug Found and Fixed
While testing duplicate detection, two identical-looking entries weren't being flagged as duplicates. Root cause: the Dictionary key built from *existing* Processed-sheet rows used Excel's default date-to-text conversion, while the key built from a *new* form submission used an explicit `Format(date, "yyyy-mm-dd")`. The two text representations of the same date didn't match as strings, so the lookup silently failed. Fixed by applying the same explicit date format on both sides before comparing. (This is a genuinely common category of bug — comparing values that look equivalent but aren't identical as text.)

## Screenshots
See [`docs/screenshots/`](docs/screenshots) — UserForm, Processed sheet (valid/invalid flags), duplicate-detection warning, and the Summary sheet with generated chart.

## Data
Synthetic sample timesheet entries, deliberately including invalid rows (missing fields, out-of-range hours, invalid billable flag) to demonstrate the validation logic catching real issues.

## Tools Used
Excel VBA (Class Modules, UserForms, Scripting.Dictionary via Microsoft Scripting Runtime reference)

---
*Self-directed portfolio project — not affiliated with any employer or client data.*
