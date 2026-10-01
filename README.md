# Contract Automation

[![tests](https://github.com/vincal848/ContractAutomation/actions/workflows/tests.yml/badge.svg)](https://github.com/vincal848/ContractAutomation/actions/workflows/tests.yml)

This project came out of a process a small Dallas valuations firm wanted: pull
monthly revenue and expense data out of a folder of client workbooks, compute
a few standard metrics, and fill a Word contract template per client without
retyping anything by hand. The values here have been anonymized and
simplified so as not to reproduce the original firm's data; everything in
`sample/` is generic, round, fake numbers.

Since there is no real contract template to build against, the macro here
writes its output straight to a `.docx` per row rather than into anything
further downstream. In a real deployment, these values would likely feed
inline into a templated legal document instead of standing alone.

I have since rebuilt the macro. The original shipped as two draft modules
with the same sub names, one of which does not compile (`Continue Do` is not
valid VBA), and the one that does compile fills the right bookmarks for the
first contract row and silently empties ones for every row after it, because
it reuses a single open Word document across the whole loop.

![Original chart output, with the chart bug described below](docs/img/result_image_sample.png)
*The screenshot this repository's sample data is reconstructed from. The
legend reads "Net Profit $40,000.00" and "Revenue Growth (%) N/A" instead of
the column headers, and the chart only plots 8 of the 12 months in the table
above it -- both symptoms of the `SetSourceData` bug described below.*

## At a glance

| | |
|---|---|
| **Macro** | `src/ReportAutomation.bas` -- one module, five subs |
| **Inputs** | `DataSources` sheet (paths to source workbooks), `RawData` (consolidated rows), `TemplateInfo` (template path + output folder) |
| **Outputs** | `CalculatedMetrics` sheet with a table and a chart; one `.docx` contract per `RawData` row |
| **Validation** | 11 pytest tests: `metrics.py` pinned against the sample screenshot's numbers, including the zero-previous-revenue case; static checks on the `.bas` text itself |
| **Stack** | VBA (Excel + late-bound Word), Python for tests and the sample-data generator |

## What it does

```mermaid
flowchart LR
    DS[DataSources sheet] --> AFR[AutomateFinancialReporting]
    AFR -->|copies each source workbook's data| RD[RawData sheet]
    RD --> CM[CalculateMetrics]
    CM --> CMS[CalculatedMetrics sheet]
    CMS --> FSR[FormatSummaryReport: table]
    CMS --> CC[CreateCharts: column chart]
    RD --> TTW[TransferToWordTemplate]
    TI[TemplateInfo sheet] --> TTW
    TTW -->|one .docx per row| OUT[Output folder]
```

`AutomateFinancialReporting` opens each workbook listed on `DataSources`,
copies its first sheet's data into `RawData`, then calls `CalculateMetrics`
(net profit, month-over-month revenue growth, profit margin, total revenue),
`FormatSummaryReport` (bold header, `ListObject` table), and `CreateCharts`
(a clustered column chart of net profit and revenue growth). Separately,
`TransferToWordTemplate` opens the template named in `TemplateInfo!B1`,
fills its bookmarks from each `RawData` row, and saves one contract per row
into the folder named in `TemplateInfo!B2`.

## What was wrong

**Both draft modules define subs with the same names.** `TransferToWordTemplate`
and `AutomateFinancialReporting` exist in both `legacy/Financial_report_module.bas`
and `legacy/finautomation_templateexceltoword.bas`. Importing both into one
workbook raises "ambiguous name detected" and neither can run. Fixed by
keeping one module, `src/ReportAutomation.bas`
(`tests/test_bas_lint.py::test_no_duplicate_sub_or_function_names`).

**The improved draft does not compile.** `AutomateFinancialReporting` uses
`Continue Do` to skip to the next `DataSources` row, which is a VB.NET/C#
construct that does not exist in VBA. Fixed with a `GoTo` to a loop-end label,
which is the idiomatic VBA way to skip an iteration
(`tests/test_bas_lint.py::test_no_continue_do_which_is_not_valid_vba`).

**`TransferToWordTemplate` only ever fills the first contract.** It opens one
Word document before the loop and reuses it for every row. A bookmark's
`Range` collapses to its filled text the moment it is written, so rows after
the first find the bookmarks already consumed; `SaveAs2` then just writes
copies of row one's contract under different file names. Fixed by creating a
fresh `Documents.Add(Template:=templatePath)` per row.

**The DataSources row counter is really the RawData row counter.**
`AutomateFinancialReporting` reads the next source path with
`wsSources.Cells(currentRow, 1)`, where `currentRow` tracks the next free row
in `RawData`. A source workbook that contributes anything other than exactly
one row desyncs the two sheets, so later paths are skipped or re-read. Fixed
with two independent counters, `sourceRow` and `destRow`.

**`ListObjects.Add` fails on a second run.** It always adds a new table,
which raises "a table already exists on this worksheet" once one is already
there. Fixed by unlisting any existing table in `FormatSummaryReport` before
adding a fresh one.

**Revenue growth divides by zero.** `CalculateMetrics` divides by the
previous row's revenue with no check for zero. Fixed by writing `"N/A"`
(`None` in the Python mirror) whenever the previous revenue is zero, the same
way the macro already handles the first row
(`tests/test_metrics.py::test_zero_previous_revenue_gives_none_instead_of_dividing_by_zero`).

**The chart mislabels its own series.** `CreateCharts` calls `SetSourceData`
on a hardcoded `Range("A1:B10")` with no `PlotBy` argument. Because both
columns hold numbers, Excel's own guess at rows-vs-columns is unreliable: the
screenshot above shows the first data row's values used as legend text
("Net Profit $40,000.00", "Revenue Growth (%) N/A") instead of the header
row, and the hardcoded `B10` only ever covered 9 of the 12 months. Fixed by
reading the real last row, passing `PlotBy:=xlColumns` explicitly, and naming
each series from its header cell.

**Hardcoded paths.** `C:\Path\To\Your\Template.docx`, `C:\Path\To\Save\Output.docx`,
and `C:\Contracts\` are all literal strings in the originals. Fixed by reading
the template path and output folder from `TemplateInfo!B1` / `!B2`, and
creating the output folder if it does not exist
(`tests/test_bas_lint.py::test_no_hardcoded_c_drive_paths`).

## Setup

1. Open the workbook in Excel, then Alt+F11 to open the VBA editor.
2. File > Import File..., and import `src/ReportAutomation.bas`. No reference
   to the Word object model is required -- `TransferToWordTemplate` creates
   Word with late binding (`CreateObject("Word.Application")`).
3. Save the workbook as macro-enabled, `.xlsm`.
4. Add four sheets, named exactly:
   - `DataSources` -- column A, from row 2 down: full paths to source workbooks.
   - `RawData` -- cleared and refilled by every run.
   - `CalculatedMetrics` -- filled by `CalculateMetrics` / `FormatSummaryReport` / `CreateCharts`.
   - `TemplateInfo` -- `B1` = path to the Word contract template, `B2` = output folder for generated contracts.
5. In the Word template, add bookmarks named `RegionBookmark`, `MonthBookmark`,
   `RevenueBookmark`, `ExpensesBookmark`, `NetProfitBookmark`,
   `CustomerNameBookmark`, `CompanyNameBookmark` wherever those values belong.
6. Run `AutomateFinancialReporting` to consolidate and compute metrics, then
   `TransferToWordTemplate` to generate the contracts.

`sample/FinancialReport_sample.xlsx` and `sample/ContractTemplate.docx` are a
ready-made example with this layout already filled in (`RawData` has 12
months of the same fake numbers as the screenshot above, and `TemplateInfo`
points at `ContractTemplate.docx` with relative paths you'll need to make
absolute). We can't ship a `.xlsm` generated from Python -- Excel's macro
project binary isn't something openpyxl can write -- so import
`src/ReportAutomation.bas` into the sample workbook yourself per the steps
above to try it end to end.

## Repository guide

| Path | Contents |
|---|---|
| `src/ReportAutomation.bas` | The consolidated, compiling module |
| `legacy/Financial_report_module.bas` | Original draft 1, annotated. Not imported; known broken |
| `legacy/finautomation_templateexceltoword.bas` | Original draft 2, annotated. Not imported; does not compile |
| `metrics.py` | Python mirror of `CalculateMetrics`, so its arithmetic can be pinned by tests |
| `sample/make_sample.py` | Regenerates the two sample files below |
| `sample/FinancialReport_sample.xlsx` | Sample workbook: `DataSources`, `RawData` (12 months), `TemplateInfo` |
| `sample/ContractTemplate.docx` | Sample Word template with all seven bookmarks |
| `tests/test_metrics.py` | Known-value tests against the sample screenshot's numbers |
| `tests/test_bas_lint.py` | Static checks on the `.bas` text: compiles cleanly, no duplicate subs, no hardcoded paths |
| `docs/img/result_image_sample.png` | The original buggy chart output |
| `finautomation_results.pdf` | The same output as a PDF |

## Notes

- `TransferToWordTemplate` runs Word invisibly (`wordApp.Visible = False`) and
  quits it in both the normal path and the error handler, so a failed run
  doesn't leave an orphaned `WINWORD.EXE` process.
- `EnsureFolderExists` creates the output folder one path segment at a time,
  since `MkDir` on its own fails if more than the last segment is missing.
- VBA can't run in CI, so `tests/test_bas_lint.py` only checks the text of
  `src/ReportAutomation.bas` -- it is not a substitute for opening the file in
  the VBA editor and compiling it.
