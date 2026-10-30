# Settings in config.json

`Setup.cmd` creates `config.json` from `config.example.json` and asks for the most important values. Everything else can stay as it is. Open `config.json` in Notepad to change a value, then run `Setup.cmd` again; it checks the file and says what is wrong. Paths may be relative to the project folder or complete. JSON needs every backslash twice, so `C:\QA\complaints.xlsx` is written `"C:\\QA\\complaints.xlsx"` or `"C:/QA/complaints.xlsx"`. Setup writes paths correctly when you answer its questions. Keys that start with `_help` are explanations and are ignored.

| Setting | What it does |
| --- | --- |
| `Language` | `de` or `en`: language of the messages, the check results in the report and the review list |
| `Workbook.Path`, `Workbook.Sheet` | Your Excel complaint list and the worksheet with the complaints |
| `Workbook.HeaderRow` | Row with the column names (1 if they are in the first row) |
| `Workbook.Columns` | Which of your columns holds what. Required: `CaseId` (complaint number) and `ReceivedDate`. Optional, `""` if your list does not have it: `Category`, `Description`, `Product`, `ProductGroup`, `Batch`, `Severity`, `Status`, `SampleReceived`, `ReportType`. The prompt, the similar-complaint counts and the report type use these. |
| `Selection.From`, `Selection.To` | Complaints received in this period get a report (`yyyy-MM-dd`; `""` = no limit) |
| `Selection.Statuses` | Only complaints with one of these statuses, for example `["Open"]`. Empty list = any status. |
| `Template.Fields` | Placeholders in your Word template that come from the Excel list: placeholder name -> column name. See [word-template.md](word-template.md). |
| `Template.DateFormat`, `Template.EmptyValue` | How dates are written (`dd.MM.yyyy`) and what an empty cell becomes |
| `Template.Styles` | Paragraph styles of your template for headings and list lines inside the AI text |
| `Reports.Short`, `Reports.Long` | Per report type: `Title`, `TemplatePath` (your Word template), `PromptPath`, `Sections` (heading the AI must write -> placeholder in the template), `Marker`, `TypeValues` and the required sentences |
| `Reports.DefaultType` | Report type for complaints without marker file and without a value in the report type column: `Short`, `Long` or `""` (skip them) |
| `Review.DraftNotice` | Text of the draft notice in each report |
| `Review.DocumentStatus` | Value of the Word document property "Status" |
| `Output.Path` | Folder for `reports/`, `review-list.xlsx`, `audit/` and `runs/` |
| `CaseFolders.Path` | Optional folder with one subfolder per complaint (named like the complaint number) for documents and the marker file `short.txt` or `long.txt`. `""` = none. |
| `SimilarComplaints` | Look-back window and thresholds for the frequency level (rare, occasional, frequent, very frequent) |
| `Prompt.MaskContactData` | Removes e-mail addresses and phone numbers from the text that goes into the prompt |
| `Run.StopAfterFailuresInARow` | The run stops after this many failed complaints in a row |
| `Copilot.PreferredModel` | Model name exactly as the Copilot model menu shows it. Required for live mode; the run stops if the menu does not offer it. |
| `Copilot.Upload` | Which files of a case folder are attached to the Copilot chat (extensions, number, size) |
| `Copilot.*Seconds` | Waiting times of the live automation |

## Short or long report

For each complaint, in this order:

1. An empty file `short.txt` or `long.txt` in its case folder.
2. The column named in `Workbook.Columns.ReportType`, with a value from `Reports.Short.TypeValues` (`short`, `kurz`) or `Reports.Long.TypeValues` (`long`, `lang`).
3. `Reports.DefaultType`.

Without any of these the complaint is skipped. The run and the review list say which complaints were skipped and why.

## Prompts

`prompts/short-report.txt` and `prompts/long-report.txt` hold the instructions for the AI. The headings they ask for must match `Reports.*.Sections[].Heading`. The prompts contain process rules of the original workflow (for example in `Reports.*.ProcessFacts`: "The batch record was reviewed and showed no findings."). Replace them with what your own process guarantees before real use.
