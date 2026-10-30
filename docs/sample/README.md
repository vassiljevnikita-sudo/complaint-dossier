# Sample run

Two reports from one demo run on the synthetic sample data, so you can read them without running anything. Both are drafts that nobody has reviewed. The AI sections hold the answer written in advance for that complaint in `samples/demo-responses/`, not output from Copilot. The layout is the example template in `templates/`.

| Report | Complaint | Template |
| --- | --- | --- |
| [SYN-2026-007, PDF](reports/SYN-2026-007-long-report-DRAFT.pdf), 3 pages | Foreign fibre in a cream, rated critical, no returned sample | Long report |
| [SYN-2026-006, PDF](reports/SYN-2026-006-short-report-DRAFT.pdf), 2 pages | Carton crushed in transport, with a CSV file in the case folder | Short report |

[complaint-dossier-sample-DRAFT.pdf](complaint-dossier-sample-DRAFT.pdf) is a copy of the SYN-2026-007 PDF under a stable name for links from elsewhere.

The folder has the same layout as the `out/` folder of a run:

| Path | Content |
| --- | --- |
| `reports/<no.>-<type>-report-DRAFT.docx` | The reports, as Word files |
| `review-list.xlsx` | One row per report with its check results and empty columns for reviewer, date, release and comment |
| `audit/<no.>/audit.json` | SHA-256 of every input and output of this report, checks, events |
| `audit/<no.>/prompt.txt`, `response.md`, `case-data.json` | The files whose hashes audit.json records |
| `runs/<run id>.json` | Summary of the run |

## How it was made

- The run: `./Invoke-ComplaintDossier.ps1 -Demo -CaseId SYN-2026-006,SYN-2026-007 -OutputPath out/sample -NoOpen`, step "Sample run of two complaints for docs/sample (PowerShell 7)" of CI run [41080343988](https://github.com/vassiljevnikita-sudo/complaint-dossier/actions/runs/41080343988) for pull request [#9](https://github.com/vassiljevnikita-sudo/complaint-dossier/pull/9), `windows-latest`, PowerShell 7.6.7, run ID `20261030T152519Z-335b42`. The files are copied unchanged from the build artifact `complaint-dossier-sample` of that run. Its Pester step failed only because this folder still held the files of the previous code version.
- The PDFs: rendered from the .docx files with LibreOffice 24.2, `soffice --headless --convert-to pdf <report>.docx`. They are not listed in audit.json.

## How it stays current

The test "Committed sample in docs/sample" runs on every push. It checks the files against the hashes in their audit.json. Then it runs the demo again for SYN-2026-007 and SYN-2026-006 and compares the input hashes, the case data, prompt and response hashes, and the report text apart from run ID and time. If a code change alters a report, CI fails until this folder is regenerated: run the command above, copy `reports/`, `audit/`, `runs/` and `review-list.xlsx` here, render the PDFs again, copy the SYN-2026-007 PDF to `complaint-dossier-sample-DRAFT.pdf` and update the run ID and links in this file.
