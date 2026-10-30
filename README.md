# Complaint Dossier

Complaint Dossier writes a draft report for every complaint in your Excel list, in your own Word template, with Microsoft 365 Copilot. You review each draft in Word afterwards.

**Kurzfassung (Deutsch):** Werkzeug für die Qualitätssicherung, das zu jeder Reklamation aus einer Excel-Liste einen Berichtsentwurf in der eigenen Word-Vorlage schreibt. Die Vorlage bekommt Platzhalter wie `{{Charge}}` oder `{{Bewertung}}`, die Excel-Liste ist das Datenblatt. Zwei Befehle: `Setup.cmd` prüft Rechner, Liste und Vorlage und sagt in einfachen Worten, was fehlt; `Run.cmd` erstellt die Berichte für alle Reklamationen eines Zeitraums, zum Beispiel eines Jahres, die noch keinen haben. Microsoft 365 Copilot schreibt die Berichtstexte. Geprüft wird danach in Word, nicht im Terminal: jeder Bericht ist als Entwurf markiert und enthält die Ergebnisse der automatischen Prüfungen, eine Excel-Prüfliste hat Spalten für "Geprüft von" und "Geprüft am". Entworfen nach den Grundsätzen von GAMP 5, nicht validiert. Alle Beispieldaten sind synthetisch.

![Setup.cmd checks the list and the templates, Run.cmd writes ten draft reports, a second Run.cmd finds nothing left to do](docs/images/demo.gif)

## What you need

- Windows with the Microsoft 365 Copilot app, signed in with a Copilot licence. Windows PowerShell is part of Windows; you need no admin rights, no Python, and no Word or Excel for the tool itself.
- Your Word report template (.docx). You type placeholders into it, for example `{{Batch}}` where the batch number goes and `{{Assessment}}` where the AI text goes. [How to prepare your template](docs/word-template.md).
- Your Excel complaint list (.xlsx), one row per complaint. Any column names; you tell the tool which column holds what.

## Three steps

1. Download: on GitHub, "Code" > "Download ZIP", then unpack the folder (or `git clone`).
2. Prepare: put the placeholders into your Word template, then run `Setup.cmd` (double-click it, or type `.\Setup.cmd` in a terminal in that folder; if Windows asks whether to run a downloaded file, choose "Run"). It asks for your Excel list, the worksheet, which column is which, your templates, the period and the Copilot model, writes `config.json` and checks everything. Fix what it lists and run it again until it says "Ready".
3. Open the Copilot app, then run `Run.cmd`. It writes a report for every complaint of the period that has none yet. You can stop it at any time; the next `Run.cmd` continues with the complaints still missing.

## Try it without Copilot

```powershell
.\Setup.cmd -Demo
.\Run.cmd -Demo
```

Demo mode uses the synthetic list in `samples/` (10 complaints, invented batch numbers like `SYN-LOT-26A019`) and an answer written in advance for each complaint instead of Copilot. On Linux or macOS: `pwsh ./Invoke-ComplaintDossier.ps1 -Demo`.

## What you get

- `out/reports/`: one Word file per complaint, for example `SYN-2026-007-long-report-DRAFT.docx`, in your template. [Sample report as PDF](docs/sample/complaint-dossier-sample-DRAFT.pdf).
- `out/review-list.xlsx`: one row per report with its automated check results and empty columns "Reviewed by", "Reviewed on", "Released (yes/no)" and "Comment". A second sheet lists complaints that got no report and why.
- `out/audit/<complaint no.>/`: the prompt, the AI answer, the data used and an audit log with SHA-256 hashes. You do not need to open these.

## How review works

Every report is a draft: it says so in the draft notice (in the header of the example template), in the Word property "Status" and in its file name. Open it in Word, check every statement against the case file, correct what is wrong, and release it the way your quality system requires, for example with a signature or in your document system. The report shows the automated checks it passed or failed; failed checks also appear at the top of the run summary. Enter who checked it and when in the review list. The tool has no approval step of its own.

## Limitations

- A prototype, designed along GAMP 5 principles and not validated. [Design note](docs/design-note.md): intended use, GAMP 5 reasoning, risks, what a company would still need to do.
- The live path drives the Copilot app's user interface. App updates or another UI language can break it. It was tested by hand on one computer ([live test](#live-copilot-tests)); CI cannot test it.
- The prompts carry process rules of the original workflow (`Reports.*.ProcessFacts`, for example "The batch record was reviewed and showed no findings."). Replace them with what your process guarantees.
- Masking removes only e-mail addresses and phone numbers from the prompt text. Files uploaded to Copilot are sent unchanged.

## My role

Complaint Dossier is a public rebuild of a tool I first built in my working-student job in pharmaceutical quality assurance, where it turned the complaint list into individual report drafts in the company template, and I reviewed each one. For this version I set the requirements and the scope; the two report prompts follow the prompts I wrote for the original tool, translated and without anything company-specific. An Claude Opus 5.5 rewrote the code here based on the original project but without company data. The data is synthetic and the templates are neutral. This repo should work on a PC without admin rights and the ability to install any form of node.js or python on the system. 

## Tech stack

PowerShell (Windows PowerShell 5.1 or PowerShell 7), OpenXML read and written directly (no Word or Excel needed), Windows UI Automation for the Microsoft 365 Copilot desktop app, Pester 5, GitHub Actions on `windows-latest`, gitleaks.

---

## Technical details

### The two commands

| Command | What it does |
| --- | --- |
| `Setup.cmd` | Unblocks the downloaded files, checks PowerShell, Windows and the Copilot app, creates `config.json` from `config.example.json` and asks the setup questions (again with `-Reconfigure`). Then it checks the Excel list (worksheet, header row, every mapped column, unreadable rows, how many complaints fall into the period), each Word template (every placeholder known, every AI section placed, draft notice and check results present), the Copilot model and the output folder. Exit code 0 when ready. |
| `Run.cmd` | Runs the same checks and refuses to start on any error. Then, one complaint after another: builds the prompt, gets the answer (Copilot or demo), runs the automated checks, fills the template, writes the audit log. Prints progress with the time left, and at the end the summary, the report folder and the review list, and opens the report folder. Options: `-Demo`, `-CaseId <no.>`, `-From`/`-To` (yyyy-MM-dd), `-Regenerate`, `-NoOpen`. |

Both `.cmd` files start `powershell.exe -ExecutionPolicy Bypass -File ...`, so no execution policy change is needed, and they keep the window open when they were double-clicked. `Invoke-ComplaintDossier.ps1` is the script behind `Run.cmd`; without `-Demo` it runs in live mode with `config.json`.

### Which complaints get a report

- Every complaint received between `Selection.From` and `Selection.To`, with a status from `Selection.Statuses` (empty = any). With `-CaseId`: exactly these complaints.
- Short or long report: an empty file `short.txt` or `long.txt` in the case folder, else the column `Workbook.Columns.ReportType` (`short`/`kurz`, `long`/`lang`), else `Reports.DefaultType`. Without any of them the complaint is skipped and listed.
- A complaint is done once `out/audit/<no.>/audit.json` exists. That file is written last, so an interrupted complaint is generated again on the next run. `-CaseId <no.> -Regenerate` writes a report again and moves the earlier one to `out/replaced/`.
- Complaints run one after another. A failed complaint does not stop the others; three failures in a row stop the run (`Run.StopAfterFailuresInARow`), because then Copilot is usually not usable. The exit code is then 2.

All settings: [docs/configuration.md](docs/configuration.md).

### Short and long report

Both prompts are in `prompts/`. They ask for fixed headings in a fixed order; `Reports.*.Sections` maps each heading to a placeholder of the template.

| | Short report | Long report |
| --- | --- | --- |
| Sections | Basic complaint data, 1 to 4: defect, statistical and historical classification, assessment, conclusion | Basic complaint data, 1 to 9: defect, same batch, same product, returned sample, batch record, analyses and retained samples, manufacturing defect, corrective actions, risk and conclusion |
| Sentences the tool requires word for word | Section 2 opens with "The reported defect is classified as <level>.", with the frequency level the tool computed | "No complaint sample was available for the investigation." when the list says no sample was received |
| Ending | | "The complaint is classified as justified." or "... not justified." |

The tool, not the model, computes the similar complaints: other complaints with the same batch in the whole list, and with the same product group and category in the 12 months up to the complaint. The frequency level (rare, occasional, frequent, very frequent) follows from that count (`SimilarComplaints.FrequencyLevels`).

### Automated checks

For each answer: not empty; not a copy of the prompt; all required sections present; names its own complaint number; every cited complaint number exists in the list; required sentences present; a long report ends with a classification. They check form, not content. Results go into the report (`{{CheckSummary}}`), the review list and audit.json.

### Live mode with Microsoft 365 Copilot

For each complaint the tool opens a new temporary chat, selects the model from `Copilot.PreferredModel` (it stops if the menu does not offer it and never falls back to another model), attaches the complaint's files through the app's own file dialog, types the prompt into the chat input, reads it back and sends only if it is complete, waits for the whole answer and reads it through UI Automation. Everything uses UI Automation patterns or window messages to the Copilot window, so it needs neither the foreground window nor the clipboard. It runs in the same Windows session as the app; do not type into the Copilot window during a run. Each step is logged in `out/audit/<no.>/copilot-status.json`. Uploaded files (`Copilot.Upload.Extensions`) are sent unchanged, so review or redact them first.

### Template filling

The tool reads and writes the .docx parts (document, headers, footers) directly. A placeholder is matched on the visible text of its paragraph, so Word's habit of splitting text into several runs does not matter; the value takes the format of the placeholder's first character. AI sections replace their placeholder paragraph with as many paragraphs as needed, in the placeholder's paragraph format, with the template's list and heading styles. An unknown placeholder or one left unfilled stops that report instead of writing a broken file. Rules for template authors: [docs/word-template.md](docs/word-template.md).

### What is tested where

| What | Where | How |
| --- | --- | --- |
| Excel reading (date cells, header row, column mapping, unreadable rows), selection by period, status and report type, resuming, similar complaints, prompt, masking, uploads, checks, template filling (split runs, block placeholders, styles, headers, text boxes, section breaks, .dotx), review list (kept reviewer entries, German headers), audit hashes, Setup questions and messages | CI, `windows-latest`, every push | Pester 5 under PowerShell 7 (`tests/`) |
| Model selection logic, including abort on a missing model or group | CI | Pester with mocked UI elements |
| Round-trip script aborts cleanly without a model and without a Copilot window | CI | Pester, starts the script in Windows PowerShell 5.1 |
| `Setup.cmd -Demo` and `Run.cmd -Demo` as a user starts them, under Windows PowerShell 5.1; demo run of all 10 complaints and a second run that generates nothing | CI | Workflow steps |
| Demo reports, review list and audit logs as downloadable files | CI | Artifact `complaint-dossier-demo` (needs a GitHub login) |
| Committed samples in `docs/sample/` still match what the code produces | CI | Pester compares the report text and the hashes |
| Employer denylist and secrets over the full history | CI | Denylist from a repository secret, gitleaks |
| Live Copilot round trip | Manual, see below | Not reproducible in CI |

### Live Copilot tests

Run against the Microsoft 365 Copilot desktop app (German UI) on one Windows 11 computer under Windows PowerShell 5.1, with the model "Quick Response" and synthetic complaints only, on 30 October 2026. Both commands were started the way a QA user starts them, `Setup.cmd -NonInteractive` and then `Run.cmd -From 2026-04-01 -To 2026-04-30`, which selects SYN-2026-006 and SYN-2026-007. They ran in the signed-in Windows session through a scheduled task, while the remote desktop session was disconnected and the Copilot window was not in the foreground.

| Run | Code | What happened |
| --- | --- | --- |
| 15:11 to 15:15 UTC | Commit `fa49a19` of pull request #9 | Passed. Setup reported "Ready" with the Copilot app found and the window open. SYN-2026-006 (short report, `shipment-log.csv` uploaded): new temporary chat, model selected and confirmed, prompt typed and read back, answer read through UI Automation, all 6 checks passed. SYN-2026-007 (long report): all 7 checks passed. Copilot sent the numbered headings as list items without their number; each section still landed in its placeholder of the example templates. About 2 minutes per complaint. |
| 15:16 to 15:22 UTC | Commit `deb27bb` | SYN-2026-006 passed again (6 of 6). For SYN-2026-007 Copilot answered with its daily usage limit message ("Sie haben Ihr Tageslimit erreicht ...") instead of a report. The checks failed 4 of 7, and the run summary named the report. |

The second run led to the last change: an answer with none of the required sections now counts as a failed complaint without a report, so the next `Run.cmd` tries it again, and three such answers in a row stop the run. CI tests this with a prepared answer; it has not run live yet, because the usage limit was still active.

Earlier live tests of the Copilot automation (pull request #8, same day) covered a missing model: the run stopped with "Configured model not found in the Copilot menu" before anything was sent. The tests cover one computer, one app version and one UI language; they are not a validation.

### Project layout

```text
Setup.cmd, Setup.ps1            first command: questions, config.json, checks
Run.cmd                         second command: starts Invoke-ComplaintDossier.ps1
Invoke-ComplaintDossier.ps1     the run: selection, one report per complaint, review list
config.example.json             example settings with explanations (_help)
src/ComplaintDossier.psm1       Excel reading, selection, prompt, checks, reports, audit, review list, setup checks
src/WordTemplate.ps1            fills placeholders in .docx/.dotx templates without Word
src/messages.json               messages in English and German
src/Copilot/                    UI automation for the Copilot desktop app (Windows PowerShell 5.1)
prompts/                        short and long report prompt
templates/                      neutral example templates for the short and the long report
samples/                        synthetic complaint list, case folders, one prepared answer per complaint
tools/                          generators for the sample data and the example templates
tests/                          Pester tests
docs/                           template guide, settings, design note, sample run
```

## License

MIT, see [LICENSE](LICENSE).
