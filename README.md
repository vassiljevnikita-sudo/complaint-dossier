# Complaint Dossier

A complaint trend review means counting cases, reading case files and writing a summary by hand. Complaint Dossier reads the cases from Excel, computes the statistics, lets Microsoft 365 Copilot draft the assessment and builds a Word report that stays a DRAFT until a named reviewer approves it.

Status: the demo mode is tested in CI on every push. It uses a prewritten answer instead of Copilot. The hand-off to the Copilot desktop app is implemented, but the full live round trip has not been confirmed against the real app yet ([details](#live-copilot-test-30-october-2026)).

[Sample report (PDF)](docs/sample/complaint-dossier-sample-DRAFT.pdf) · [all files of the sample run](docs/sample/) · [Design note: GAMP 5 categories, risks, what validation would still need](docs/design-note.md)

**Kurzfassung (Deutsch):** PowerShell-Werkzeug für die Trendauswertung von Reklamationen. Es liest Fälle aus einer Excel-Datei, berechnet Statistiken, übergibt einen Prompt an die Desktop-App von Microsoft 365 Copilot und erzeugt einen Word-Bericht und ein Audit-Log (SHA-256-Hashes). Der Bericht bleibt ein Entwurf, bis eine prüfende Person ihn freigibt. Der Demo-Modus ist in CI getestet. Die Übergabe an Copilot ist umgesetzt, der vollständige Live-Ablauf mit der echten App ist aber noch nicht bestätigt. Entworfen nach den Grundsätzen von GAMP 5; nicht validiert. Alle Beispieldaten sind synthetisch.

![Demo run: draft report, shared batch from the statistics, approval, audit log](docs/images/demo.gif)

## Try it

Needs PowerShell 7 (Windows, Linux or macOS) or Windows PowerShell 5.1. No Excel, Word or Copilot needed for the demo.

```powershell
git clone https://github.com/vassiljevnikita-sudo/complaint-dossier
cd complaint-dossier
./Invoke-ComplaintDossier.ps1 -Demo
./Approve-ComplaintDossier.ps1 -RunPath out/<run id> -Reviewer 'Your Name'
```

The demo uses the 10 synthetic cases in `samples/` (with invented batch numbers in the format `SYN-LOT-26A019`) and a prewritten answer instead of Copilot. CI runs the same command on every push. The output of one such run is committed in [docs/sample/](docs/sample/), so you can read the report without running anything.

## My role

Complaint Dossier is a public reimplementation of a workflow I first built in my working-student job in pharmaceutical quality assurance. For this version I set the requirements and the scope. An AI coding agent (Claude) wrote the code, tests and documentation, and the commits name it as co-author. The data is synthetic and the report template is neutral. Nothing from the original environment is in this repository.

## Tech stack

PowerShell 7 or Windows PowerShell 5.1 (the demo runs on either; the live Copilot automation always runs in Windows PowerShell 5.1), OpenXML read and written directly (no Office COM), Windows UI Automation for the Microsoft 365 Copilot desktop app, Pester 5, GitHub Actions on `windows-latest`, gitleaks.

## Limitations

- A prototype, designed along GAMP 5 principles. It is not validated, and the reviewer name is typed, not an authenticated electronic signature. See [docs/design-note.md](docs/design-note.md).
- The live Copilot path drives the app's user interface. App updates or another UI language can break it. CI cannot test it, because GitHub runners have no Copilot licence, and the full live round trip has not been confirmed yet. See [What is tested where](#what-is-tested-where).
- audit.json detects changed files through SHA-256 hashes, but the log itself is an ordinary file without access control. Approval reads the results of the automated checks from audit.json, so whoever can edit that file can also hide a failed check.
- Masking covers case descriptions and the case document text inserted into the prompt, and there it removes only e-mail addresses and phone numbers. Uploaded attachments are sent to Copilot unchanged and must be reviewed or redacted separately. Masking is not anonymisation.

## How to run

### Demo mode

`./Invoke-ComplaintDossier.ps1 -Demo` reads `config.example.json` and writes one folder per run to `out/<run id>/`:

| File | Content |
| --- | --- |
| `complaint-dossier-<run id>-DRAFT.docx` | The report: statistics, case list with product group and batch, AI-drafted assessment, automated checks, approval block |
| `prompt.txt` | The exact prompt, as it would go to Copilot |
| `response.md` | The answer (in demo mode the prewritten sample) |
| `statistics.json` | Counts by category, product group and month; batches with more than one case |
| `audit.json` | SHA-256 of every input and output, timestamps, checks, reviewer field |

![First page of the generated draft report](docs/images/report-page1.png)

### Live mode with Microsoft 365 Copilot

Live mode needs Windows, the Microsoft 365 Copilot desktop app signed in with a Copilot licence, and an unlocked desktop. Windows PowerShell 5.1 is enough; the UI automation always runs in it.

1. Copy `config.example.json` to `config.json`. `config.json` is in `.gitignore`.
2. Point `Workbook`, `Documents` and `Copilot.Upload.Path` at your data. Relative paths resolve against the project folder.
3. Set `Copilot.PreferredModel` to the model name exactly as the Copilot model menu shows it. If the model sits in a submenu, put the submenu name into `Copilot.PreferredModelGroup`. If the menu does not offer the model, the run stops. It never falls back to another model.
4. Check the files in `Copilot.Upload.Path`. Masking covers case descriptions and text inserted into the prompt. Uploaded attachments are sent unchanged, so review or redact them before the run.
5. Open the Copilot app, then run:

```powershell
./Invoke-ComplaintDossier.ps1 -Live -CaseId SYN-2026-007
```

The tool opens a new chat, selects the configured model, attaches every file from `Copilot.Upload.Path` through the app's own file dialog, sends the prompt, waits for the complete answer and copies it with the chat's copy button. Do not use mouse or keyboard while it runs. Each step is logged in `copilot-status.json` and in `audit.json`.

### Approval

```powershell
./Approve-ComplaintDossier.ps1 -RunPath out/<run id> -Reviewer 'A. Reviewer' -Comment 'Checked against the case files.'
```

Approval recomputes the hashes of prompt, response, statistics and draft report and refuses if anything changed after the run. It writes an approved copy without the DRAFT marking, adds reviewer, time and the new hash to `audit.json`, and keeps the draft.

Approval also refuses while any automated check in `audit.json` has failed, for example when the answer cites a case ID that is not in the workbook or misses a required section. It lists the failed checks, and the run stays DRAFT. Usually the right step is a new run. If the reviewer decides the failure is acceptable, they have to say why:

```powershell
./Approve-ComplaintDossier.ps1 -RunPath out/<run id> -Reviewer 'A. Reviewer' -AcceptFailedChecks -Justification 'SYN-2026-099 is a typo for SYN-2026-009; checked against the case list.'
```

The justification must not be empty. The approved report then says "APPROVED WITH A DOCUMENTED EXCEPTION" and lists the accepted checks with the justification. `audit.json` stores the same under `review.exception`.

### Your own data

The workbook columns are mapped in `Workbook.Columns`, so the tool works with other column names. `CaseId`, `ReceivedDate`, `Category` and `Description` are required; `ProductGroup`, `Batch`, `Severity` and `Status` are optional. Case documents are text files in `<Documents.Path>/<case id>/`. The prompt lives in `prompts/dossier-prompt.txt`, the report layout in `templates/report-template.docx` (regenerate it with `tools/New-ReportTemplate.ps1`, or use your own .docx with a `{{REPORT_BODY}}` paragraph). `tools/New-SyntheticCases.ps1` regenerates the sample data.

## What is tested where

| What | Where | How |
| --- | --- | --- |
| Excel reading, statistics, prompt, masking, response checks, report, audit hashes, approval and its refusals (changed files, failed checks, override without justification) | CI, `windows-latest`, every push | Pester 5 under PowerShell 7 (`tests/`) |
| Model selection logic, including abort on a missing model or group | CI | Pester with mocked UI elements |
| Round-trip script aborts cleanly without a model and without a Copilot window | CI | Pester, starts the script in Windows PowerShell 5.1 |
| Demo run and approval under Windows PowerShell 5.1 | CI | Workflow step |
| Demo report as a downloadable .docx | CI | Artifact `complaint-dossier-demo` (needs a GitHub login) |
| Committed sample in `docs/sample/` still matches what the code produces | CI | Pester compares the report text and the hashes in its audit.json |
| Employer denylist and secrets over the full history | CI | Denylist from a repository secret, gitleaks |
| Live Copilot round trip | Manual, see below | Not reproducible in CI |

### Live Copilot test (30 October 2026)

Run against the Microsoft 365 Copilot desktop app (package version 19.2610.35011.0, German UI) under Windows PowerShell 5.1, with the synthetic case SYN-2026-007 only. The end-to-end run is not finished yet.

- Passed: the tool found the Copilot window, brought it to the foreground and opened a new chat through UI Automation. Reading the app's controls showed how the model menu names its entries ("Quick Response Answers right away"). The model matching was changed to handle that form, and it now also checks the switcher label after the click.
- Failed: the first live run stopped at the temporary-chat toggle ("did not switch on"). The tool aborted as designed and sent nothing to Copilot. The state check now looks the button up again after the click, because the app re-renders it.
- Not run: the Windows desktop session was disconnected before the next attempt, so model selection, file upload, prompt sending and answer capture have not yet been confirmed against the real app. Until they are, treat live mode as untested.

## Project layout

```text
Invoke-ComplaintDossier.ps1     entry point: -Demo or -Live
Approve-ComplaintDossier.ps1    reviewer approval
src/ComplaintDossier.psm1       Excel reading, statistics, prompt, checks, Word report, audit, approval
src/Copilot/                    UI automation for the Copilot desktop app (Windows PowerShell 5.1)
prompts/  templates/            prompt template and neutral Word template
samples/                        10 synthetic cases, case documents, upload file, sample answer
tools/                          generators for the sample data and the Word template
tests/                          Pester tests
docs/design-note.md             GAMP 5 category reasoning, risks, what validation would still need
docs/sample/                    output of one demo run: DRAFT report (.docx and PDF), audit.json, prompt, answer
```

## License

MIT, see [LICENSE](LICENSE).
