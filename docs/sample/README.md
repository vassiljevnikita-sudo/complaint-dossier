# Sample reports

Two of the six reports from one demo run on the synthetic sample data, so you can read them without running anything. Both are DRAFTs: nobody has approved them. The report text in section 4 is the prewritten sample answer for that case from `samples/demo-responses/`, not output from Copilot.

| Folder | Complaint | Report |
| --- | --- | --- |
| [SYN-2026-007/](SYN-2026-007/) | Foreign fibre in a cream, rated critical, no returned sample | Long report, [PDF](SYN-2026-007/SYN-2026-007-long-report-DRAFT.pdf), 3 pages |
| [SYN-2026-006/](SYN-2026-006/) | Carton crushed in transport, with a CSV file in the case folder | Short report, [PDF](SYN-2026-006/SYN-2026-006-short-report-DRAFT.pdf), 2 pages |

[complaint-dossier-sample-DRAFT.pdf](complaint-dossier-sample-DRAFT.pdf) is a copy of the SYN-2026-007 PDF under a stable name for links from elsewhere.

Each folder holds the files of one report as the tool wrote them:

| File | Content |
| --- | --- |
| `<case id>-<type>-report-DRAFT.docx` | The report |
| `audit.json` | SHA-256 of every input and output of this report, checks, events, empty review field |
| `prompt.txt`, `response.md`, `case-data.json` | The files whose hashes audit.json records |

## How it was made

- The run: `./Invoke-ComplaintDossier.ps1 -Demo -OutputPath out/demo`, step "Demo run, one report per open case (PowerShell 7)" of CI run [41036469722](https://github.com/vassiljevnikita-sudo/complaint-dossier/actions/runs/41036469722) for pull request [#6](https://github.com/vassiljevnikita-sudo/complaint-dossier/pull/6), `windows-latest`, PowerShell 7.6.7, run ID `20261030T095212Z-0d886d`. The two folders are copied unchanged from the build artifact of that run. Its Pester step failed only because this folder did not exist yet.
- The PDFs: rendered from the .docx files with LibreOffice 24.2, `soffice --headless --convert-to pdf <report>.docx`. They are not listed in audit.json.

## How it stays current

The test "Committed sample in docs/sample" runs on every push. It checks the files in both folders against the hashes in their audit.json. Then it runs the demo again for SYN-2026-007 and SYN-2026-006 and compares the input hashes, the case data, prompt and response hashes, and the report text apart from run ID and timestamps. If a code change alters a report, CI fails until this folder is regenerated:

```powershell
./Invoke-ComplaintDossier.ps1 -Demo -OutputPath out/sample
```

Copy the two case folders into this folder, render the PDFs again as above and copy the SYN-2026-007 PDF to `complaint-dossier-sample-DRAFT.pdf`. Then update the run ID and links in this file.
