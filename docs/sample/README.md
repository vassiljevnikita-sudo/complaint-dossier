# Sample report

The output of one demo run on the synthetic sample data, so you can read a report without running anything. The report is a DRAFT: nobody has approved it. Section 3 holds the prewritten sample answer from `samples/demo-response.md`, not output from Copilot.

| File | Content |
| --- | --- |
| [complaint-dossier-sample-DRAFT.pdf](complaint-dossier-sample-DRAFT.pdf) | The report as PDF, 2 pages |
| [complaint-dossier-20261030T023604Z-7f1975-DRAFT.docx](complaint-dossier-20261030T023604Z-7f1975-DRAFT.docx) | The report as the tool wrote it |
| [audit.json](audit.json) | SHA-256 of every input and output, checks, events, empty review field |
| [prompt.txt](prompt.txt), [response.md](response.md), [statistics.json](statistics.json) | The files whose hashes audit.json records |

## How it was made

- The run: CI run [40992565280](https://github.com/vassiljevnikita-sudo/complaint-dossier/actions/runs/40992565280) on commit [8647e60](https://github.com/vassiljevnikita-sudo/complaint-dossier/commit/8647e60ef83777978cac42dc52f508178a322694) (pull request #2), `windows-latest`, PowerShell 7.6.7, step "Demo run (PowerShell 7)": `./Invoke-ComplaintDossier.ps1 -Demo -OutputPath out/demo`. The .docx, audit.json, prompt, response and statistics are the build artifact of that run, unchanged.
- The PDF: rendered from the .docx with LibreOffice 24.2, `soffice --headless --convert-to pdf <report>.docx`, then renamed. It is not listed in audit.json.

## How it stays current

The test "Committed sample in docs/sample" runs on every push. It checks the files in this folder against the hashes in audit.json. Then it runs the demo again and compares the input hashes, the prompt, statistics and response hashes, and the report text apart from run ID and timestamps. If a code change alters the report, CI fails until this folder is regenerated:

```powershell
./Invoke-ComplaintDossier.ps1 -Demo -OutputPath out/sample
```

Copy the run files into this folder and render the PDF again as above. Then update the run ID, commit and run links in this file.
