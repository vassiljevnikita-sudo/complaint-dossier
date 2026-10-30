# Sample report

The output of one demo run on the synthetic sample data, so you can read a report without running anything. The report is a DRAFT: nobody has approved it. Section 3 holds the prewritten sample answer from `samples/demo-response.md`, not output from Copilot.

| File | Content |
| --- | --- |
| [complaint-dossier-sample-DRAFT.pdf](complaint-dossier-sample-DRAFT.pdf) | The report as PDF, 3 pages |
| [complaint-dossier-20261030T032116Z-7cb872-DRAFT.docx](complaint-dossier-20261030T032116Z-7cb872-DRAFT.docx) | The report as the tool wrote it |
| [audit.json](audit.json) | SHA-256 of every input and output, checks, events, empty review field |
| [prompt.txt](prompt.txt), [response.md](response.md), [statistics.json](statistics.json) | The files whose hashes audit.json records |

## How it was made

- The run: `./Invoke-ComplaintDossier.ps1 -Demo -OutputPath out/demo`, step "Demo run (PowerShell 7)" of CI run [40996837685](https://github.com/vassiljevnikita-sudo/complaint-dossier/actions/runs/40996837685) for pull request [#3](https://github.com/vassiljevnikita-sudo/complaint-dossier/pull/3), `windows-latest`, PowerShell 7.6.7. That run used the code that pull request #3 merged into `main`; the later commit in that pull request only changed this folder and the README screenshot `docs/images/report-page1.png`. The .docx, audit.json, prompt, response and statistics are the build artifact of that run, unchanged.
- The PDF: rendered from the .docx with LibreOffice 24.2, `soffice --headless --convert-to pdf <report>.docx`, then renamed. It is not listed in audit.json.

## How it stays current

The test "Committed sample in docs/sample" runs on every push. It checks the files in this folder against the hashes in audit.json. Then it runs the demo again and compares the input hashes, the prompt, statistics and response hashes, and the report text apart from run ID and timestamps. If a code change alters the report, CI fails until this folder is regenerated:

```powershell
./Invoke-ComplaintDossier.ps1 -Demo -OutputPath out/sample
```

Copy the run files into this folder and render the PDF again as above. Then update the run ID, commit and run links in this file.
