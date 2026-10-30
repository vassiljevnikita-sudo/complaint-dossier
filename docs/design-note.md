# Design note: Complaint Dossier and GAMP 5

Complaint Dossier is a prototype. It is designed along GAMP 5 principles, but it is not validated, and no tool can be validated on its own. Validation belongs to a company, its intended use and its quality system. This note explains the design decisions and lists what a company would still have to do before using it with real complaint data.

## Intended use

The tool drafts the report for individual product complaints. Its input is the complaint list (one row per complaint) and the case folder of each complaint. For each selected complaint it writes one draft report. A trained reviewer reads each draft, checks it against the case file, corrects or rejects it, and decides what goes into the complaint record. Only the complaints and their reviewed reports are part of the complaint process. The tool only prepares drafts.

The tool does not decide whether a complaint is justified, does not close cases, does not start CAPAs and does not write into any GxP system. The classification sentence at the end of a long report is a proposal that the reviewer accepts or changes.

## Why the tool sits outside the regulated record

The record is the report a reviewer has checked and approved, not the tool's output. That is the main GAMP 5 argument for the design:

- Every report is reviewed on its own. There is no batch approval of several reports.
- Each report has its own audit.json, so the review of one complaint does not depend on any other.
- A report cannot leave DRAFT status without a named reviewer, and not while an automated check has failed, unless the reviewer writes down why the failure is acceptable.
- The draft cannot be edited and then approved. If the reviewer disagrees, the options are a review comment, a new draft or writing the report by hand in the company's own system.

Under this intended use the tool has an indirect GxP impact. Its output influences a quality record, but a person makes every decision and approves every report.

## GAMP 5 software categories

| Component | Category | Reasoning |
| --- | --- | --- |
| Complaint Dossier scripts (PowerShell) | 5, custom application | Written for this workflow. Case selection, similar-complaint counts, prompt assembly, checks, report generation and audit log are custom code, so they need full lifecycle documentation and testing. |
| Microsoft 365 Copilot (desktop app and service) | Treated as a supplier-provided cloud service; the language model is a non-deterministic component | The company configures it at tenant level but cannot test the model itself. The same prompt can give different answers. ISPE GAMP 5 Second Edition (2022), Appendix D11 "Artificial Intelligence and Machine Learning", describes a risk-based approach for such components. Here that means controls around the model's output: fixed prompts, computed facts, automated checks and a review of every report. |
| Excel workbook as complaint list | 3 for Excel itself, the workbook is data | The tool reads the .xlsx file directly without Excel. The workbook structure is fixed by the configuration (sheet, header row, column mapping). |
| Prompts, process facts, Word report template | Configuration and data, under document control | The prompts define report structure and rules. `Reports.*.ProcessFacts` states facts the model must treat as given (see risks). The template only carries layout and two anchors. |
| Windows, PowerShell | 1, infrastructure | Qualified as part of the workstation, not of this tool. |

## Risk view

The risks below are the ones considered in the design. A real assessment would use the company's risk method (for example an FMEA) and its own severity scale.

| Risk | Effect | Control in the tool | Remaining gap |
| --- | --- | --- | --- |
| Wrong case selected or report written for the wrong case | A report describes another complaint | Selection by status or explicit case ID; one folder per case; the prompt holds only that case's row, documents and uploads. A check fails when the answer does not cite the report's own case ID. | Case folders are matched by case ID only. A misfiled document in the wrong case folder goes into the wrong prompt. |
| Wrong similar-complaint counts or frequency level | History of the batch or product misstated | Counts and level are computed by the tool, not the model, and stored in case-data.json. Pester tests check the windows and thresholds. The short report must repeat the level word for word. | The thresholds are a configuration choice, not a validated statistical method. Column mapping errors in a real workbook are only caught if a required column is missing. |
| Model invents facts, findings or cases | Reviewer reads false statements | The prompts forbid it and require missing evidence to be named as missing. Automated checks flag cited case IDs that are not in the complaint list, missing sections, missing required sentences and a long report that does not end with a classification. Approval is refused while any check fails, unless the reviewer gives a written justification. | Invented facts without a case ID are not detected automatically. Only the human review catches them. |
| Process facts are wrong for a case | A report states that a batch record review found nothing when it was not done | The facts are visible in config, prompt and case-data.json, and the README names them as process assumptions of the original workflow. | The tool cannot know whether the facts are true. A company must replace them with what its own process guarantees, or remove them. |
| Answer copied from the wrong place | The prompt or an older answer ends up in the report | Every run uses a new temporary chat. The copy step checks that the clipboard changed and rejects the prompt. A check fails when the answer contains the data headings of the prompt. | The copy button is found through the app's accessibility tree; an app update can change it. |
| Prompt lands with the wrong model | Unknown model behaviour | The model comes from the configuration. If the Copilot menu does not offer it, the run stops. It never falls back to another model. | The model name in the menu says nothing about the model version behind it. Microsoft can change it without notice. |
| UI automation acts on the wrong window or control | Wrong file uploaded, prompt sent elsewhere | The window is found by its title and process tree. The upload helper only accepts a file dialog owned by the Copilot window and confirms each file as a chat attachment before the prompt is sent. Ambiguous matches stop the run. | UI automation depends on the app version, the UI language and a displayed desktop. |
| Personal data sent to the model | Data protection issue | E-mail addresses and phone numbers are masked by default in the case description and in case document text inserted into the prompt. The prompts tell the model not to copy contact data into the report. | Masking is not anonymisation. Names, addresses and free-text details pass through. Uploaded files are sent unchanged. |
| Report changed after review | Approved record differs from what was reviewed | SHA-256 of case data, prompt, response and draft report are stored at run time. Approval recomputes them and refuses on any mismatch. The approved report is a new file with its own hash. | audit.json is a plain file. Whoever can edit the run folder can rewrite the hashes too. |
| Reviewer approves without reading | AI text enters the record unchecked | Each report says DRAFT in the page header until approval, approval is per report, and the approval block names what must be checked. | No tool can prove that someone read a document. This is a training and procedure topic. |

## Human-in-the-loop control

1. Every report is written with `DRAFT - NOT REVIEWED - NOT FOR USE` in the page header, with status, reviewer and approval time set to "not approved".
2. The AI-drafted section is labelled as such and states which model was configured (or that demo mode used a prewritten text).
3. The automated checks are printed in each report, so the reviewer sees what the tool did and did not check.
4. `Approve-ComplaintDossier.ps1 -Path <run>/<case id>` is the only way to remove the DRAFT marking of a report. It needs a reviewer name, verifies the hashes, writes an approved copy, and records reviewer, time, OS account, the hash of the reviewed draft and the hash of the approved report in that report's audit.json. The reviewer name is typed, so this approval record is not an electronic signature. A second approval is refused.
5. Approval is refused while any automated check of the report has failed. A reviewer who accepts the failures has to pass `-AcceptFailedChecks` with a non-empty `-Justification`. The approved report then carries "APPROVED WITH A DOCUMENTED EXCEPTION" in the page header and prints the accepted checks with the justification.
6. A complaint with an approved report is not drafted again by later runs.

## What a company would still need to do

This list is the difference between a prototype and a system in GxP use. None of it is done in this repository.

- **Requirements and risk assessment.** User requirements for the intended use, a documented risk assessment, and a decision on the validation approach (for example a CSA-style, risk-based approach).
- **Prompts and process facts.** Review the prompts against the company's complaint SOP. Replace or remove the process facts; decide which classification rules (for example "incomplete data means not justified" in the short report) the company actually wants.
- **Computerised system validation.** A validation plan and report. The Pester tests here can serve as a starting point for OQ test scripts, but they run on synthetic data and are not executed under a test protocol.
- **IQ.** Installation on controlled workstations: fixed versions of PowerShell, the Copilot app and the scripts, with checksums of the installed files.
- **OQ.** Tests of each function against the requirements, including the live Copilot path on the qualified app version and UI language, and the failure paths (missing model, missing window, upload rejected, answer not copied).
- **PQ.** A trial period with real, representative complaints, comparing the drafts with reports written by hand.
- **ALCOA+ and electronic records.** The reviewer name is typed, not authenticated, so it is not an electronic signature. audit.json is not protected against changes. A real deployment needs authenticated signatures, access control, write-once or versioned storage, a trusted time source, retention and backup (EU GMP Annex 11, 21 CFR Part 11 where applicable). The approved report would normally be filed in the company's document or complaint system, which then holds the record.
- **Copilot tenant governance.** Supplier assessment of Microsoft, data processing agreement, data residency, which users may use Copilot and which models, retention and logging of prompts and answers (for example in Microsoft Purview), sensitivity labels, and a rule for what data may be sent at all.
- **Change control and periodic review.** Copilot app updates, model changes and Microsoft UI changes must trigger an impact assessment. Model behaviour should be re-checked on a schedule with a fixed set of test complaints.
- **Procedures and training.** An SOP for the review step, including what "checked" means and when a failed automated check may be accepted as an exception, and training records for reviewers.
- **Data protection.** A data protection impact assessment. The masking in this tool is a convenience, not a control a DPO can rely on.
