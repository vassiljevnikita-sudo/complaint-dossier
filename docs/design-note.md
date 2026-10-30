# Design note: Complaint Dossier and GAMP 5

Complaint Dossier is a prototype. It is designed along GAMP 5 principles, but it is not validated, and no tool can be validated on its own. Validation belongs to a company, its intended use and its quality system. This note explains the design decisions and lists what a company would still have to do before using it with real complaint data.

## Intended use

The tool drafts the report for individual product complaints. Its inputs are the company's own Word report template, the complaint list in Excel (one row per complaint) and, optionally, a case folder per complaint. In one run it writes a draft report for every complaint of a period that has no report yet, for example all complaints of a year. A trained reviewer then opens each draft in Word, checks it against the case file, corrects or rejects it, and finalises it the way the company's quality system requires. Only the complaints and their reviewed reports are part of the complaint process. The tool only prepares drafts.

The tool does not decide whether a complaint is justified, does not close cases, does not start CAPAs, does not release anything and does not write into any GxP system. The classification sentence at the end of a long report is a proposal that the reviewer accepts or changes.

## Why the tool output sits outside the regulated record

The record is the report a reviewer has checked and released in the company's quality system, not the file the tool wrote. Until that review, a generated report is a draft outside the regulated record. That is the main GAMP 5 argument for the design: the human review of each generated report in Word, before it is used, is the control.

- Every report is reviewed on its own. The tool generates many reports in one run, but there is no batch release: each draft carries its own draft notice and its own check results.
- The review happens where the reviewer normally works, in Word, with the company's own template. The reviewer can correct the text directly.
- Release happens in the company's quality system: a signature on the report, the release step of a document management system, or whatever the complaint SOP defines. That is where authenticated signatures, access control and record retention already exist.

Under this intended use the tool has an indirect GxP impact. Its output influences a quality record, but a person reviews and releases every report.

### Why there is no approval step in the tool any more

An earlier version had a terminal command that "approved" one report: it checked the file hashes, removed the DRAFT marking and stored a typed reviewer name. It was dropped for three reasons. A typed name in a terminal looks like an electronic signature but is not one. A QA reviewer who is not technical should not have to use a terminal to do their review. And the release belongs to the company's quality system, which has to hold the record anyway. The hashes stay in the audit log, so a company can still check that the tool's inputs and outputs were not changed afterwards.

## GAMP 5 software categories

| Component | Category | Reasoning |
| --- | --- | --- |
| Complaint Dossier scripts (PowerShell) | 5, custom application | Written for this workflow. Case selection, similar-complaint counts, prompt assembly, checks, template filling, review list and audit log are custom code, so they need full lifecycle documentation and testing. |
| Microsoft 365 Copilot (desktop app and service) | Treated as a supplier-provided cloud service; the language model is a non-deterministic component | The company configures it at tenant level but cannot test the model itself. The same prompt can give different answers. ISPE GAMP 5 Second Edition (2022), Appendix D11 "Artificial Intelligence and Machine Learning", describes a risk-based approach for such components. Here that means controls around the model's output: fixed prompts, computed facts, automated checks and a review of every report. |
| Excel workbook as complaint list | 3 for Excel itself, the workbook is data | The tool reads the .xlsx file directly without Excel. The configuration names the sheet, the header row and which column holds what; Setup checks that these columns exist. |
| Word report template | Configuration, under document control | The company's own template. The tool fills placeholders and leaves everything else unchanged. Setup checks that every placeholder is known and every AI section has a place. |
| Prompts, process facts | Configuration, under document control | The prompts define report structure and rules. `Reports.*.ProcessFacts` states facts the model must treat as given (see risks). |
| Windows, PowerShell, Word | 1, infrastructure | Qualified as part of the workstation, not of this tool. Word is only used by the reviewer; the tool writes the .docx files without it. |

## Risk view

The risks below are the ones considered in the design. A real assessment would use the company's risk method (for example an FMEA) and its own severity scale.

| Risk | Effect | Control in the tool | Remaining gap |
| --- | --- | --- | --- |
| Wrong complaint selected or report written for the wrong complaint | A report describes another complaint | Selection by period, status or explicit complaint number; one prompt per complaint with only that row, its documents and its uploads; one report file and one audit folder per complaint. A check fails when the answer does not name the report's own complaint number. | Case folders are matched by complaint number only. A misfiled document in the wrong case folder goes into the wrong prompt. |
| Template field filled from the wrong column | Wrong batch, product or date in the report | Setup refuses column names that do not exist in the list and suggests the closest one. The values come straight from the Excel cell; the AI does not touch them. case-data.json stores every filled value. | A mapping to an existing but wrong column (for example an order number mapped as batch) is not detected. The first reports after setting up a template must be checked field by field. |
| Wrong similar-complaint counts or frequency level | History of the batch or product misstated | Counts and level are computed by the tool, not the model, stored in case-data.json and shown in the report. Pester tests check the windows and thresholds. The short report must repeat the level word for word. | The thresholds are a configuration choice, not a validated statistical method. |
| Model invents facts, findings or cases | Reviewer reads false statements | The prompts forbid it and require missing evidence to be named as missing. Automated checks flag cited complaint numbers that are not in the list, missing sections, missing required sentences and a long report that does not end with a classification. The results are printed in every report and listed in the review list, failed checks first. | Invented facts without a complaint number are not detected automatically. Only the review catches them. |
| AI text lands in the wrong place of the template, or is lost | Report misses a section or mixes sections | Each section of the answer goes to its own placeholder, matched by the heading the prompt requires. Setup refuses a template in which a section has no place. A report with a placeholder left unfilled is not written. | If the model writes text outside the required headings, that text only appears with `{{ReportText}}`. |
| Process facts are wrong for a case | A report states that a batch record review found nothing when it was not done | The facts are visible in config, prompt and case-data.json, and the README names them as process assumptions of the original workflow. | The tool cannot know whether the facts are true. A company must replace them with what its own process guarantees, or remove them. |
| Answer read from the wrong place | The prompt or an older answer ends up in the report | Every complaint uses a new temporary chat. The answer is read from the newest answer element of the app's UI Automation tree, only after a new answer exists and its text stayed the same over several polls. Text that is the prompt is rejected, and a check fails when the answer contains the data headings of the prompt. | The answer element is found through the app's accessibility tree; an app update can change it. |
| Prompt lands with the wrong model | Unknown model behaviour | The model comes from the configuration. If the Copilot menu does not offer it, the run stops. It never falls back to another model. | The model name in the menu says nothing about the model version behind it. Microsoft can change it without notice. |
| UI automation acts on the wrong window or control | Wrong file uploaded, prompt sent elsewhere or incomplete | The window is found by its title and process tree. Typed text goes as window messages to the Copilot web view, and the prompt is sent only after the chat input reads back the complete prompt. The upload helper only accepts a file dialog owned by the Copilot window. Three failed complaints in a row stop the run. | UI automation depends on the app version and the UI language. |
| Personal data sent to the model | Data protection issue | E-mail addresses and phone numbers are masked by default in the description and the case document text inserted into the prompt. The prompts tell the model not to copy contact data into the report. | Masking is not anonymisation. Names, addresses and free-text details pass through. Uploaded files are sent unchanged. |
| A draft is used without review | AI text enters the record unchecked | Every report carries the draft notice (in the template's `{{DraftNotice}}` or as a red first line), the Word property "Status" says draft, the file name ends in `-DRAFT`, and the review list has empty "Reviewed by", "Reviewed on" and "Released" columns for every report. | The draft is an ordinary Word file. The tool cannot stop anyone from removing the notice and using the file. The complaint SOP must say that only released reports enter the record. |
| Reviewer releases without reading | AI text enters the record unchecked | One draft per complaint, the check results inside the report, failed checks at the top of the list and in the run summary. | No tool can prove that someone read a document. This is a training and procedure topic. A run of a whole year creates many reports at once; the review must be planned for that volume. |
| Files changed after generation | The draft under review is not what the tool produced | audit.json stores SHA-256 of the case data, prompt, answer and generated report. Edits during review change the report file; that is expected and the hash shows the state before review. | audit.json is a plain file. Whoever can edit the output folder can rewrite the hashes too. |

## Human-in-the-loop control

1. Every generated report is a draft: draft notice from `Review.DraftNotice`, Word document status from `Review.DocumentStatus`, file name ending in `-DRAFT.docx`.
2. The AI-written sections are labelled as such in the template, and `{{Model}}` states which model was configured (or that demo mode used a prewritten text).
3. The automated checks are printed in each report, with failed checks first, and they are listed per report in `review-list.xlsx`. They check form, not content, and the report says so.
4. The reviewer reviews each report in Word against the case file, corrects it and finalises it as the quality system requires. The tool takes no part in that step.
5. `review-list.xlsx` has empty columns "Reviewed by", "Reviewed on", "Released (yes/no)" and "Comment" for the team. It is a working list for planning the review, not a signature. The tool keeps these entries when it writes the list again.
6. A complaint with a generated report is not drafted again by later runs. Regenerating needs an explicit `-Regenerate`, and the earlier draft and its audit log are kept under `out/replaced/`.

## What a company would still need to do

This list is the difference between a prototype and a system in GxP use. None of it is done in this repository.

- **Requirements and risk assessment.** User requirements for the intended use, a documented risk assessment, and a decision on the validation approach (for example a CSA-style, risk-based approach).
- **Template and mapping under document control.** The Word template, the column mapping in config.json and the prompts are controlled documents. A change to any of them needs a change record and a check of the first reports.
- **Prompts and process facts.** Review the prompts against the company's complaint SOP. Replace or remove the process facts; decide which classification rules (for example "incomplete data means not justified" in the short report) the company actually wants.
- **Computerised system validation.** A validation plan and report. The Pester tests here can serve as a starting point for OQ test scripts, but they run on synthetic data and are not executed under a test protocol.
- **IQ.** Installation on controlled workstations: fixed versions of PowerShell, the Copilot app and the scripts, with checksums of the installed files.
- **OQ.** Tests of each function against the requirements, including the live Copilot path on the qualified app version and UI language, and the failure paths (missing model, missing window, upload rejected, answer not read).
- **PQ.** A trial period with real, representative complaints, comparing the drafts with reports written by hand.
- **Review and release procedure.** An SOP for the review in Word: what "checked" means, how failed automated checks are handled, how a released report is marked and filed, and that drafts never enter the complaint record. Training records for reviewers.
- **ALCOA+ and electronic records.** The output folder and audit.json are not protected against changes. A real deployment needs access control, versioned storage, retention and backup for whatever it keeps (EU GMP Annex 11, 21 CFR Part 11 where applicable). The released report is normally filed in the company's document or complaint system, which then holds the record and the signature.
- **Copilot tenant governance.** Supplier assessment of Microsoft, data processing agreement, data residency, which users may use Copilot and which models, retention and logging of prompts and answers (for example in Microsoft Purview), sensitivity labels, and a rule for what data may be sent at all.
- **Change control and periodic review.** Copilot app updates, model changes and Microsoft UI changes must trigger an impact assessment. Model behaviour should be re-checked on a schedule with a fixed set of test complaints.
- **Data protection.** A data protection impact assessment. The masking in this tool is a convenience, not a control a DPO can rely on.
