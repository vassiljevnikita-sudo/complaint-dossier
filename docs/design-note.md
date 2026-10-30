# Design note: Complaint Dossier and GAMP 5

Complaint Dossier is a prototype. It is designed along GAMP 5 principles, but it is not validated, and no tool can be validated on its own. Validation belongs to a company, its intended use and its quality system. This note explains the design decisions and lists what a company would still have to do before using it with real complaint data.

## Intended use

The tool prepares a draft for a periodic complaint trend review, for example as input to a product quality review. A qualified person reads the draft, checks it against the case records and decides what goes into the official record. The tool does not classify complaints, does not decide on root causes or CAPAs and does not write into any GxP system.

Under this intended use the tool has an indirect GxP impact. Its output influences a quality decision, but a person makes that decision and signs it.

## GAMP 5 software categories

| Component | Category | Reasoning |
| --- | --- | --- |
| Complaint Dossier scripts (PowerShell) | 5, custom application | Written for this workflow. The statistics, prompt assembly, report generation and audit log are custom code, so they need full lifecycle documentation and testing. |
| Microsoft 365 Copilot (desktop app and service) | Treated as a supplier-provided cloud service; the language model is a non-deterministic component | The company configures it at tenant level but cannot test the model itself. The same prompt can give different answers. GAMP 5 Second Edition (Appendix D11) asks for controls around such AI components instead of validating them like deterministic code. |
| Excel workbook as case database | 3 for Excel itself, the workbook is data | The tool reads the .xlsx file directly without Excel. The workbook structure is fixed by the configuration (sheet, header row, column mapping). |
| Word report template | Data, under document control | The template only carries layout and two anchors. Content comes from the tool. |
| Windows, PowerShell | 1, infrastructure | Qualified as part of the workstation, not of this tool. |

## Risk view

The risks below are the ones I considered when designing the tool. A real assessment would use the company's risk method (for example an FMEA) and its own severity scale.

| Risk | Effect | Control in the tool | Remaining gap |
| --- | --- | --- | --- |
| Wrong statistics | Trend under- or overstated | Counts are computed by the tool, not by the model. Pester tests check counts, shares and empty months. The prompt tells the model not to recalculate. | Column mapping errors in a real workbook are only caught if a required column is missing. |
| Model invents cases, numbers or causes | Reviewer reads false statements | The prompt forbids it. An automated check rejects case IDs that are not in the input. The report labels the section as AI-drafted. The reviewer must check every statement. | Invented facts without a case ID are not detected automatically. Only the human review catches them. |
| Model answer incomplete | Missing sections in the review | Check for the five required sections; result shown in the report and in audit.json. | A section can exist and still be empty or shallow. |
| Prompt lands in an old chat or with the wrong model | Context from other conversations, unknown model behaviour | Live mode always opens a new chat. The model comes from the configuration. If the Copilot menu does not offer it, the run stops. It never falls back to another model. | The model name in the menu says nothing about the model version behind it. Microsoft can change it without notice. |
| UI automation acts on the wrong window or control | Wrong file uploaded, prompt sent elsewhere | The window is found by its title and process tree. The upload helper only accepts a file dialog owned by the Copilot window, and confirms each file as a chat attachment before the prompt is sent. Ambiguous matches stop the run. | UI automation depends on the app version and the UI language. An app update can break it. |
| Personal data sent to the model | Data protection issue | E-mail addresses and phone numbers in descriptions and documents are masked by default. | Masking is not anonymisation. Names, addresses and free-text details pass through. |
| Report changed after review | Signed record differs from what was reviewed | SHA-256 of prompt, response, statistics and draft report are stored at run time. Approval recomputes them and refuses on any mismatch. The signed report is a new file with its own hash. | audit.json is a plain file. Whoever can edit the run folder can rewrite the hashes too. |
| Reviewer signs without reading | AI text enters the record unchecked | The report says DRAFT in the page header until approval, and the sign-off block names what must be checked. | No tool can prove that someone read a document. This is a training and procedure topic. |

## Human-in-the-loop control

The design treats the model as a drafting aid and the reviewer as the only decision point.

1. Every run produces a report marked `DRAFT - NOT REVIEWED - NOT FOR USE` in the page header, with status, reviewer and signature fields set to "not signed".
2. The AI-drafted section is labelled as such and states which model was configured (or that demo mode used a prewritten text).
3. The automated checks are printed in the report, so the reviewer sees what was and was not checked by the tool.
4. `Approve-ComplaintDossier.ps1` is the only way to remove the DRAFT marking. It needs a reviewer name, verifies the hashes, writes a signed copy, and records reviewer, time, OS account, the hash of the reviewed draft and the hash of the signed report in audit.json. A second approval of the same run is refused.
5. The draft cannot be edited and then approved. If the reviewer disagrees with the draft, the options are a review comment or a new run. This keeps the signed report identical in content to what the reviewer saw.

## What a company would still need to do

This list is the difference between a prototype and a system in GxP use. None of it is done in this repository.

- **Requirements and risk assessment.** User requirements for the intended use, a documented risk assessment, and a decision on the validation approach (for example a CSA-style, risk-based approach).
- **Computerised system validation.** A validation plan and report. The Pester tests here can serve as a starting point for OQ test scripts, but they run on synthetic data and are not executed under a test protocol.
- **IQ.** Installation on controlled workstations: fixed versions of PowerShell, the Copilot app and the scripts, with checksums of the installed files.
- **OQ.** Tests of each function against the requirements, including the live Copilot path on the qualified app version and UI language, and the failure paths (missing model, missing window, upload rejected).
- **PQ.** A trial period with real, representative complaint data, comparing the drafts with reviews written by hand.
- **ALCOA+ and electronic records.** The reviewer name is typed, not authenticated, so it is not an electronic signature. audit.json is not protected against changes. A real deployment needs authenticated signatures, access control, write-once or versioned storage, a trusted time source, retention and backup (EU GMP Annex 11, 21 CFR Part 11 where applicable).
- **Copilot tenant governance.** Supplier assessment of Microsoft, data processing agreement, data residency, which users may use Copilot and which models, retention and logging of prompts and answers (for example in Microsoft Purview), sensitivity labels, and a rule for what data may be sent at all.
- **Change control and periodic review.** Copilot app updates, model changes and Microsoft UI changes must trigger an impact assessment. Model behaviour should be re-checked on a schedule with a fixed set of test cases.
- **Procedures and training.** An SOP for the review step, including what "checked" means, and training records for reviewers.
- **Data protection.** A data protection impact assessment. The masking in this tool is a convenience, not a control a DPO can rely on.
