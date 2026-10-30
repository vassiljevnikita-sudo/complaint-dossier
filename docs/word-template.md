# Your Word template

Complaint Dossier writes each report into your own Word template. You keep your layout, logo, header, footer, fonts and tables. Where the tool should put something, you type a placeholder.

Two example templates show how it looks: [templates/short-report-template.docx](../templates/short-report-template.docx) and [templates/long-report-template.docx](../templates/long-report-template.docx). Open one in Word next to this page.

## Placeholders in plain words

- A placeholder is a name between two pairs of curly braces: `{{Batch}}`.
- Type it as normal text, anywhere: in a paragraph, a table cell, the header, the footer or a text box.
- The name may contain letters, digits and `_`, and starts with a letter. No spaces, umlauts or hyphens in the name: `{{ComplaintDate}}`, not `{{Complaint date}}`.
- Upper and lower case do not matter, and neither do spaces inside the braces: `{{batch}}` and `{{ Batch }}` both work.
- The filled-in text takes the font, size and colour of the placeholder. Make `{{Batch}}` bold in the template, and the batch number is bold in every report.
- Word sometimes stores a word in several pieces, for example after a spelling correction or when you typed part of it later. The tool reads the text the way you see it, so such a placeholder still works.
- The same placeholder may appear several times, for example `{{CaseId}}` in the title and in the header.
- Save the template as `.docx` (or `.dotx`). Old `.doc` files do not work; open them in Word and save them as `.docx`.

After every change to the template, run `Setup.cmd`. It lists every placeholder it found and tells you if one is unknown, for example because of a typo, and which name was probably meant.

## Which placeholders exist

### Values from your Excel list

You choose these names yourself in `config.json` under `Template.Fields`: the placeholder name on the left, the column of your Excel list on the right.

```json
"Fields": {
  "CaseId": "Complaint no.",
  "ComplaintDate": "Date received",
  "Batch": "Batch"
}
```

With this, `{{ComplaintDate}}` becomes the content of the column "Date received" of that complaint's row. Dates are written as set in `Template.DateFormat` (for example `dd.MM.yyyy`). An empty cell becomes the text in `Template.EmptyValue` (for example `n/a`). You can add as many fields as your list has columns.

### Text written by the AI

Each section of the AI answer goes into its own placeholder. The names are set in `config.json` under `Reports.Short.Sections` and `Reports.Long.Sections`, next to the heading the AI must write:

```json
{ "Heading": "3. Assessment of the available information", "Placeholder": "Assessment" }
```

`{{Assessment}}` then receives the text the AI wrote under "3. Assessment of the available information", without the heading itself, because your template has its own heading there.

Put each of these placeholders in a paragraph of its own. The tool then replaces that paragraph with as many paragraphs as the text needs. Every new paragraph gets the paragraph format of the placeholder paragraph (indent, spacing, alignment). List lines get the bullet style named in `Template.Styles.Bullet` (in Word: "List Bullet" / "Aufzählungszeichen"), if your template has it; otherwise they start with a bullet character. If such a placeholder shares its paragraph with other text, the AI text is inserted as plain text with line breaks, and `Setup.cmd` warns you.

Instead of one placeholder per section you can also use `{{ReportText}}` once: it receives the whole AI answer with its headings. Headings then get the style named in `Template.Styles.Heading`.

Every AI section needs a place. If a section has no placeholder in the template and there is no `{{ReportText}}`, `Setup.cmd` reports it and the run does not start, so no AI text is silently lost.

### Filled in by the tool

| Placeholder | Becomes |
| --- | --- |
| `{{DraftNotice}}` | The draft notice from `Review.DraftNotice`, for example "DRAFT ... Not reviewed." Put it in the header so it is on every page. |
| `{{CheckSummary}}` | The results of the automated checks, one line per check, failed checks first. Own paragraph. |
| `{{SimilarComplaints}}` | Earlier complaints with the same batch, and with the same product group and category in the last 12 months, as the tool counted them. Own paragraph. |
| `{{CaseDocuments}}` | Which documents of the case folder went into the prompt or were uploaded to Copilot. Own paragraph. |
| `{{ReportText}}` | The whole AI answer with headings (instead of one placeholder per section). Own paragraph. |
| `{{ReportTitle}}` | `Reports.Short.Title` or `Reports.Long.Title` |
| `{{ReportType}}` | "short report" or "long report" |
| `{{FrequencyLevel}}` | rare, occasional, frequent or very frequent, as the tool computed it |
| `{{GeneratedAt}}` | Date and time of generation |
| `{{Model}}` | Which AI wrote the text (or that the demo used a prewritten answer) |
| `{{RunId}}` | ID of the run, the same as in the audit log |

If the template has no `{{DraftNotice}}`, the tool puts the draft notice as a red first line at the top of each report. If it has no `{{CheckSummary}}`, the tool adds the check results at the end. A report is never without either.

## What the tool does not change

Everything else in the template stays as it is: styles, tables, images, fields such as page numbers, headers and footers. The tool also sets the Word document property "Status" (File > Info) to the value of `Review.DocumentStatus`, for example "Draft - to be reviewed".

## Not supported

- Word content controls and mail-merge fields. Type placeholders as plain text instead.
- Macro templates (`.docm`, `.dotm`).
- Placeholders in footnotes, endnotes and comments.
