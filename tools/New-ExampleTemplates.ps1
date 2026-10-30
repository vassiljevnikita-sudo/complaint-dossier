<#
.SYNOPSIS
  Writes the two neutral example Word templates in templates/.

.DESCRIPTION
  templates/short-report-template.docx and templates/long-report-template.docx
  show how a report template for Complaint Dossier looks: normal Word text with
  placeholders such as {{CaseId}}, {{Batch}} or {{Assessment}}. Your own
  company template works the same way; see docs/word-template.md.
  One placeholder, {{ComplaintDate}}, is split over two runs with a spell-check
  mark in between, the way Word stores text that was edited. The tool must
  still find it, and the tests check that it does.
#>
[CmdletBinding()]
param([string]$OutputFolder)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path (Join-Path $root 'src') 'ComplaintDossier.psm1') -Force
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'templates' }
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null

$w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
$x = { param($t) [Security.SecurityElement]::Escape([string]$t) }
function P([string]$Text, [string]$Style = '', [string]$RunProps = '') {
    $ppr = ''
    if ($Style) { $ppr = "<w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr>" }
    $rpr = ''
    if ($RunProps) { $rpr = "<w:rPr>$RunProps</w:rPr>" }
    "<w:p>$ppr<w:r>$rpr<w:t xml:space=`"preserve`">$(& $x $Text)</w:t></w:r></w:p>"
}
function Cell([string]$Inner, [int]$Width, [switch]$Shade) {
    $shd = ''
    if ($Shade) { $shd = '<w:shd w:val="clear" w:color="auto" w:fill="EDF2F8"/>' }
    "<w:tc><w:tcPr><w:tcW w:w=`"$Width`" w:type=`"dxa`"/>$shd</w:tcPr>$Inner</w:tc>"
}
function Row([string]$Label, [string]$ValueXml) {
    '<w:tr>' + (Cell (P $Label 'TableLabel') 3000 -Shade) + (Cell $ValueXml 6638) + '</w:tr>'
}
function Table([string[]]$Rows) {
    '<w:tbl><w:tblPr><w:tblStyle w:val="FormTable"/><w:tblW w:w="9638" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="6638"/></w:tblGrid>' + ($Rows -join '') + '</w:tbl>'
}

$parts = [ordered]@{}
$parts['[Content_Types].xml'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>
'@
$parts['_rels/.rels'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>
'@
$parts['docProps/core.xml'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:title>Complaint report</dc:title>
<dc:creator>Complaint Dossier example template</dc:creator>
</cp:coreProperties>
'@
$parts['word/_rels/document.xml.rels'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
'@
$parts['word/styles.xml'] = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles $w>
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="21"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="100" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="60"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="2" w:color="8EAADB"/></w:pBdr><w:spacing w:before="300" w:after="100"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="22"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:spacing w:after="40"/><w:ind w:left="360" w:hanging="360"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Note"><w:name w:val="Note"/><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:color w:val="595959"/><w:sz w:val="19"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TableLabel"><w:name w:val="Table Label"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr><w:rPr><w:b/><w:sz w:val="19"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TableValue"><w:name w:val="Table Value"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr><w:rPr><w:sz w:val="19"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:sz w:val="18"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:color w:val="7F7F7F"/><w:sz w:val="16"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="FormTable"><w:name w:val="Form Table"/><w:tblPr><w:tblBorders>
<w:top w:val="single" w:sz="4" w:color="A6A6A6"/><w:left w:val="single" w:sz="4" w:color="A6A6A6"/><w:bottom w:val="single" w:sz="4" w:color="A6A6A6"/><w:right w:val="single" w:sz="4" w:color="A6A6A6"/><w:insideH w:val="single" w:sz="4" w:color="A6A6A6"/><w:insideV w:val="single" w:sz="4" w:color="A6A6A6"/>
</w:tblBorders><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>
"@
$parts['word/numbering.xml'] = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering $w><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="$([char]0x2022)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>
"@
# Header: draft notice on the left, report and complaint number on the right.
$parts['word/header1.xml'] = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr $w><w:p><w:pPr><w:pStyle w:val="Header"/><w:tabs><w:tab w:val="right" w:pos="9638"/></w:tabs></w:pPr><w:r><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr><w:t>{{DraftNotice}}</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">{{ReportTitle}}, complaint {{CaseId}}</w:t></w:r></w:p></w:hdr>
"@
$footerRun = '<w:rPr><w:color w:val="7F7F7F"/><w:sz w:val="16"/></w:rPr>'
$parts['word/footer1.xml'] = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr $w><w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r>$footerRun<w:t xml:space="preserve">Generated {{GeneratedAt}} by Complaint Dossier (run {{RunId}}, AI: {{Model}}). Page </w:t></w:r><w:r>$footerRun<w:fldChar w:fldCharType="begin"/></w:r><w:r>$footerRun<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r>$footerRun<w:fldChar w:fldCharType="separate"/></w:r><w:r>$footerRun<w:t>1</w:t></w:r><w:r>$footerRun<w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>
"@

# The complaint date placeholder is stored the way Word often stores edited text:
# split over two runs with a spell-check mark in between.
$splitDate = '<w:p><w:pPr><w:pStyle w:val="TableValue"/></w:pPr><w:r><w:t>{{Complaint</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>Date}}</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>'
$dataTable = Table @(
    (Row 'Complaint number' (P '{{CaseId}}' 'TableValue')),
    (Row 'Date received' $splitDate),
    (Row 'Product' (P '{{Product}}' 'TableValue')),
    (Row 'Product group' (P '{{ProductGroup}}' 'TableValue')),
    (Row 'Batch' (P '{{Batch}}' 'TableValue')),
    (Row 'Defect category' (P '{{Category}}' 'TableValue')),
    (Row 'Severity' (P '{{Severity}}' 'TableValue')),
    (Row 'Returned sample received' (P '{{SampleReceived}}' 'TableValue')),
    (Row 'Description from the complaint list' (P '{{Description}}' 'TableValue'))
)
$reviewTable = Table @(
    (Row 'Checked against the case file by' (P '' 'TableValue')),
    (Row 'Date' (P '' 'TableValue')),
    (Row 'Changes made during review' (P '' 'TableValue')),
    (Row 'Released for the complaint record (yes/no)' (P '' 'TableValue')),
    (Row 'Signature' (P '' 'TableValue'))
)
$sectPr = '<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId3"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1247" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="510" w:gutter="0"/></w:sectPr>'

$templates = [ordered]@{
    'short-report-template.docx' = @{
        Intro    = 'Short report. The sections "Basic complaint data" and 1 to 4 were drafted by AI from the complaint list and the case file. Check every statement against the case file before release.'
        Sections = @(
            @('Basic complaint data', 'BasicData'),
            @('1. Technical description and categorisation of the defect', 'DefectDescription'),
            @('2. Statistical and historical classification', 'History'),
            @('3. Assessment of the available information', 'Assessment'),
            @('4. Conclusion', 'Conclusion'))
    }
    'long-report-template.docx'  = @{
        Intro    = 'Long report. The sections "Basic complaint data" and 1 to 9 were drafted by AI from the complaint list and the case file. Check every statement against the case file before release.'
        Sections = @(
            @('Basic complaint data', 'BasicData'),
            @('1. Technical description and categorisation of the defect', 'DefectDescription'),
            @('2. Previous complaints for the same batch', 'SameBatch'),
            @('3. Previous complaints for the same product', 'SameProduct'),
            @('4. Examination of the returned complaint sample', 'SampleExamination'),
            @('5. Result of the batch record review', 'BatchRecordReview'),
            @('6. Result of the review of analyses and retained samples', 'AnalysesRetainedSamples'),
            @('7. Assessment of a possible manufacturing defect', 'ManufacturingDefect'),
            @('8. Status of existing or initiated corrective actions', 'CorrectiveActions'),
            @('9. Risk assessment and overall conclusion', 'RiskAndConclusion'))
    }
}

foreach ($name in $templates.Keys) {
    $t = $templates[$name]
    $body = New-Object System.Text.StringBuilder
    [void]$body.Append((P 'Complaint report {{CaseId}}' 'Title'))
    [void]$body.Append((P $t.Intro 'Note'))
    [void]$body.Append((P 'Complaint data (from the complaint list)' 'Heading1'))
    [void]$body.Append($dataTable)
    foreach ($section in $t.Sections) {
        [void]$body.Append((P $section[0] 'Heading1'))
        [void]$body.Append((P ('{{' + $section[1] + '}}')))
    }
    [void]$body.Append((P 'Review and release' 'Heading1'))
    [void]$body.Append((P 'Filled in by the reviewer. Replace this block with the release step your quality system requires.' 'Note'))
    [void]$body.Append($reviewTable)
    [void]$body.Append((P 'For the reviewer: automated checks' 'Heading1'))
    [void]$body.Append((P '{{CheckSummary}}'))
    [void]$body.Append((P 'Similar complaints computed by the tool' 'Heading2'))
    [void]$body.Append((P '{{SimilarComplaints}}'))
    [void]$body.Append((P 'Case documents used' 'Heading2'))
    [void]$body.Append((P '{{CaseDocuments}}'))
    $parts['word/document.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<w:document $w><w:body>$($body.ToString())$sectPr</w:body></w:document>"
    $path = Join-Path $OutputFolder $name
    New-ZipPackage -Path $path -Parts $parts
    Write-Host "Example template written: $path"
}
