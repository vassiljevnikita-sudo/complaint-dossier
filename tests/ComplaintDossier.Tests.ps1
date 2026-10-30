# Pester 5 tests for the core module, the Word template filling and the two commands.
# Run: Invoke-Pester ./tests

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    Import-Module (Join-Path (Join-Path $script:Root 'src') 'ComplaintDossier.psm1') -Force
    $script:ExampleConfig = Join-Path $script:Root 'config.example.json'
    $script:AllCases = @('SYN-2026-001', 'SYN-2026-002', 'SYN-2026-003', 'SYN-2026-004', 'SYN-2026-005', 'SYN-2026-006', 'SYN-2026-007', 'SYN-2026-008', 'SYN-2026-009', 'SYN-2026-010')
    $script:W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

    function New-TestConfig {
        # Writes a copy of config.example.json with changes applied by $Edit.
        # The output folder stays inside TestDrive.
        param([scriptblock]$Edit = {})
        $config = Get-Content -Raw $script:ExampleConfig | ConvertFrom-Json
        $config.Output.Path = Join-Path $TestDrive ('out-' + [guid]::NewGuid().ToString('N'))
        foreach ($key in 'Workbook', 'CaseFolders', 'Demo') {
            $value = $config.$key.(@{ Workbook = 'Path'; CaseFolders = 'Path'; Demo = 'ResponsePath' }[$key])
            $config.$key.(@{ Workbook = 'Path'; CaseFolders = 'Path'; Demo = 'ResponsePath' }[$key]) = Join-Path $script:Root $value
        }
        foreach ($type in 'Short', 'Long') {
            $config.Reports.$type.TemplatePath = Join-Path $script:Root $config.Reports.$type.TemplatePath
            $config.Reports.$type.PromptPath = Join-Path $script:Root $config.Reports.$type.PromptPath
        }
        & $Edit $config
        $path = Join-Path $TestDrive ("config-{0}.json" -f [guid]::NewGuid().ToString('N'))
        $config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $path -Encoding utf8
        $path
    }
    function New-TestCase([string]$Id, [string]$Date, [string]$Category, [string]$Group = 'G', [string]$Batch = '') {
        [pscustomobject]@{ CaseId = $Id; ReceivedDate = [datetime]$Date; Category = $Category; Product = ''; ProductGroup = $Group; Batch = $Batch; Severity = ''; Status = 'Open'; SampleReceived = ''; Description = ''; ReportType = '' }
    }
    function Invoke-Demo([string]$Config, [string[]]$Ids, [hashtable]$Extra = @{}) {
        $arguments = @{ Demo = $true; ConfigPath = $Config; PassThru = $true }
        if ($Ids) { $arguments.CaseId = $Ids }
        foreach ($k in $Extra.Keys) { $arguments[$k] = $Extra[$k] }
        & (Join-Path $script:Root 'Invoke-ComplaintDossier.ps1') @arguments 6> $null
    }
    function New-TestWorkbook([string]$Path, [string[]]$Header, [object[]]$Rows, [string]$Sheet = 'List') {
        # A small .xlsx with inline strings; [datetime] values become date cells.
        $cells = { param($r, $values)
            $out = for ($i = 0; $i -lt $values.Count; $i++) {
                $ref = ConvertTo-CellReference $i $r
                $v = $values[$i]
                if ($null -eq $v) { continue }
                if ($v -is [datetime]) { "<c r=`"$ref`" s=`"1`"><v>$($v.ToOADate())</v></c>" }
                else { "<c r=`"$ref`" t=`"inlineStr`"><is><t>$([Security.SecurityElement]::Escape([string]$v))</t></is></c>" }
            }
            "<row r=`"$r`">" + ($out -join '') + '</row>'
        }
        $rowsXml = (& $cells 1 $Header)
        $n = 2
        foreach ($row in $Rows) { $rowsXml += (& $cells $n $row); $n++ }
        $ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        New-ZipPackage -Path $Path -Parts ([ordered]@{
                '[Content_Types].xml'        = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
                '_rels/.rels'                = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
                'xl/workbook.xml'            = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><workbook $ns><sheets><sheet name=`"$Sheet`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
                'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
                'xl/styles.xml'              = '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>'
                'xl/worksheets/sheet1.xml'   = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><worksheet $ns><sheetData>$rowsXml</sheetData></worksheet>"
            })
        $Path
    }
    function New-TestDocx([string]$Path, [string]$Body, [string]$Header = '', [string]$Styles = '', [switch]$Template) {
        # A minimal Word document. $Body is the XML inside w:body.
        $main = 'document'
        if ($Template) { $main = 'template' }
        $parts = [ordered]@{
            '[Content_Types].xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><Types xmlns=`"http://schemas.openxmlformats.org/package/2006/content-types`"><Default Extension=`"rels`" ContentType=`"application/vnd.openxmlformats-package.relationships+xml`"/><Default Extension=`"xml`" ContentType=`"application/xml`"/><Override PartName=`"/word/document.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.wordprocessingml.$main.main+xml`"/><Override PartName=`"/docProps/core.xml`" ContentType=`"application/vnd.openxmlformats-package.core-properties+xml`"/></Types>"
            '_rels/.rels'         = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
            'docProps/core.xml'   = '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>t</dc:title></cp:coreProperties>'
            'word/document.xml'   = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:document $script:W><w:body>$Body<w:sectPr/></w:body></w:document>"
        }
        if ($Header) { $parts['word/header1.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:hdr $script:W>$Header</w:hdr>" }
        if ($Styles) { $parts['word/styles.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:styles $script:W>$Styles</w:styles>" }
        New-ZipPackage -Path $Path -Parts $parts
        $Path
    }
    function Get-DocxPart([string]$Path, [string]$Name) {
        $zip = [IO.Compression.ZipFile]::OpenRead($Path)
        try { $reader = New-Object IO.StreamReader($zip.GetEntry($Name).Open()); try { $reader.ReadToEnd() } finally { $reader.Dispose() } } finally { $zip.Dispose() }
    }
}

Describe 'Configuration' {
    It 'loads config.example.json in demo mode and resolves paths against the project root' {
        $config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root -Mode Demo
        $config.Paths.Workbook | Should -Be ([IO.Path]::GetFullPath((Join-Path $script:Root 'samples/complaint-list.xlsx')))
        $config.Paths.Templates.Long | Should -Be ([IO.Path]::GetFullPath((Join-Path $script:Root 'templates/long-report-template.docx')))
        $config.Paths.Reports | Should -Be (Join-Path ([IO.Path]::GetFullPath((Join-Path $script:Root 'out'))) 'reports')
    }
    It 'refuses live mode when no model is configured' {
        { Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root -Mode Live } | Should -Throw '*Copilot.PreferredModel*never picks a model*'
    }
    It 'points to Setup when the configuration file is missing' {
        { Import-DossierConfig -Path (Join-Path $TestDrive 'nope.json') -Root $script:Root } | Should -Throw '*Setup.cmd*'
    }
    It 'names a missing Word template' {
        $path = New-TestConfig { param($c) $c.Reports.Long.TemplatePath = 'templates/missing.docx' }
        { Import-DossierConfig -Path $path -Root $script:Root } | Should -Throw '*missing.docx*'
    }
    It 'refuses a .doc template' {
        $doc = Join-Path $TestDrive 'old.doc'
        Set-Content -LiteralPath $doc -Value 'x'
        $path = New-TestConfig { param($c) $c.Reports.Short.TemplatePath = $doc }
        { Import-DossierConfig -Path $path -Root $script:Root } | Should -Throw '*not a .docx*'
    }
    It 'refuses equal markers and placeholder names the tool fills itself' {
        { Import-DossierConfig -Path (New-TestConfig { param($c) $c.Reports.Long.Marker = 'short.txt' }) -Root $script:Root } | Should -Throw '*different file names*'
        { Import-DossierConfig -Path (New-TestConfig { param($c) $c.Template.Fields | Add-Member -NotePropertyName CheckSummary -NotePropertyValue 'Batch' }) -Root $script:Root } | Should -Throw '*{{CheckSummary}}*filled by the tool*'
    }
    It 'refuses a period that is not a date' {
        { Import-DossierConfig -Path (New-TestConfig { param($c) $c.Selection.From = '01.01.2026' }) -Root $script:Root } | Should -Throw '*Selection.From*yyyy-MM-dd*'
    }
    It 'writes the same JSON bytes on every PowerShell version' {
        ConvertTo-DossierJson ([ordered]@{ a = 1; b = 'x"y'; c = @(); d = $null; e = $true; f = @('u', 2.5); g = [ordered]@{} }) |
            Should -Be "{`n  `"a`": 1,`n  `"b`": `"x\`"y`",`n  `"c`": [],`n  `"d`": null,`n  `"e`": true,`n  `"f`": [`n    `"u`",`n    2.5`n  ],`n  `"g`": {}`n}"
    }
}

Describe 'Reading the complaint list' {
    BeforeAll { $script:Config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root }

    It 'reads all 10 synthetic complaints without Excel, through the column names in Workbook.Columns' {
        $cases = @(Import-DossierCases -Config $script:Config)
        $cases.Count | Should -Be 10
        $cases[0].CaseId | Should -Be 'SYN-2026-001'
        $cases[0].ReceivedDate | Should -Be ([datetime]'2026-01-12')
        $cases[0].ReportType | Should -Be 'short'
        $cases[6].Severity | Should -Be 'Critical'
        $cases[6].Product | Should -Be 'SYN Cream 30 g'
        $cases[6].Batch | Should -Be 'SYN-LOT-26D002'
    }
    It 'returns date cells as dates' {
        $cases = @(Import-DossierCases -Config $script:Config)
        $cases[0].Row.'Date received' | Should -BeOfType [datetime]
    }
    It 'uses only batch numbers in the synthetic SYN-LOT- format' {
        @(Import-DossierCases -Config $script:Config | Where-Object { $_.Batch -notmatch '^SYN-LOT-\d\d[A-D]\d{3}$' }).Count | Should -Be 0
    }
    It 'keeps the synthetic notice in the first row of the sheet' {
        $rows = @(Read-XlsxRows -Path $script:Config.Paths.Workbook -Sheet 'Complaints')
        $rows[0].Values[0] | Should -Match '^SYNTHETIC TEST DATA'
    }
    It 'names the column when the mapping does not match the workbook' {
        $config = Import-DossierConfig -Path (New-TestConfig { param($c) $c.Workbook.Columns.Batch = 'Lot' }) -Root $script:Root
        { Import-DossierCases -Config $config } | Should -Throw '*"Lot"*Workbook.Columns.Batch*'
    }
    It 'names the sheet when it does not exist' {
        $config = Import-DossierConfig -Path (New-TestConfig { param($c) $c.Workbook.Sheet = 'Nope' }) -Root $script:Root
        { Import-DossierCases -Config $config } | Should -Throw '*"Nope"*Complaints, About*'
    }
    It 'ignores rows without a complaint number or date and lists them' {
        $xlsx = New-TestWorkbook (Join-Path $TestDrive 'gaps.xlsx') @('No', 'Date', 'Category', 'Text') @(
            , @('A-1', [datetime]'2026-02-01', 'Leak', 'one')
            , @($null, [datetime]'2026-02-02', 'Leak', 'no number')
            , @('A-3', 'sometime', 'Leak', 'bad date')
            , @('A-4', '05.03.2026', 'Leak', 'text date'))
        $config = Import-DossierConfig -Path (New-TestConfig { param($c) $c.Workbook.Path = $xlsx; $c.Workbook.Sheet = 'List'; $c.Workbook.HeaderRow = 1
                $c.Workbook.Columns = [pscustomobject]@{ CaseId = 'No'; ReceivedDate = 'Date'; Category = 'Category'; Description = 'Text' } }) -Root $script:Root
        $cases = @(Import-DossierCases -Config $config)
        @($cases | ForEach-Object CaseId) -join ',' | Should -Be 'A-1,A-4'
        $cases[1].ReceivedDate | Should -Be ([datetime]'2026-03-05')
        $warnings = & (Get-Module ComplaintDossier) { $script:ImportWarnings }
        @($warnings).Count | Should -Be 2
        $warnings[1] | Should -Match 'A-3.*sometime'
    }
    It 'refuses duplicate complaint numbers' {
        $xlsx = New-TestWorkbook (Join-Path $TestDrive 'dup.xlsx') @('No', 'Date') @(
            , @('A-1', [datetime]'2026-02-01')
            , @('A-1', [datetime]'2026-02-02'))
        $config = Import-DossierConfig -Path (New-TestConfig { param($c) $c.Workbook.Path = $xlsx; $c.Workbook.Sheet = 'List'; $c.Workbook.HeaderRow = 1
                $c.Workbook.Columns = [pscustomobject]@{ CaseId = 'No'; ReceivedDate = 'Date' } }) -Root $script:Root
        { Import-DossierCases -Config $config } | Should -Throw '*more than once*A-1*'
    }
    It 'finds the header row below a title row' {
        Find-XlsxHeaderRow @(Read-XlsxRows -Path $script:Config.Paths.Workbook -Sheet 'Complaints') | Should -Be 2
    }
}

Describe 'Selection' {
    BeforeAll {
        $script:Config = Import-DossierConfig -Path (New-TestConfig) -Root $script:Root
        $script:Cases = @(Import-DossierCases -Config $script:Config)
    }
    It 'selects every complaint of the period, with the report type from the marker file or the list column' {
        $selection = @(Select-DossierCases -Config $script:Config -Cases $script:Cases)
        @($selection | Where-Object State -eq 'todo').Count | Should -Be 10
        ($selection | Where-Object { $_.Case.CaseId -eq 'SYN-2026-001' }).Type | Should -Be 'Short'
        ($selection | Where-Object { $_.Case.CaseId -eq 'SYN-2026-001' }).Reason | Should -Be 'column "Report type"'
        ($selection | Where-Object { $_.Case.CaseId -eq 'SYN-2026-007' }).Type | Should -Be 'Long'
        ($selection | Where-Object { $_.Case.CaseId -eq 'SYN-2026-007' }).Reason | Should -Be 'marker file long.txt'
    }
    It 'narrows to the period and the statuses' {
        @(Select-DossierCases -Config $script:Config -Cases $script:Cases -From '2026-04-01' -To '2026-04-30' | ForEach-Object { $_.Case.CaseId }) -join ',' | Should -Be 'SYN-2026-006,SYN-2026-007'
        $open = Import-DossierConfig -Path (New-TestConfig { param($c) $c.Selection.Statuses = @('open') }) -Root $script:Root
        @(Select-DossierCases -Config $open -Cases $script:Cases).Count | Should -Be 6
    }
    It 'selects exactly the given complaints and rejects unknown numbers' {
        @(Select-DossierCases -Config $script:Config -Cases $script:Cases -CaseId 'SYN-2026-003').Count | Should -Be 1
        { Select-DossierCases -Config $script:Config -Cases $script:Cases -CaseId 'SYN-2099-001' } | Should -Throw '*SYN-2099-001*'
    }
    It 'skips a complaint whose folder holds both markers, or whose report type value is unknown' {
        $root = Join-Path $TestDrive 'cases'
        New-Item -ItemType Directory -Force -Path (Join-Path $root 'SYN-2026-005') | Out-Null
        'x' | Set-Content (Join-Path $root 'SYN-2026-005/short.txt')
        'x' | Set-Content (Join-Path $root 'SYN-2026-005/long.txt')
        $config = Import-DossierConfig -Path (New-TestConfig { param($c) $c.CaseFolders.Path = $root }) -Root $script:Root
        $cases = @($script:Cases | ForEach-Object { $_.PSObject.Copy() })
        ($cases | Where-Object CaseId -eq 'SYN-2026-001').ReportType = 'medium'
        $selection = @(Select-DossierCases -Config $config -Cases $cases)
        ($selection | Where-Object { $_.Case.CaseId -eq 'SYN-2026-005' }).Reason | Should -Match 'both short.txt and long.txt'
        ($selection | Where-Object { $_.Case.CaseId -eq 'SYN-2026-001' }).Reason | Should -Match 'unknown report type "medium"'
        ($selection | Where-Object { $_.Case.CaseId -eq 'SYN-2026-006' }).State | Should -Be 'skip'
    }
    It 'uses Reports.DefaultType when neither a marker nor the column decides' {
        $config = Import-DossierConfig -Path (New-TestConfig { param($c) $c.CaseFolders.Path = ''; $c.Workbook.Columns.ReportType = ''; $c.Reports.DefaultType = 'Long' }) -Root $script:Root
        $cases = @(Import-DossierCases -Config $config)
        @(Select-DossierCases -Config $config -Cases $cases | Where-Object Type -eq 'Long').Count | Should -Be 10
    }
}

Describe 'Similar complaints' {
    It 'counts same batch over the whole list and same product group and category in the lookback window' {
        $cases = @(
            (New-TestCase 'A-1' '2025-05-01' 'Leak' 'G' 'L1'),
            (New-TestCase 'A-2' '2025-07-01' 'Leak' 'G' 'L2'),
            (New-TestCase 'A-3' '2026-05-20' 'Leak' 'G' 'L3'),
            (New-TestCase 'A-4' '2026-06-01' 'Leak' 'G' 'L1'),
            (New-TestCase 'A-5' '2026-06-10' 'Leak' 'G' 'L4'),
            (New-TestCase 'A-6' '2026-06-01' 'Print' 'G' 'L4'))
        $similar = Get-SimilarComplaints -Case $cases[3] -Cases $cases -LookbackMonths 12 -RecentDays 30
        @($similar.SameBatch | ForEach-Object CaseId) -join ',' | Should -Be 'A-1'
        @($similar.SameProductGroupCategory | ForEach-Object CaseId) -join ',' | Should -Be 'A-2,A-3'
        $similar.RecentCount | Should -Be 1
        $similar.LookbackFrom | Should -Be '2025-06-02'
        $similar.FrequencyLevel | Should -Be 'occasional'
    }
    It 'gives the sample data the expected facts' {
        $config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root
        $cases = @(Import-DossierCases -Config $config)
        $case = $cases | Where-Object CaseId -eq 'SYN-2026-010'
        $sentences = Get-SimilarComplaintSentences $case (Get-SimilarComplaints -Case $case -Cases $cases)
        $sentences[0] | Should -Be 'Same batch: 1 other case in the complaint list has batch SYN-LOT-26A019: SYN-2026-009.'
        $sentences[1] | Should -Match '^Same product group and category \(Tablets, Packaging damage\): 2 other cases received from 2025-06-03 to 2026-06-02.*none of them in the last 30 days\.$'
    }
}

Describe 'Per-case prompt' {
    BeforeAll {
        $script:Config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root
        $script:Cases = @(Import-DossierCases -Config $script:Config)
        function Get-Prompt([string]$Id, [string]$Type) {
            $case = $script:Cases | Where-Object CaseId -eq $Id
            $similar = Get-SimilarComplaints -Case $case -Cases $script:Cases
            New-CasePrompt -Template (Read-Utf8File $script:Config.Paths.Prompts.$Type) -Case $case -Similar $similar `
                -Documents @(Get-CaseDocuments -Config $script:Config -CaseId $Id) -RequiredStatements @(Get-RequiredStatements -Config $script:Config -Case $case -Type $Type -Similar $similar) `
                -ProcessFacts @($script:Config.Settings.Reports.$Type.ProcessFacts)
        }
    }
    It 'holds only this case, its similar complaints and its own documents' {
        $prompt = Get-Prompt 'SYN-2026-010' 'Long'
        $prompt | Should -Not -Match '\{\{[A-Z_]+\}\}'
        $prompt | Should -Match 'Write the complete complaint report for case SYN-2026-010 straight away'
        $prompt | Should -Match '(?m)^- Retained samples did not need to be checked\.$'
        $prompt | Should -Match '\| Batch \| SYN-LOT-26A019 \|'
        $prompt | Should -Match '--- Document sample-check\.txt ---'
        $prompt | Should -Match '\| SYN-2026-009 \| 2026-05-20 \| Label or print \|'
        $prompt | Should -Not -Match 'SYN-2026-007'
        $prompt | Should -Not -Match 'long\.txt'
    }
    It 'masks e-mail addresses and phone numbers by default' {
        $prompt = Get-Prompt 'SYN-2026-010' 'Long'
        $prompt | Should -Not -Match 'example\.org'
        $prompt | Should -Not -Match '5550 1234'
        $prompt | Should -Match '\[phone removed\]'
    }
    It 'demands the no-sample sentence in a long report when no sample was received' {
        Get-Prompt 'SYN-2026-007' 'Long' | Should -Match '(?m)^- No complaint sample was available for the investigation\.$'
        Get-Prompt 'SYN-2026-010' 'Long' | Should -Match 'section 4\):\nNone for this case\.'
    }
    It 'demands the frequency sentence with the level the tool computed in a short report' {
        Get-Prompt 'SYN-2026-006' 'Short' | Should -Match '(?m)^- The reported defect is classified as occasional\.$'
        Get-Prompt 'SYN-2026-005' 'Short' | Should -Match '(?m)^- The reported defect is classified as rare\.$'
        Get-Prompt 'SYN-2026-006' 'Short' | Should -Match '(?m)^- The batch record was reviewed and showed no findings\.$'
    }
    It 'leaves the upload-only CSV out of the prompt text' {
        Get-Prompt 'SYN-2026-006' 'Short' | Should -Not -Match 'shipment-log|SHP-0408'
    }
    It 'rejects a template with an unknown placeholder' {
        $case = $script:Cases[0]
        { New-CasePrompt -Template 'x {{NOPE}}' -Case $case -Similar (Get-SimilarComplaints -Case $case -Cases $script:Cases) } | Should -Throw '*{{NOPE}}*'
    }
}

Describe 'Uploads per case' {
    It 'takes the upload files from the case folder, without markers and text documents' {
        $config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root
        @(Get-UploadFiles -Config $config -CaseId 'SYN-2026-006' | ForEach-Object Name) -join ',' | Should -Be 'shipment-log.csv'
        @(Get-UploadFiles -Config $config -CaseId 'SYN-2026-007').Count | Should -Be 0
    }
    It 'refuses more files than Copilot.Upload.MaxFiles' {
        $config = Import-DossierConfig -Path (New-TestConfig { param($c) $c.Copilot.Upload.Extensions = @('.txt', '.csv'); $c.Copilot.Upload.MaxFiles = 1 }) -Root $script:Root
        { Get-UploadFiles -Config $config -CaseId 'SYN-2026-006' } | Should -Throw '*MaxFiles allows 1*'
    }
}

Describe 'Response checks' {
    BeforeAll {
        $script:Sections = @('Complaint', 'Conclusion')
        $script:Known = @('SYN-1', 'SYN-2')
    }
    It 'passes an answer with all sections, its own case ID and known case IDs' {
        $checks = Test-ReportResponse -Response "## Complaint`nSee [SYN-1] and [SYN-2].`n## Conclusion`nOpen [1]." -Prompt 'p' -CaseId 'SYN-1' -RequiredSections $script:Sections -KnownCaseIds $script:Known
        @($checks | Where-Object { -not $_.Passed }).Count | Should -Be 0
    }
    It 'fails on a missing section, a missing own case ID and an invented case ID' {
        $checks = Test-ReportResponse -Response "## Complaint`nSee [SYN-9]." -Prompt 'p' -CaseId 'SYN-1' -RequiredSections $script:Sections -KnownCaseIds $script:Known
        ($checks | Where-Object Name -eq 'Required sections present').Detail | Should -Be 'missing: Conclusion'
        ($checks | Where-Object Name -eq 'Report names its own case ID').Passed | Should -BeFalse
        ($checks | Where-Object Name -eq 'Cited case IDs exist').Detail | Should -Be 'not in the complaint list: SYN-9'
    }
    It 'accepts the own case ID written without brackets, but not as part of another ID' {
        $plain = Test-ReportResponse -Response "## Complaint`nCase ID: SYN-1" -Prompt 'p' -CaseId 'SYN-1' -RequiredSections @() -KnownCaseIds $script:Known
        ($plain | Where-Object Name -eq 'Report names its own case ID').Detail | Should -Be 'SYN-1 named'
        $other = Test-ReportResponse -Response "## Complaint`nCase ID: SYN-10" -Prompt 'p' -CaseId 'SYN-1' -RequiredSections @() -KnownCaseIds $script:Known
        ($other | Where-Object Name -eq 'Report names its own case ID').Passed | Should -BeFalse
    }
    It 'fails when the answer is the prompt or contains its data headings' {
        $checks = Test-ReportResponse -Response 'same text' -Prompt 'same text' -CaseId 'SYN-1' -RequiredSections @() -KnownCaseIds @()
        ($checks | Where-Object Name -eq 'Response is not a copy of the prompt').Passed | Should -BeFalse
        $checks = Test-ReportResponse -Response "## Complaint [SYN-1]`n### Case documents`nNo case documents." -Prompt 'other' -CaseId 'SYN-1' -RequiredSections @() -KnownCaseIds $script:Known
        ($checks | Where-Object Name -eq 'Response is not a copy of the prompt').Detail | Should -Match 'contains prompt heading: ### Case documents'
    }
    It 'turns plain-text section names, as the Copilot copy button returns them, into headings' {
        $plain = "Complaint`r`n`r`nCase [SYN-1] was received.`r`n`r`nConclusion:`r`nOpen.`r`nComplaint text that is not a heading."
        $text = ConvertTo-SectionHeadings $plain $script:Sections
        $text | Should -Match '(?m)^## Complaint$'
        $text | Should -Match '(?m)^## Conclusion$'
        $text | Should -Match '(?m)^Complaint text that is not a heading\.$'
        $checks = Test-ReportResponse -Response $text -Prompt 'p' -CaseId 'SYN-1' -RequiredSections $script:Sections -KnownCaseIds $script:Known
        @($checks | Where-Object { -not $_.Passed }).Count | Should -Be 0
    }
    It 'recognises numbered headings that arrive as list items without their number' {
        $sections = @('Basic complaint data', '1. Technical description', '2. Conclusion')
        $text = ConvertTo-SectionHeadings "Basic complaint data`n- Technical description`nText.`n2. Conclusion:`n## Conclusion" $sections
        $text | Should -Be "## Basic complaint data`n## 1. Technical description`nText.`n## 2. Conclusion`n## 2. Conclusion"
    }
    It 'fails when the answer does not end with a classification' {
        $classes = @('The complaint is classified as justified.', 'The complaint is classified as not justified.')
        $ok = Test-ReportResponse -Response "## Complaint [SYN-1]`nThe complaint is classified as not justified.`n" -Prompt 'p' -CaseId 'SYN-1' -RequiredSections @() -KnownCaseIds $script:Known -ClassificationStatements $classes
        ($ok | Where-Object Name -eq 'Conclusion ends with a classification').Passed | Should -BeTrue
        $bad = Test-ReportResponse -Response "## Complaint [SYN-1]`nThe complaint is classified as justified. More text." -Prompt 'p' -CaseId 'SYN-1' -RequiredSections @() -KnownCaseIds $script:Known -ClassificationStatements $classes
        ($bad | Where-Object Name -eq 'Conclusion ends with a classification').Passed | Should -BeFalse
        $none = Test-ReportResponse -Response '## Complaint [SYN-1]' -Prompt 'p' -CaseId 'SYN-1' -RequiredSections @() -KnownCaseIds $script:Known
        @($none | Where-Object Name -eq 'Conclusion ends with a classification').Count | Should -Be 0
    }
    It 'fails when a required statement is missing' {
        $checks = Test-ReportResponse -Response '## Complaint [SYN-1]' -Prompt 'p' -CaseId 'SYN-1' -RequiredSections @() -KnownCaseIds $script:Known -RequiredStatements @('No returned sample was available for examination.')
        ($checks | Where-Object Name -eq 'Required statements present').Detail | Should -Be 'missing: No returned sample was available for examination.'
    }
}


Describe 'Word template filling' {
    BeforeAll {
        $script:Styles = '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="berschrift2"><w:name w:val="heading 2"/></w:style><w:style w:type="paragraph" w:styleId="Aufzhlungszeichen"><w:name w:val="List Bullet"/></w:style>'
        function Invoke-Fill([string]$Template, [hashtable]$Values = @{}, [hashtable]$Blocks = @{}, [string]$Notice = '', [string]$Appendix = '') {
            $out = Join-Path $TestDrive ('out-' + [guid]::NewGuid().ToString('N') + '.docx')
            Invoke-WordTemplate -TemplatePath $Template -OutputPath $out -Values $Values -Blocks $Blocks -Styles @{ Heading = 'heading 2'; Bullet = 'List Bullet' } -Notice $Notice -Appendix $Appendix -ContentStatus 'Draft'
            $out
        }
    }
    It 'fills a placeholder that Word split over several runs, with the format of its first character' {
        $body = '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Batch: {{Ba</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:i/></w:rPr><w:t>tch</w:t></w:r><w:r><w:t>}} end</w:t></w:r></w:p>'
        $out = Invoke-Fill (New-TestDocx (Join-Path $TestDrive 'split.docx') $body) @{ Batch = 'L-7' }
        Get-DocxText $out | Should -Be 'Batch: L-7 end'
        Get-DocxPart $out 'word/document.xml' | Should -Match '<w:r><w:rPr><w:b /></w:rPr><w:t xml:space="preserve">Batch: L-7</w:t></w:r>'
    }
    It 'matches names without regard to case and spaces inside the braces' {
        $out = Invoke-Fill (New-TestDocx (Join-Path $TestDrive 'case.docx') '<w:p><w:r><w:t>{{ batch }} / {{BATCH}}</w:t></w:r></w:p>') @{ Batch = 'L-7' }
        Get-DocxText $out | Should -Be 'L-7 / L-7'
    }
    It 'turns line breaks of a text value into Word line breaks' {
        $out = Invoke-Fill (New-TestDocx (Join-Path $TestDrive 'br.docx') '<w:p><w:r><w:t>{{Description}}</w:t></w:r></w:p>') @{ Description = "one`ntwo" }
        Get-DocxPart $out 'word/document.xml' | Should -Match '<w:t xml:space="preserve">one</w:t><w:br /><w:t xml:space="preserve">two</w:t>'
    }
    It 'replaces a block placeholder by paragraphs in the template format, with the template list and heading styles' {
        $body = '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>{{Assessment}}</w:t></w:r></w:p>'
        $out = Invoke-Fill (New-TestDocx (Join-Path $TestDrive 'block.docx') $body -Styles $script:Styles) @{} @{ Assessment = "First **bold** part.`n- item one`n## Sub heading`nLast." }
        $xml = Get-DocxPart $out 'word/document.xml'
        Get-DocxText $out | Should -Be "First bold part.`nitem one`nSub heading`nLast."
        ([regex]::Matches($xml, '<w:jc w:val="both" />')).Count | Should -Be 4
        $xml | Should -Match '<w:pStyle w:val="Aufzhlungszeichen" />'
        $xml | Should -Match '<w:pStyle w:val="berschrift2" /><w:jc w:val="both" /></w:pPr><w:r><w:t xml:space="preserve">Sub heading'
        $xml | Should -Match '<w:rPr><w:b /><w:sz w:val="22" /></w:rPr><w:t xml:space="preserve">bold</w:t>'
    }
    It 'writes bullets and bold headings itself when the template lacks the styles' {
        $out = Invoke-Fill (New-TestDocx (Join-Path $TestDrive 'nostyle.docx') '<w:p><w:r><w:t>{{X}}</w:t></w:r></w:p>') @{} @{ X = "- a`n## H" }
        Get-DocxText $out | Should -Be ([string][char]0x2022 + " a`nH")
        Get-DocxPart $out 'word/document.xml' | Should -Match '<w:r><w:rPr><w:b /></w:rPr><w:t xml:space="preserve">H</w:t>'
    }
    It 'inserts a block placeholder that shares its paragraph as text with line breaks' {
        $out = Invoke-Fill (New-TestDocx (Join-Path $TestDrive 'inline.docx') '<w:p><w:r><w:t>Result: {{X}}</w:t></w:r></w:p>') @{} @{ X = "- a`n- b" }
        Get-DocxText $out | Should -Be ('Result: ' + [char]0x2022 + ' a' + [char]0x2022 + ' b')
    }
    It 'fills placeholders in headers, table cells and text boxes, and keeps a section break once' {
        $body = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>{{A}}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
        '<w:p><w:r><w:t>box: </w:t></w:r><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>{{A}}</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>' +
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="1"/></w:sectPr></w:pPr><w:r><w:t>{{X}}</w:t></w:r></w:p>'
        $out = Invoke-Fill (New-TestDocx (Join-Path $TestDrive 'places.docx') $body -Header '<w:p><w:r><w:t>{{A}}</w:t></w:r></w:p>') @{ A = 'v' } @{ X = "p1`np2" }
        Get-DocxText $out | Should -Be "v`nbox: v`np1`np2`nv"
        $xml = Get-DocxPart $out 'word/document.xml'
        ([regex]::Matches($xml, '<w:pgSz w:w="1" />')).Count | Should -Be 1
        $xml | Should -Match 'p2</w:t></w:r></w:p>$|<w:sectPr><w:pgSz w:w="1" /></w:sectPr></w:pPr><w:r><w:t xml:space="preserve">p2'
    }
    It 'refuses an unknown placeholder and leaves no file behind' {
        $template = New-TestDocx (Join-Path $TestDrive 'unknown.docx') '<w:p><w:r><w:t>{{Nope}}</w:t></w:r></w:p>'
        $out = Join-Path $TestDrive 'unknown-out.docx'
        { Invoke-WordTemplate -TemplatePath $template -OutputPath $out -Values @{} -Blocks @{} } | Should -Throw '*{{Nope}}*'
        $out | Should -Not -Exist
    }
    It 'adds the draft notice and the check results when the template has no place for them, and sets the document status' {
        $out = Invoke-Fill (New-TestDocx (Join-Path $TestDrive 'plain.docx') '<w:p><w:r><w:t>Body</w:t></w:r></w:p>') @{} @{} 'DRAFT notice' "## Checks`n- passed: x"
        Get-DocxText $out | Should -Be ("DRAFT notice`nBody`nChecks`n" + [char]0x2022 + ' passed: x')
        Get-DocxPart $out 'docProps/core.xml' | Should -Match '<cp:contentStatus>Draft</cp:contentStatus>'
    }
    It 'turns a .dotx template into a normal document' {
        $template = New-TestDocx (Join-Path $TestDrive 'tpl.dotx') '<w:p><w:r><w:t>{{A}}</w:t></w:r></w:p>' -Template
        $out = Invoke-Fill $template @{ A = 'x' }
        Get-DocxPart $out '[Content_Types].xml' | Should -Match 'wordprocessingml\.document\.main\+xml'
    }
    It 'lists the placeholders of a template and whether each stands alone in its paragraph' {
        $found = @(Get-WordPlaceholders (Join-Path $script:Root 'templates/short-report-template.docx'))
        ($found | Where-Object Name -eq 'ComplaintDate').Alone | Should -BeTrue
        ($found | Where-Object Name -eq 'Assessment').Alone | Should -BeTrue
        @($found | Where-Object { $_.Name -eq 'CaseId' -and $_.Part -eq 'header1' }).Count | Should -Be 1
    }
}

Describe 'Demo run: one draft report per complaint, in the Word template' {
    BeforeAll {
        $script:Config = New-TestConfig
        $script:Result = Invoke-Demo $script:Config
        $script:Out = $script:Result.OutputPath
        $script:Report007 = Join-Path $script:Out 'reports/SYN-2026-007-long-report-DRAFT.docx'
    }
    It 'generates a report for every complaint of the period, short or long as decided per complaint' {
        @($script:Result.Complaints | Where-Object state -eq 'generated').Count | Should -Be 10
        $files = @(Get-ChildItem (Join-Path $script:Out 'reports') -Filter *.docx | ForEach-Object Name | Sort-Object)
        $files.Count | Should -Be 10
        $files | Should -Contain 'SYN-2026-001-short-report-DRAFT.docx'
        $files | Should -Contain 'SYN-2026-010-long-report-DRAFT.docx'
        @($script:Result.Complaints | Where-Object failedChecks -gt 0).Count | Should -Be 0
    }
    It 'fills the template: draft notice in the header, list values, AI sections in their places, checks for the reviewer' {
        $text = Get-DocxText $script:Report007
        $text | Should -Match 'DRAFT generated by Complaint Dossier\. Not reviewed\.'
        $text | Should -Match '(?m)^22\.04\.2026$'
        $text | Should -Match '(?m)^SYN Cream 30 g$'
        $text | Should -Match '(?ms)^4\. Examination of the returned complaint sample\nNo complaint sample was available for the investigation\.'
        $text | Should -Match '(?ms)^9\. Risk assessment and overall conclusion\nA foreign fibre in a cream.*The complaint is classified as not justified\.\nReview and release'
        $text | Should -Match 'All 7 automated checks passed\.'
        $text | Should -Match 'Same batch: no other case in the complaint list has batch SYN-LOT-26D002\.'
        $text | Should -Match 'intake-note\.txt: text inserted into the prompt'
        $text | Should -Not -Match '\{\{'
        $text | Should -Not -Match 'SYNTHETIC SAMPLE ANSWER'
    }
    It 'marks the document status as draft' {
        Get-DocxPart $script:Report007 'docProps/core.xml' | Should -Match '<cp:contentStatus>Draft - to be reviewed</cp:contentStatus>'
    }
    It 'records SHA-256 hashes that match the files, without any review step' {
        $audit = Read-JsonFile (Join-Path $script:Out 'audit/SYN-2026-007/audit.json')
        $audit.status | Should -Be 'DRAFT'
        $audit.reportType | Should -Be 'long'
        foreach ($name in 'caseData', 'prompt', 'response', 'report') {
            Get-Sha256 (Join-Path $script:Out $audit.artifacts.$name.file) | Should -Be $audit.artifacts.$name.sha256 -Because $name
        }
        @($audit.inputs | ForEach-Object role) | Should -Contain 'word-template'
        ($audit.inputs | Where-Object role -eq 'word-template').path | Should -Be 'templates/long-report-template.docx'
        $audit.review.where | Should -Match 'in Word'
        @($audit.events | ForEach-Object event) -join ',' | Should -Be 'case-started,prompt-built,response-received,checks-run,report-written'
    }
    It 'keeps the case data of each complaint, including the template values' {
        $data = Read-JsonFile (Join-Path $script:Out 'audit/SYN-2026-001/case-data.json')
        $data.reportTypeSource | Should -Be 'column "Report type"'
        $data.templateFields.ComplaintDate | Should -Be '12.01.2026'
        $data.similarComplaints.FrequencyLevel | Should -Be 'rare'
    }
    It 'skips every complaint with a report on the next run' {
        $again = Invoke-Demo $script:Config
        @($again.Complaints | Where-Object state -eq 'generated').Count | Should -Be 0
        @($again.Complaints | Where-Object state -eq 'done').Count | Should -Be 10
    }
    It 'generates one report again with -Regenerate and keeps the earlier one under replaced/' {
        $again = Invoke-Demo $script:Config @('SYN-2026-006') @{ Regenerate = $true }
        @($again.Complaints | Where-Object state -eq 'generated').Count | Should -Be 1
        @(Get-ChildItem (Join-Path $script:Out 'replaced') -Recurse -Filter 'SYN-2026-006-short-report-DRAFT.docx').Count | Should -Be 1
        Join-Path $script:Out 'reports/SYN-2026-006-short-report-DRAFT.docx' | Should -Exist
    }
    It 'records a failed complaint, goes on with the others, and stops after three failures in a row' {
        $answers = Join-Path $TestDrive 'answers'
        New-Item -ItemType Directory -Force -Path $answers | Out-Null
        Get-ChildItem (Join-Path $script:Root 'samples/demo-responses') | Where-Object Name -notin 'SYN-2026-002.md', 'SYN-2026-005.md', 'SYN-2026-006.md', 'SYN-2026-007.md' | Copy-Item -Destination $answers
        $config = New-TestConfig { param($c) $c.Demo.ResponsePath = $answers }
        $result = Invoke-Demo $config
        $LASTEXITCODE | Should -Be 2
        @($result.Complaints | Where-Object state -eq 'generated' | ForEach-Object caseId) -join ',' | Should -Be 'SYN-2026-001,SYN-2026-003,SYN-2026-004'
        @($result.Complaints | Where-Object state -eq 'failed' | ForEach-Object caseId) -join ',' | Should -Be 'SYN-2026-002,SYN-2026-005,SYN-2026-006,SYN-2026-007'
        @($result.Complaints | Where-Object state -eq 'not-started').Count | Should -Be 3
        ($result.Complaints | Where-Object caseId -eq 'SYN-2026-002').detail | Should -Match 'No prewritten demo answer'
        Join-Path $result.OutputPath 'audit/SYN-2026-002/audit.json' | Should -Not -Exist
    }
}

Describe 'Review list' {
    BeforeAll {
        $script:Config = New-TestConfig
        $script:Result = Invoke-Demo $script:Config
        $script:List = $script:Result.ReviewList
    }
    It 'lists every generated report with its checks and empty reviewer columns' {
        $rows = @(Read-XlsxSheet -Path $script:List -Sheet 'Reports')
        $rows.Count | Should -Be 10
        $rows[6].'Complaint no.' | Should -Be 'SYN-2026-007'
        $rows[6].'Report file' | Should -Be 'SYN-2026-007-long-report-DRAFT.docx'
        $rows[6].'Automated checks' | Should -Be '7 of 7 passed'
        $rows[6].'Reviewed by' | Should -BeNullOrEmpty
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Released (yes/no)'
    }
    It 'keeps what the QA team entered when the list is written again' {
        $sheet = Get-DocxPart $script:List 'xl/worksheets/sheet1.xml'
        $sheet = $sheet.Replace('<row r="8">', '<row r="8"><c r="J8" t="inlineStr"><is><t>A. Reviewer</t></is></c>')
        $parts = [ordered]@{}
        $zip = [IO.Compression.ZipFile]::OpenRead($script:List)
        try { foreach ($e in $zip.Entries) { $r = New-Object IO.StreamReader($e.Open()); $parts[$e.FullName] = $r.ReadToEnd(); $r.Dispose() } } finally { $zip.Dispose() }
        $parts['xl/worksheets/sheet1.xml'] = $sheet
        New-ZipPackage -Path $script:List -Parts $parts
        Invoke-Demo $script:Config @('SYN-2026-006') @{ Regenerate = $true } | Out-Null
        $rows = @(Read-XlsxSheet -Path $script:List -Sheet 'Reports')
        ($rows | Where-Object 'Complaint no.' -eq 'SYN-2026-007').'Reviewed by' | Should -Be 'A. Reviewer'
    }
    It 'lists complaints without a report on a second sheet' {
        $config = New-TestConfig { param($c) $c.Workbook.Columns.ReportType = '' }
        $result = Invoke-Demo $config
        $rows = @(Read-XlsxSheet -Path $result.ReviewList -Sheet 'Not generated')
        @($rows | ForEach-Object 'Complaint no.') -join ',' | Should -Be 'SYN-2026-001,SYN-2026-002,SYN-2026-003,SYN-2026-004'
        $rows[0].Reason | Should -Match 'no report type'
    }
    It 'speaks German when Language is de' {
        $config = New-TestConfig { param($c) $c.Language = 'de' }
        $result = Invoke-Demo $config @('SYN-2026-007')
        $rows = @(Read-XlsxSheet -Path $result.ReviewList -Sheet 'Berichte')
        $rows[0].("Gepr$([char]0xFC)ft von") | Should -BeNullOrEmpty
        @($rows[0].PSObject.Properties.Name) | Should -Contain "Gepr$([char]0xFC)ft von"
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Freigegeben (ja/nein)'
        Get-DocxText (Join-Path $result.OutputPath 'reports/SYN-2026-007-long-report-DRAFT.docx') | Should -Match 'Alle 7 automatischen Pr.fungen bestanden\.'
        Set-DossierLanguage en
    }
}

Describe 'Setup' {
    It 'reports the example configuration as ready for demo mode' {
        $output = & pwsh -NoProfile -File (Join-Path $script:Root 'Setup.ps1') -Demo -NonInteractive 2>&1
        $LASTEXITCODE | Should -Be 0
        ($output -join "`n") | Should -Match 'Ready for demo mode'
    }
    It 'names an unknown placeholder, suggests the right one, and a section without a place' {
        $template = Join-Path $TestDrive 'typo.docx'
        $body = '<w:p><w:r><w:t>{{Bacth}}</w:t></w:r></w:p><w:p><w:r><w:t>{{BasicData}}</w:t></w:r></w:p>'
        New-TestDocx $template $body | Out-Null
        $config = Read-DossierConfig -Path (New-TestConfig { param($c) $c.Reports.Short.TemplatePath = $template }) -Root $script:Root
        $errors = @(Test-DossierSetup -Config $config | Where-Object Level -eq 'error')
        ($errors.Text -join "`n") | Should -Match 'Unknown placeholder \{\{Bacth\}\}.*Did you mean "\{\{Batch\}\}"'
        ($errors.Text -join "`n") | Should -Match 'The AI section "4\. Conclusion" has no place'
    }
    It 'names a column that is missing in the list and suggests the closest one' {
        $config = Read-DossierConfig -Path (New-TestConfig { param($c) $c.Template.Fields.Batch = 'Batsh' }) -Root $script:Root
        $errors = @(Test-DossierSetup -Config $config | Where-Object Level -eq 'error')
        $errors[0].Text | Should -Be 'The list has no column "Batsh" (used for {{Batch}}). Did you mean "Batch"?'
        $errors[0].Fix | Should -Match 'Template\.Fields\.Batch'
    }
    It 'asks for the columns of an unknown list and writes them into config.json' {
        $xlsx = New-TestWorkbook (Join-Path $TestDrive 'own-list.xlsx') @('Rek-Nr', 'Eingang', 'Fehlerart', 'Beschreibung', 'Artikel', 'Charge') @(
            , @('R-1', [datetime]'2026-03-01', 'Leak', 'Bottle leaking', 'Item A', 'C-1'))
        $configPath = Join-Path $TestDrive 'setup-config.json'
        $answers = @('en', $xlsx, '',
            '1', '2', '3', '4', '5', '', '6', '', '', '', '',
            '1', '2', '5', '', '6', '3', '', '', '4',
            '', '', '', '', '-', 'Quick Response')
        & (Join-Path $script:Root 'Setup.ps1') -ConfigPath $configPath -Answers $answers *> $null
        $saved = Get-Content -Raw $configPath | ConvertFrom-Json
        $saved.Workbook.Path | Should -Be $xlsx
        $saved.Workbook.Sheet | Should -Be 'List'
        $saved.Workbook.HeaderRow | Should -Be 1
        $saved.Workbook.Columns.CaseId | Should -Be 'Rek-Nr'
        $saved.Workbook.Columns.Batch | Should -Be 'Charge'
        $saved.Workbook.Columns.ProductGroup | Should -Be ''
        $saved.Template.Fields.ComplaintDate | Should -Be 'Eingang'
        $saved.Template.Fields.Description | Should -Be 'Beschreibung'
        $saved.CaseFolders.Path | Should -Be ''
        $saved.Copilot.PreferredModel | Should -Be 'Quick Response'
    }
}

Describe 'Committed sample in docs/sample' {
    BeforeAll {
        $script:Sample = Join-Path (Join-Path $script:Root 'docs') 'sample'
        $script:Fresh = (Invoke-Demo (New-TestConfig) @('SYN-2026-007', 'SYN-2026-006')).OutputPath
    }
    It 'holds the sample PDFs that other pages link to' {
        Join-Path $script:Sample 'complaint-dossier-sample-DRAFT.pdf' | Should -Exist
        Join-Path $script:Sample 'reports/SYN-2026-007-long-report-DRAFT.pdf' | Should -Exist
        Join-Path $script:Sample 'reports/SYN-2026-006-short-report-DRAFT.pdf' | Should -Exist
        Join-Path $script:Sample 'review-list.xlsx' | Should -Exist
    }
    It 'is a demo draft with all checks passed, matching the hashes in its audit.json (<_>)' -ForEach @('SYN-2026-007', 'SYN-2026-006') {
        $audit = Read-JsonFile (Join-Path $script:Sample "audit/$_/audit.json")
        $audit.mode | Should -Be 'demo'
        $audit.status | Should -Be 'DRAFT'
        @($audit.checks | Where-Object { -not $_.passed }).Count | Should -Be 0
        foreach ($name in 'caseData', 'prompt', 'response', 'report') {
            $artifact = $audit.artifacts.$name
            Get-Sha256 (Join-Path $script:Sample $artifact.file) | Should -Be $artifact.sha256 -Because $artifact.file
        }
    }
    It 'matches the inputs and the report the current code produces (<_>)' -ForEach @('SYN-2026-007', 'SYN-2026-006') {
        $sample = Read-JsonFile (Join-Path $script:Sample "audit/$_/audit.json")
        $fresh = Read-JsonFile (Join-Path $script:Fresh "audit/$_/audit.json")
        foreach ($item in $fresh.inputs | Where-Object role -ne 'config') {
            ($sample.inputs | Where-Object path -eq $item.path).sha256 | Should -Be $item.sha256 -Because $item.path
        }
        foreach ($name in 'caseData', 'prompt', 'response') {
            $sample.artifacts.$name.sha256 | Should -Be $fresh.artifacts.$name.sha256 -Because $name
        }
        # Run ID and generation time differ between runs; everything else must match.
        $normalize = {
            param($path, $runId)
            ((Get-DocxText $path).Replace($runId, '<run id>')) -replace '\d\d\.\d\d\.\d{4} \d\d:\d\d', '<time>'
        }
        $expected = & $normalize (Join-Path $script:Sample $sample.artifacts.report.file) $sample.runId
        & $normalize (Join-Path $script:Fresh $fresh.artifacts.report.file) $fresh.runId | Should -Be $expected
    }
}

Describe 'Generators' {
    It 'reproduce the committed complaint list, case folders and example templates' {
        $out = Join-Path $TestDrive 'samples'
        & (Join-Path $script:Root 'tools/New-SyntheticCases.ps1') -OutputPath $out 6> $null
        Get-Sha256 (Join-Path $out 'complaint-list.xlsx') | Should -Be (Get-Sha256 (Join-Path $script:Root 'samples/complaint-list.xlsx'))
        $committedRoot = Join-Path $script:Root 'samples/case-folders'
        $expected = @(Get-ChildItem -Recurse -File $committedRoot | ForEach-Object { $_.FullName.Substring($committedRoot.Length).Replace('\', '/') + '=' + (Get-Sha256 $_.FullName) } | Sort-Object)
        $generatedRoot = Join-Path $out 'case-folders'
        $actual = @(Get-ChildItem -Recurse -File $generatedRoot | ForEach-Object { $_.FullName.Substring($generatedRoot.Length).Replace('\', '/') + '=' + (Get-Sha256 $_.FullName) } | Sort-Object)
        $actual -join "`n" | Should -Be ($expected -join "`n")
        $templates = Join-Path $TestDrive 'templates'
        & (Join-Path $script:Root 'tools/New-ExampleTemplates.ps1') -OutputFolder $templates 6> $null
        foreach ($name in 'short-report-template.docx', 'long-report-template.docx') {
            Get-Sha256 (Join-Path $templates $name) | Should -Be (Get-Sha256 (Join-Path $script:Root "templates/$name")) -Because $name
        }
    }
}

Describe 'Live handoff outside Windows' -Skip:($env:OS -eq 'Windows_NT') {
    It 'stops with a clear message instead of trying UI Automation' {
        $config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root
        { Invoke-CopilotHandoff -Config $config -PromptFile 'p' -ResponseFile 'r' -RunPath $TestDrive } | Should -Throw '*Use -Demo*'
    }
}
