# Pester 5 tests for the core module and the two entry scripts.
# Run: Invoke-Pester ./tests

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    Import-Module (Join-Path (Join-Path $script:Root 'src') 'ComplaintDossier.psm1') -Force
    $script:ExampleConfig = Join-Path $script:Root 'config.example.json'
    $script:OpenCases = @('SYN-2026-005', 'SYN-2026-006', 'SYN-2026-007', 'SYN-2026-008', 'SYN-2026-009', 'SYN-2026-010')

    function New-TestConfig {
        # Writes a copy of config.example.json with changes applied by $Edit.
        # Run output and the search for approved reports stay inside TestDrive.
        param([scriptblock]$Edit = {})
        $config = Get-Content -Raw $script:ExampleConfig | ConvertFrom-Json
        $config.Report.OutputPath = Join-Path $TestDrive ('out-' + [guid]::NewGuid().ToString('N'))
        & $Edit $config
        $path = Join-Path $TestDrive ("config-{0}.json" -f [guid]::NewGuid().ToString('N'))
        $config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $path -Encoding utf8
        $path
    }
    function New-TestCase([string]$Id, [string]$Date, [string]$Category, [string]$Group = 'G', [string]$Batch = '') {
        [pscustomobject]@{ CaseId = $Id; ReceivedDate = [datetime]$Date; Category = $Category; ProductGroup = $Group; Batch = $Batch; Severity = ''; Status = 'Open'; SampleReceived = ''; Description = '' }
    }
    function Invoke-Demo([string]$Config, [string]$Out, [string[]]$Ids) {
        $arguments = @{ Demo = $true; ConfigPath = $Config; OutputPath = $Out; PassThru = $true }
        if ($Ids) { $arguments.CaseId = $Ids }
        & (Join-Path $script:Root 'Invoke-ComplaintDossier.ps1') @arguments 6> $null
    }
}

Describe 'Configuration' {
    It 'loads config.example.json in demo mode and resolves paths against the project root' {
        $config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root -Mode Demo
        $config.Paths.Workbook | Should -Be ([IO.Path]::GetFullPath((Join-Path $script:Root 'samples/complaint-list.xlsx')))
        $config.Paths.Prompts.Long | Should -Be ([IO.Path]::GetFullPath((Join-Path $script:Root 'prompts/long-report.txt')))
    }
    It 'refuses live mode when no model is configured' {
        { Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root -Mode Live } |
            Should -Throw '*Copilot.PreferredModel*never picks a model*'
    }
    It 'reports a missing configuration file with a hint' {
        { Import-DossierConfig -Path (Join-Path $TestDrive 'nope.json') -Root $script:Root } | Should -Throw '*Copy config.example.json*'
    }
    It 'fails when a configured prompt does not exist' {
        $path = New-TestConfig { param($c) $c.Reports.Long.PromptPath = 'prompts/missing.txt' }
        { Import-DossierConfig -Path $path -Root $script:Root } | Should -Throw '*Reports.Long.PromptPath*'
    }
    It 'fails when both report types use the same marker' {
        $path = New-TestConfig { param($c) $c.Reports.Long.Marker = 'short.txt' }
        { Import-DossierConfig -Path $path -Root $script:Root } | Should -Throw '*must differ*'
    }
}

Describe 'Reading the complaint list' {
    BeforeAll { $script:Config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root }

    It 'reads all 10 synthetic cases without Excel' {
        $cases = @(Import-DossierCases -Config $script:Config)
        $cases.Count | Should -Be 10
        $cases[0].CaseId | Should -Be 'SYN-2026-001'
        $cases[0].ReceivedDate | Should -Be ([datetime]'2026-01-12')
        $cases[6].Severity | Should -Be 'Critical'
        $cases[6].ProductGroup | Should -Be 'Cream'
        $cases[6].Batch | Should -Be 'SYN-LOT-26D002'
        $cases[6].SampleReceived | Should -Be 'No'
    }
    It 'uses only batch numbers in the synthetic SYN-LOT- format' {
        foreach ($case in Import-DossierCases -Config $script:Config) { $case.Batch | Should -Match '^SYN-LOT-\d\d[A-Z]\d{3}$' -Because $case.CaseId }
    }
    It 'keeps the synthetic notice in the first row of the sheet' {
        $rows = Read-XlsxSheet -Path $script:Config.Paths.Workbook -Sheet 'Complaints' -HeaderRow 1
        ($rows[0].PSObject.Properties.Name -join ' ') | Should -Match 'SYNTHETIC TEST DATA'
    }
    It 'names the column when the mapping does not match the workbook' {
        $path = New-TestConfig { param($c) $c.Workbook.Columns.Category = 'Kind' }
        $config = Import-DossierConfig -Path $path -Root $script:Root
        { Import-DossierCases -Config $config } | Should -Throw "*Column 'Kind'*"
    }
    It 'names the sheet when it does not exist' {
        $path = New-TestConfig { param($c) $c.Workbook.Sheet = 'Missing' }
        $config = Import-DossierConfig -Path $path -Root $script:Root
        { Import-DossierCases -Config $config } | Should -Throw "*'Missing' not found*Complaints, About*"
    }
}

Describe 'Case selection' {
    BeforeAll {
        $script:Config = Import-DossierConfig -Path (New-TestConfig) -Root $script:Root
        $script:Cases = @(Import-DossierCases -Config $script:Config)
    }
    It 'selects all open cases and reads the report type from the marker file' {
        $selection = @(Select-DossierCases -Config $script:Config -Cases $script:Cases)
        @($selection | ForEach-Object { $_.Case.CaseId }) -join ',' | Should -Be ($script:OpenCases -join ',')
        @($selection | Where-Object SkipReason).Count | Should -Be 0
        ($selection | Where-Object { $_.Case.CaseId -eq 'SYN-2026-007' }).Type | Should -Be 'Long'
        ($selection | Where-Object { $_.Case.CaseId -eq 'SYN-2026-006' }).Type | Should -Be 'Short'
    }
    It 'selects a closed case when its ID is given, and skips it without a marker' {
        $selection = @(Select-DossierCases -Config $script:Config -Cases $script:Cases -CaseId 'SYN-2026-001')
        $selection.Count | Should -Be 1
        $selection[0].SkipReason | Should -Match 'no report type marker \(short\.txt or long\.txt\)'
    }
    It 'rejects unknown case IDs' {
        { Select-DossierCases -Config $script:Config -Cases $script:Cases -CaseId 'SYN-2026-999' } | Should -Throw '*SYN-2026-999*'
    }
    It 'skips a case whose folder holds both markers' {
        $folders = Join-Path $TestDrive 'folders-both'
        Copy-Item -Recurse (Join-Path $script:Root 'samples/case-folders') $folders
        Set-Content -LiteralPath (Join-Path $folders 'SYN-2026-006/long.txt') -Value 'SYNTHETIC marker'
        $config = Import-DossierConfig -Path (New-TestConfig { param($c) $c.CaseFolders.Path = $folders }) -Root $script:Root
        $item = @(Select-DossierCases -Config $config -Cases $script:Cases -CaseId 'SYN-2026-006')[0]
        $item.SkipReason | Should -Match 'both short\.txt and long\.txt'
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
        ($checks | Where-Object Name -eq 'Report cites its own case ID').Passed | Should -BeFalse
        ($checks | Where-Object Name -eq 'Cited case IDs exist').Detail | Should -Be 'not in the complaint list: SYN-9'
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

Describe 'Demo run: one DRAFT report per open case' {
    BeforeAll {
        $script:Config = New-TestConfig
        $script:Run = Join-Path $TestDrive 'run'
        $script:Result = Invoke-Demo $script:Config $script:Run
        $script:Report007 = Join-Path $script:Run 'SYN-2026-007/SYN-2026-007-long-report-DRAFT.docx'
        $script:Audit007 = Read-JsonFile (Join-Path $script:Run 'SYN-2026-007/audit.json')
    }
    It 'drafts every open case in its own folder, with short or long report from the marker' {
        @($script:Result.Cases | Where-Object status -eq 'drafted').Count | Should -Be 6
        foreach ($id in $script:OpenCases) {
            foreach ($name in 'prompt.txt', 'response.md', 'case-data.json', 'audit.json') { Join-Path $script:Run "$id/$name" | Should -Exist }
        }
        Join-Path $script:Run 'SYN-2026-006/SYN-2026-006-short-report-DRAFT.docx' | Should -Exist
        $script:Report007 | Should -Exist
        Join-Path $script:Run 'SYN-2026-001' | Should -Not -Exist
        (Read-JsonFile (Join-Path $script:Run 'run.json')).cases.Count | Should -Be 6
    }
    It 'marks the report as DRAFT in the page header and in the approval block' {
        $text = Get-DocxText $script:Report007
        $markers = Get-DossierMarkers
        foreach ($key in $markers.Keys) { $text | Should -Match ([regex]::Escape($markers[$key])) }
        $text | Should -Match 'Demo: prewritten sample answer, Copilot not used'
        $text | Should -Match 'Complaint report \(long\) SYN-2026-007'
    }
    It 'shows the complaint record, the similar complaints and the case documents of this case only' {
        $text = Get-DocxText (Join-Path $script:Run 'SYN-2026-010/SYN-2026-010-long-report-DRAFT.docx')
        $text | Should -Match 'Same batch: 1 other case in the complaint list has batch SYN-LOT-26A019: SYN-2026-009\.'
        $text | Should -Match 'sample-check\.txt'
        $text | Should -Match '\[phone removed\]'
        $text | Should -Not -Match '5550 1234'
        $text | Should -Not -Match 'SYN-LOT-26D002'
    }
    It 'records SHA-256 hashes that match the files' {
        $script:Audit007.schema | Should -Be 'complaint-dossier/audit/v3'
        $script:Audit007.caseId | Should -Be 'SYN-2026-007'
        $script:Audit007.reportType | Should -Be 'long'
        $script:Audit007.status | Should -Be 'DRAFT'
        $script:Audit007.review.reviewer | Should -BeNullOrEmpty
        foreach ($name in 'caseData', 'prompt', 'response', 'report') {
            $artifact = $script:Audit007.artifacts.$name
            Get-Sha256 (Join-Path $script:Run "SYN-2026-007/$($artifact.file)") | Should -Be $artifact.sha256
        }
        ($script:Audit007.inputs | Where-Object role -eq 'complaint-list').sha256 | Should -Be (Get-Sha256 (Join-Path $script:Root 'samples/complaint-list.xlsx'))
        @($script:Audit007.inputs | Where-Object role -eq 'case-document' | ForEach-Object path) -join ',' |
            Should -Be 'samples/case-folders/SYN-2026-007/intake-note.txt,samples/case-folders/SYN-2026-007/lab-request.txt'
        ($script:Audit007.inputs | Where-Object role -eq 'demo-response').path | Should -Be 'samples/demo-responses/SYN-2026-007.md'
        @($script:Audit007.checks | Where-Object { -not $_.passed }).Count | Should -Be 0
        ($script:Audit007.checks | Where-Object name -eq 'Required statements present').detail | Should -Match 'present: No complaint sample was available for the investigation\.'
        ($script:Audit007.checks | Where-Object name -eq 'Conclusion ends with a classification').passed | Should -BeTrue
    }
    It 'logs input paths relative to the project, without user folders' {
        foreach ($item in $script:Audit007.inputs | Where-Object role -ne 'config') { $item.path | Should -Not -Match '^(/|[A-Za-z]:)' }
    }
    It 'drafts only the given case with -CaseId' {
        $result = Invoke-Demo $script:Config (Join-Path $TestDrive 'single') @('SYN-2026-009')
        @($result.Cases).Count | Should -Be 1
        $result.Cases[0].reportType | Should -Be 'short'
    }
    It 'records a case as failed and goes on when its demo answer is missing' {
        $answers = Join-Path $TestDrive 'answers'
        Copy-Item -Recurse (Join-Path $script:Root 'samples/demo-responses') $answers
        Remove-Item (Join-Path $answers 'SYN-2026-008.md')
        $config = New-TestConfig { param($c) $c.Demo.ResponsePath = $answers }
        $result = Invoke-Demo $config (Join-Path $TestDrive 'missing-answer')
        $LASTEXITCODE | Should -Be 2
        ($result.Cases | Where-Object caseId -eq 'SYN-2026-008').status | Should -Be 'failed'
        @($result.Cases | Where-Object status -eq 'drafted').Count | Should -Be 5
    }
}

Describe 'Approval of one report' {
    BeforeAll {
        $script:Config = New-TestConfig
        $script:Run = Join-Path ((Get-Content -Raw $script:Config | ConvertFrom-Json).Report.OutputPath) 'run-a'
        Invoke-Demo $script:Config $script:Run | Out-Null
        $script:Case = Join-Path $script:Run 'SYN-2026-007'
        $script:Before = Read-JsonFile (Join-Path $script:Case 'audit.json')
    }
    It 'refuses approval of a run folder and points to the case folders' {
        { Approve-DossierReport -Path $script:Run -Reviewer 'R. Viewer' } | Should -Throw '*approve each report in its case folder*'
    }
    It 'refuses approval when the response changed after the run' {
        $copy = Join-Path $TestDrive 'tampered'
        Copy-Item -Recurse $script:Case $copy
        Add-Content -LiteralPath (Join-Path $copy 'response.md') -Value 'edited later'
        { Approve-DossierReport -Path $copy -Reviewer 'R. Viewer' } | Should -Throw '*Hash mismatch for response.md*'
        (Read-JsonFile (Join-Path $copy 'audit.json')).status | Should -Be 'DRAFT'
    }
    It 'refuses approval when the case data changed after the run' {
        $copy = Join-Path $TestDrive 'tampered-data'
        Copy-Item -Recurse $script:Case $copy
        Add-Content -LiteralPath (Join-Path $copy 'case-data.json') -Value ' '
        { Approve-DossierReport -Path $copy -Reviewer 'R. Viewer' } | Should -Throw '*Hash mismatch for case-data.json*'
    }
    It 'refuses approval without a reviewer name' {
        { Approve-DossierReport -Path $script:Case -Reviewer ' ' } | Should -Throw '*Reviewer name is required*'
    }
    It 'approves one report: removes DRAFT, records reviewer and time, hashes the approved report, leaves the other reports DRAFT' {
        & (Join-Path $script:Root 'Approve-ComplaintDossier.ps1') -Path $script:Case -Reviewer 'R. Viewer' 6> $null
        $audit = Read-JsonFile (Join-Path $script:Case 'audit.json')
        $audit.status | Should -Be 'APPROVED'
        $audit.review.reviewer | Should -Be 'R. Viewer'
        $audit.review.approvedAtUtc | Should -Match '^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ$'
        $audit.review.draftSha256 | Should -Be $script:Before.artifacts.report.sha256
        $audit.events[0].atUtc | Should -Be $script:Before.events[0].atUtc
        $audit.events[-1].event | Should -Be 'approved'
        $approved = Join-Path $script:Case $audit.artifacts.approvedReport.file
        $audit.artifacts.approvedReport.file | Should -Be 'SYN-2026-007-long-report.docx'
        Get-Sha256 $approved | Should -Be $audit.artifacts.approvedReport.sha256
        $text = Get-DocxText $approved
        $text | Should -Not -MatchExactly 'DRAFT'
        $text | Should -Match 'Approved by: R\. Viewer'
        $text | Should -Match 'Approved at \(UTC\): \d{4}-'
        $text | Should -Not -Match '(?i)\b(signed|signs|sign-off)\b'
        $text | Should -Match 'Status: APPROVED'
        $text | Should -Match 'Accepted check failures: none \(all automated checks passed\)'
        $audit.review.exception | Should -BeNullOrEmpty
        Join-Path $script:Case $audit.artifacts.report.file | Should -Exist
        (Read-JsonFile (Join-Path $script:Run 'SYN-2026-010/audit.json')).status | Should -Be 'DRAFT'
    }
    It 'refuses a second approval' {
        { Approve-DossierReport -Path $script:Case -Reviewer 'Someone Else' } | Should -Throw '*only DRAFT reports*'
    }
    It 'does not draft an approved case again' {
        $result = Invoke-Demo $script:Config (Join-Path $TestDrive 'after-approval')
        $item = $result.Cases | Where-Object caseId -eq 'SYN-2026-007'
        $item.status | Should -Be 'skipped'
        $item.detail | Should -Match 'approved report already exists'
        @($result.Cases | Where-Object status -eq 'drafted').Count | Should -Be 5
    }
}

Describe 'Approval when automated checks failed' {
    BeforeAll {
        # Demo run whose answer for SYN-2026-007 cites a case ID that is not in the complaint list.
        $answers = Join-Path $TestDrive 'answers-bad'
        Copy-Item -Recurse (Join-Path $script:Root 'samples/demo-responses') $answers
        $file = Join-Path $answers 'SYN-2026-007.md'
        Write-Utf8File $file ((Read-Utf8File $file).Replace('In case [SYN-2026-007] a dark fibre', 'In case [SYN-2026-007] (compare [SYN-2026-099]) a dark fibre'))
        $config = New-TestConfig { param($c) $c.Demo.ResponsePath = $answers }
        $script:FailedRun = Join-Path $TestDrive 'failed-run'
        Invoke-Demo $config $script:FailedRun @('SYN-2026-007', 'SYN-2026-006') | Out-Null
        function Copy-FailedReport([string]$Name) {
            $copy = Join-Path $TestDrive $Name
            Copy-Item -Recurse (Join-Path $script:FailedRun 'SYN-2026-007') $copy
            $copy
        }
    }
    It 'records the unknown case ID as a failed check' {
        $audit = Read-JsonFile (Join-Path $script:FailedRun 'SYN-2026-007/audit.json')
        $check = $audit.checks | Where-Object name -eq 'Cited case IDs exist'
        $check.passed | Should -BeFalse
        $check.detail | Should -Be 'not in the complaint list: SYN-2026-099'
    }
    It 'refuses normal approval, names the failed check and leaves the report DRAFT' {
        $path = Copy-FailedReport 'refused'
        $before = Get-Content -Raw (Join-Path $path 'audit.json')
        { Approve-DossierReport -Path $path -Reviewer 'R. Viewer' -Comment 'Looks fine.' } |
            Should -Throw '*Approval refused: 1 automated check(s) failed: Cited case IDs exist (not in the complaint list: SYN-2026-099)*stays DRAFT*'
        Get-Content -Raw (Join-Path $path 'audit.json') | Should -Be $before
        @(Get-ChildItem -LiteralPath $path -Filter *.docx).Count | Should -Be 1
    }
    It 'still approves the other report of the same run' {
        Approve-DossierReport -Path (Join-Path $script:FailedRun 'SYN-2026-006') -Reviewer 'R. Viewer' | Out-Null
        (Read-JsonFile (Join-Path $script:FailedRun 'SYN-2026-006/audit.json')).status | Should -Be 'APPROVED'
    }
    It 'refuses the override without a justification' {
        $path = Copy-FailedReport 'no-justification'
        { Approve-DossierReport -Path $path -Reviewer 'R. Viewer' -AcceptFailedChecks -Justification '  ' } | Should -Throw '*needs a -Justification*'
        (Read-JsonFile (Join-Path $path 'audit.json')).status | Should -Be 'DRAFT'
    }
    It 'refuses a justification without the override switch' {
        $path = Copy-FailedReport 'justification-only'
        { Approve-DossierReport -Path $path -Reviewer 'R. Viewer' -Justification 'reason' } | Should -Throw '*only used together with -AcceptFailedChecks*'
    }
    It 'refuses the override when all checks passed' {
        $run = Join-Path $TestDrive 'clean-run'
        Invoke-Demo (New-TestConfig) $run @('SYN-2026-005') | Out-Null
        { Approve-DossierReport -Path (Join-Path $run 'SYN-2026-005') -Reviewer 'R. Viewer' -AcceptFailedChecks -Justification 'reason' } | Should -Throw '*All automated checks passed*'
    }
    It 'records an override with justification in audit.json and in the approved report' {
        $path = Copy-FailedReport 'override'
        $reason = 'SYN-2026-099 is a typo for SYN-2026-009; checked against the complaint list.'
        & (Join-Path $script:Root 'Approve-ComplaintDossier.ps1') -Path $path -Reviewer 'R. Viewer' -AcceptFailedChecks -Justification $reason 6> $null
        $audit = Read-JsonFile (Join-Path $path 'audit.json')
        $audit.status | Should -Be 'APPROVED'
        $audit.review.exception.justification | Should -Be $reason
        @($audit.review.exception.acceptedFailedChecks).Count | Should -Be 1
        $audit.review.exception.acceptedFailedChecks[0].name | Should -Be 'Cited case IDs exist'
        $audit.events[-1].event | Should -Be 'approved-with-exception'
        $text = Get-DocxText (Join-Path $path $audit.artifacts.approvedReport.file)
        $text | Should -Match 'APPROVED WITH A DOCUMENTED EXCEPTION by R\. Viewer'
        $text | Should -Match 'Status: APPROVED with a documented exception'
        $text | Should -Match ([regex]::Escape("Accepted check failures: DOCUMENTED EXCEPTION. Failed checks: Cited case IDs exist (not in the complaint list: SYN-2026-099). Justification: $reason"))
        $text | Should -Not -MatchExactly 'DRAFT'
    }
}

Describe 'Committed sample in docs/sample' {
    BeforeAll {
        $script:Sample = Join-Path (Join-Path $script:Root 'docs') 'sample'
        $script:SampleIds = @('SYN-2026-007', 'SYN-2026-006')
        $script:Fresh = Join-Path $TestDrive 'sample-check'
        Invoke-Demo (New-TestConfig) $script:Fresh $script:SampleIds | Out-Null
    }
    It 'holds the main sample PDF that other pages link to' {
        Join-Path $script:Sample 'complaint-dossier-sample-DRAFT.pdf' | Should -Exist
        Join-Path $script:Sample 'SYN-2026-007/SYN-2026-007-long-report-DRAFT.pdf' | Should -Exist
        Join-Path $script:Sample 'SYN-2026-006/SYN-2026-006-short-report-DRAFT.pdf' | Should -Exist
    }
    It 'is an unapproved demo report with all checks passed, matching the hashes in its audit.json (<_>)' -ForEach @('SYN-2026-007', 'SYN-2026-006') {
        $folder = Join-Path $script:Sample $_
        $audit = Read-JsonFile (Join-Path $folder 'audit.json')
        $audit.mode | Should -Be 'demo'
        $audit.status | Should -Be 'DRAFT'
        @($audit.checks | Where-Object { -not $_.passed }).Count | Should -Be 0
        foreach ($name in 'caseData', 'prompt', 'response', 'report') {
            $artifact = $audit.artifacts.$name
            Get-Sha256 (Join-Path $folder $artifact.file) | Should -Be $artifact.sha256 -Because $artifact.file
        }
    }
    It 'matches the inputs and the report the current code produces (<_>)' -ForEach @('SYN-2026-007', 'SYN-2026-006') {
        $folder = Join-Path $script:Sample $_
        $sample = Read-JsonFile (Join-Path $folder 'audit.json')
        $fresh = Read-JsonFile (Join-Path (Join-Path $script:Fresh $_) 'audit.json')
        foreach ($item in $fresh.inputs | Where-Object role -ne 'config') {
            ($sample.inputs | Where-Object path -eq $item.path).sha256 | Should -Be $item.sha256 -Because $item.path
        }
        foreach ($name in 'caseData', 'prompt', 'response') {
            $sample.artifacts.$name.sha256 | Should -Be $fresh.artifacts.$name.sha256 -Because $name
        }
        # Run ID and generation time differ between runs; everything else must match.
        $normalize = {
            param($path, $runId)
            ((Get-DocxText $path).Replace($runId, '<run id>')) -replace '\d{4}-\d\d-\d\d \d\d:\d\d:\d\d', '<time>'
        }
        $expected = & $normalize (Join-Path $folder $sample.artifacts.report.file) $sample.runId
        & $normalize (Join-Path (Join-Path $script:Fresh $_) $fresh.artifacts.report.file) $fresh.runId | Should -Be $expected
    }
}

Describe 'Report template' {
    It 'fails when the template has no body anchor' {
        $template = Join-Path $TestDrive 'bad.docx'
        New-ZipPackage -Path $template -Parts ([ordered]@{ 'word/document.xml' = '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>no anchor</w:t></w:r></w:p></w:body></w:document>' })
        $case = New-TestCase 'A-1' '2026-01-05' 'X'
        $similar = Get-SimilarComplaints -Case $case -Cases @($case)
        $run = [pscustomobject]@{ RunId = 'r'; ReportType = 'Short report'; GeneratedAt = 'now'; ModeText = 'm'; Model = 'none'; ResponseNote = 'n'; WorkbookSha256 = 'a'; PromptSha256 = 'b'; ResponseSha256 = 'c' }
        $out = Join-Path $TestDrive 'out.docx'
        { New-CaseReport -TemplatePath $template -OutputPath $out -Title 't' -Run $run -Case $case -Similar $similar -Response 'r' -Checks @([pscustomobject]@{ Name = 'c'; Passed = $true; Detail = '' }) } |
            Should -Throw '*{{REPORT_BODY}}*'
        $out | Should -Not -Exist
    }
}

Describe 'Synthetic data generator' {
    It 'reproduces the committed complaint list and case folders' {
        $out = Join-Path $TestDrive 'samples'
        & (Join-Path $script:Root 'tools/New-SyntheticCases.ps1') -OutputPath $out 6> $null
        $config = Import-DossierConfig -Path (New-TestConfig) -Root $script:Root
        $committed = Import-DossierCases -Config $config | ConvertTo-Json
        $config.Paths.Workbook = Join-Path $out 'complaint-list.xlsx'
        Import-DossierCases -Config $config | ConvertTo-Json | Should -Be $committed
        $committedRoot = Join-Path $script:Root 'samples/case-folders'
        $expected = @(Get-ChildItem -Recurse -File $committedRoot | ForEach-Object { $_.FullName.Substring($committedRoot.Length).Replace('\', '/') + '=' + (Get-Sha256 $_.FullName) } | Sort-Object)
        $generatedRoot = Join-Path $out 'case-folders'
        $actual = @(Get-ChildItem -Recurse -File $generatedRoot | ForEach-Object { $_.FullName.Substring($generatedRoot.Length).Replace('\', '/') + '=' + (Get-Sha256 $_.FullName) } | Sort-Object)
        $actual -join "`n" | Should -Be ($expected -join "`n")
    }
}

Describe 'Live handoff outside Windows' -Skip:($env:OS -eq 'Windows_NT') {
    It 'stops with a clear message instead of trying UI Automation' {
        $config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root
        { Invoke-CopilotHandoff -Config $config -PromptFile 'p' -ResponseFile 'r' -RunPath $TestDrive } | Should -Throw '*Use -Demo*'
    }
}
