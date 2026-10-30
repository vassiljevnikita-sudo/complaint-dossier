# Pester 5 tests for the core module and the two entry scripts.
# Run: Invoke-Pester ./tests

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    Import-Module (Join-Path (Join-Path $script:Root 'src') 'ComplaintDossier.psm1') -Force
    $script:ExampleConfig = Join-Path $script:Root 'config.example.json'

    function New-TestConfig {
        # Writes a copy of config.example.json with changes applied by $Edit.
        param([scriptblock]$Edit)
        $config = Get-Content -Raw $script:ExampleConfig | ConvertFrom-Json
        & $Edit $config
        $path = Join-Path $TestDrive ("config-{0}.json" -f [guid]::NewGuid().ToString('N'))
        $config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $path -Encoding utf8
        $path
    }
    function New-TestCase([string]$Id, [string]$Date, [string]$Category) {
        [pscustomobject]@{ CaseId = $Id; ReceivedDate = [datetime]$Date; Category = $Category; ProductGroup = ''; Severity = ''; Status = ''; Description = '' }
    }
}

Describe 'Configuration' {
    It 'loads config.example.json in demo mode and resolves paths against the project root' {
        $config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root -Mode Demo
        $config.Paths.Workbook | Should -Be ([IO.Path]::GetFullPath((Join-Path $script:Root 'samples/cases.xlsx')))
    }
    It 'refuses live mode when no model is configured' {
        { Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root -Mode Live } |
            Should -Throw '*Copilot.PreferredModel*never picks a model*'
    }
    It 'reports a missing configuration file with a hint' {
        { Import-DossierConfig -Path (Join-Path $TestDrive 'nope.json') -Root $script:Root } | Should -Throw '*Copy config.example.json*'
    }
    It 'fails when a configured file does not exist' {
        $path = New-TestConfig { param($c) $c.Prompt.TemplatePath = 'prompts/missing.txt' }
        { Import-DossierConfig -Path $path -Root $script:Root } | Should -Throw '*PromptTemplate*'
    }
}

Describe 'Reading cases from the workbook' {
    BeforeAll { $script:Config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root }

    It 'reads all 10 synthetic cases without Excel' {
        $cases = Import-DossierCases -Config $script:Config
        $cases.Count | Should -Be 10
        $cases[0].CaseId | Should -Be 'SYN-2026-001'
        $cases[0].ReceivedDate | Should -Be ([datetime]'2026-01-12')
        $cases[6].Severity | Should -Be 'Critical'
    }
    It 'keeps the synthetic notice in the first row of the sheet' {
        $rows = Read-XlsxSheet -Path $script:Config.Paths.Workbook -Sheet 'Cases' -HeaderRow 1
        ($rows[0].PSObject.Properties.Name -join ' ') | Should -Match 'SYNTHETIC TEST DATA'
    }
    It 'filters by case ID and rejects unknown IDs' {
        @(Import-DossierCases -Config $script:Config -CaseId 'SYN-2026-007').Count | Should -Be 1
        { Import-DossierCases -Config $script:Config -CaseId 'SYN-2026-999' } | Should -Throw '*SYN-2026-999*'
    }
    It 'names the column when the mapping does not match the workbook' {
        $path = New-TestConfig { param($c) $c.Workbook.Columns.Category = 'Kind' }
        $config = Import-DossierConfig -Path $path -Root $script:Root
        { Import-DossierCases -Config $config } | Should -Throw "*Column 'Kind'*"
    }
    It 'names the sheet when it does not exist' {
        $path = New-TestConfig { param($c) $c.Workbook.Sheet = 'Missing' }
        $config = Import-DossierConfig -Path $path -Root $script:Root
        { Import-DossierCases -Config $config } | Should -Throw "*'Missing' not found*Cases, About*"
    }
}

Describe 'Statistics' {
    It 'counts by category with shares, largest first' {
        $config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root
        $stats = Get-DossierStatistics -Cases (Import-DossierCases -Config $config)
        $stats.Total | Should -Be 10
        $stats.ByCategory[0].Category | Should -Be 'Packaging damage'
        $stats.ByCategory[0].Count | Should -Be 5
        $stats.ByCategory[0].Share | Should -Be 50
        ($stats.ByMonth | ForEach-Object { "$($_.Month)=$($_.Count)" }) -join ' ' |
            Should -Be '2026-01=2 2026-02=1 2026-03=2 2026-04=2 2026-05=2 2026-06=1'
    }
    It 'lists months without cases as 0' {
        $stats = Get-DossierStatistics -Cases @((New-TestCase 'A-1' '2026-01-05' 'X'), (New-TestCase 'A-2' '2026-03-07' 'X'))
        ($stats.ByMonth | ForEach-Object { $_.Count }) -join ',' | Should -Be '1,0,1'
    }
}

Describe 'Prompt' {
    BeforeAll {
        $script:Config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root
        $script:Cases = @(Import-DossierCases -Config $script:Config)
        $script:Stats = Get-DossierStatistics -Cases $script:Cases
        $script:Docs = @(Get-CaseDocuments -Config $script:Config -Cases $script:Cases)
        $script:Template = Read-Utf8File $script:Config.Paths.PromptTemplate
    }
    It 'fills every placeholder' {
        $prompt = New-DossierPrompt -Template $script:Template -Cases $script:Cases -Statistics $script:Stats -Documents $script:Docs
        $prompt | Should -Not -Match '\{\{[A-Z_]+\}\}'
        $prompt | Should -Match '\| Packaging damage \| 5 \| 50\.0 \|'
        $prompt | Should -Match 'Document lab-request\.txt for case SYN-2026-007'
    }
    It 'masks e-mail addresses and phone numbers by default' {
        $prompt = New-DossierPrompt -Template $script:Template -Cases $script:Cases -Statistics $script:Stats -Documents $script:Docs
        $prompt | Should -Not -Match 'example\.org'
        $prompt | Should -Not -Match '5550 1234'
        $prompt | Should -Match '\[phone removed\]'
    }
    It 'rejects a template with an unknown placeholder' {
        { New-DossierPrompt -Template 'x {{NOPE}}' -Cases $script:Cases -Statistics $script:Stats } | Should -Throw '*{{NOPE}}*'
    }
}

Describe 'Response checks' {
    BeforeAll {
        $script:Sections = @('Summary', 'Data gaps')
        $script:Known = @('SYN-1', 'SYN-2')
    }
    It 'passes a response with all sections and known case IDs' {
        $checks = Test-DossierResponse -Response "## Summary`nSee [SYN-1].`n## Data gaps`nNone [1]." -Prompt 'p' -RequiredSections $script:Sections -KnownCaseIds $script:Known
        @($checks | Where-Object { -not $_.Passed }).Count | Should -Be 0
    }
    It 'fails on a missing section and on an invented case ID' {
        $checks = Test-DossierResponse -Response "## Summary`nSee [SYN-9]." -Prompt 'p' -RequiredSections $script:Sections -KnownCaseIds $script:Known
        ($checks | Where-Object Name -eq 'Required sections present').Detail | Should -Be 'missing: Data gaps'
        ($checks | Where-Object Name -eq 'Cited case IDs exist').Detail | Should -Be 'not in the input: SYN-9'
    }
    It 'fails when the response is the prompt' {
        $checks = Test-DossierResponse -Response 'same text' -Prompt 'same text' -RequiredSections @() -KnownCaseIds @()
        ($checks | Where-Object Name -eq 'Response differs from prompt').Passed | Should -BeFalse
    }
}

Describe 'Demo run, report and approval' {
    BeforeAll {
        $script:Run = Join-Path $TestDrive 'run'
        $result = & (Join-Path $script:Root 'Invoke-ComplaintDossier.ps1') -Demo -OutputPath $script:Run -PassThru 6> $null
        $script:Result = $result
        $script:Audit = Read-JsonFile (Join-Path $script:Run 'audit.json')
    }
    It 'writes the draft report, prompt, response, statistics and audit log' {
        foreach ($name in 'prompt.txt', 'response.md', 'statistics.json', 'audit.json') { Join-Path $script:Run $name | Should -Exist }
        $script:Result.Report | Should -Match '-DRAFT\.docx$'
        $script:Result.Report | Should -Exist
    }
    It 'marks the report as DRAFT in the page header and in the sign-off block' {
        $text = Get-DocxText $script:Result.Report
        $markers = Get-DossierMarkers
        foreach ($key in $markers.Keys) { $text | Should -Match ([regex]::Escape($markers[$key])) }
        $text | Should -Match 'Demo: prewritten sample response, Copilot not used'
        $text | Should -Match 'Packaging damage'
    }
    It 'records SHA-256 hashes that match the files' {
        $script:Audit.schema | Should -Be 'complaint-dossier/audit/v1'
        $script:Audit.status | Should -Be 'DRAFT'
        $script:Audit.review.reviewer | Should -BeNullOrEmpty
        foreach ($name in 'prompt', 'statistics', 'response', 'report') {
            $artifact = $script:Audit.artifacts.$name
            Get-Sha256 (Join-Path $script:Run $artifact.file) | Should -Be $artifact.sha256
        }
        ($script:Audit.inputs | Where-Object role -eq 'workbook').sha256 | Should -Be (Get-Sha256 (Join-Path $script:Root 'samples/cases.xlsx'))
        @($script:Audit.inputs | Where-Object role -eq 'case-document').Count | Should -Be 6
        @($script:Audit.checks | Where-Object { -not $_.passed }).Count | Should -Be 0
    }
    It 'logs input paths relative to the project, without user folders' {
        foreach ($item in $script:Audit.inputs) { $item.path | Should -Not -Match '^(/|[A-Za-z]:)' }
    }
    It 'refuses approval when the response changed after the run' {
        $copy = Join-Path $TestDrive 'tampered'
        Copy-Item -Recurse $script:Run $copy
        Add-Content -LiteralPath (Join-Path $copy 'response.md') -Value 'edited later'
        { Approve-DossierRun -RunPath $copy -Reviewer 'R. Viewer' } | Should -Throw '*Hash mismatch for response.md*'
        (Read-JsonFile (Join-Path $copy 'audit.json')).status | Should -Be 'DRAFT'
    }
    It 'refuses approval without a reviewer name' {
        { Approve-DossierRun -RunPath $script:Run -Reviewer ' ' } | Should -Throw '*Reviewer name is required*'
    }
    It 'signs the run: removes DRAFT, records reviewer and time, hashes the signed report' {
        & (Join-Path $script:Root 'Approve-ComplaintDossier.ps1') -RunPath $script:Run -Reviewer 'R. Viewer' 6> $null
        $audit = Read-JsonFile (Join-Path $script:Run 'audit.json')
        $audit.status | Should -Be 'APPROVED'
        $audit.review.reviewer | Should -Be 'R. Viewer'
        $audit.review.signedAtUtc | Should -Match '^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ$'
        $audit.review.draftSha256 | Should -Be $script:Audit.artifacts.report.sha256
        $audit.events[0].atUtc | Should -Be $script:Audit.events[0].atUtc
        $audit.events[-1].event | Should -Be 'approved'
        $signed = Join-Path $script:Run $audit.artifacts.approvedReport.file
        $signed | Should -Not -Match 'DRAFT'
        Get-Sha256 $signed | Should -Be $audit.artifacts.approvedReport.sha256
        $text = Get-DocxText $signed
        $text | Should -Not -MatchExactly 'DRAFT'
        $text | Should -Match 'Reviewer: R\. Viewer'
        $text | Should -Match 'Status: APPROVED'
        Join-Path $script:Run $audit.artifacts.report.file | Should -Exist
    }
    It 'refuses a second approval' {
        { Approve-DossierRun -RunPath $script:Run -Reviewer 'Someone Else' } | Should -Throw '*only DRAFT runs*'
    }
}

Describe 'Report template' {
    It 'fails when the template has no body anchor' {
        $template = Join-Path $TestDrive 'bad.docx'
        New-ZipPackage -Path $template -Parts ([ordered]@{ 'word/document.xml' = '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>no anchor</w:t></w:r></w:p></w:body></w:document>' })
        $stats = Get-DossierStatistics -Cases @(New-TestCase 'A-1' '2026-01-05' 'X')
        $run = [pscustomobject]@{ RunId = 'r'; GeneratedAt = 'now'; ModeText = 'm'; Model = 'none'; ResponseNote = 'n'; WorkbookSha256 = 'a'; PromptSha256 = 'b'; ResponseSha256 = 'c' }
        $out = Join-Path $TestDrive 'out.docx'
        { New-DossierReport -TemplatePath $template -OutputPath $out -Title 't' -Run $run -Cases @(New-TestCase 'A-1' '2026-01-05' 'X') -Statistics $stats -Response 'r' -Checks @([pscustomobject]@{ Name = 'c'; Passed = $true; Detail = '' }) } |
            Should -Throw '*{{REPORT_BODY}}*'
        $out | Should -Not -Exist
    }
}

Describe 'Synthetic data generator' {
    It 'reproduces the committed sample cases' {
        $out = Join-Path $TestDrive 'samples'
        & (Join-Path $script:Root 'tools/New-SyntheticCases.ps1') -OutputPath $out 6> $null
        $path = New-TestConfig { param($c) }
        $config = Import-DossierConfig -Path $path -Root $script:Root
        $committed = Import-DossierCases -Config $config | ConvertTo-Json
        $config.Paths.Workbook = Join-Path $out 'cases.xlsx'
        Import-DossierCases -Config $config | ConvertTo-Json | Should -Be $committed
        foreach ($file in Get-ChildItem -Recurse -File $out -Filter *.txt) {
            (Get-Content -LiteralPath $file.FullName -TotalCount 1) | Should -Match '^SYNTHETIC TEST DATA'
        }
    }
}

Describe 'Live handoff outside Windows' -Skip:($env:OS -eq 'Windows_NT') {
    It 'stops with a clear message instead of trying UI Automation' {
        $config = Import-DossierConfig -Path $script:ExampleConfig -Root $script:Root
        { Invoke-CopilotHandoff -Config $config -PromptFile 'p' -ResponseFile 'r' -RunPath $TestDrive } | Should -Throw '*Use -Demo*'
    }
}
