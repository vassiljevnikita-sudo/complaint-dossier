<#
.SYNOPSIS
  Builds a complaint trend dossier: Excel cases -> statistics -> one prompt ->
  Microsoft 365 Copilot (or a prewritten demo response) -> DRAFT Word report
  with a JSON audit log (SHA-256 hashes).

.EXAMPLE
  ./Invoke-ComplaintDossier.ps1 -Demo
  Runs on the synthetic sample data with the prewritten sample response.
  Needs no Excel, Word or Copilot.

.EXAMPLE
  ./Invoke-ComplaintDossier.ps1 -Live -CaseId SYN-2026-007
  Sends one synthetic case to the Microsoft 365 Copilot desktop app
  (Windows only, unlocked desktop, Copilot window open). Reads config.json.
#>
[CmdletBinding()]
param(
    # Use the prewritten sample response instead of Copilot.
    [switch]$Demo,
    # Hand the prompt to the Microsoft 365 Copilot desktop app.
    [switch]$Live,
    # Defaults: config.example.json for -Demo, config.json for -Live.
    [string]$ConfigPath,
    # Restrict the run to these case IDs.
    [string[]]$CaseId,
    # Folder for this run. Default: <Report.OutputPath>/<run id>.
    [string]$OutputPath,
    # Return the run result as an object.
    [switch]$PassThru
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = $PSScriptRoot
Import-Module (Join-Path (Join-Path $root 'src') 'ComplaintDossier.psm1') -Force

if ($Demo -eq $Live) { throw 'Choose exactly one mode: -Demo or -Live.' }
$mode = 'Live'
if ($Demo) { $mode = 'Demo' }
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $root 'config.json'
    if ($Demo) { $ConfigPath = Join-Path $root 'config.example.json' }
}
# Resolve against the PowerShell location; .NET would use the process directory.
$ConfigPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ConfigPath)

$events = New-Object System.Collections.Generic.List[object]
function Add-Event([string]$Name, [string]$Detail = '') {
    $events.Add([pscustomobject]@{ atUtc = (Get-UtcStamp); event = $Name; detail = $Detail })
}
Add-Event 'run-started' "mode=$mode"

$config = Import-DossierConfig -Path $ConfigPath -Root $root -Mode $mode
$runId = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ') + '-' + ([guid]::NewGuid().ToString('N').Substring(0, 6))
if (-not $OutputPath) { $OutputPath = Join-Path $config.Paths.Output $runId }
$runPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ((Test-Path -LiteralPath $runPath) -and @(Get-ChildItem -LiteralPath $runPath -Force).Count) { throw "Output folder is not empty: $runPath" }
New-Item -ItemType Directory -Force -Path $runPath | Out-Null
Write-Host "[run] $runId ($mode mode)"

$inputs = New-Object System.Collections.Generic.List[object]
function Add-Input([string]$Role, [string]$Path) {
    $inputs.Add([pscustomobject]@{ role = $Role; path = (Get-RelativeDisplayPath $root $Path); sha256 = (Get-Sha256 $Path) })
}
Add-Input 'config' $config.Paths.Config
Add-Input 'workbook' $config.Paths.Workbook
Add-Input 'prompt-template' $config.Paths.PromptTemplate
Add-Input 'report-template' $config.Paths.ReportTemplate

# 1. Cases, statistics, documents
$cases = @(Import-DossierCases -Config $config -CaseId $CaseId)
$statistics = Get-DossierStatistics -Cases $cases
$documents = @(Get-CaseDocuments -Config $config -Cases $cases)
foreach ($doc in $documents) { Add-Input 'case-document' $doc.Path }
Add-Event 'cases-read' "$($cases.Count) cases, $($documents.Count) documents"
Write-Host "[data] $($cases.Count) cases, $($documents.Count) case documents, $($statistics.ByCategory.Count) categories, $($statistics.ByMonth.Count) months"

$uploads = @()
if ($mode -eq 'Live') {
    $uploads = @(Get-UploadFiles -Config $config)
    foreach ($file in $uploads) { Add-Input 'upload' $file.FullName }
}

# 2. Prompt
$mask = $true
if ($null -ne (Get-ConfigValue $config.Settings.Prompt 'MaskContactData')) { $mask = [bool]$config.Settings.Prompt.MaskContactData }
$prompt = New-DossierPrompt -Template (Read-Utf8File $config.Paths.PromptTemplate) -Cases $cases -Statistics $statistics `
    -Documents $documents -Attachments @($uploads | ForEach-Object { $_.Name }) -Mask $mask
$promptFile = Join-Path $runPath 'prompt.txt'
Write-Utf8File $promptFile $prompt
$statisticsFile = Join-Path $runPath 'statistics.json'
# JSON with LF line ends, so the hash is the same on every platform.
Write-Utf8File $statisticsFile (($statistics | ConvertTo-Json -Depth 5).Replace("`r`n", "`n") + "`n")
Add-Event 'prompt-built' "$($prompt.Length) characters"

# 3. Response
$responseFile = Join-Path $runPath 'response.md'
$copilotStatus = $null
if ($mode -eq 'Demo') {
    Add-Input 'demo-response' $config.Paths.DemoResponse
    Write-Utf8File $responseFile (Read-Utf8File $config.Paths.DemoResponse)
    $modelText = 'none (demo mode)'
    $modeText = 'Demo: prewritten sample response, Copilot not used'
    $responseNote = 'Demo mode: the text below is a prewritten sample response for the synthetic data. It was not produced by Copilot.'
    Add-Event 'response-received' 'demo response copied'
} else {
    $cp = $config.Settings.Copilot
    $modelText = [string]$cp.PreferredModel
    if ([string](Get-ConfigValue $cp 'PreferredModelGroup')) { $modelText = "$($cp.PreferredModelGroup) / $modelText" }
    $modeText = 'Live: Microsoft 365 Copilot desktop app'
    $responseNote = "Drafted by Microsoft 365 Copilot (configured model: $modelText). Check every statement against the case records."
    Write-Host "[copilot] handing the prompt to '$($cp.WindowTitle)', model '$modelText', $($uploads.Count) upload(s)"
    $copilotStatus = Invoke-CopilotHandoff -Config $config -PromptFile $promptFile -ResponseFile $responseFile -UploadFiles $uploads -RunPath $runPath
    Add-Event 'response-received' "copilot, model=$modelText, uploads=$($uploads.Count)"
}
$response = Read-Utf8File $responseFile

# 4. Checks and report
$checks = @(Test-DossierResponse -Response $response -Prompt $prompt -RequiredSections @($config.Settings.Prompt.RequiredSections) -KnownCaseIds @($cases | ForEach-Object { $_.CaseId }))
foreach ($check in $checks) {
    $mark = 'FAIL'
    if ($check.Passed) { $mark = 'pass' }
    $detail = ''
    if ($check.Detail) { $detail = " ($($check.Detail))" }
    Write-Host ("[check] {0,-4} {1}{2}" -f $mark, $check.Name, $detail)
}
$title = 'Complaint trend dossier'
if ([string](Get-ConfigValue $config.Settings.Report 'Title')) { $title = $config.Settings.Report.Title }
$reportName = "complaint-dossier-$runId-DRAFT.docx"
$reportFile = Join-Path $runPath $reportName
$run = [pscustomobject]@{
    RunId          = $runId
    GeneratedAt    = [DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss')
    ModeText       = $modeText
    Model          = $modelText
    ResponseNote   = $responseNote
    WorkbookSha256 = (Get-Sha256 $config.Paths.Workbook)
    PromptSha256   = (Get-Sha256 $promptFile)
    ResponseSha256 = (Get-Sha256 $responseFile)
}
New-DossierReport -TemplatePath $config.Paths.ReportTemplate -OutputPath $reportFile -Title $title -Run $run `
    -Cases $cases -Statistics $statistics -Response $response -Checks $checks
Add-Event 'report-written' $reportName

# 5. Audit log (SHA-256 hashes)
$artifact = { param($name) [pscustomobject]@{ file = $name; sha256 = (Get-Sha256 (Join-Path $runPath $name)) } }
$audit = [pscustomobject][ordered]@{
    schema    = 'complaint-dossier/audit/v2'
    tool      = [pscustomobject]@{ name = 'Complaint Dossier'; version = (Get-DossierVersion) }
    runId     = $runId
    mode      = $mode.ToLowerInvariant()
    status    = 'DRAFT'
    operator  = [pscustomobject]@{ osAccount = [Environment]::UserName; host = 'PowerShell ' + $PSVersionTable.PSVersion.ToString() + ' (' + $PSVersionTable.PSEdition + ')' }
    model     = [pscustomobject]@{ configured = $modelText; copilot = $copilotStatus }
    caseIds   = @($cases | ForEach-Object { $_.CaseId })
    inputs    = $inputs.ToArray()
    artifacts = [pscustomobject][ordered]@{
        prompt     = & $artifact 'prompt.txt'
        statistics = & $artifact 'statistics.json'
        response   = & $artifact 'response.md'
        report     = & $artifact $reportName
    }
    checks    = @($checks | ForEach-Object { [pscustomobject]@{ name = $_.Name; passed = $_.Passed; detail = $_.Detail } })
    review    = [pscustomobject]@{ reviewer = $null; approvedAtUtc = $null; comment = $null }
    events    = $events.ToArray()
}
$auditFile = Join-Path $runPath 'audit.json'
Save-AuditLog $audit $auditFile

$displayRun = Get-RelativeDisplayPath $root $runPath -FullPathOutside
Write-Host "[report] $displayRun/$reportName"
Write-Host "[audit]  $displayRun/audit.json (report sha256 $($audit.artifacts.report.sha256))"
$failedCount = @($checks | Where-Object { -not $_.Passed }).Count
if ($failedCount) {
    Write-Host "[next]   status DRAFT. $failedCount automated check(s) failed, so normal approval is refused. Start a new run, or after review accept the failures as a documented exception:"
    Write-Host "         ./Approve-ComplaintDossier.ps1 -RunPath $displayRun -Reviewer 'Your Name' -AcceptFailedChecks -Justification '<reason>'"
} else {
    Write-Host "[next]   status DRAFT. After review: ./Approve-ComplaintDossier.ps1 -RunPath $displayRun -Reviewer 'Your Name'"
}
if ($PassThru) {
    [pscustomobject]@{ RunId = $runId; RunPath = $runPath; Report = $reportFile; Audit = $auditFile; Checks = $checks }
}
