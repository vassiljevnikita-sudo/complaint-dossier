<#
.SYNOPSIS
  Drafts one complaint report per case. Reads the complaint list (Excel),
  selects the cases, builds a prompt from each case's row and case folder,
  hands it to Microsoft 365 Copilot (or uses a prewritten demo answer) and
  writes one DRAFT Word report with its own audit log (SHA-256 hashes) per case.

.DESCRIPTION
  Without -CaseId every open case of the complaint list is drafted (status
  listed in Selection.OpenStatuses). The report type comes from a marker file
  in the case folder: short.txt or long.txt. Cases without a marker, and cases
  that already have an approved report, are skipped.
  Each report stays DRAFT until a reviewer approves it with
  Approve-ComplaintDossier.ps1.

.EXAMPLE
  ./Invoke-ComplaintDossier.ps1 -Demo
  Drafts the open synthetic sample cases with prewritten sample answers.
  Needs no Excel, Word or Copilot.

.EXAMPLE
  ./Invoke-ComplaintDossier.ps1 -Live -CaseId SYN-2026-006
  Sends one synthetic case to the Microsoft 365 Copilot desktop app
  (Windows only, unlocked desktop, Copilot window open). Reads config.json.
#>
[CmdletBinding()]
param(
    # Use the prewritten sample answers instead of Copilot.
    [switch]$Demo,
    # Hand each prompt to the Microsoft 365 Copilot desktop app.
    [switch]$Live,
    # Defaults: config.example.json for -Demo, config.json for -Live.
    [string]$ConfigPath,
    # Draft only these cases (whatever their status). Default: all open cases.
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
$config = Import-DossierConfig -Path $ConfigPath -Root $root -Mode $mode

# 1. Complaint list and case selection
$cases = @(Import-DossierCases -Config $config)
$selection = @(Select-DossierCases -Config $config -Cases $cases -CaseId $CaseId)
$todo = @($selection | Where-Object { -not $_.SkipReason })
$scope = 'all open cases'
if ($CaseId) { $scope = 'case(s) ' + ($CaseId -join ', ') }
Write-Host "[list]   $($cases.Count) cases in the complaint list; $($selection.Count) selected ($scope); $($todo.Count) to draft"
foreach ($item in $selection | Where-Object { $_.SkipReason }) { Write-Host "[skip]   $($item.Case.CaseId): $($item.SkipReason)" }
if ($todo.Count -eq 0) {
    Write-Host '[done]   nothing to draft.'
    if ($CaseId) { exit 1 }
    return
}

$runId = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ') + '-' + ([guid]::NewGuid().ToString('N').Substring(0, 6))
if (-not $OutputPath) { $OutputPath = Join-Path $config.Paths.Output $runId }
$runPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ((Test-Path -LiteralPath $runPath) -and @(Get-ChildItem -LiteralPath $runPath -Force).Count) { throw "Output folder is not empty: $runPath" }
New-Item -ItemType Directory -Force -Path $runPath | Out-Null
$displayRun = Get-RelativeDisplayPath $root $runPath -FullPathOutside
Write-Host "[run]    $runId ($mode mode) -> $displayRun"

# 2. One report per case. Cases run one after another, because the live path
#    drives one Copilot window. A failed case does not stop the others.
$results = New-Object System.Collections.Generic.List[object]
foreach ($item in $selection) {
    $id = $item.Case.CaseId
    $typeName = $null
    if ($item.Type) { $typeName = $item.Type.ToLowerInvariant() }
    if ($item.SkipReason) {
        $results.Add([pscustomobject][ordered]@{ caseId = $id; reportType = $typeName; status = 'skipped'; detail = $item.SkipReason; folder = $null; failedChecks = 0 })
        continue
    }
    Write-Host ''
    Write-Host "[case]   $id ($typeName report)"
    try {
        $report = New-DossierCaseReport -Config $config -Selected $item -AllCases $cases -RunId $runId -RunPath $runPath
        foreach ($check in $report.Checks) {
            $mark = 'FAIL'
            if ($check.Passed) { $mark = 'pass' }
            $detail = ''
            if ($check.Detail) { $detail = " ($($check.Detail))" }
            Write-Host ("[check]  {0,-4} {1}{2}" -f $mark, $check.Name, $detail)
        }
        $failed = @($report.Checks | Where-Object { -not $_.Passed }).Count
        Write-Host "[report] $displayRun/$id/$([IO.Path]::GetFileName($report.Report))"
        $results.Add([pscustomobject][ordered]@{ caseId = $id; reportType = $typeName; status = 'drafted'; detail = [IO.Path]::GetFileName($report.Report); folder = $id; failedChecks = $failed })
    } catch {
        Write-Host "[error]  $id`: $($_.Exception.Message)"
        $results.Add([pscustomobject][ordered]@{ caseId = $id; reportType = $typeName; status = 'failed'; detail = $_.Exception.Message; folder = $null; failedChecks = 0 })
    }
}

# 3. Run summary (not part of any report's audit log)
$summary = [pscustomobject][ordered]@{
    schema = 'complaint-dossier/run/v1'
    runId  = $runId
    mode   = $mode.ToLowerInvariant()
    scope  = $scope
    cases  = $results.ToArray()
}
Write-Utf8File (Join-Path $runPath 'run.json') (($summary | ConvertTo-Json -Depth 5).Replace("`r`n", "`n") + "`n")

$drafted = @($results | Where-Object { $_.status -eq 'drafted' })
$failedCases = @($results | Where-Object { $_.status -eq 'failed' })
Write-Host ''
Write-Host "[done]   $($drafted.Count) report(s) drafted, $(@($results | Where-Object { $_.status -eq 'skipped' }).Count) skipped, $($failedCases.Count) failed. Every report is DRAFT."
foreach ($r in $drafted) {
    if ($r.failedChecks) {
        Write-Host "[next]   $($r.caseId): $($r.failedChecks) automated check(s) failed, so normal approval is refused. Draft it again, or after review: ./Approve-ComplaintDossier.ps1 -Path $displayRun/$($r.caseId) -Reviewer 'Your Name' -AcceptFailedChecks -Justification '<reason>'"
    } else {
        Write-Host "[next]   $($r.caseId): review the report, then ./Approve-ComplaintDossier.ps1 -Path $displayRun/$($r.caseId) -Reviewer 'Your Name'"
    }
}
if ($PassThru) {
    [pscustomobject]@{ RunId = $runId; RunPath = $runPath; Cases = $results.ToArray() }
}
if ($failedCases.Count) { exit 2 }
