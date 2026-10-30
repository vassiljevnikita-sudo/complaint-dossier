<#
.SYNOPSIS
  Generates a draft complaint report for every complaint of the Excel list
  that has no report yet, in your own Word template. Run.cmd starts this.

.DESCRIPTION
  Reads config.json, checks it (the same checks as Setup.cmd), selects the
  complaints received in the configured period and generates one Word report
  per complaint, one after another. Complaints with a report from an earlier
  run are skipped, so an interrupted run continues where it stopped.

  Every report is a draft. Review happens afterwards in Word: the reviewer
  checks each report against the case file and finalises it as the quality
  system requires. The run writes out/review-list.xlsx with one row per report,
  its automated check results and empty columns for reviewer and date.

.EXAMPLE
  .\Run.cmd
  Live mode with config.json: Microsoft 365 Copilot writes the report text.

.EXAMPLE
  .\Run.cmd -Demo
  Demo mode with config.example.json and prewritten answers. Needs no Copilot.

.EXAMPLE
  ./Invoke-ComplaintDossier.ps1 -CaseId SYN-2026-007 -Regenerate
  Generates the report of one complaint again. The earlier report and its
  audit log move to out/replaced/.
#>
[CmdletBinding()]
param(
    # Use the prewritten sample answers instead of Copilot (default config: config.example.json).
    [switch]$Demo,
    # Hand each prompt to the Microsoft 365 Copilot desktop app. This is the default without -Demo.
    [switch]$Live,
    # Default: config.json, or config.example.json with -Demo.
    [string]$ConfigPath,
    # Only these complaints, whatever their date and status.
    [string[]]$CaseId,
    # Generate the report again even if one exists. The old one moves to out/replaced/.
    [switch]$Regenerate,
    # Override Selection.From / Selection.To (yyyy-MM-dd).
    [string]$From,
    [string]$To,
    # Output folder. Default: Output.Path of the configuration.
    [string]$OutputPath,
    # Do not open the reports folder at the end.
    [switch]$NoOpen,
    # Return the run result as an object.
    [switch]$PassThru
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = $PSScriptRoot
Import-Module (Join-Path (Join-Path $root 'src') 'ComplaintDossier.psm1') -Force

if ($Demo -and $Live) { throw 'Choose one mode: -Demo or -Live.' }
# Run.cmd passes "-CaseId A,B" as one text; split it into complaint numbers.
if ($CaseId) { $CaseId = @($CaseId | ForEach-Object { $_ -split '[,;\s]+' } | Where-Object { $_ }) }
$mode = 'Live'
if ($Demo) { $mode = 'Demo' }
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $root 'config.json'
    if ($Demo) { $ConfigPath = Join-Path $root 'config.example.json' }
}
# Resolve against the PowerShell location; .NET would use the process directory.
$ConfigPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ConfigPath)
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf) -and -not $Demo) {
    Set-DossierLanguage (Get-DefaultLanguage)
    Write-Host (Get-DossierText 'RunNoConfig')
    exit 1
}
$config = Read-DossierConfig -Path $ConfigPath -Root $root -Mode $mode -OutputPath $OutputPath
function Say([string]$Key, [object[]]$Arguments = @()) { Write-Host (Get-DossierText $Key -f $Arguments) }
$display = { param($path) Get-RelativeDisplayPath $root $path -FullPathOutside }

Write-Host ''
Say 'RunTitle' @((Get-DossierVersion), (Get-DossierText ('Mode' + $mode)))

# 1. Same checks as Setup. Errors stop the run before anything is generated.
$findings = @(Test-DossierSetup -Config $config -From $From -To $To)
$errors = @($findings | Where-Object { $_.Level -eq 'error' })
if ($errors.Count) {
    Say 'RunBlocked' @($errors.Count)
    foreach ($e in $errors) {
        Write-Host "  - $($e.Text)"
        if ($e.Fix) { Write-Host "    $($e.Fix)" }
    }
    Say 'RunSetupHint'
    exit 1
}
if ($mode -eq 'Live') {
    if ($env:OS -ne 'Windows_NT') { Say 'LiveNeedsWindows'; exit 1 }
    if (-not (Test-CopilotWindowOpen ([string]$config.Settings.Copilot.WindowTitle))) { Say 'CopilotNotOpen' @($config.Settings.Copilot.WindowTitle); exit 1 }
}

# 2. Complaint list and selection
$cases = @(Import-DossierCases -Config $config)
$period = Get-SelectionPeriod $config $From $To
if ($CaseId) { $selection = @(Select-DossierCases -Config $config -Cases $cases -CaseId $CaseId -Regenerate:$Regenerate) }
else { $selection = @(Select-DossierCases -Config $config -Cases $cases -From $period.From -To $period.To -Regenerate:$Regenerate) }
$todo = @($selection | Where-Object { $_.State -eq 'todo' })
$done = @($selection | Where-Object { $_.State -eq 'done' })
$skip = @($selection | Where-Object { $_.State -eq 'skip' })
Say 'RunList' @((& $display $config.Paths.Workbook), $config.Settings.Workbook.Sheet, $cases.Count)
if ($CaseId) { Say 'RunScopeIds' @(($CaseId -join ', ')) }
else {
    $fromText = $period.From; if (-not $fromText) { $fromText = '...' }
    $toText = $period.To; if (-not $toText) { $toText = '...' }
    Say 'RunScopePeriod' @($fromText, $toText, $selection.Count)
}
Say 'RunPlan' @($todo.Count, $done.Count, $skip.Count)
foreach ($group in @($skip | Group-Object Reason)) { Say 'RunSkipGroup' @($group.Count, $group.Name, ((@($group.Group | ForEach-Object { $_.Case.CaseId })) -join ', ')) }
Say 'RunWhere' @((& $display $config.Paths.Reports))

$runId = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ') + '-' + ([guid]::NewGuid().ToString('N').Substring(0, 6))
$results = New-Object System.Collections.Generic.List[object]
foreach ($item in @($done + $skip)) {
    $results.Add([pscustomobject][ordered]@{ caseId = $item.Case.CaseId; receivedDate = $item.Case.ReceivedDate.ToString('yyyy-MM-dd'); reportType = $null; state = $item.State; detail = $item.Reason; report = $null; failedChecks = 0 })
}

# 3. One report per complaint, one after another (the live path drives one
#    Copilot window). A failed complaint does not stop the others; three
#    failures in a row stop the run, because then usually Copilot is not usable.
$stopAfter = 3
$runSettings = Get-ConfigValue $config.Settings 'Run'
if ($null -ne (Get-ConfigValue $runSettings 'StopAfterFailuresInARow')) { $stopAfter = [int]$runSettings.StopAfterFailuresInARow }
$inARow = 0
$stopped = $false
$started = Get-Date
$n = 0
if ($todo.Count) { Write-Host '' }
foreach ($item in $todo) {
    $n++
    $id = $item.Case.CaseId
    $typeWidth = [math]::Max((Get-DossierText 'TypeShort').Length, (Get-DossierText 'TypeLong').Length)
    $typeText = (Get-DossierText ('Type' + $item.Type)).PadRight($typeWidth)
    $format = '[{0,' + ([string]$todo.Count).Length + '}/{1}] {2}  {3}'
    $line = $format -f $n, $todo.Count, $id, $typeText
    Write-Host -NoNewline "$line ... "
    $entry = [pscustomobject][ordered]@{ caseId = $id; receivedDate = $item.Case.ReceivedDate.ToString('yyyy-MM-dd'); reportType = $item.Type.ToLowerInvariant(); state = ''; detail = ''; report = $null; failedChecks = 0 }
    try {
        if ($Regenerate) { Move-GeneratedReport -Config $config -CaseId $id }
        $report = New-DossierCaseReport -Config $config -Selected $item -AllCases $cases -RunId $runId
        $failed = @($report.Checks | Where-Object { -not $_.Passed })
        $entry.state = 'generated'
        $entry.report = 'reports/' + [IO.Path]::GetFileName($report.Report)
        $entry.failedChecks = $failed.Count
        if ($failed.Count) {
            Write-Host (Get-DossierText 'CaseDoneFailedChecks' -f $failed.Count, $report.Checks.Count, ((@($failed | ForEach-Object { Get-CheckDisplayName $_ })) -join '; ')) -ForegroundColor Yellow
        } else {
            Write-Host (Get-DossierText 'CaseDone' -f $report.Checks.Count) -ForegroundColor Green
        }
        $inARow = 0
    } catch {
        $entry.state = 'failed'
        $entry.detail = $_.Exception.Message
        Write-Host (Get-DossierText 'CaseFailed' -f $_.Exception.Message) -ForegroundColor Red
        $inARow++
    }
    $results.Add($entry)
    if ($mode -eq 'Live' -and $n -lt $todo.Count) {
        $perCase = ((Get-Date) - $started).TotalSeconds / $n
        Say 'TimeLeft' @([math]::Ceiling($perCase * ($todo.Count - $n) / 60))
    }
    if ($inARow -ge $stopAfter -and $n -lt $todo.Count) {
        Say 'StoppedAfterFailures' @($inARow)
        $stopped = $true
        foreach ($rest in @($todo | Select-Object -Skip $n)) {
            $results.Add([pscustomobject][ordered]@{ caseId = $rest.Case.CaseId; receivedDate = $rest.Case.ReceivedDate.ToString('yyyy-MM-dd'); reportType = $rest.Type.ToLowerInvariant(); state = 'not-started'; detail = (Get-DossierText 'NotStarted'); report = $null; failedChecks = 0 })
        }
        break
    }
}

# 4. Run summary, review list
$generated = @($results | Where-Object { $_.state -eq 'generated' })
$failedCases = @($results | Where-Object { $_.state -eq 'failed' })
New-Item -ItemType Directory -Force -Path $config.Paths.Runs | Out-Null
Write-JsonFile (Join-Path $config.Paths.Runs "$runId.json") ([ordered]@{
        schema   = 'complaint-dossier/run/v2'
        runId    = $runId
        mode     = $mode.ToLowerInvariant()
        period   = [ordered]@{ from = $period.From; to = $period.To }
        caseIds  = @($CaseId)
        stopped  = $stopped
        complaints = $results.ToArray()
    })
$notGenerated = @($results | Where-Object { $_.state -in 'skip', 'failed', 'not-started' } | ForEach-Object {
        [pscustomobject]@{ CaseId = $_.caseId; ReceivedDate = $_.receivedDate; State = (Get-DossierText ('State_' + $_.state.Replace('-', ''))); Reason = $_.detail }
    })
$reviewList = Update-ReviewList -Config $config -NotGenerated $notGenerated

Write-Host ''
Say 'Summary' @($generated.Count, $failedCases.Count, $done.Count, $skip.Count)
$withFailedChecks = @($generated | Where-Object { $_.failedChecks -gt 0 })
if ($withFailedChecks.Count) { Say 'SummaryFailedChecks' @($withFailedChecks.Count, ((@($withFailedChecks | ForEach-Object { $_.caseId })) -join ', ')) }
if ($failedCases.Count) { Say 'SummaryFailedCases' @(((@($failedCases | ForEach-Object { $_.caseId })) -join ', ')) }
Say 'SummaryReports' @((& $display $config.Paths.Reports))
Say 'SummaryReviewList' @((& $display $reviewList))
if ($generated.Count) { Say 'NextStep' }
elseif (-not $todo.Count -and $done.Count) { Say 'NothingToDo' }

if (-not $NoOpen -and -not $PassThru -and $generated.Count -and $env:OS -eq 'Windows_NT' -and -not $env:CI -and [Environment]::UserInteractive) {
    try { Invoke-Item -LiteralPath $config.Paths.Reports } catch { }
}
if ($PassThru) {
    [pscustomobject]@{ RunId = $runId; OutputPath = $config.Paths.Output; ReviewList = $reviewList; Complaints = $results.ToArray() }
}
if ($failedCases.Count -or $stopped) { exit 2 }
