<#
.SYNOPSIS
  Approves one reviewed complaint report: checks the hashes, removes the DRAFT
  marking, records reviewer and time in the report and in its audit.json.

.DESCRIPTION
  Each case of a run has its own folder <run folder>/<case id> with one report
  and one audit.json. Run this for one report, after you have checked its
  AI-drafted text against the case file. The draft stays in the folder; the
  approved copy is written next to it and its SHA-256 is added to audit.json.
  If case data, prompt, response or draft report changed after the run,
  approval is refused.

  Approval is also refused while any automated check of the report has failed.
  To approve anyway, pass -AcceptFailedChecks and a -Justification. The approved
  report and audit.json then record the failed checks and the justification
  as a documented exception.

.EXAMPLE
  ./Approve-ComplaintDossier.ps1 -Path out/20261030T101500Z-1a2b3c/SYN-2026-007 -Reviewer 'A. Reviewer' -Comment 'Checked against the case file.'

.EXAMPLE
  ./Approve-ComplaintDossier.ps1 -Path out/20261030T101500Z-1a2b3c/SYN-2026-007 -Reviewer 'A. Reviewer' -AcceptFailedChecks -Justification 'SYN-2026-099 is a typo for SYN-2026-009; checked against the complaint list.'
#>
[CmdletBinding()]
param(
    # Folder of one report: <run folder>/<case id>.
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$Reviewer,
    [string]$Comment = '',
    # Approve although automated checks failed. Needs -Justification.
    [switch]$AcceptFailedChecks,
    # Why the failed checks are acceptable. Printed in the approved report and stored in audit.json.
    [string]$Justification = ''
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Join-Path $PSScriptRoot 'src') 'ComplaintDossier.psm1') -Force
$result = Approve-DossierReport -Path ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)) -Reviewer $Reviewer -Comment $Comment `
    -AcceptFailedChecks:$AcceptFailedChecks -Justification $Justification
Write-Host "[approved] $($result.CaseId) (run $($result.RunId)) by $Reviewer at $($result.ApprovedAt)"
if ($result.Exception) {
    Write-Host "[exception] $(@($result.Exception.acceptedFailedChecks).Count) failed check(s) accepted. Justification: $($result.Exception.justification)"
}
Write-Host "[report]   $(Get-RelativeDisplayPath $PSScriptRoot $result.Report -FullPathOutside)"
Write-Host "[sha256]   $($result.Sha256)"
