<#
.SYNOPSIS
  Approves a reviewed dossier run: checks the hashes, removes the DRAFT marking,
  records reviewer and time in the report and in audit.json.

.DESCRIPTION
  Run this only after you have checked the AI-drafted section of the report
  against the case records. The draft report stays in the run folder; the
  approved copy is written next to it and its SHA-256 is added to audit.json.
  If prompt, response, statistics or draft report changed after the run,
  approval is refused.

  Approval is also refused while any automated check of the run has failed.
  To approve anyway, pass -AcceptFailedChecks and a -Justification. The approved
  report and audit.json then record the failed checks and the justification
  as a documented exception.

.EXAMPLE
  ./Approve-ComplaintDossier.ps1 -RunPath out/20261030T101500Z-1a2b3c -Reviewer 'A. Reviewer' -Comment 'Checked against case files.'

.EXAMPLE
  ./Approve-ComplaintDossier.ps1 -RunPath out/20261030T101500Z-1a2b3c -Reviewer 'A. Reviewer' -AcceptFailedChecks -Justification 'Section "Data gaps" is missing; the gaps are listed in the review comment instead.'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$RunPath,
    [Parameter(Mandatory)][string]$Reviewer,
    [string]$Comment = '',
    # Approve although automated checks failed. Needs -Justification.
    [switch]$AcceptFailedChecks,
    # Why the failed checks are acceptable. Printed in the approved report and stored in audit.json.
    [string]$Justification = ''
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Join-Path $PSScriptRoot 'src') 'ComplaintDossier.psm1') -Force
$result = Approve-DossierRun -RunPath ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RunPath)) -Reviewer $Reviewer -Comment $Comment `
    -AcceptFailedChecks:$AcceptFailedChecks -Justification $Justification
Write-Host "[approved] run $($result.RunId) by $Reviewer at $($result.ApprovedAt)"
if ($result.Exception) {
    Write-Host "[exception] $(@($result.Exception.acceptedFailedChecks).Count) failed check(s) accepted. Justification: $($result.Exception.justification)"
}
Write-Host "[report]   $(Get-RelativeDisplayPath $PSScriptRoot $result.Report -FullPathOutside)"
Write-Host "[sha256]   $($result.Sha256)"
