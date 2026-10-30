<#
.SYNOPSIS
  Signs a reviewed dossier run: checks the hashes, removes the DRAFT marking,
  records reviewer and time in the report and in audit.json.

.DESCRIPTION
  Run this only after you have checked the AI-drafted section of the report
  against the case records. The draft report stays in the run folder; the
  signed copy is written next to it and its SHA-256 is added to audit.json.
  If prompt, response, statistics or draft report changed after the run,
  approval is refused.

.EXAMPLE
  ./Approve-ComplaintDossier.ps1 -RunPath out/20261030T101500Z-1a2b3c -Reviewer 'A. Reviewer' -Comment 'Checked against case files.'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$RunPath,
    [Parameter(Mandatory)][string]$Reviewer,
    [string]$Comment = ''
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Join-Path $PSScriptRoot 'src') 'ComplaintDossier.psm1') -Force
$result = Approve-DossierRun -RunPath ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RunPath)) -Reviewer $Reviewer -Comment $Comment
Write-Host "[approved] run $($result.RunId) signed by $Reviewer at $($result.SignedAt)"
Write-Host "[report]   $(Get-RelativeDisplayPath $PSScriptRoot $result.Report -FullPathOutside)"
Write-Host "[sha256]   $($result.Sha256)"
