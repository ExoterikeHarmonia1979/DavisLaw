<#
  One-time: registers an Entra ID app for PnP PowerShell interactive login and
  prints its Client ID (for Copy-ClientFolders.ps1 -ClientId).

  Run as a Global Admin / Application Admin (grants admin consent in the browser popup).
  Requires PowerShell 7.4+ and PnP.PowerShell 3.x.

  Example:
    ./Register-PnPApp.ps1 -Tenant dgvisalaw.onmicrosoft.com
#>
param(
  [Parameter(Mandatory)][string]$Tenant,          # e.g. contoso.onmicrosoft.com
  [string]$ApplicationName = "PnP DavisLaw"
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell -MinimumVersion 2.0

# Delegated only: acts as the signed-in user, never more than that user can access.
$out = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName $ApplicationName -Tenant $Tenant `
  -SharePointDelegatePermissions "AllSites.FullControl"

$clientId = [regex]::Match(($out | Format-List | Out-String), '[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}').Value
if (-not $clientId) { $out; throw "Could not find Client ID in output above." }

$clientId | Set-Clipboard
Write-Host "Client ID (copied to clipboard):"
$clientId
