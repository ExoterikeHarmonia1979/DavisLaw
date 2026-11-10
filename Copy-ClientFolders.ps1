<#
  Server-side copy/move of top-level client folders between two SharePoint Online
  sites in the same tenant. Uses SharePoint copy/move jobs (no download/upload).

  Folder names come from the "Name" column of FolderNames.xlsx (next to this script,
  or -NamesFile). Rows whose job finishes without errors are deleted from the sheet;
  failed rows stay so a rerun retries only those. Close the file in Excel first.

  Prereqs (PowerShell 7.4+):
    Install-Module PnP.PowerShell -Scope CurrentUser
    Install-Module ImportExcel   -Scope CurrentUser   # reads/writes .xlsx, no Excel needed

  Example:
    ./Copy-ClientFolders.ps1 -SourceSite https://TENANT.sharepoint.com/sites/OldSite `
      -TargetSite https://TENANT.sharepoint.com/sites/Corporate

  Default is COPY (source untouched, only latest version copied).
  Add -Move to keep version history and Modified/Modified By (removes source).
  -TargetLibrary may include a subfolder, e.g. "Shared Documents/Clients".
  Limits per job: 30,000 files / 100 GB. Split bigger folders.
  Delete partial leftovers in the target first (e.g. from a failed browser upload).
#>
param(
  [Parameter(Mandatory)][string]$SourceSite,
  [Parameter(Mandatory)][string]$TargetSite,
  [string]$ClientId = "e95d7cff-fddf-42f3-8e80-06d5a5154f7a",
  [string]$SourceLibrary = "Shared Documents",
  [string]$TargetLibrary = "Shared Documents",
  [string]$NamesFile = (Join-Path $PSScriptRoot "FolderNames.xlsx"),
  [switch]$Move
)

$ErrorActionPreference = "Stop"

# Install required modules only if missing.
foreach ($m in 'PnP.PowerShell', 'ImportExcel') {
  if (-not (Get-Module -ListAvailable $m)) { Install-Module $m -Scope CurrentUser -Force }
}

Import-Module ImportExcel

# Fail before queueing anything if the sheet can't be saved later (e.g. open in Excel).
try { [IO.File]::Open($NamesFile, 'Open', 'ReadWrite', 'None').Close() }
catch { throw "Cannot open '$NamesFile' for writing (missing, or open in Excel?): $_" }

$Folders = Import-Excel $NamesFile | ForEach-Object { "$($_.Name)".Trim() } |
  Where-Object { $_ } | Select-Object -Unique
if (-not $Folders) { Write-Host "No folder names in '$NamesFile'. Nothing to do."; return }

Connect-PnPOnline -Url $SourceSite -ClientId $ClientId -Interactive
$targetPath = ([uri]$TargetSite).AbsolutePath.TrimEnd('/') + "/$TargetLibrary"

$jobs = foreach ($folder in $Folders) {
  $p = @{
    SourceUrl           = "$SourceLibrary/$folder"
    TargetUrl           = $targetPath
    AllowSchemaMismatch = $true
    NoWait              = $true
    Force               = $true
  }
  Write-Host "Queueing $(if ($Move) {'move'} else {'copy'}): $folder"
  try {
    [pscustomobject]@{ Folder = $folder; Job = $(if ($Move) { Move-PnPFile @p } else { Copy-PnPFile @p }) }
  } catch {
    Write-Warning "Failed to queue '$folder' (stays in sheet): $_"
  }
}

$log = Join-Path $PSScriptRoot "copy-log-$(Get-Date -Format yyyyMMdd-HHmmss).json"
$results = foreach ($j in $jobs) {
  Write-Host "Waiting: $($j.Folder)"
  $status = Receive-PnPCopyMoveJobStatus -Job $j.Job -Wait
  $ok = ($status | ConvertTo-Json -Depth 6) -notmatch 'JobError|JobFatalError'
  if (-not $ok) { Write-Warning "'$($j.Folder)' finished with errors (stays in sheet). See log." }
  [pscustomobject]@{ Folder = $j.Folder; Succeeded = $ok; Status = $status }
}
$results | ConvertTo-Json -Depth 6 | Out-File $log

# Delete succeeded rows in place (keeps other columns and formatting).
$done = @($results | Where-Object Succeeded | ForEach-Object Folder)
if ($done) {
  $pkg = Open-ExcelPackage -Path $NamesFile
  try {
    $ws = $pkg.Workbook.Worksheets[1]
    $col = 1..$ws.Dimension.End.Column | Where-Object { "$($ws.Cells[1, $_].Value)".Trim() -eq 'Name' } | Select-Object -First 1
    for ($r = $ws.Dimension.End.Row; $r -ge 2; $r--) {
      if ("$($ws.Cells[$r, $col].Value)".Trim() -in $done) { $ws.DeleteRow($r) }
    }
  } finally { Close-ExcelPackage $pkg }
}

Write-Host "Done. $($done.Count) of $($Folders.Count) folder(s) succeeded and were removed from '$NamesFile'."
Write-Host "Per-folder job logs: $log"
