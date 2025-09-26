param(
  [Parameter(Mandatory=$true)][ValidateSet("test","prod")]
  [string]$Env
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoPath = "C:\Projects\pm_automation_v1"
Set-Location $RepoPath
[Environment]::CurrentDirectory = $RepoPath

# Pick the right env file
$ConfFile = if ($Env -eq "test") { ".clasp.test.json" } else { ".clasp.prod.json" }
$Active   = ".clasp.json"

if (-not (Test-Path $ConfFile)) { throw "Missing $ConfFile" }
Copy-Item -LiteralPath $ConfFile -Destination $Active -Force

# Read active config
$conf = Get-Content -LiteralPath $Active -Raw | ConvertFrom-Json
$scriptId = $conf.scriptId
$rootRel  = $conf.rootDir

# Convert clasp-style rootDir to a Windows path
function Convert-RootDirToPath([string]$root) {
  $rel = ($root -replace '^\.\/','' -replace '/','\')
  if ([string]::IsNullOrWhiteSpace($rel)) { return $RepoPath }
  return (Join-Path -Path $RepoPath -ChildPath $rel)
}
$codePath = Convert-RootDirToPath $rootRel

# If the configured root does not contain appsscript.json, auto-detect
if (-not (Test-Path (Join-Path -Path $codePath -ChildPath 'appsscript.json'))) {
  $found = Get-ChildItem -Path $RepoPath -Recurse -Filter "appsscript.json" | Select-Object -First 1
  if (-not $found) { throw "Could not find appsscript.json anywhere in the repo." }
  $codePath = Split-Path -Path $found.FullName -Parent
}

Write-Host "Environment: $Env"
Write-Host "Target Script ID: $scriptId"
Write-Host "Code root: $codePath"

# Get latest code
git fetch --all
git checkout main
git pull --ff-only origin main

# Push to Apps Script
clasp push --force

Write-Host ("{0} deploy complete." -f $Env.ToUpper())