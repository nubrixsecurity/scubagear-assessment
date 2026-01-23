<# 
run-scubagear.ps1
Runner script (called by start-scubagear.ps1):

- Assumes start-scubagear.ps1 already created %TEMP%\scuba and downloaded THIS file there.
- Downloads invoke-scubagear.ps1 + two logos from GitHub into %TEMP%\scuba
- Creates run folder structure under C:\temp\SCuBA\run-<stamp>
- Calls invoke-scubagear.ps1 with -CompanyName and -RunRoot
- Cleans up %TEMP%\scuba at the very end (force)

Important change vs previous version:
- DO NOT delete %TEMP%\scuba at the start, or you'd delete run-scubagear.ps1 while it's executing.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CompanyName,

    # Base folder where run-<stamp> will be created
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$BaseRunPath = "C:\temp\SCuBA",

    # Repo settings
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$RepoOwner = "nubrixsecurity",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$RepoName = "scubagear-assessment",

    # Branch can be main/master/etc.
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$RepoBranch = "main",

    # Script + logo file names in repo root
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$InvokeScriptName = "invoke-scubagear.ps1",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$CisaLogoName = "cisa_logo.png",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$CompanyLogoName = "company_logo.png"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Helpers
function New-RunStamp {
    # run-HHmmss (matches your example like run-192442)
    return ("run-" + (Get-Date).ToString("HHmmss"))
}
#endregion Helpers

#region Paths
$tempRoot   = Join-Path $env:TEMP "scuba"
$invokePath = Join-Path $tempRoot $InvokeScriptName
$cisaPath   = Join-Path $tempRoot $CisaLogoName
$coPath     = Join-Path $tempRoot $CompanyLogoName

$runStamp   = New-RunStamp
$runRoot    = Join-Path $BaseRunPath $runStamp
$runReport  = Join-Path $runRoot "Report"
$runIndiv   = Join-Path $runRoot "IndividualReports"
#endregion Paths

try {
    #region Prepare temp download folder (DO NOT delete it here)
    # start-scubagear.ps1 downloads THIS script into %TEMP%\scuba.
    # If we remove the folder at the beginning, we delete ourselves mid-execution.
    if (-not (Test-Path $tempRoot)) {
        New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null
    }
    #endregion Prepare temp download folder

    #region Create run folder structure in C:\temp\SCuBA\run-<stamp>
    if (-not (Test-Path $BaseRunPath)) {
        New-Item -Path $BaseRunPath -ItemType Directory -Force | Out-Null
    }

    if (Test-Path $runRoot) {
        throw "Run folder already exists (unexpected): $runRoot"
    }
    New-Item -Path $runRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $runReport -ItemType Directory -Force | Out-Null
    New-Item -Path $runIndiv -ItemType Directory -Force | Out-Null
    #endregion Create run folder structure

    #region Download invoke script + logos from GitHub into %TEMP%\scuba
    $baseRaw = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$RepoBranch"
    $invokeUrl = "$baseRaw/$InvokeScriptName"
    $cisaUrl   = "$baseRaw/$CisaLogoName"
    $coUrl     = "$baseRaw/$CompanyLogoName"

    Write-Host "Downloading invoke script: $invokeUrl"
    Invoke-WebRequest -Uri $invokeUrl -OutFile $invokePath -UseBasicParsing

    Write-Host "Downloading CISA logo: $cisaUrl"
    Invoke-WebRequest -Uri $cisaUrl -OutFile $cisaPath -UseBasicParsing

    Write-Host "Downloading company logo: $coUrl"
    Invoke-WebRequest -Uri $coUrl -OutFile $coPath -UseBasicParsing

    if (-not (Test-Path $invokePath)) { throw "Invoke script download failed: $invokePath" }
    if (-not (Test-Path $cisaPath))   { Write-Warning "CISA logo missing after download: $cisaPath (report will omit if not found)" }
    if (-not (Test-Path $coPath))     { Write-Warning "Company logo missing after download: $coPath (report will omit if not found)" }
    #endregion Download

    #region Execute main script
    Write-Host "Starting SCuBA run..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $invokePath `
        -CompanyName $CompanyName `
        -RunRoot $runRoot `
        -CisaLogoPath $cisaPath `
        -CompanyLogoPath $coPath `
        -ExportPdf

    Write-Host "Run completed. Output folder:"
    Write-Host "  $runRoot"
    #endregion Execute main script
}
finally {
    #region Cleanup temp download folder
    # Remove the entire %TEMP%\scuba folder at the end (includes run-scubagear.ps1)
    if (Test-Path $tempRoot) {
        Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
    #endregion Cleanup
}
