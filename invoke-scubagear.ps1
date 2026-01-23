<#
invoke-scubagear.ps1

Main script:
- Installs/initializes ScubaGear (if missing)
- Runs SCuBA (forces output into %WINDIR%\System32)
- Finds newest M365BaselineConformance_* folder created by the run
- Copies SCuBA outputs into RunRoot (C:\temp\SCuBA\run-<stamp>)
- Builds consolidated HTML report in RunRoot\Report
- Fixes per-report HTML (Light Mode removed, broken CISA logo fixed, results row colors, CAP table flattened)
- Ensures consolidated links resolve by creating RunRoot\Report\IndividualReports
- Optional: Export to PDF (Edge print-to-PDF workflow if you’re using it)

Notes:
- "CAP-only" styling means we only enforce special layout/styling inside Conditional Access Policies section.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$CompanyName,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$RunRoot,

    [Parameter()]
    [string]$CisaLogoPath,

    [Parameter()]
    [string]$CompanyLogoPath,

    [Parameter()]
    [switch]$ExportPdf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Helpers

function Test-PathSafe {
    param([string]$Path)
    try { return (Test-Path -LiteralPath $Path) } catch { return $false }
}

function Get-Base64DataUri {
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )
    if (-not (Test-PathSafe $Path)) { return $null }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $b64 = [Convert]::ToBase64String($bytes)
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    $mime = switch ($ext) {
        ".png"  { "image/png" }
        ".jpg"  { "image/jpeg" }
        ".jpeg" { "image/jpeg" }
        ".gif"  { "image/gif" }
        default { "application/octet-stream" }
    }
    return "data:$mime;base64,$b64"
}

function Ensure-Dir {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-PathSafe $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Copy-DirContents {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$Dest
    )
    Ensure-Dir $Dest
    Get-ChildItem -LiteralPath $Source -Force | ForEach-Object {
        $target = Join-Path $Dest $_.Name
        if ($_.PSIsContainer) {
            Copy-Item -LiteralPath $_.FullName -Destination $target -Recurse -Force
        } else {
            Copy-Item -LiteralPath $_.FullName -Destination $target -Force
        }
    }
}

function Read-FileText {
    param([Parameter(Mandatory=$true)][string]$Path)
    return [System.IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
}

function Write-FileText {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Content
    )
    [System.IO.File]::WriteAllText($Path, $Content, [Text.Encoding]::UTF8)
}

#endregion Helpers

#region Paths
Ensure-Dir $RunRoot

$reportRoot    = Join-Path $RunRoot "Report"
$indivRoot     = Join-Path $RunRoot "IndividualReports"
$reportIndiv   = Join-Path $reportRoot "IndividualReports"

Ensure-Dir $reportRoot

# consolidated filenames
$todayShort = (Get-Date).ToString("yyyy-MM-dd")
$consolidatedHtml = Join-Path $reportRoot ("SCuBA-Report-$todayShort.html")
$consolidatedPdf  = Join-Path $reportRoot ("SCuBA-Report-$todayShort.pdf")
#endregion Paths

#region Install + init ScubaGear
if (-not (Get-Module -ListAvailable -Name ScubaGear)) {
    Write-Host "Installing ScubaGear from PSGallery..."
    Install-Module -Name ScubaGear -Force -Scope CurrentUser
}

Write-Host "Initializing SCuBA dependencies..."
Initialize-SCuBA

Write-Host "SCuBA version:"
Invoke-SCuBA -Version | Out-Host
#endregion Install + init ScubaGear

#region Run SCuBA (force output to System32, prevents Documents output)
$system32 = "$env:WINDIR\System32"
$runStart = Get-Date
$priorLocation = Get-Location

try {
    Set-Location -Path $system32

    Write-Host "Running SCuBA assessment (all products)..."
    Invoke-SCuBA -ProductNames * -Quiet
}
finally {
    Set-Location -Path $priorLocation
}
#endregion Run SCuBA

#region Locate output folder in System32 (newest since run start)
Write-Host "Searching System32 for newest SCuBA output folder: M365BaselineConformance_*"

$matches = Get-ChildItem -Path $system32 -Directory -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -like "M365BaselineConformance_*" -and $_.LastWriteTime -ge $runStart.AddMinutes(-5)
    } |
    Sort-Object LastWriteTime -Descending

if (-not $matches -or $matches.Count -eq 0) {
    # fallback: pick newest overall
    $matches = Get-ChildItem -Path $system32 -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "M365BaselineConformance_*" } |
        Sort-Object LastWriteTime -Descending
}

if (-not $matches -or $matches.Count -eq 0) {
    throw "No SCuBA output folder found in $system32 matching: M365BaselineConformance_*"
}

$sourceFolder = $matches | Select-Object -First 1
Write-Host "Using SCuBA output folder: $($sourceFolder.FullName)"
#endregion Locate output

#region Copy SCuBA outputs into RunRoot
# This preserves SCuBA’s folder structure under your run-stamp root.
Copy-DirContents -Source $sourceFolder.FullName -Dest $RunRoot

# If SCuBA created IndividualReports at run root, make sure it exists
if (-not (Test-PathSafe $indivRoot)) {
    # Some SCuBA versions output individual reports elsewhere; try to find them
    $possible = Get-ChildItem -Path $RunRoot -Directory -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq "IndividualReports" } |
        Sort-Object FullName |
        Select-Object -First 1

    if ($possible) {
        $indivRoot = $possible.FullName
    }
}

# IMPORTANT: create Report\IndividualReports copy so consolidated links work
# (This fixes your "File not found" when clicking DefenderReport.html etc.)
if (Test-PathSafe $indivRoot) {
    Ensure-Dir $reportIndiv
    Copy-DirContents -Source $indivRoot -Dest $reportIndiv
}
#endregion Copy outputs

#region Build consolidated report HTML (with embedded logos, correct links)
$cisaData = $null
$coData   = $null

if ($CisaLogoPath)   { $cisaData = Get-Base64DataUri -Path $CisaLogoPath }
if ($CompanyLogoPath){ $coData   = Get-Base64DataUri -Path $CompanyLogoPath }

# Pull baseline summary from existing SCuBA HTML if present
$baselineReportsHtml = Join-Path $RunRoot "BaselineReports.html"

$baselineSection = ""
if (Test-PathSafe $baselineReportsHtml) {
    $baselineSection = Read-FileText $baselineReportsHtml
} else {
    $baselineSection = "<p>BaselineReports.html not found in run output.</p>"
}

# Replace links so they point to Report\IndividualReports (relative to consolidated html in Report)
# SCuBA sometimes links like: IndividualReports\DefenderReport.html or ./IndividualReports/...
$baselineSection = $baselineSection `
    -replace '(?i)href\s*=\s*"(?:\./)?IndividualReports[/\\]', 'href="IndividualReports/' `
    -replace '(?i)href\s*=\s*"(?:\.\./)?IndividualReports[/\\]', 'href="IndividualReports/'

# Convert baseline summary "Details" cell lines into bullets + tighten spacing (only on summary table)
# (Lightweight heuristic: replace <br> separated counts in Details column with <ul><li>..</li>)
$baselineSection = $baselineSection -replace '(?is)<td([^>]*)>\s*(\d+\s+passes?)<br\s*/?>\s*(\d+\s+warnings?)<br\s*/?>\s*(\d+\s+failures?)<br\s*/?>\s*(\d+\s+manual\s+checks?)\s*</td>',
'<td$1><ul class="tight-bullets"><li>$2</li><li>$3</li><li>$4</li><li>$5</li></ul></td>'

# Header / logos layout
$logoRow = @"
<div class="top-logos">
  <div class="logo-left">
    $(if ($cisaData) { "<img class='logo-cisa' src='$cisaData' alt='CISA Logo' />" } else { "" })
  </div>
  <div class="logo-right">
    $(if ($coData) { "<img class='logo-company' src='$coData' alt='Company Logo' />" } else { "" })
  </div>
</div>
"@

$consolidated = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>SCuBA Report - $todayShort</title>
  <style>
    /* Global: force white background in PDF */
    html, body { background: #ffffff !important; color: #000; margin: 0; padding: 0; }
    @page { margin: 18mm; }

    /* Simple header / footer print feel */
    .doc { padding: 22px 26px; }
    h1, h2 { color: #005a9e; }

    .top-logos { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom: 12px; }
    .logo-cisa { width: 110px; height: 110px; object-fit: contain; } /* square, not distorted */
    .logo-company { height: 48px; object-fit: contain; } /* will appear larger than before */

    .tight-bullets { margin: 0; padding-left: 18px; }
    .tight-bullets li { margin: 0; line-height: 1.15; }

    /* Remove any Light Mode label or toggle anywhere */
    .toggle, .toggle-switch, .toggle-container, #toggle, #toggleSwitch { display:none !important; }
    .toggle-label, .toggleText, .mode-label, .modeText { display:none !important; }
    /* Specific text node fallback: hide elements that contain "Light mode" if wrapped */
    .light-mode, #lightMode { display:none !important; }

    /* Result row coloring across whole row */
    tr.result-pass td { background:#e6f4e6 !important; }     /* light green */
    tr.result-fail td { background:#f7dede !important; }     /* light red */
    tr.result-warning td { background:#fff6d6 !important; }  /* light yellow */
    tr.result-na td { background:#efefef !important; }       /* light gray */

    /* Conditional Access Policies: CAP-only styling */
    .cap-only { }
    .cap-only table { table-layout: fixed; width: 100%; }
    .cap-only th, .cap-only td { word-wrap: break-word; vertical-align: top; }
    /* Make Users + Conditions narrower so Session Controls fits */
    .cap-only col.col-users { width: 18%; }
    .cap-only col.col-conditions { width: 18%; }
    .cap-only col.col-name { width: 18%; }
    .cap-only col.col-state { width: 6%; }
    .cap-only col.col-apps { width: 16%; }
    .cap-only col.col-block { width: 14%; }
    .cap-only col.col-session { width: 10%; }

    /* Remove expand/collapse buttons and chevron column if present */
    .cap-only .expandButtons, .cap-only .expand-buttons, .cap-only .controls { display:none !important; }
    .cap-only button { display:none !important; }
  </style>
</head>
<body>
  <div class="doc">
    $logoRow

    <h1>SCuBA M365 Secure Configuration Baseline Assessment</h1>
    <div>Consolidated report package</div>

    <div style="margin-top:14px;">
      <div><b>Client:</b> $CompanyName</div>
      <div><b>Report date:</b> $todayShort</div>
    </div>

    <h2 style="margin-top:24px;">Baseline Conformance Summary</h2>

    $baselineSection
  </div>
</body>
</html>
"@

Write-FileText -Path $consolidatedHtml -Content $consolidated
#endregion Build consolidated report

#region Fix individual reports (Light Mode removal, results row coloring, CISA logo, CAP flatten)
if (Test-PathSafe $reportIndiv) {

    $htmlFiles = Get-ChildItem -Path $reportIndiv -Filter *.html -File -ErrorAction SilentlyContinue

    foreach ($f in $htmlFiles) {
        $t = Read-FileText $f.FullName

        # remove "Light mode" label + toggle block if present
        $t = $t -replace '(?is)<div[^>]*>\s*Light\s*mode\s*</div>', ''
        $t = $t -replace '(?is)Light\s*mode', ''  # blunt fallback

        # fix broken CISA logo references (replace any CISA image with embedded CISA)
        if ($cisaData) {
            $t = $t -replace '(?is)<img[^>]+alt="Cybersecurity and Infrastructure Security Agency Logo"[^>]*>', "<img src='$cisaData' alt='CISA Logo' style='width:110px;height:110px;object-fit:contain;'/>"
            $t = $t -replace '(?is)<img[^>]+alt="CISA Logo"[^>]*>', "<img src='$cisaData' alt='CISA Logo' style='width:110px;height:110px;object-fit:contain;'/>"
        }

        # apply result row class based on Result cell text
        # This is heuristic but works well for SCuBA’s tables
        $t = $t -replace '(?is)<tr>(\s*<td[^>]*>[^<]*</td>\s*<td[^>]*>[^<]*</td>\s*<td[^>]*>\s*Pass\s*</td>)', '<tr class="result-pass">$1'
        $t = $t -replace '(?is)<tr>(\s*<td[^>]*>[^<]*</td>\s*<td[^>]*>[^<]*</td>\s*<td[^>]*>\s*Fail\s*</td>)', '<tr class="result-fail">$1'
        $t = $t -replace '(?is)<tr>(\s*<td[^>]*>[^<]*</td>\s*<td[^>]*>[^<]*</td>\s*<td[^>]*>\s*Warning\s*</td>)', '<tr class="result-warning">$1'
        $t = $t -replace '(?is)<tr>(\s*<td[^>]*>[^<]*</td>\s*<td[^>]*>[^<]*</td>\s*<td[^>]*>\s*N/?A\s*</td>)', '<tr class="result-na">$1'

        # Conditional Access Policies section:
        # - remove expand/collapse buttons
        # - remove chevron/expand column if present
        # - force “expanded view” by removing JS-driven collapsers and leaving full content visible
        if ($t -match '(?is)Conditional Access Policies') {

            # remove expand/collapse button blocks
            $t = $t -replace '(?is)<div[^>]*>\s*<button[^>]*>\s*\+\s*Expand all.*?</div>', ''
            $t = $t -replace '(?is)<div[^>]*>\s*<button[^>]*>\s*-\s*Collapse all.*?</div>', ''
            $t = $t -replace '(?is)<button[^>]*>\s*\+\s*Expand all\s*</button>', ''
            $t = $t -replace '(?is)<button[^>]*>\s*-\s*Collapse all\s*</button>', ''

            # remove chevron column header/cells (common patterns)
            $t = $t -replace '(?is)<th[^>]*>\s*</th>', ''  # empty header used for chevron col
            $t = $t -replace '(?is)<td[^>]*>\s*(?:&gt;|>|<span[^>]*class="[^"]*(?:chevron|arrow)[^"]*"[^>]*>.*?</span>)\s*</td>', ''

            # wrap CAP table in cap-only container and force columns via colgroup
            # Insert a colgroup after <table> for CAP tables that have these headers
            if ($t -match '(?is)<table[^>]*>.*?<th[^>]*>\s*Name\s*</th>.*?<th[^>]*>\s*Session\s*Controls\s*</th>') {
                $t = $t -replace '(?is)(<table[^>]*>)', @"
<div class="cap-only">
$1
<colgroup>
  <col class="col-name">
  <col class="col-state">
  <col class="col-users">
  <col class="col-apps">
  <col class="col-conditions">
  <col class="col-block">
  <col class="col-session">
</colgroup>
"@
                # close wrapper before </body>
                $t = $t -replace '(?is)</table>', '</table></div>'
            }
        }

        # Ensure our CSS exists in head (and NO raw JS printed)
        if ($t -match '(?is)</head>') {
            $injectCss = @"
<style>
  html, body { background:#ffffff !important; }
  .toggle, .toggle-switch, .toggle-container { display:none !important; }
  tr.result-pass td { background:#e6f4e6 !important; }
  tr.result-fail td { background:#f7dede !important; }
  tr.result-warning td { background:#fff6d6 !important; }
  tr.result-na td { background:#efefef !important; }
</style>
"@
            $t = $t -replace '(?is)</head>', ($injectCss + "`n</head>")
        }

        Write-FileText -Path $f.FullName -Content $t
    }
}
#endregion Fix individual reports

#region Export PDF (optional)
# NOTE: Your environment/tooling decides how you export.
# If you already have a working PDF export method, keep using it.
# This block is a placeholder to show where it would occur.
if ($ExportPdf) {
    Write-Host "ExportPdf switch set. Consolidated HTML created at:"
    Write-Host "  $consolidatedHtml"
    Write-Host "If you have an automated HTML->PDF method, call it here to output:"
    Write-Host "  $consolidatedPdf"
}
#endregion Export PDF

Write-Host "Done. Consolidated report:"
Write-Host "  $consolidatedHtml"
Write-Host "Individual reports (linked correctly under Report\IndividualReports):"
Write-Host "  $reportIndiv"
