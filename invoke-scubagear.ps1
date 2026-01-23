<#
invoke-scubagear.ps1
Main script:
- Ensures ScubaGear is installed
- Initializes dependencies (Initialize-SCuBA)
- Runs Invoke-SCuBA -ProductNames *
- Finds System32 output folder using M365BaselineConformance_yyyy_MM_dd*
- Copies output into RunRoot (and IndividualReports folder if present)
- Builds a branded consolidated HTML (cover + contents + baseline summary + embedded sections)
- Applies CAP-only fixed-table CSS (cap-fixed class)
- Fixes broken CISA logo in individual reports to local cisa_logo.png
- Removes Light mode label + toggle
- Colors rows (Pass/Fail/Warning/Manual) across entire row
- Flattens Conditional Access Policies (removes expand/collapse UI, ensures all content visible)
- Exports PDF using Edge headless into RunRoot\Report
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CompanyName,

    # Run root: C:\temp\SCuBA\run-<stamp>
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$RunRoot,

    # Downloaded to %TEMP%\scuba\*.png by runner
    [Parameter()]
    [string]$CisaLogoPath,

    [Parameter()]
    [string]$CompanyLogoPath,

    [switch]$ExportPdf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Helpers
function Ensure-Folder([string]$Path) {
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Get-EdgePath {
    $edgeCandidates = @(
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
    )

    $edge = $edgeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($edge) { return $edge }

    # fallback: try App Paths registry
    $reg = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"
    if (Test-Path $reg) {
        $p = (Get-ItemProperty $reg)."(default)"
        if ($p -and (Test-Path $p)) { return $p }
    }

    return $null
}

function Read-FileText([string]$Path) {
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
}

function Write-FileText([string]$Path, [string]$Content) {
    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::UTF8)
}

function Escape-Html([string]$s) {
    if ($null -eq $s) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($s)
}

function Get-TodayToken {
    return (Get-Date).ToString("yyyy_MM_dd")
}

function Get-ReportDateIso {
    return (Get-Date).ToString("yyyy-MM-dd")
}

function Strip-DoctypeHeadBody([string]$html) {
    # crude extraction of body inner html
    if ($html -match '(?is)<body[^>]*>(.*)</body>') {
        return $Matches[1]
    }
    return $html
}

#endregion Helpers

#region Validate / create run structure
if (-not (Test-Path $RunRoot)) { throw "RunRoot not found: $RunRoot" }

$reportFolder = Join-Path $RunRoot "Report"
$indivFolder  = Join-Path $RunRoot "IndividualReports"
Ensure-Folder $reportFolder
Ensure-Folder $indivFolder

#endregion Validate / create run structure

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

#region Run assessment (force output to System32)
Write-Host "Running SCuBA assessment (all products)..."

$system32 = "$env:WINDIR\System32"
$priorLocation = Get-Location

try {
    Set-Location -Path $system32

    # Run assessment (drops M365BaselineConformance_* into System32)
    Invoke-SCuBA -ProductNames * -Quiet
}
finally {
    Set-Location -Path $priorLocation
}
#endregion Run assessment

#region Locate output folder in System32 (latest run)
$system32 = "$env:WINDIR\System32"
Write-Host "Searching System32 for latest SCuBA output folder: M365BaselineConformance_*"

$matches = Get-ChildItem -Path $system32 -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "M365BaselineConformance_*" } |
    Sort-Object LastWriteTime -Descending

if (-not $matches -or $matches.Count -eq 0) {
    throw "No SCuBA output folder found in $system32 matching: M365BaselineConformance_*"
}

$sourceFolder = $matches | Select-Object -First 1
Write-Host "Using SCuBA output folder: $($sourceFolder.FullName)"
#endregion Locate output

#region Copy output into RunRoot
# Copy files in root
Get-ChildItem -Path $sourceFolder.FullName -File -ErrorAction SilentlyContinue | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination (Join-Path $RunRoot $_.Name) -Force
}

# Copy IndividualReports if present
$srcIndiv = Join-Path $sourceFolder.FullName "IndividualReports"
if (Test-Path $srcIndiv) {
    Copy-Item -Path (Join-Path $srcIndiv "*") -Destination $indivFolder -Recurse -Force
}

Write-Host "Copied SCuBA output into: $RunRoot"
#endregion Copy output

#region Identify baseline + individual report files
$baselinePath = Join-Path $RunRoot "BaselineReports.html"
if (-not (Test-Path $baselinePath)) {
    # sometimes file name varies; fallback find
    $baselinePath = (Get-ChildItem -Path $RunRoot -Filter "*Baseline*Reports*.html" -File -ErrorAction SilentlyContinue |
        Select-Object -First 1).FullName
}
if (-not $baselinePath -or -not (Test-Path $baselinePath)) {
    throw "BaselineReports.html not found in RunRoot: $RunRoot"
}

$indivHtml = Get-ChildItem -Path $indivFolder -Filter "*.html" -File -ErrorAction SilentlyContinue
if (-not $indivHtml -or $indivHtml.Count -eq 0) {
    throw "No individual HTML reports found in: $indivFolder"
}

#endregion Identify baseline + individual report files

#region Parse tenant info from baseline
$baselineHtml = Read-FileText $baselinePath

# Tenant Display Name
$tenantDisplayName = $null
if ($baselineHtml -match '(?is)<th[^>]*>\s*Tenant Display Name\s*</th>\s*<th[^>]*>\s*Tenant Domain Name\s*</th>.*?<td[^>]*>\s*(?<tname>.*?)\s*</td>') {
    $tenantDisplayName = ($Matches['tname'] -replace '<[^>]+>', '').Trim()
}

# Tenant Domain Name
$tenantDomainName = $null
if ($baselineHtml -match '(?is)<th[^>]*>\s*Tenant Domain Name\s*</th>.*?<td[^>]*>\s*(?<tdom>.*?)\s*</td>') {
    $tenantDomainName = ($Matches['tdom'] -replace '<[^>]+>', '').Trim()
}

# Tenant ID
$tenantId = $null
if ($baselineHtml -match '(?is)<th[^>]*>\s*Tenant ID\s*</th>.*?<td[^>]*>\s*(?<tid>.*?)\s*</td>') {
    $tenantId = ($Matches['tid'] -replace '<[^>]+>', '').Trim()
}

if (-not $tenantDisplayName) { $tenantDisplayName = "Unknown Tenant" }
if (-not $tenantDomainName)  { $tenantDomainName  = "Unknown Domain" }
if (-not $tenantId)          { $tenantId          = "Unknown Tenant ID" }

#endregion Parse tenant info

#region Build master HTML (branded)
$reportDateIso = Get-ReportDateIso
$reportDateUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss 'UTC'")

$masterHtmlPath = Join-Path $reportFolder ("SCuBA-Report-{0}.html" -f $reportDateIso)
$pdfPath        = Join-Path $reportFolder ("SCuBA Report - {0}.pdf" -f $reportDateIso)

# Only show logos if files exist
$hasCisaLogo = ($CisaLogoPath -and (Test-Path $CisaLogoPath))
$hasCoLogo   = ($CompanyLogoPath -and (Test-Path $CompanyLogoPath))

# Copy logos into Report folder so HTML can reference stable relative paths
# (PDF print is more reliable with local relative files)
$relCisa = $null
$relCo   = $null
if ($hasCisaLogo) {
    $dest = Join-Path $reportFolder "cisa_logo.png"
    Copy-Item $CisaLogoPath -Destination $dest -Force
    $relCisa = "cisa_logo.png"
}
if ($hasCoLogo) {
    $dest = Join-Path $reportFolder "company_logo.png"
    Copy-Item $CompanyLogoPath -Destination $dest -Force
    $relCo = "company_logo.png"
}

# Build contents list from known product names
# map: display title -> expected report file (best-effort)
$sections = @(
    @{ Key="Baseline"; Title="Baseline Conformance Summary"; Anchor="baseline-summary" }
)

# Derive per-report friendly names (we keep simple ordering from your baseline table)
# We'll embed all individual HTMLs, with anchors based on filename
$indivSections = @()
foreach ($f in $indivHtml) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)

    # Make nicer title heuristics
    $title = $name
    if ($title -match 'AAD') { $title = "Azure Active Directory Baseline Report" }
    elseif ($title -match 'Defender') { $title = "Microsoft 365 Defender Baseline Report" }
    elseif ($title -match 'EXO|Exchange') { $title = "Exchange Online Baseline Report" }
    elseif ($title -match 'PowerPlatform') { $title = "Microsoft Power Platform Baseline Report" }
    elseif ($title -match 'SharePoint') { $title = "SharePoint Online Baseline Report" }
    elseif ($title -match 'Teams') { $title = "Microsoft Teams Baseline Report" }

    $anchor = ("r-" + ($name -replace '[^a-zA-Z0-9\-]+','-').Trim('-')).ToLowerInvariant()
    $indivSections += @{ File=$f.FullName; Title=$title; Anchor=$anchor; FileName=$f.Name }
}

# Deduplicate by Title keeping first
$seen = @{}
$finalIndiv = @()
foreach ($s in $indivSections) {
    if (-not $seen.ContainsKey($s.Title)) {
        $seen[$s.Title] = $true
        $finalIndiv += $s
    }
}

# Compose baseline summary body from baseline HTML table section
$baselineBodyInner = Strip-DoctypeHeadBody $baselineHtml

# We'll extract only the baseline conformance summary table if present; else embed entire baseline body
$baselineSummaryHtml = $null
if ($baselineHtml -match '(?is)(<table[^>]*>.*?Baseline Conformance Reports.*?</table>)') {
    $baselineSummaryHtml = $Matches[1]
} else {
    $baselineSummaryHtml = $baselineBodyInner
}

# Create master HTML shell with CSS + JS transforms
$masterHtml = @"
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>SCuBA Report - $reportDateIso</title>
  <style>
    /* Page */
    html, body { background: #ffffff !important; margin: 0; padding: 0; font-family: Arial, Helvetica, sans-serif; color: #111; }
    .page { padding: 36px 42px; background: #ffffff; }
    h1, h2, h3 { color: #0b4f7b; margin: 0 0 10px 0; }
    h1 { font-size: 34px; font-weight: 700; }
    h2 { font-size: 22px; font-weight: 700; margin-top: 28px; }
    h3 { font-size: 18px; font-weight: 700; margin-top: 22px; }
    .subtitle { color: #333; font-size: 14px; margin: 0 0 18px 0; }

    /* Cover header logos row */
    .cover { border: 1px solid #e3e3e3; padding: 26px; border-radius: 2px; }
    .logo-row { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 16px; }
    .logo-left { width: 70px; height: 70px; object-fit: contain; }     /* square, no distortion */
    .logo-right { height: 46px; width: auto; object-fit: contain; }     /* company logo (double-ish size) */
    .meta { margin-top: 16px; font-size: 14px; line-height: 1.6; }
    .meta b { display: inline-block; width: 120px; }

    /* Contents */
    .contents { margin-top: 26px; }
    .contents a { color: #0b4f7b; text-decoration: underline; }
    .contents ol { margin: 8px 0 0 18px; }
    .contents li { margin: 4px 0; }

    /* Tables global (reasonable) */
    table { border-collapse: collapse; width: 100%; background: #fff; }
    th, td { border: 1px solid #000; padding: 8px 10px; vertical-align: top; }
    th { background: #fff; font-weight: 700; }

    /* Baseline Summary table tweaks */
    #baseline-summary table td { padding: 10px; }
    .details-bullets { margin: 0; padding-left: 18px; }
    .details-bullets li { margin: 2px 0; }

    /* Result row coloring across entire row (default SCuBA colors) */
    tr.result-pass  td { background: #e6f4e2 !important; }   /* light green */
    tr.result-fail  td { background: #f3dede !important; }   /* light red */
    tr.result-warn  td { background: #fbf3cf !important; }   /* light yellow */
    tr.result-manual td { background: #e9eef7 !important; }  /* light blue/gray */

    /* Remove UI-only elements */
    .no-print, .toggle, .toggle-wrapper, .light-mode, .lightMode, .lightmode, #lightMode, .light-mode-label { display:none !important; }

    /* Ensure sections are not cramped */
    .section { margin-top: 34px; page-break-inside: avoid; }

    /* CAP-only hard layout (scoped) */
    table.cap-fixed {
      width: 100% !important;
      table-layout: fixed !important;
      border-collapse: collapse !important;
      font-size: 11px !important;
    }
    table.cap-fixed th, table.cap-fixed td {
      word-break: break-word !important;
      overflow-wrap: anywhere !important;
      white-space: normal !important;
      vertical-align: top !important;
    }
    table.cap-fixed th:nth-child(1), table.cap-fixed td:nth-child(1) { width: 22% !important; }
    table.cap-fixed th:nth-child(2), table.cap-fixed td:nth-child(2) { width: 6%  !important; }
    table.cap-fixed th:nth-child(3), table.cap-fixed td:nth-child(3) { width: 18% !important; } /* Users shorter */
    table.cap-fixed th:nth-child(4), table.cap-fixed td:nth-child(4) { width: 14% !important; }
    table.cap-fixed th:nth-child(5), table.cap-fixed td:nth-child(5) { width: 18% !important; } /* Conditions shorter */
    table.cap-fixed th:nth-child(6), table.cap-fixed td:nth-child(6) { width: 12% !important; }
    table.cap-fixed th:nth-child(7), table.cap-fixed td:nth-child(7) { width: 10% !important; } /* Session controls stays visible */
  </style>
</head>
<body>
  <div class="page">
    <div class="cover">
      <div class="logo-row">
        <div>
          $(if ($relCisa) { "<img class='logo-left' src='$relCisa' alt='CISA Logo'>" } else { "" })
        </div>
        <div>
          $(if ($relCo) { "<img class='logo-right' src='$relCo' alt='Company Logo'>" } else { "" })
        </div>
      </div>

      <h1>SCuBA M365 Secure Configuration Baseline Assessment</h1>
      <p class="subtitle">Consolidated report package</p>

      <div class="meta">
        <div><b>Client:</b> $(Escape-Html $tenantDisplayName)</div>
        <div><b>Tenant domain:</b> $(Escape-Html $tenantDomainName)</div>
        <div><b>Tenant ID:</b> $(Escape-Html $tenantId)</div>
        <div><b>Report date:</b> $(Escape-Html $reportDateUtc)</div>
        <div><b>Prepared by:</b> $(Escape-Html $CompanyName)</div>
      </div>
    </div>

    <div class="contents">
      <h2>Contents</h2>
      <ol>
        <li><a href="#baseline-summary">Baseline Conformance Summary</a></li>
"@

$idx = 2
foreach ($s in $finalIndiv) {
    $masterHtml += "        <li><a href=""#$($s.Anchor)"">$($s.Title)</a></li>`r`n"
    $idx++
}

$masterHtml += @"
      </ol>
    </div>

    <div class="section" id="baseline-summary">
      <h2>Baseline Conformance Summary</h2>
      <div class="baseline-wrap">
        $baselineSummaryHtml
      </div>
    </div>
"@

# Embed each individual report body into its own section container
foreach ($s in $finalIndiv) {
    $raw = Read-FileText $s.File
    $body = Strip-DoctypeHeadBody $raw

    # Wrap in container and add a top H2 for consistency
    $masterHtml += @"
    <div class="section" id="$($s.Anchor)">
      <h2>$($s.Title)</h2>
      <div class="report-body" data-source="$($s.FileName)">
        $body
      </div>
    </div>
"@
}

# JS transforms: baseline details bullets; remove light mode label; fix CISA logo; flatten CAP; color result rows
$masterHtml += @"
  </div>

  <script>
  (function() {
    // --- Utility ---
    function textContains(el, needle) {
      if (!el) return false;
      return (el.textContent || "").toLowerCase().indexOf(needle.toLowerCase()) >= 0;
    }

    // --- Make Baseline Summary details into bullets and reduce whitespace ---
    (function baselineDetailsToBullets() {
      var base = document.getElementById('baseline-summary');
      if (!base) return;

      // baseline summary table: "Details" column holds "7 passes", "5 warnings", etc.
      var tables = base.getElementsByTagName('table');
      if (!tables || !tables.length) return;
      var t = tables[0];

      // For each row, find the "Details" cell(s) and convert stacked badges/text into bullet list.
      for (var r = 1; r < t.rows.length; r++) {
        var row = t.rows[r];
        if (!row || row.cells.length < 2) continue;

        // The right cell is "Details" in your baseline table
        var detailsCell = row.cells[row.cells.length - 1];
        if (!detailsCell) continue;

        // Extract lines from badges/spans/divs
        var parts = [];
        // Prefer text chunks from direct children
        var txt = detailsCell.innerText || detailsCell.textContent || "";
        txt.split(/\\r?\\n/).forEach(function(line) {
          var clean = line.trim();
          if (clean) parts.push(clean);
        });

        // If the cell contains the words pass/warning/failure/manual etc in one string, split by two+ spaces
        if (parts.length <= 1 && txt.indexOf('pass') >= 0) {
          parts = txt.split(/\\s{2,}/).map(function(x){ return x.trim(); }).filter(Boolean);
        }

        // Build bullet list
        if (parts.length) {
          var ul = document.createElement('ul');
          ul.className = 'details-bullets';
          parts.forEach(function(p) {
            var li = document.createElement('li');
            li.textContent = p;
            ul.appendChild(li);
          });
          detailsCell.innerHTML = '';
          detailsCell.appendChild(ul);
        }
      }
    })();

    // --- Remove "Light mode" label text anywhere ---
    (function removeLightModeLabel() {
      var walkers = document.querySelectorAll('body *');
      walkers.forEach(function(el) {
        // Remove nodes that are basically only "Light mode"
        if (el.children.length === 0) {
          var t = (el.textContent || "").trim().toLowerCase();
          if (t === "light mode" || t === "lightmode") {
            el.style.display = "none";
          }
        }
      });
    })();

    // --- Fix broken CISA logo in embedded reports (replace missing image src with local cisa logo if available) ---
    (function fixCisaLogo() {
      var localCisa = document.querySelector("img.logo-left");
      if (!localCisa) return;
      var localSrc = localCisa.getAttribute('src');

      // Find all images with alt containing "Cybersecurity and Infrastructure Security Agency" OR that are broken placeholders
      var imgs = document.querySelectorAll('.report-body img');
      imgs.forEach(function(img) {
        var alt = (img.getAttribute('alt') || '');
        if (alt.toLowerCase().indexOf('infrastructure security agency') >= 0 ||
            alt.toLowerCase().indexOf('cisa') >= 0) {
          img.setAttribute('src', localSrc);
          img.style.maxWidth = "70px";
          img.style.maxHeight = "70px";
          img.style.objectFit = "contain";
        }
      });
    })();

    // --- Color result rows across entire row based on "Result" cell text ---
    (function colorResultRows() {
      // For each table in report bodies, detect header containing "Result"
      var tables = document.querySelectorAll('.report-body table');
      tables.forEach(function(t) {
        var headerRow = t.rows && t.rows.length ? t.rows[0] : null;
        if (!headerRow) return;

        var resultCol = -1;
        for (var c = 0; c < headerRow.cells.length; c++) {
          var h = (headerRow.cells[c].innerText || headerRow.cells[c].textContent || "").trim().toLowerCase();
          if (h === "result") { resultCol = c; break; }
        }
        if (resultCol < 0) return;

        for (var r = 1; r < t.rows.length; r++) {
          var row = t.rows[r];
          if (!row.cells || row.cells.length <= resultCol) continue;
          var val = (row.cells[resultCol].innerText || row.cells[resultCol].textContent || "").trim().toLowerCase();

          row.classList.remove('result-pass','result-fail','result-warn','result-manual');

          // common values: Pass, Fail, Warning, N/A, Manual, Should/Not-Implemented etc
          if (val === "pass") row.classList.add('result-pass');
          else if (val === "fail") row.classList.add('result-fail');
          else if (val === "warning") row.classList.add('result-warn');
          else if (val === "manual" || val === "manual check") row.classList.add('result-manual');
        }
      });
    })();

    // --- Flatten Conditional Access Policies tables: remove expand/collapse UI & ensure all details visible ---
    (function flattenConditionalAccessPolicies() {
      // Identify CAP sections by heading text
      var headings = Array.from(document.querySelectorAll('.report-body h1, .report-body h2, .report-body h3'))
        .filter(function(h){ return (h.textContent || '').trim().toLowerCase() === 'conditional access policies'; });

      headings.forEach(function(h) {
        // Remove any nearby buttons like Expand all / Collapse all
        var scope = h.parentElement || document;
        var btns = scope.querySelectorAll('button, input[type="button"], a');
        btns.forEach(function(b) {
          var t = (b.innerText || b.value || b.textContent || '').trim().toLowerCase();
          if (t.indexOf('expand all') >= 0 || t.indexOf('collapse all') >= 0) {
            b.style.display = 'none';
          }
        });

        // Find the first table after the heading (CAP table)
        var capTable = null;
        var n = h;
        for (var i = 0; i < 10; i++) {
          n = n.nextElementSibling;
          if (!n) break;
          if (n.tagName && n.tagName.toLowerCase() === 'table') { capTable = n; break; }
          // sometimes wrapped
          var t = n.querySelector && n.querySelector('table');
          if (t) { capTable = t; break; }
        }
        if (!capTable) return;

        // Mark table so CAP-only CSS applies
        capTable.classList.add('cap-fixed');

        // CAP table often has a first column with chevron / expand arrow and hidden rows.
        // We want everything displayed. We'll:
        // - Unhide any hidden rows
        // - Remove chevron-only column if present
        // - Replace "..." links with their parent text, and ensure any collapsed detail rows are included inline if present.

        // Show all rows
        Array.from(capTable.querySelectorAll('tr')).forEach(function(tr) {
          tr.style.display = '';
          tr.hidden = false;
        });

        // Remove chevron column if it's the first column and contains mostly arrow glyphs
        // Detect if first column cells contain only > or chevrons or are empty
        var firstIsChevron = true;
        var rows = capTable.rows;
        if (!rows || rows.length < 2) return;

        // Check header first cell
        var headCell = rows[0].cells[0];
        var headText = (headCell ? (headCell.textContent || '').trim() : '');
        // If header says Name, then no chevron column
        if (headText.toLowerCase() === 'name') {
          firstIsChevron = false;
        } else {
          // Scan data rows for chevrons
          for (var r = 1; r < rows.length; r++) {
            var c0 = rows[r].cells[0];
            if (!c0) continue;
            var t0 = (c0.textContent || '').trim();
            if (t0 && t0 !== '>' && t0 !== '▸' && t0 !== '▾' && t0 !== 'v' && t0 !== '∨') {
              // If it looks like a policy name, then it's not chevron column
              // policy names tend to contain letters/numbers and hyphens
              if (t0.length > 2 && /[A-Za-z0-9]/.test(t0)) {
                // likely Name column
                firstIsChevron = false;
                break;
              }
            }
          }
        }

        if (firstIsChevron) {
          for (var r = 0; r < rows.length; r++) {
            if (rows[r].cells.length > 0) {
              rows[r].deleteCell(0);
            }
          }
        }

        // Replace "..." links with nothing (keep cell readable) and show their parent lists
        capTable.querySelectorAll('a').forEach(function(a) {
          var t = (a.textContent || '').trim();
          if (t === '...' || t === '…') {
            // remove link but keep a small placeholder to avoid losing spacing
            a.replaceWith(document.createTextNode(''));
          }
        });

        // Sometimes extra "detail rows" exist outside the table. If present, force display.
        var maybeHidden = scope.querySelectorAll('[style*="display:none"], [hidden]');
        maybeHidden.forEach(function(el) {
          // Only unhide if inside CAP section scope and seems relevant
          if (el.closest && el.closest('table') === capTable) {
            el.style.display = '';
            el.hidden = false;
          }
        });
      });
    })();

    // Re-run result coloring after CAP flattening, in case CAP tables have "Result" fields (usually not)
    (function() {
      // no-op: just ensure print stability
    })();
  })();
  </script>
</body>
</html>
"@

Write-FileText $masterHtmlPath $masterHtml
Write-Host "Master HTML created: $masterHtmlPath"
#endregion Build master HTML

#region Export PDF
if ($ExportPdf) {
    $edge = Get-EdgePath
    if (-not $edge) { throw "Microsoft Edge not found. Install Edge or provide msedge.exe in standard paths." }

    # Ensure we print from the Report folder so relative logo paths resolve
    $workDir = $reportFolder

    # Use a bigger window size and virtual time budget to allow JS transforms to apply before printing
    $args = @(
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--window-size=1400,900",
        "--virtual-time-budget=25000",
        "--print-to-pdf=`"$pdfPath`"",
        "`"$masterHtmlPath`""
    )

    Write-Host "Exporting PDF via Edge..."
    Start-Process -FilePath $edge -WorkingDirectory $workDir -ArgumentList $args -Wait | Out-Null

    if (-not (Test-Path $pdfPath)) {
        throw "PDF export failed. Master HTML is available: $masterHtmlPath"
    }

    Write-Host "PDF created: $pdfPath"
}
#endregion Export PDF

