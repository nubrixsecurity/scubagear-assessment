[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ScubaContainerSasUrl
)

$ErrorActionPreference = "Stop"

function Write-Err {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "[ERROR] $Message"
}

if ($PSVersionTable.PSEdition -ne "Desktop" -or $PSVersionTable.PSVersion.Major -ne 5) {
    Write-Err "This assessment must be run using Windows PowerShell 5.1. Please open Windows PowerShell, not PowerShell 7 / pwsh."
    exit 1
}

# ---------------------------
# Resolve package URL from container SAS
# ---------------------------

$qIndex = $ScubaContainerSasUrl.IndexOf("?")

if ($qIndex -lt 0) {
    Write-Err "ScubaContainerSasUrl does not contain a SAS query string ('?'). Please provide a valid container SAS URL."
    exit 1
}

$baseUrl = $ScubaContainerSasUrl.Substring(0, $qIndex).TrimEnd("/")
$sasPart = $ScubaContainerSasUrl.Substring($qIndex)

$PackageSasUrl = "$baseUrl/prod/scubagear-prod-package.zip$sasPart"

# ---------------------------
# Local temp paths
# ---------------------------

$root = Join-Path $env:TEMP "nubrix-scubagear"
$packagePath = Join-Path $env:TEMP "nubrix-scubagear-prod-package.zip"

try {
    if (Test-Path -LiteralPath $root) {
        Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path -LiteralPath $packagePath) {
        Remove-Item -LiteralPath $packagePath -Force -ErrorAction SilentlyContinue
    }
}
catch {}

New-Item -Path $root -ItemType Directory -Force | Out-Null

function Download-WithSasErrorHandling {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$OutFile,
        [Parameter(Mandatory = $true)][string]$FriendlyName
    )

    $oldPP = $ProgressPreference
    $ProgressPreference = "SilentlyContinue"

    try {
        Invoke-WebRequest -Uri $Uri -OutFile $OutFile -ErrorAction Stop

        try {
            Unblock-File -LiteralPath $OutFile -ErrorAction SilentlyContinue
        }
        catch {}

        return $true
    }
    catch {
        $msg = $_.Exception.Message

        if ($msg -match "403|AuthenticationFailed|Authorization") {
            Write-Err "Failed to download $FriendlyName. The link may have expired. Please request a refreshed link and try again."
        }
        elseif ($msg -match "409|404|NotFound") {
            Write-Err "Failed to download $FriendlyName. Expected blob path: prod/scubagear-prod-package.zip."
        }
        else {
            Write-Err "Failed to download $FriendlyName. $msg"
        }

        return $false
    }
    finally {
        $ProgressPreference = $oldPP
    }
}

try {
    if (-not (Download-WithSasErrorHandling -Uri $PackageSasUrl -OutFile $packagePath -FriendlyName "scubagear-prod-package.zip")) {
        exit 1
    }

    Write-Host "Extracting assessment package..."
    Expand-Archive -LiteralPath $packagePath -DestinationPath $root -Force

    $invokePath = Join-Path $root "invoke-scubagear.ps1"

    if (-not (Test-Path -LiteralPath $invokePath)) {
        throw "invoke-scubagear.ps1 was not found after extracting the assessment package."
    }

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File $invokePath
}
finally {
    try {
        if (Test-Path -LiteralPath $root) {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }

        if (Test-Path -LiteralPath $packagePath) {
            Remove-Item -LiteralPath $packagePath -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Warning "Unable to fully clean up temporary assessment files."
    }
}
