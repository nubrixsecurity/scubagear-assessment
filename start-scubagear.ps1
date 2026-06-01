[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ScubaContainerSasUrl,
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
# Resolve invoke script URL from container SAS
# ---------------------------

$qIndex = $ScubaContainerSasUrl.IndexOf("?")

if ($qIndex -lt 0) {
    Write-Err "ScubaContainerSasUrl does not contain a SAS query string ('?'). Please provide a valid container SAS URL."
    exit 1
}

$baseUrl = $ScubaContainerSasUrl.Substring(0, $qIndex).TrimEnd("/")
$sasPart = $ScubaContainerSasUrl.Substring($qIndex)

$InvokeScubaSasUrl = "$baseUrl/prod/invoke-scubagear.ps1$sasPart"

# ---------------------------
# Local temp paths
# ---------------------------

$root = Join-Path $env:TEMP "nubrix-scubagear"
New-Item -Path $root -ItemType Directory -Force | Out-Null

$invokePath = Join-Path $root "invoke-scubagear.ps1"

function Download-WithSasErrorHandling {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$OutFile,
        [Parameter(Mandatory = $true)][string]$FriendlyName
    )

    $oldPP = $ProgressPreference
    $ProgressPreference = "SilentlyContinue"

    try {
        if (Test-Path -LiteralPath $OutFile) {
            Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
        }

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
        elseif ($msg -match "409") {
            Write-Err "Failed to download $FriendlyName. The SAS URL appears to point to the container, but the script could not derive or access the expected blob path: prod/invoke-scubagear.ps1."
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

if (-not (Download-WithSasErrorHandling -Uri $InvokeScubaSasUrl -OutFile $invokePath -FriendlyName "invoke-scubagear.ps1")) {
    exit 1
}

try {
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File $invokePath
}
finally {
    try {
        if (Test-Path -LiteralPath $invokePath) {
            Remove-Item -LiteralPath $invokePath -Force -ErrorAction SilentlyContinue
        }
    }
    catch {}
}
