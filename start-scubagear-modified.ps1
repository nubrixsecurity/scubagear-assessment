[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ScubaContainerSasUrl
)

$ErrorActionPreference = "Stop"

function Write-Err {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "[ERROR] $Message"
}

if (
    $PSVersionTable.PSEdition -ne "Desktop" -or
    $PSVersionTable.PSVersion.Major -ne 5
) {
    Write-Err "This assessment must be run using Windows PowerShell 5.1. Please open Windows PowerShell, not PowerShell 7 / pwsh."
    exit 1
}

$qIndex = $ScubaContainerSasUrl.IndexOf("?")

if ($qIndex -lt 0) {
    Write-Err "ScubaContainerSasUrl does not contain a SAS query string ('?'). Please provide a valid assessment link."
    exit 1
}

$baseUrl = $ScubaContainerSasUrl.Substring(0, $qIndex).TrimEnd("/")
$sasPart = $ScubaContainerSasUrl.Substring($qIndex)

$PackageSasUrl = "$baseUrl/prod/scubagear-prod-package.zip$sasPart"

$root = Join-Path $env:TEMP "nubrix-scubagear"
$packagePath = Join-Path $env:TEMP "nubrix-scubagear-prod-package.zip"

try {
    if (Test-Path -LiteralPath $root) {
        Remove-Item `
            -LiteralPath $root `
            -Recurse `
            -Force `
            -ErrorAction SilentlyContinue
    }

    if (Test-Path -LiteralPath $packagePath) {
        Remove-Item `
            -LiteralPath $packagePath `
            -Force `
            -ErrorAction SilentlyContinue
    }
}
catch {}

New-Item `
    -Path $root `
    -ItemType Directory `
    -Force | Out-Null

function Download-WithSasErrorHandling {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [string]$OutFile,

        [Parameter(Mandatory = $true)]
        [string]$FriendlyName
    )

    $oldPP = $ProgressPreference
    $ProgressPreference = "SilentlyContinue"

    try {
        Invoke-WebRequest `
            -Uri $Uri `
            -OutFile $OutFile `
            -ErrorAction Stop

        try {
            Unblock-File `
                -LiteralPath $OutFile `
                -ErrorAction SilentlyContinue
        }
        catch {}

        return $true
    }
    catch {
        $msg = $_.Exception.Message

        if ($msg -match "403|AuthenticationFailed|Authorization") {
            Write-Err "Failed to download $FriendlyName. The assessment link may have expired. Please request a refreshed link and try again."
        }
        elseif ($msg -match "409|404|NotFound") {
            Write-Err "Failed to download $FriendlyName. The assessment package could not be found."
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
    Write-Host "Extracting assessment package..."

    if (
        -not (
            Download-WithSasErrorHandling `
                -Uri $PackageSasUrl `
                -OutFile $packagePath `
                -FriendlyName "assessment package"
        )
    ) {
        exit 1
    }

    Expand-Archive `
        -LiteralPath $packagePath `
        -DestinationPath $root `
        -Force

    $invokePath = Join-Path $root "invoke-scubagear.ps1"
    $localModuleRoot = Join-Path $root "modules"

    if (-not (Test-Path -LiteralPath $invokePath)) {
        throw "The assessment package is missing a required component. Please request a refreshed package and try again."
    }

    # ==============================
    # MICROSOFT GRAPH AUTHENTICATION
    # ==============================

    if (Test-Path -LiteralPath $localModuleRoot) {
        $modulePaths = @(
            $localModuleRoot
            $env:PSModulePath -split [IO.Path]::PathSeparator
        ) |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_)
            } |
            Select-Object -Unique

        $env:PSModulePath = $modulePaths -join [IO.Path]::PathSeparator
    }

    $graphAuthenticationModule = Get-Module `
        -ListAvailable `
        -Name "Microsoft.Graph.Authentication" |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $graphAuthenticationModule) {
        Write-Host "Installing required Microsoft Graph module..."

        try {
            Install-Module `
                -Name "Microsoft.Graph.Authentication" `
                -Scope CurrentUser `
                -Force `
                -AllowClobber `
                -ErrorAction Stop
        }
        catch {
            throw "Unable to install the required Microsoft Graph authentication module. $($_.Exception.Message)"
        }
    }

    try {
        Import-Module `
            Microsoft.Graph.Authentication `
            -Force `
            -ErrorAction Stop

        Disconnect-MgGraph `
            -ErrorAction SilentlyContinue

        Connect-MgGraph `
            -UseDeviceCode `
            -ContextScope CurrentUser `
            -NoWelcome `
            -ErrorAction Stop
    }
    catch {
        throw "Unable to authenticate to Microsoft Graph. $($_.Exception.Message)"
    }

    powershell.exe `
        -NoProfile `
        -ExecutionPolicy Bypass `
        -File $invokePath
}
finally {
    try {
        if (Test-Path -LiteralPath $root) {
            Remove-Item `
                -LiteralPath $root `
                -Recurse `
                -Force `
                -ErrorAction SilentlyContinue
        }

        if (Test-Path -LiteralPath $packagePath) {
            Remove-Item `
                -LiteralPath $packagePath `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Warning "Unable to fully clean up temporary assessment files."
    }
}
