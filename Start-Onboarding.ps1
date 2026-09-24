<#
    Start-Onboarding.ps1
    Entry point for the PC Onboarding application.

    Steps, unattended, with on-screen progress and a transcript log:
      1. Scan for and remove any installed Microsoft Office products
         (Click-to-Run, MSI, per-language ARP entries, and the OneNote
         for Windows 10 AppX package).
      2. Remove preloaded/consumer Teams and the new Outlook app, ahead of
         the licensed Microsoft 365 deployment. OneDrive is left alone.
      3. Remove any preloaded McAfee products (Total Protection, LiveSafe,
         WebAdvisor, Safe Connect, the Store app), then run McAfee's own
         MCPR removal tool to clear the leftovers their uninstallers strand.
         MCPR is a CAPTCHA-gated wizard and needs someone at the keyboard -
         pass -SkipMcprCleanup to keep the run fully unattended.
      4. Download and silently install Google Chrome Enterprise.
      5. Download and silently install Adobe Acrobat Reader.
      6. Optional apps - prompted for interactively at startup (Dropbox,
         Slack, Google Drive, Cisco Secure Client). Answering "no" to all
         of them runs just the default set above.

    Usage:
        powershell.exe -ExecutionPolicy Bypass -File Start-Onboarding.ps1
#>

[CmdletBinding()]
param(
    [switch]$WhatIf,

    # Skips only the interactive MCPR pass; the silent McAfee uninstalls
    # in step 3 still run.
    [switch]$SkipMcprCleanup
)

$ErrorActionPreference = 'Stop'
$scriptRoot = $PSScriptRoot
$logDir = Join-Path $scriptRoot 'Logs'
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory | Out-Null }

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$transcriptPath = Join-Path $logDir "Onboarding_$timestamp.log"

# --- Require elevation; self-relaunch if needed ---
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host 'Relaunching with administrator privileges...' -ForegroundColor Yellow
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"") + $args
    Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs
    exit
}

# Padding so the blue Write-Progress bar doesn't cover the first lines of output.
1..6 | ForEach-Object { Write-Host '' }

Start-Transcript -Path $transcriptPath | Out-Null
Import-Module (Join-Path $scriptRoot 'Modules\OfficeRemoval.psm1') -Force
Import-Module (Join-Path $scriptRoot 'Modules\BundledAppRemoval.psm1') -Force
Import-Module (Join-Path $scriptRoot 'Modules\McAfeeRemoval.psm1') -Force
Import-Module (Join-Path $scriptRoot 'Modules\AppInstalls.psm1') -Force
Import-Module (Join-Path $scriptRoot 'Modules\OptionalAppInstalls.psm1') -Force

function Write-Step {
    param([string]$Message, [string]$Color = 'Cyan')
    Write-Host "`n=== $Message ===" -ForegroundColor $Color
}

function Read-YesNo {
    param([string]$Prompt)
    $answer = Read-Host "$Prompt (y/N)"
    return $answer -match '^[Yy]'
}

# --- Ask up front: default set only, or also optional apps? ---
Write-Step 'PC Onboarding Setup'
Write-Host 'Always runs: remove Office, remove preloaded Teams/new Outlook, remove McAfee, install Google Chrome Enterprise, install Adobe Acrobat Reader.'

$installDropbox = $false
$installSlack = $false
$installGoogleDrive = $false
$installCisco = $false
$ciscoInstallerPath = $null

if (Read-YesNo 'Also install any optional apps (Dropbox, Slack, Google Drive, Cisco Secure Client)?') {
    $installDropbox = Read-YesNo '  Install Dropbox?'
    $installSlack = Read-YesNo '  Install Slack?'
    $installGoogleDrive = Read-YesNo '  Install Google Drive?'
    $installCisco = Read-YesNo '  Install Cisco Secure Client?'

    if ($installCisco) {
        # Cisco Secure Client has no public download - it's gated behind a
        # Cisco.com account and your org's own entitlement/VPN headend - so
        # it has to come from a locally supplied installer instead.
        do {
            $ciscoInstallerPath = Read-Host '    Path to the Cisco Secure Client installer (.msi)'
            if (-not (Test-Path $ciscoInstallerPath)) {
                Write-Host "    File not found: $ciscoInstallerPath" -ForegroundColor Red
            }
        } while (-not (Test-Path $ciscoInstallerPath))
    }
}

try {
    Write-Step 'PC Onboarding - Step 1: Remove Microsoft Office'

    Write-Host 'Scanning installed applications for Office products...'
    $officeProducts = Get-InstalledOffice

    if (-not $officeProducts -or $officeProducts.Count -eq 0) {
        Write-Host 'No Office installations detected. Nothing to remove.' -ForegroundColor Green
    }
    else {
        Write-Host "Found $($officeProducts.Count) Office product(s):"
        $officeProducts | ForEach-Object { Write-Host "  - $($_.Name) [$($_.Type)]" }

        $total = $officeProducts.Count
        $index = 0
        $failed = [System.Collections.Generic.List[string]]::new()

        foreach ($product in $officeProducts) {
            $index++
            $percent = [int](($index - 1) / $total * 100)
            Write-Progress -Activity 'Removing Office installations' `
                -Status "($index of $total) $($product.Name)" `
                -PercentComplete $percent

            if ($WhatIf) {
                Write-Host "[WhatIf] Would remove: $($product.Name)" -ForegroundColor DarkYellow
                continue
            }

            Write-Host "Removing $($product.Name)..." -NoNewline
            $msiLog = Join-Path $logDir "Uninstall_$($product.ProductId -replace '[{}]','')_$timestamp.log"
            $success = Remove-OfficeInstallation -Product $product -LogPath $msiLog

            if ($success) {
                Write-Host ' Done.' -ForegroundColor Green
            }
            else {
                Write-Host ' FAILED.' -ForegroundColor Red
                $failed.Add($product.Name)
            }
        }

        Write-Progress -Activity 'Removing Office installations' -Completed

        Write-Step 'Office Removal Summary'
        Write-Host "Processed: $total"
        Write-Host "Succeeded: $($total - $failed.Count)" -ForegroundColor Green
        if ($failed.Count -gt 0) {
            Write-Host "Failed: $($failed.Count)" -ForegroundColor Red
            $failed | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        }

        # Re-scan to confirm removal.
        Write-Host "`nVerifying removal..."
        $remaining = Get-InstalledOffice
        if ($remaining.Count -eq 0) {
            Write-Host 'Verification passed: no Office installations remain.' -ForegroundColor Green
        }
        else {
            Write-Host "Verification found $($remaining.Count) remaining product(s) (may require a reboot to fully clear):" -ForegroundColor Yellow
            $remaining | ForEach-Object { Write-Host "  - $($_.Name)" }
        }
    }

    Write-Step 'PC Onboarding - Step 2: Remove Teams and new Outlook'

    Write-Host 'Scanning for Teams and new Outlook installs...'
    $bundledApps = Get-InstalledBundledApps

    if (-not $bundledApps -or $bundledApps.Count -eq 0) {
        Write-Host 'None detected. Nothing to remove.' -ForegroundColor Green
    }
    else {
        Write-Host "Found $($bundledApps.Count) app(s):"
        $bundledApps | ForEach-Object { Write-Host "  - $($_.Name) [$($_.Type)]" }

        $bundledTotal = $bundledApps.Count
        $bundledIndex = 0
        $bundledFailed = [System.Collections.Generic.List[string]]::new()

        foreach ($app in $bundledApps) {
            $bundledIndex++
            $percent = [int](($bundledIndex - 1) / $bundledTotal * 100)
            Write-Progress -Activity 'Removing Teams/Outlook' `
                -Status "($bundledIndex of $bundledTotal) $($app.Name)" `
                -PercentComplete $percent

            if ($WhatIf) {
                Write-Host "[WhatIf] Would remove: $($app.Name)" -ForegroundColor DarkYellow
                continue
            }

            Write-Host "Removing $($app.Name)..." -NoNewline
            $bundledLog = Join-Path $logDir "Uninstall_$($app.ProductId -replace '[{}:\\]','_')_$timestamp.log"
            $success = Remove-BundledAppInstallation -Product $app -LogPath $bundledLog

            if ($success) {
                Write-Host ' Done.' -ForegroundColor Green
            }
            else {
                Write-Host ' FAILED.' -ForegroundColor Red
                $bundledFailed.Add($app.Name)
            }
        }

        Write-Progress -Activity 'Removing Teams/Outlook' -Completed

        Write-Step 'Teams/Outlook Removal Summary'
        Write-Host "Processed: $bundledTotal"
        Write-Host "Succeeded: $($bundledTotal - $bundledFailed.Count)" -ForegroundColor Green
        if ($bundledFailed.Count -gt 0) {
            Write-Host "Failed: $($bundledFailed.Count)" -ForegroundColor Red
            $bundledFailed | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        }
    }

    Write-Step 'PC Onboarding - Step 3: Remove McAfee'

    Write-Host 'Scanning for McAfee products...'
    $mcafeeProducts = Get-InstalledMcAfee

    if (-not $mcafeeProducts -or $mcafeeProducts.Count -eq 0) {
        Write-Host 'No McAfee installations detected. Nothing to remove.' -ForegroundColor Green
    }
    else {
        Write-Host "Found $($mcafeeProducts.Count) McAfee product(s):"
        $mcafeeProducts | ForEach-Object { Write-Host "  - $($_.Name) [$($_.Type)]" }

        $mcafeeTotal = $mcafeeProducts.Count
        $mcafeeIndex = 0
        $mcafeeFailed = [System.Collections.Generic.List[string]]::new()

        foreach ($product in $mcafeeProducts) {
            $mcafeeIndex++
            $percent = [int](($mcafeeIndex - 1) / $mcafeeTotal * 100)
            Write-Progress -Activity 'Removing McAfee products' `
                -Status "($mcafeeIndex of $mcafeeTotal) $($product.Name)" `
                -PercentComplete $percent

            if ($WhatIf) {
                Write-Host "[WhatIf] Would remove: $($product.Name)" -ForegroundColor DarkYellow
                continue
            }

            Write-Host "Removing $($product.Name)..." -NoNewline
            $mcafeeLog = Join-Path $logDir "Uninstall_$($product.ProductId -replace '[{}:\\]','_')_$timestamp.log"
            $success = Remove-McAfeeInstallation -Product $product -LogPath $mcafeeLog

            if ($success) {
                Write-Host ' Done.' -ForegroundColor Green
            }
            else {
                Write-Host ' FAILED.' -ForegroundColor Red
                $mcafeeFailed.Add($product.Name)
            }
        }

        Write-Progress -Activity 'Removing McAfee products' -Completed

        Write-Step 'McAfee Removal Summary'
        Write-Host "Processed: $mcafeeTotal"
        Write-Host "Succeeded: $($mcafeeTotal - $mcafeeFailed.Count)" -ForegroundColor Green
        if ($mcafeeFailed.Count -gt 0) {
            Write-Host "Failed: $($mcafeeFailed.Count)" -ForegroundColor Red
            $mcafeeFailed | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        }

        # --- MCPR cleanup pass ---
        # McAfee's own uninstallers habitually strand services, drivers and
        # registry keys that block a later clean reinstall, so MCPR runs even
        # when every uninstall above reported success.
        if ($WhatIf) {
            Write-Host '[WhatIf] Would download and run the McAfee removal tool (MCPR).' -ForegroundColor DarkYellow
        }
        elseif ($SkipMcprCleanup) {
            Write-Host 'Skipping MCPR cleanup (-SkipMcprCleanup).' -ForegroundColor DarkYellow
        }
        else {
            Write-Host "`nDownloading the McAfee removal tool (MCPR)..."
            Write-Host 'NOTE: MCPR is a wizard with a CAPTCHA - it needs someone at the keyboard.' -ForegroundColor Yellow
            Write-Host '      Pass -SkipMcprCleanup for a fully unattended run.' -ForegroundColor Yellow

            $mcprSuccess = Invoke-McAfeeRemovalTool
            if ($mcprSuccess) {
                Write-Host 'MCPR completed.' -ForegroundColor Green
                Write-Host 'A reboot is required to finish clearing McAfee drivers and services.' -ForegroundColor Yellow
            }
            else {
                Write-Host 'MCPR did not complete successfully.' -ForegroundColor Red
            }
        }

        # Re-scan to confirm removal.
        Write-Host "`nVerifying removal..."
        $mcafeeRemaining = Get-InstalledMcAfee
        if ($mcafeeRemaining.Count -eq 0) {
            Write-Host 'Verification passed: no McAfee installations remain.' -ForegroundColor Green
        }
        else {
            Write-Host "Verification found $($mcafeeRemaining.Count) remaining product(s) (may require a reboot to fully clear):" -ForegroundColor Yellow
            $mcafeeRemaining | ForEach-Object { Write-Host "  - $($_.Name)" }
        }
    }

    Write-Step 'PC Onboarding - Step 4: Install Google Chrome Enterprise'
    if ($WhatIf) {
        Write-Host '[WhatIf] Would download and install Google Chrome Enterprise.' -ForegroundColor DarkYellow
    }
    else {
        Write-Host 'Downloading and installing Google Chrome Enterprise...' -NoNewline
        $chromeLog = Join-Path $logDir "ChromeInstall_$timestamp.log"
        $chromeSuccess = Install-ChromeEnterprise -LogPath $chromeLog
        if ($chromeSuccess) {
            Write-Host ' Done.' -ForegroundColor Green
        }
        else {
            Write-Host ' FAILED.' -ForegroundColor Red
        }
    }

    Write-Step 'PC Onboarding - Step 5: Install Adobe Acrobat Reader'
    if ($WhatIf) {
        Write-Host '[WhatIf] Would download and install Adobe Acrobat Reader.' -ForegroundColor DarkYellow
    }
    else {
        Write-Host 'Downloading and installing Adobe Acrobat Reader...' -NoNewline
        $adobeSuccess = Install-AdobeReader
        if ($adobeSuccess) {
            Write-Host ' Done.' -ForegroundColor Green
        }
        else {
            Write-Host ' FAILED.' -ForegroundColor Red
        }
    }

    Write-Step 'PC Onboarding - Step 6: Optional Applications'

    if (-not $installDropbox -and -not $installSlack -and -not $installGoogleDrive -and -not $installCisco) {
        Write-Host 'None selected. Skipping.' -ForegroundColor Green
    }

    if ($installDropbox) {
        if ($WhatIf) {
            Write-Host '[WhatIf] Would download and install Dropbox.' -ForegroundColor DarkYellow
        }
        else {
            Write-Host 'Downloading and installing Dropbox...' -NoNewline
            $dropboxSuccess = Install-Dropbox
            if ($dropboxSuccess) { Write-Host ' Done.' -ForegroundColor Green }
            else { Write-Host ' FAILED.' -ForegroundColor Red }
        }
    }

    if ($installSlack) {
        if ($WhatIf) {
            Write-Host '[WhatIf] Would download and install Slack.' -ForegroundColor DarkYellow
        }
        else {
            Write-Host 'Downloading and installing Slack...' -NoNewline
            $slackSuccess = Install-Slack
            if ($slackSuccess) { Write-Host ' Done.' -ForegroundColor Green }
            else { Write-Host ' FAILED.' -ForegroundColor Red }
        }
    }

    if ($installGoogleDrive) {
        if ($WhatIf) {
            Write-Host '[WhatIf] Would download and install Google Drive.' -ForegroundColor DarkYellow
        }
        else {
            Write-Host 'Downloading and installing Google Drive...' -NoNewline
            $googleDriveSuccess = Install-GoogleDrive
            if ($googleDriveSuccess) { Write-Host ' Done.' -ForegroundColor Green }
            else { Write-Host ' FAILED.' -ForegroundColor Red }
        }
    }

    if ($installCisco) {
        if ($WhatIf) {
            Write-Host "[WhatIf] Would install Cisco Secure Client from $ciscoInstallerPath." -ForegroundColor DarkYellow
        }
        else {
            Write-Host "Installing Cisco Secure Client from $ciscoInstallerPath..." -NoNewline
            $ciscoLog = Join-Path $logDir "CiscoSecureClientInstall_$timestamp.log"
            $ciscoSuccess = Install-CiscoSecureClient -InstallerPath $ciscoInstallerPath -LogPath $ciscoLog
            if ($ciscoSuccess) { Write-Host ' Done.' -ForegroundColor Green }
            else { Write-Host ' FAILED.' -ForegroundColor Red }
        }
    }
}
finally {
    Stop-Transcript | Out-Null
    Write-Host "`nLog saved to: $transcriptPath"
}
