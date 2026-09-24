<#
    Start-Onboarding.ps1
    Entry point for the PC Onboarding application.

    Startup questions, then steps with on-screen progress and a transcript log:
      1. Rename the PC (optional) and join it - chosen at startup: Microsoft
         Entra ID, an Active Directory domain, or both (domain join + hybrid
         Entra join).
      2. Scan for and remove any installed Microsoft Office products
         (Click-to-Run, MSI, per-language ARP entries, and the OneNote
         for Windows 10 AppX package).
      3. Remove preloaded/consumer Teams and the new Outlook app, ahead of
         the licensed Microsoft 365 deployment. OneDrive is left alone.
      4. Remove any preloaded McAfee products (Total Protection, LiveSafe,
         WebAdvisor, Safe Connect, the Store app), then run McAfee's own
         MCPR removal tool to clear the leftovers their uninstallers strand.
         MCPR is a CAPTCHA-gated wizard and needs someone at the keyboard -
         pass -SkipMcprCleanup to keep the run fully unattended.
      5. Download and silently install Google Chrome Enterprise.
      6. Download and silently install Adobe Acrobat Reader.
      7. Optional apps - prompted for interactively at startup (Dropbox,
         Slack, Google Drive, Cisco Secure Client, Firefox, Zoom). Answering "no" to all
         of them runs just the default set above.
      8. Windows Update - chosen at startup: install updates as part of the
         run (time counted, updates listed in the summary), or start a
         "Check for updates" and let them finish in the background.

    Usage:
        powershell.exe -ExecutionPolicy Bypass -File Start-Onboarding.ps1
#>

[CmdletBinding()]
param(
    [switch]$WhatIf,

    # Skips only the interactive MCPR pass; the silent McAfee uninstalls
    # in step 4 still run.
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
    # -NoExit keeps the elevated window open afterward so the summary can be read.
    $argList = @('-NoExit', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"") + $args
    Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs
    exit
}

# Clock starts before the startup questions so they count toward runtime.
$startTime = Get-Date

# Padding so the blue Write-Progress bar doesn't cover the first lines of output.
1..6 | ForEach-Object { Write-Host '' }

Start-Transcript -Path $transcriptPath | Out-Null
Import-Module (Join-Path $scriptRoot 'Modules\DeviceJoin.psm1') -Force
Import-Module (Join-Path $scriptRoot 'Modules\OfficeRemoval.psm1') -Force
Import-Module (Join-Path $scriptRoot 'Modules\BundledAppRemoval.psm1') -Force
Import-Module (Join-Path $scriptRoot 'Modules\McAfeeRemoval.psm1') -Force
Import-Module (Join-Path $scriptRoot 'Modules\AppInstalls.psm1') -Force
Import-Module (Join-Path $scriptRoot 'Modules\OptionalAppInstalls.psm1') -Force
Import-Module (Join-Path $scriptRoot 'Modules\WindowsUpdates.psm1') -Force

function Write-Step {
    param([string]$Message, [string]$Color = 'Cyan')
    Write-Host "`n=== $Message ===" -ForegroundColor $Color
}

function Read-YesNo {
    param([string]$Prompt)
    $answer = Read-Host "$Prompt (y/N)"
    return $answer -match '^[Yy]'
}

function Read-DomainName {
    # Re-asks until something is entered - an empty name would fail parameter
    # binding in Join-ADDomain and abort the whole run.
    do { $name = (Read-Host '  Domain name (e.g. contoso.local)').Trim() } while (-not $name)
    return $name
}

function Read-NewComputerName {
    # Returns $null to keep the current name. Windows computer names are 1-15
    # letters, numbers, or hyphens, not all numbers, and no leading/trailing hyphen.
    param([string]$CurrentName)
    while ($true) {
        $name = (Read-Host "Rename this PC? Current name is $CurrentName. Enter a new name, or press Enter to keep it").Trim()
        if (-not $name -or $name -eq $CurrentName) { return $null }
        if ($name -match '^(?![0-9]+$)(?!-)[A-Za-z0-9-]{1,15}(?<!-)$') { return $name }
        Write-Host '  Names must be 1-15 letters, numbers, or hyphens, not all numbers, and can''t start or end with a hyphen.' -ForegroundColor Red
    }
}

function Format-Minutes {
    param([int]$Minutes)
    if ($Minutes -lt 60) { return "$Minutes min" }
    $hours = [math]::Floor($Minutes / 60)
    $rest = $Minutes % 60
    if ($rest -eq 0) { return "$hours hr" }
    return "$hours hr $rest min"
}

function Get-SummaryName {
    # AppX entries carry the raw package name, e.g. "Microsoft Teams (new)
    # (MSTeams_25.1_x64__8wekyb3d8bbwe)". Package names always contain an
    # underscore, so strip that trailing parenthetical for a client-readable name.
    param([string]$Name)
    return $Name -replace ' \((provisioned, )?[^()]*_[^()]*\)$', ''
}

# Collected across every step for the end-of-run ticket summary.
$summaryRemoved = [System.Collections.Generic.List[string]]::new()
$summaryInstalled = [System.Collections.Generic.List[string]]::new()
$summaryFailed = [System.Collections.Generic.List[string]]::new()
$summaryJoined = [System.Collections.Generic.List[string]]::new()
$restartNeeded = $false
$renamed = $false
$updatesStarted = $false
$summaryUpdates = [System.Collections.Generic.List[string]]::new()
$updatesChecked = $false
$runCompleted = $false

# --- Ask up front: default set only, or also optional apps? ---
Write-Step 'PC Onboarding Setup'
Write-Host 'Always runs: remove Office, remove preloaded Teams/new Outlook, remove McAfee, install Google Chrome Enterprise, install Adobe Acrobat Reader.'

$joinStatus = Get-JoinStatus

Write-Host ''
$newName = Read-NewComputerName -CurrentName $env:COMPUTERNAME

Write-Host ''
Write-Host 'How should this PC be joined?'
Write-Host '  1) Microsoft Entra ID (Entra joined)'
Write-Host '  2) Active Directory domain (domain joined)'
Write-Host '  3) Both - domain join, then hybrid Entra join'
do {
    $joinChoice = (Read-Host 'Enter 1, 2, or 3 (or just press Enter to skip joining)').Trim()
} while ($joinChoice -and $joinChoice -notin '1', '2', '3')

$domainName = $null
$domainCredential = $null
if ($joinChoice -in '2', '3') {
    $domainName = Read-DomainName
    $domainCredential = Get-Credential -Message "Account allowed to join computers to $domainName (e.g. CONTOSO\admin)"
}

# Renaming a PC that's already in a domain also renames its account in AD,
# which needs domain rights - ask now rather than stall mid-run.
if ($newName -and $joinStatus.DomainJoined -and -not $domainCredential) {
    $domainCredential = Get-Credential -Message "Account allowed to rename computers in $($joinStatus.Domain) (e.g. CONTOSO\admin)"
}

Write-Host ''
$installDropbox = $false
$installSlack = $false
$installGoogleDrive = $false
$installCisco = $false
$installFirefox = $false
$installZoom = $false
$ciscoInstallerPath = $null

if (Read-YesNo 'Also install any optional apps (Dropbox, Slack, Google Drive, Cisco Secure Client, Firefox, Zoom)?') {
    $installDropbox = Read-YesNo '  Install Dropbox?'
    $installSlack = Read-YesNo '  Install Slack?'
    $installGoogleDrive = Read-YesNo '  Install Google Drive?'
    $installFirefox = Read-YesNo '  Install Firefox?'
    $installZoom = Read-YesNo '  Install Zoom?'
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

# Asked here with the other questions rather than at step 8, so the run never
# sits waiting on a prompt (with the clock running) after the tech walks away.
Write-Host ''
Write-Host 'Windows Update at the end of the run:'
Write-Host '  1) Install updates as part of the run - update time counts toward runtime, updates listed in the summary'
Write-Host '  2) Start updates in the background - summary is ready right away, updates finish afterward'
do {
    $updateChoice = (Read-Host 'Enter 1 or 2 (or just press Enter for 2)').Trim()
} while ($updateChoice -and $updateChoice -notin '1', '2')
$installUpdatesInRun = $updateChoice -eq '1'

try {
    Write-Step 'PC Onboarding - Step 1: Rename and Join Device'

    if (-not $newName -and -not $joinChoice) {
        Write-Host 'No rename or join selected. Skipping.' -ForegroundColor Green
    }

    # Windows Home can still be renamed, just not joined.
    $joinSupported = $true
    if ($joinChoice -and -not (Test-JoinSupportedEdition)) {
        Write-Host 'This PC is running Windows Home, which cannot join a domain or Microsoft Entra ID.' -ForegroundColor Red
        Write-Host 'Upgrade it to Windows Pro, then join it manually.' -ForegroundColor Red
        $summaryFailed.Add('Join device (Windows Home edition cannot be joined; upgrade to Pro)')
        $joinSupported = $false
    }

    # Set once the rename has been dealt with (done, or folded into the domain
    # join); $renamed is only set when it actually succeeded.
    $renameHandled = $false

    # --- Domain join: options 2 and 3 (also applies the rename, if one was asked for) ---
    $domainReady = $false
    if ($joinSupported -and $joinChoice -in '2', '3') {
        if ($joinStatus.DomainJoined) {
            Write-Host "Already joined to domain $($joinStatus.Domain). Skipping domain join." -ForegroundColor Green
            $domainReady = $true
        }
        elseif ($WhatIf) {
            $asName = if ($newName) { " as $newName" } else { '' }
            Write-Host "[WhatIf] Would join domain $domainName$asName." -ForegroundColor DarkYellow
            $renameHandled = [bool]$newName
        }
        else {
            # Retry loop: a typo in the domain name or password, or the PC not
            # yet being on the client's network/VPN, is the usual failure.
            while ($true) {
                $asName = if ($newName) { " as $newName" } else { '' }
                Write-Host "Joining domain $domainName$asName..." -NoNewline
                if ($domainCredential -and (Join-ADDomain -DomainName $domainName -Credential $domainCredential -NewName $newName)) {
                    Write-Host ' Done (takes effect after restart).' -ForegroundColor Green
                    $domainReady = $true
                    if ($newName) { $renameHandled = $true; $renamed = $true }
                    break
                }
                Write-Host ' FAILED.' -ForegroundColor Red
                if (-not (Read-YesNo 'Try again (you can re-enter the domain name and account)?')) { break }
                $domainName = Read-DomainName
                $domainCredential = Get-Credential -Message "Account allowed to join computers to $domainName (e.g. CONTOSO\admin)"
            }

            if ($domainReady) {
                $summaryJoined.Add("Domain: $domainName")
                $restartNeeded = $true
            }
            else {
                $summaryFailed.Add("Join domain $domainName")
            }
        }
    }

    # --- Rename on its own: no domain join this run, or the join failed ---
    if ($newName -and -not $renameHandled) {
        if ($WhatIf) {
            Write-Host "[WhatIf] Would rename this PC to $newName." -ForegroundColor DarkYellow
        }
        else {
            Write-Host "Renaming this PC to $newName..." -NoNewline
            $renameCredential = if ($joinStatus.DomainJoined) { $domainCredential } else { $null }
            if (Rename-Device -NewName $newName -DomainCredential $renameCredential) {
                Write-Host ' Done (takes effect after restart).' -ForegroundColor Green
                $renamed = $true
                $restartNeeded = $true
            }
            else {
                Write-Host ' FAILED.' -ForegroundColor Red
                $summaryFailed.Add("Rename PC to $newName")
            }
        }
    }

    if ($joinSupported) {
        # --- Hybrid Entra join: option 3 ---
        # Can't be done directly from here: after the domain join takes effect,
        # Windows registers the device itself - but only if the client's
        # Microsoft Entra Connect is set up for hybrid join.
        if ($joinChoice -eq '3' -and $domainReady) {
            if ($joinStatus.EntraJoined) {
                Write-Host "Already joined to Microsoft Entra ID ($($joinStatus.TenantName)). Skipping." -ForegroundColor Green
            }
            else {
                Write-Host 'Hybrid Entra join completes automatically after restart, if the client''s Microsoft Entra Connect' -ForegroundColor Yellow
                Write-Host 'is set up for hybrid join. Confirm later with: dsregcmd /status  (AzureAdJoined : YES)' -ForegroundColor Yellow
                $summaryJoined.Add('Microsoft Entra ID (hybrid): completes after restart')
            }
        }

        # --- Entra join: option 1 ---
        if ($joinChoice -eq '1') {
            if ($joinStatus.EntraJoined) {
                Write-Host "Already joined to Microsoft Entra ID ($($joinStatus.TenantName)). Skipping." -ForegroundColor Green
            }
            elseif ($joinStatus.DomainJoined) {
                Write-Host "This PC is joined to domain $($joinStatus.Domain), so it can't be Entra joined directly." -ForegroundColor Red
                Write-Host 'Re-run and choose option 3 (hybrid) instead.' -ForegroundColor Red
                $summaryFailed.Add('Join Microsoft Entra ID (PC is domain joined; needs hybrid join)')
            }
            elseif ($WhatIf) {
                Write-Host '[WhatIf] Would open the Microsoft Entra join screen.' -ForegroundColor DarkYellow
            }
            else {
                if ($renamed) {
                    # The new name only takes effect after the restart, so the
                    # join below registers the device under the current name.
                    Write-Host "Note: Entra ID may list this PC as $env:COMPUTERNAME until it restarts as $newName." -ForegroundColor Yellow
                }
                do {
                    Start-EntraJoin
                    Write-Host 'Settings > Access work or school is opening. In that window:' -ForegroundColor Yellow
                    Write-Host '  1. Click Connect.'
                    Write-Host '  2. Click "Join this device to Microsoft Entra ID" (link at the bottom).'
                    Write-Host '  3. Sign in with the user''s work account and confirm the organization.'
                    Write-Host '  4. Click Done. Do NOT restart yet - this script will say when.'
                    $null = Read-Host 'Press Enter here once the join is finished'

                    $joinStatus = Get-JoinStatus
                    if ($joinStatus.EntraJoined) {
                        Write-Host "Joined to Microsoft Entra ID ($($joinStatus.TenantName))." -ForegroundColor Green
                        break
                    }
                    Write-Host 'This PC does not show as Entra joined yet.' -ForegroundColor Red
                } while (Read-YesNo 'Open the join screen and try again?')

                if ($joinStatus.EntraJoined) {
                    $tenantLabel = if ($joinStatus.TenantName) { "Microsoft Entra ID: $($joinStatus.TenantName)" } else { 'Microsoft Entra ID' }
                    $summaryJoined.Add($tenantLabel)
                    $restartNeeded = $true
                }
                else {
                    $summaryFailed.Add('Join Microsoft Entra ID')
                }
            }
        }
    }

    Write-Step 'PC Onboarding - Step 2: Remove Microsoft Office'

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
                $summaryRemoved.Add((Get-SummaryName $product.Name))
            }
            else {
                Write-Host ' FAILED.' -ForegroundColor Red
                $failed.Add($product.Name)
                $summaryFailed.Add("Remove $(Get-SummaryName $product.Name)")
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

    Write-Step 'PC Onboarding - Step 3: Remove Teams and new Outlook'

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
                $summaryRemoved.Add((Get-SummaryName $app.Name))
            }
            else {
                Write-Host ' FAILED.' -ForegroundColor Red
                $bundledFailed.Add($app.Name)
                $summaryFailed.Add("Remove $(Get-SummaryName $app.Name)")
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

    Write-Step 'PC Onboarding - Step 4: Remove McAfee'

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
                $summaryRemoved.Add((Get-SummaryName $product.Name))
            }
            else {
                Write-Host ' FAILED.' -ForegroundColor Red
                $mcafeeFailed.Add($product.Name)
                $summaryFailed.Add("Remove $(Get-SummaryName $product.Name)")
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
                $summaryRemoved.Add('McAfee leftover files and services (McAfee removal tool)')
            }
            else {
                Write-Host 'MCPR did not complete successfully.' -ForegroundColor Red
                $summaryFailed.Add('McAfee leftover cleanup (McAfee removal tool)')
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

    Write-Step 'PC Onboarding - Step 5: Install Google Chrome Enterprise'
    if ($WhatIf) {
        Write-Host '[WhatIf] Would download and install Google Chrome Enterprise.' -ForegroundColor DarkYellow
    }
    else {
        Write-Host 'Downloading and installing Google Chrome Enterprise...' -NoNewline
        $chromeLog = Join-Path $logDir "ChromeInstall_$timestamp.log"
        $chromeSuccess = Install-ChromeEnterprise -LogPath $chromeLog
        if ($chromeSuccess) {
            Write-Host ' Done.' -ForegroundColor Green
            $summaryInstalled.Add('Google Chrome Enterprise')
        }
        else {
            Write-Host ' FAILED.' -ForegroundColor Red
            $summaryFailed.Add('Install Google Chrome Enterprise')
        }
    }

    Write-Step 'PC Onboarding - Step 6: Install Adobe Acrobat Reader'
    if ($WhatIf) {
        Write-Host '[WhatIf] Would download and install Adobe Acrobat Reader.' -ForegroundColor DarkYellow
    }
    else {
        Write-Host 'Downloading and installing Adobe Acrobat Reader...' -NoNewline
        $adobeSuccess = Install-AdobeReader
        if ($adobeSuccess) {
            Write-Host ' Done.' -ForegroundColor Green
            $summaryInstalled.Add('Adobe Acrobat Reader')
        }
        else {
            Write-Host ' FAILED.' -ForegroundColor Red
            $summaryFailed.Add('Install Adobe Acrobat Reader')
        }
    }

    Write-Step 'PC Onboarding - Step 7: Optional Applications'

    if (-not ($installDropbox -or $installSlack -or $installGoogleDrive -or $installFirefox -or $installZoom -or $installCisco)) {
        Write-Host 'None selected. Skipping.' -ForegroundColor Green
    }

    if ($installDropbox) {
        if ($WhatIf) {
            Write-Host '[WhatIf] Would download and install Dropbox.' -ForegroundColor DarkYellow
        }
        else {
            Write-Host 'Downloading and installing Dropbox...' -NoNewline
            $dropboxSuccess = Install-Dropbox
            if ($dropboxSuccess) { Write-Host ' Done.' -ForegroundColor Green; $summaryInstalled.Add('Dropbox') }
            else { Write-Host ' FAILED.' -ForegroundColor Red; $summaryFailed.Add('Install Dropbox') }
        }
    }

    if ($installSlack) {
        if ($WhatIf) {
            Write-Host '[WhatIf] Would download and install Slack.' -ForegroundColor DarkYellow
        }
        else {
            Write-Host 'Downloading and installing Slack...' -NoNewline
            $slackSuccess = Install-Slack
            if ($slackSuccess) { Write-Host ' Done.' -ForegroundColor Green; $summaryInstalled.Add('Slack') }
            else { Write-Host ' FAILED.' -ForegroundColor Red; $summaryFailed.Add('Install Slack') }
        }
    }

    if ($installGoogleDrive) {
        if ($WhatIf) {
            Write-Host '[WhatIf] Would download and install Google Drive.' -ForegroundColor DarkYellow
        }
        else {
            Write-Host 'Downloading and installing Google Drive...' -NoNewline
            $googleDriveSuccess = Install-GoogleDrive
            if ($googleDriveSuccess) { Write-Host ' Done.' -ForegroundColor Green; $summaryInstalled.Add('Google Drive') }
            else { Write-Host ' FAILED.' -ForegroundColor Red; $summaryFailed.Add('Install Google Drive') }
        }
    }

    if ($installFirefox) {
        if ($WhatIf) {
            Write-Host '[WhatIf] Would download and install Firefox.' -ForegroundColor DarkYellow
        }
        else {
            Write-Host 'Downloading and installing Firefox...' -NoNewline
            $firefoxSuccess = Install-Firefox -LogPath (Join-Path $logDir "FirefoxInstall_$timestamp.log")
            if ($firefoxSuccess) { Write-Host ' Done.' -ForegroundColor Green; $summaryInstalled.Add('Mozilla Firefox') }
            else { Write-Host ' FAILED.' -ForegroundColor Red; $summaryFailed.Add('Install Mozilla Firefox') }
        }
    }

    if ($installZoom) {
        if ($WhatIf) {
            Write-Host '[WhatIf] Would download and install Zoom.' -ForegroundColor DarkYellow
        }
        else {
            Write-Host 'Downloading and installing Zoom...' -NoNewline
            $zoomSuccess = Install-Zoom -LogPath (Join-Path $logDir "ZoomInstall_$timestamp.log")
            if ($zoomSuccess) { Write-Host ' Done.' -ForegroundColor Green; $summaryInstalled.Add('Zoom Workplace') }
            else { Write-Host ' FAILED.' -ForegroundColor Red; $summaryFailed.Add('Install Zoom Workplace') }
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
            if ($ciscoSuccess) { Write-Host ' Done.' -ForegroundColor Green; $summaryInstalled.Add('Cisco Secure Client') }
            else { Write-Host ' FAILED.' -ForegroundColor Red; $summaryFailed.Add('Install Cisco Secure Client') }
        }
    }

    Write-Step 'PC Onboarding - Step 8: Windows Update'
    if ($installUpdatesInRun) {
        if ($WhatIf) {
            Write-Host '[WhatIf] Would check for and install Windows updates as part of the run.' -ForegroundColor DarkYellow
        }
        else {
            Write-Host 'Checking for Windows updates (this can take a few minutes)...'
            $pendingUpdates = @()
            try {
                $pendingUpdates = Get-PendingWindowsUpdate
                $updatesChecked = $true
            }
            catch {
                Write-Host "Couldn't check for updates: $($_.Exception.Message)" -ForegroundColor Red
                $summaryFailed.Add('Check for Windows updates')
            }

            if ($updatesChecked -and $pendingUpdates.Count -eq 0) {
                Write-Host 'Windows is already up to date.' -ForegroundColor Green
            }
            elseif ($pendingUpdates.Count -gt 0) {
                Write-Host "Found $($pendingUpdates.Count) update(s):"
                $pendingUpdates | ForEach-Object { Write-Host "  - $($_.Title)" }

                $updateTotal = $pendingUpdates.Count
                $updateIndex = 0
                foreach ($update in $pendingUpdates) {
                    $updateIndex++
                    Write-Progress -Activity 'Installing Windows updates' `
                        -Status "($updateIndex of $updateTotal) $($update.Title)" `
                        -PercentComplete ([int](($updateIndex - 1) / $updateTotal * 100))

                    Write-Host "Installing ($updateIndex of $updateTotal) $($update.Title)..." -NoNewline
                    $updateResult = Install-WindowsUpdateItem -Update $update
                    if ($updateResult.Succeeded) {
                        Write-Host ' Done.' -ForegroundColor Green
                        $summaryUpdates.Add($update.Title)
                    }
                    else {
                        Write-Host ' FAILED.' -ForegroundColor Red
                        $summaryFailed.Add("Install update: $($update.Title)")
                    }
                    if ($updateResult.RebootRequired) { $restartNeeded = $true }
                }
                Write-Progress -Activity 'Installing Windows updates' -Completed
            }
        }
    }
    else {
        if ($WhatIf) {
            Write-Host '[WhatIf] Would open Windows Update and start a check for updates.' -ForegroundColor DarkYellow
        }
        else {
            try {
                Start-WindowsUpdateScan
                Write-Host 'Windows Update is checking for updates in Settings and will download and install them on its own.' -ForegroundColor Green
                $updatesStarted = $true
            }
            catch {
                Write-Host "Couldn't start Windows Update: $($_.Exception.Message)" -ForegroundColor Red
                $summaryFailed.Add('Start Windows Update')
            }
        }
    }

    $runCompleted = $true
}
finally {
    # --- Ticket summary: printed, copied to the clipboard, and saved to Logs ---
    # Lives in finally so a run that dies partway still produces billable time
    # and a record of what it got through.
    $endTime = Get-Date
    $actualMinutes = [int][math]::Ceiling(($endTime - $startTime).TotalMinutes)
    $billedMinutes = [int]([math]::Ceiling(($endTime - $startTime).TotalMinutes / 15) * 15)

    $lines = [System.Collections.Generic.List[string]]::new()
    $pcName = if ($renamed) { "$newName (renamed from $env:COMPUTERNAME)" } else { $env:COMPUTERNAME }
    $lines.Add("PC Onboarding Summary - $pcName")
    $lines.Add("Start Time: $($startTime.ToString('M/d/yyyy h:mm tt'))")
    $lines.Add("End Time:   $($endTime.ToString('M/d/yyyy h:mm tt'))")
    $lines.Add("Runtime:    $(Format-Minutes $billedMinutes) (actual $(Format-Minutes $actualMinutes), rounded up to 15-min increments)")
    if ($WhatIf) { $lines.Add('Dry run (-WhatIf): no changes were made.') }
    if (-not $runCompleted) { $lines.Add('NOTE: The run stopped early because of an error. See the log for details.') }

    $removed = @($summaryRemoved | Select-Object -Unique)
    $installed = @($summaryInstalled | Select-Object -Unique)
    $failedItems = @($summaryFailed | Select-Object -Unique)

    if ($summaryJoined.Count -gt 0) {
        $lines.Add('')
        $lines.Add('Joined:')
        $summaryJoined | ForEach-Object { $lines.Add("  - $_") }
    }
    if ($removed.Count -gt 0) {
        $lines.Add('')
        $lines.Add("Removed ($($removed.Count)):")
        $removed | ForEach-Object { $lines.Add("  - $_") }
    }
    if ($installed.Count -gt 0) {
        $lines.Add('')
        $lines.Add("Installed ($($installed.Count)):")
        $installed | ForEach-Object { $lines.Add("  - $_") }
    }
    if ($summaryUpdates.Count -gt 0) {
        $lines.Add('')
        $lines.Add("Windows Updates installed ($($summaryUpdates.Count)):")
        $summaryUpdates | ForEach-Object { $lines.Add("  - $_") }
    }
    if ($failedItems.Count -gt 0) {
        $lines.Add('')
        $lines.Add("Failed ($($failedItems.Count)):")
        $failedItems | ForEach-Object { $lines.Add("  - $_") }
    }

    if ($updatesStarted) {
        $lines.Add('')
        $lines.Add('Windows Update: started at the end of the run; updates finish installing in the background.')
    }
    elseif ($updatesChecked -and @($pendingUpdates).Count -eq 0) {
        $lines.Add('')
        $lines.Add('Windows Update: already up to date.')
    }

    $summaryText = $lines -join [Environment]::NewLine
    $summaryPath = Join-Path $logDir "Summary_$timestamp.txt"
    Set-Content -Path $summaryPath -Value $summaryText

    Write-Host "`n==================== Ticket Summary ====================" -ForegroundColor Cyan
    Write-Host $summaryText
    Write-Host '========================================================' -ForegroundColor Cyan

    try {
        Set-Clipboard -Value $summaryText
        Write-Host 'Summary copied to the clipboard - paste it into the ticket.' -ForegroundColor Green
    }
    catch {
        Write-Host "Couldn't copy to the clipboard; the summary is saved at $summaryPath." -ForegroundColor Yellow
    }
    Write-Host "Summary saved to: $summaryPath"

    if ($restartNeeded) {
        Write-Host "`nRESTART REQUIRED to finish setting up this PC (updates, rename, or join)." -ForegroundColor Yellow
        if ($updatesStarted) {
            Write-Host 'Let Windows Update finish first, so one restart applies the updates and the rename/join together.' -ForegroundColor Yellow
        }
    }
    elseif ($updatesStarted) {
        Write-Host "`nWindows Update is still running and may ask for a restart when it's done." -ForegroundColor Yellow
    }

    Stop-Transcript | Out-Null
    Write-Host "`nLog saved to: $transcriptPath"
}
