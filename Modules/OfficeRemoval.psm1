#Requires -Version 5.1
<#
    OfficeRemoval.psm1
    Detects and silently removes Microsoft Office installations
    (Click-to-Run and MSI-based) as part of PC onboarding.
#>

function Get-InstalledOffice {
    <#
        Scans registry uninstall keys and the Click-to-Run configuration
        for any Microsoft Office / Microsoft 365 / Visio / Project installs.
        Returns one object per detected product.
    #>
    [CmdletBinding()]
    param()

    $results = [System.Collections.Generic.List[object]]::new()

    # Products/keywords we want to catch, and non-Office Microsoft products
    # to explicitly exclude so we never touch things like Edge, OneDrive, etc.
    $includePattern = 'Microsoft (Office|365 Apps|Visio|Project|OneNote)\b'
    $excludePattern = 'OneDrive|Edge|Visual Studio|\.NET|Teams|SQL Server|PowerShell|Silverlight|Skype|Windows |Runtime Redistributable|Visual C\+\+|Update Health'

    # Newer suite installs brand themselves just "Microsoft 365" (no "Apps"),
    # and register one Programs & Features entry per installed language, e.g.
    # "Microsoft 365 - en-us", "Microsoft 365 - es-es". This is intentionally
    # a separate, narrow pattern (exact "Microsoft 365 - <locale>" shape)
    # rather than broadening $includePattern to bare "365\b", which would
    # also catch unrelated products like "Microsoft 365 Copilot".
    $languagePackPattern = '^Microsoft 365 - [a-z]{2}(-[a-z]{2})?$'

    # --- 1. Click-to-Run installs (Office 365 / 2019 / 2021 / 2024) ---
    $c2rConfigPath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path $c2rConfigPath) {
        $c2r = Get-ItemProperty -Path $c2rConfigPath -ErrorAction SilentlyContinue
        if ($c2r.ProductReleaseIds) {
            $productIds = $c2r.ProductReleaseIds -split ','
            foreach ($id in $productIds) {
                $results.Add([pscustomobject]@{
                    Name            = "$id (Click-to-Run)"
                    Type            = 'ClickToRun'
                    ProductId       = $id
                    Platform        = $c2r.Platform
                    ClientCulture   = $c2r.ClientCulture
                    UninstallString = $null
                })
            }
        }
    }

    # --- 2. MSI-based installs (Office 2007/2010/2013/2016 non-C2R) ---
    $uninstallKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    Get-ItemProperty -Path $uninstallKeys -ErrorAction SilentlyContinue |
        Where-Object {
            ($_.DisplayName -match $includePattern -or $_.DisplayName -match $languagePackPattern) -and
            $_.DisplayName -notmatch $excludePattern
        } |
        ForEach-Object {
            # Click-to-Run registers one Programs & Features entry per
            # installed language for some products (e.g. the standalone free
            # OneNote shows as "Microsoft OneNote - es-es", "... - pt-br",
            # etc.). Those aren't real MSI products - their UninstallString
            # already invokes OfficeClickToRun.exe - so they need the C2R
            # removal path, not msiexec.
            $type = if ($_.UninstallString -match 'OfficeClickToRun\.exe') { 'ClickToRunARP' } else { 'MSI' }
            $results.Add([pscustomobject]@{
                Name            = $_.DisplayName
                Type            = $type
                ProductId       = $_.PSChildName
                Platform        = $null
                ClientCulture   = $null
                UninstallString = $_.UninstallString
            })
        }

    # --- 3. "OneNote for Windows 10" UWP app ---
    # This is a Store/AppX package, not a registry Uninstall entry, and it can
    # be installed per-user (each profile gets its own copy - this is usually
    # what shows up as "multiple installs") plus a machine-wide provisioned
    # copy that silently reinstalls it for any new user profile.
    $oneNoteAppxName = 'Microsoft.Office.OneNote'

    Get-AppxPackage -AllUsers -Name $oneNoteAppxName -ErrorAction SilentlyContinue |
        ForEach-Object {
            $results.Add([pscustomobject]@{
                Name            = "OneNote for Windows 10 ($($_.PackageFullName))"
                Type            = 'AppX'
                ProductId       = $_.PackageFullName
                Platform        = $null
                ClientCulture   = $null
                UninstallString = $null
            })
        }

    Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -eq $oneNoteAppxName } |
        ForEach-Object {
            $results.Add([pscustomobject]@{
                Name            = "OneNote for Windows 10 (provisioned, $($_.PackageName))"
                Type            = 'AppXProvisioned'
                ProductId       = $_.PackageName
                Platform        = $null
                ClientCulture   = $null
                UninstallString = $null
            })
        }

    return $results
}

function Remove-OfficeInstallation {
    <#
        Silently removes a single detected Office product (from Get-InstalledOffice).
        Returns $true/$false for success, and writes progress via Write-Progress
        (caller supplies the outer progress context).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Product,

        [string]$LogPath
    )

    switch ($Product.Type) {

        'ClickToRun' {
            $c2rExe = Join-Path $env:CommonProgramFiles 'Microsoft Shared\ClickToRun\OfficeClickToRun.exe'
            if (-not (Test-Path $c2rExe)) {
                Write-Warning "OfficeClickToRun.exe not found; cannot remove $($Product.Name)."
                return $false
            }

            $culture = if ($Product.ClientCulture) { $Product.ClientCulture } else { 'en-us' }
            $arguments = @(
                'scenario=install'
                'scenariosubtype=ARP'
                'sourcetype=None'
                "productstoremove=$($Product.ProductId)"
                "culture=$culture"
                'DisplayLevel=False'
            ) -join ' '

            Write-Verbose "Running: $c2rExe $arguments"
            $proc = Start-Process -FilePath $c2rExe -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
            return ($proc.ExitCode -eq 0)
        }

        'ClickToRunARP' {
            # A per-language (or otherwise standalone) Click-to-Run entry -
            # its UninstallString already builds the correct
            # OfficeClickToRun.exe command line, so just run it, forcing
            # silent mode instead of guessing productstoremove/culture ourselves.
            if (-not $Product.UninstallString) {
                Write-Warning "No uninstall string recorded for $($Product.Name)."
                return $false
            }

            $cmd = $Product.UninstallString
            if ($cmd -match 'DisplayLevel=True') {
                $cmd = $cmd -replace 'DisplayLevel=True', 'DisplayLevel=False'
            }
            elseif ($cmd -notmatch 'DisplayLevel=False') {
                $cmd += ' DisplayLevel=False'
            }

            Write-Verbose "Running: $cmd"
            $proc = Start-Process -FilePath 'cmd.exe' -ArgumentList "/c $cmd" -Wait -PassThru -WindowStyle Hidden
            return ($proc.ExitCode -eq 0)
        }

        'MSI' {
            # Prefer the GUID-based product code if we have one.
            $productCode = $Product.ProductId
            $logArg = if ($LogPath) { "/l*v `"$LogPath`"" } else { '' }

            if ($productCode -match '^\{.*\}$') {
                $arguments = "/x `"$productCode`" /qn /norestart $logArg"
                $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru
                return ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010)
            }
            elseif ($Product.UninstallString) {
                # Fall back to the recorded uninstall string, forced silent.
                $uninstall = $Product.UninstallString -replace '/I', '/X'
                if ($uninstall -notmatch '/qn') { $uninstall += ' /qn /norestart' }
                $proc = Start-Process -FilePath 'cmd.exe' -ArgumentList "/c $uninstall" -Wait -PassThru
                return ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010)
            }
            else {
                Write-Warning "No usable uninstall method for $($Product.Name)."
                return $false
            }
        }

        'AppX' {
            try {
                Remove-AppxPackage -Package $Product.ProductId -AllUsers -ErrorAction Stop
                return $true
            }
            catch {
                Write-Warning "Failed to remove AppX package $($Product.ProductId): $_"
                return $false
            }
        }

        'AppXProvisioned' {
            # Removes the provisioned package so it stops reinstalling for new user profiles.
            try {
                Remove-AppxProvisionedPackage -Online -PackageName $Product.ProductId -ErrorAction Stop | Out-Null
                return $true
            }
            catch {
                Write-Warning "Failed to remove provisioned AppX package $($Product.ProductId): $_"
                return $false
            }
        }

        default {
            Write-Warning "Unknown product type for $($Product.Name)."
            return $false
        }
    }
}

Export-ModuleMember -Function Get-InstalledOffice, Remove-OfficeInstallation
