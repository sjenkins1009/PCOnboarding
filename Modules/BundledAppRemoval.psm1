#Requires -Version 5.1
<#
    BundledAppRemoval.psm1
    Detects and silently removes the consumer/preloaded Microsoft apps that
    ship on new machines and get replaced by the licensed Microsoft 365
    deployment: Teams (classic MSI + new AppX) and the new Outlook app.
    OneDrive is intentionally left alone.
#>

function Get-InstalledBundledApps {
    <#
        Scans for Teams (classic and new) and the new Outlook app.
        Note: classic per-user Teams is detected/removed for the
        *currently running* user profile only - there is no per-machine
        "-AllUsers" equivalent for that installer the way there is for
        AppX packages. This matches the typical onboarding scenario of a
        single admin profile prepping a fresh machine.
    #>
    [CmdletBinding()]
    param()

    $results = [System.Collections.Generic.List[object]]::new()

    # --- Teams: classic "Teams Machine-Wide Installer" MSI ---
    $uninstallKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    Get-ItemProperty -Path $uninstallKeys -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -eq 'Teams Machine-Wide Installer' } |
        ForEach-Object {
            $results.Add([pscustomobject]@{
                Name            = $_.DisplayName
                Type            = 'MSI'
                ProductId       = $_.PSChildName
                UninstallString = $_.UninstallString
            })
        }

    # --- Teams: classic per-user install (current user only) ---
    $teamsUpdateExe = Join-Path $env:LOCALAPPDATA 'Microsoft\Teams\Update.exe'
    if (Test-Path $teamsUpdateExe) {
        $results.Add([pscustomobject]@{
            Name            = 'Microsoft Teams (classic, current user)'
            Type            = 'TeamsClassicPerUser'
            ProductId       = $teamsUpdateExe
            UninstallString = $null
        })
    }

    # --- Teams (new) / Outlook (new): AppX-based clients, machine-wide ---
    $appxTargets = @(
        @{ PackageName = 'MSTeams';                    Label = 'Microsoft Teams (new)' }
        @{ PackageName = 'MicrosoftTeams';              Label = 'Microsoft Teams (Windows 11 chat)' }
        @{ PackageName = 'Microsoft.OutlookForWindows'; Label = 'Microsoft Outlook (new)' }
    )

    foreach ($target in $appxTargets) {
        Get-AppxPackage -AllUsers -Name $target.PackageName -ErrorAction SilentlyContinue |
            ForEach-Object {
                $results.Add([pscustomobject]@{
                    Name            = "$($target.Label) ($($_.PackageFullName))"
                    Type            = 'AppX'
                    ProductId       = $_.PackageFullName
                    UninstallString = $null
                })
            }

        Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -eq $target.PackageName } |
            ForEach-Object {
                $results.Add([pscustomobject]@{
                    Name            = "$($target.Label) (provisioned, $($_.PackageName))"
                    Type            = 'AppXProvisioned'
                    ProductId       = $_.PackageName
                    UninstallString = $null
                })
            }
    }

    return $results
}

function Remove-BundledAppInstallation {
    <#
        Silently removes a single detected app (from Get-InstalledBundledApps).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Product,

        [string]$LogPath
    )

    switch ($Product.Type) {

        'TeamsClassicPerUser' {
            $proc = Start-Process -FilePath $Product.ProductId -ArgumentList @('--uninstall', '-s') -Wait -PassThru
            return ($proc.ExitCode -eq 0)
        }

        'MSI' {
            $logArg = if ($LogPath) { "/l*v `"$LogPath`"" } else { '' }
            $arguments = "/x `"$($Product.ProductId)`" /qn /norestart $logArg"
            $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru
            return ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010)
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

Export-ModuleMember -Function Get-InstalledBundledApps, Remove-BundledAppInstallation
