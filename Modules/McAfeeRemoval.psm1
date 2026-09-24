#Requires -Version 5.1
<#
    McAfeeRemoval.psm1
    Detects and removes the McAfee security suite that ships preloaded on
    most consumer/OEM machines (Total Protection, LiveSafe, WebAdvisor,
    Safe Connect, Security Scan Plus, and the Store "McAfee Security" app).

    Two-stage, matching how McAfee actually uninstalls in the field:
      1. Silent removal via each product's own registered uninstaller.
      2. MCPR (McAfee Consumer Product Removal) as a cleanup pass for the
         leftovers those uninstallers routinely strand - orphaned services,
         drivers and registry keys that block a clean reinstall.
#>

# MCPR is fetched straight from McAfee's own CDN.
#
# A shortened alias (mcpr.notlong.com) circulates in older removal guides and
# currently 302s to this exact path. It is deliberately NOT used here: it is a
# wildcard host on third-party infrastructure, so the alias can be repointed at
# any binary at any time, and this script executes whatever comes back with
# SYSTEM-level rights. Pointing at McAfee directly fetches the identical file
# with nobody in the middle.
$script:McprUrl = 'https://download.mcafee.com/products/licensed/cust_support_patches/MCPR.exe'

# Authenticode subject must match this before the downloaded binary is run.
$script:McprExpectedPublisher = 'McAfee'

function Get-InstalledMcAfee {
    <#
        Scans registry uninstall keys (by display name *and* publisher, so
        unbranded entries like "WebAdvisor" and "Safe Connect" are caught),
        plus AppX packages for the Store-delivered McAfee app.
        Returns one object per detected product.
    #>
    [CmdletBinding()]
    param()

    $results = [System.Collections.Generic.List[object]]::new()

    # Some OEM preloads register the component without "McAfee" in the display
    # name, so the publisher field is the reliable signal for those.
    $namePattern      = 'McAfee|WebAdvisor|Safe ?Connect|LiveSafe|Total Protection'
    $publisherPattern = 'McAfee'

    $uninstallKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    Get-ItemProperty -Path $uninstallKeys -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -and
            ($_.DisplayName -match $namePattern -or $_.Publisher -match $publisherPattern)
        } |
        ForEach-Object {
            # A GUID-shaped key means a real MSI product we can hand to
            # msiexec; anything else has to go through its own recorded
            # uninstall string (McAfee's suites use custom uninstallers).
            $type = if ($_.PSChildName -match '^\{.*\}$') { 'MSI' } else { 'Uninstaller' }

            # QuietUninstallString, when a product bothers to register one, is
            # the vendor's own already-silent command line - prefer it.
            $uninstall = if ($_.QuietUninstallString) { $_.QuietUninstallString } else { $_.UninstallString }

            $results.Add([pscustomobject]@{
                Name            = $_.DisplayName
                Type            = $type
                ProductId       = $_.PSChildName
                UninstallString = $uninstall
            })
        }

    # --- Store-delivered "McAfee Security" app ---
    # Ships as a provisioned package on many OEM images, so it silently
    # reappears for every new user profile unless the provisioned copy goes too.
    Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match 'McAfee' } |
        ForEach-Object {
            $results.Add([pscustomobject]@{
                Name            = "$($_.Name) ($($_.PackageFullName))"
                Type            = 'AppX'
                ProductId       = $_.PackageFullName
                UninstallString = $null
            })
        }

    Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match 'McAfee' } |
        ForEach-Object {
            $results.Add([pscustomobject]@{
                Name            = "$($_.DisplayName) (provisioned, $($_.PackageName))"
                Type            = 'AppXProvisioned'
                ProductId       = $_.PackageName
                UninstallString = $null
            })
        }

    # Removal order matters: the bolt-on components (WebAdvisor, Safe Connect)
    # are owned by the main suite's installer, so taking the suite out first
    # leaves them stranded with a dead uninstall string. Suite goes last.
    # @() so a zero- or one-result scan still returns something the caller can
    # .Count and foreach over, the way the List does in the other modules.
    $suitePattern = 'Total Protection|LiveSafe|Internet Security|Antivirus|Personal Security'
    return @($results | Sort-Object @{ Expression = { $_.Name -match $suitePattern } })
}

function Remove-McAfeeInstallation {
    <#
        Silently removes a single detected McAfee product
        (from Get-InstalledMcAfee).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Product,

        [string]$LogPath
    )

    switch ($Product.Type) {

        'MSI' {
            $logArg = if ($LogPath) { "/l*v `"$LogPath`"" } else { '' }
            $arguments = "/x `"$($Product.ProductId)`" /qn /norestart $logArg"
            $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru
            return ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010)
        }

        'Uninstaller' {
            if (-not $Product.UninstallString) {
                Write-Warning "No uninstall string recorded for $($Product.Name)."
                return $false
            }

            $cmd = $Product.UninstallString
            # McAfee's own uninstallers are silent only when told to be, and
            # they don't share one convention - add whichever flag the recorded
            # command line is missing rather than assuming.
            if ($cmd -notmatch '/quiet|/qn|/silent|/S\b') { $cmd += ' /quiet' }

            Write-Verbose "Running: $cmd"
            $proc = Start-Process -FilePath 'cmd.exe' -ArgumentList "/c $cmd" -Wait -PassThru -WindowStyle Hidden
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

function Invoke-McAfeeRemovalTool {
    <#
        Downloads MCPR (McAfee Consumer Product Removal) from McAfee's CDN,
        verifies it is signed by McAfee, and runs it.

        NOTE: MCPR is a GUI wizard and gates itself behind a CAPTCHA, so this
        step cannot run unattended - it needs someone at the keyboard. It is
        therefore a cleanup pass *after* the silent uninstalls above, not the
        primary removal path.
    #>
    [CmdletBinding()]
    param(
        [string]$DownloadDir = (Join-Path $env:ProgramData 'PCOnboarding\Downloads'),
        [string]$Url = $script:McprUrl,
        [switch]$SkipSignatureCheck
    )

    if (-not (Test-Path $DownloadDir)) { New-Item -Path $DownloadDir -ItemType Directory -Force | Out-Null }
    $exePath = Join-Path $DownloadDir 'MCPR.exe'

    try {
        Write-Verbose "Downloading MCPR from $Url"
        $prevProgressPref = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $Url -OutFile $exePath -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = $prevProgressPref
    }
    catch {
        Write-Warning "Failed to download MCPR: $_"
        return $false
    }

    if (-not (Test-Path $exePath) -or (Get-Item $exePath).Length -eq 0) {
        Write-Warning 'MCPR download appears empty or missing.'
        return $false
    }

    # Nothing downloaded gets executed with admin rights until the signature
    # says McAfee actually produced it. A redirect that quietly starts serving
    # something else fails here instead of running.
    if (-not $SkipSignatureCheck) {
        $signature = Get-AuthenticodeSignature -FilePath $exePath

        if ($signature.Status -ne 'Valid') {
            Write-Warning "MCPR signature is not valid (status: $($signature.Status)). Refusing to run it."
            return $false
        }

        if ($signature.SignerCertificate.Subject -notmatch $script:McprExpectedPublisher) {
            Write-Warning "MCPR is signed by an unexpected publisher: $($signature.SignerCertificate.Subject). Refusing to run it."
            return $false
        }

        Write-Verbose "MCPR signature verified: $($signature.SignerCertificate.Subject)"
    }

    Write-Verbose "Running: $exePath"
    $proc = Start-Process -FilePath $exePath -Wait -PassThru

    # MCPR uses 0 for a clean run and 3010 to mean "done, needs a reboot".
    return ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010)
}

Export-ModuleMember -Function Get-InstalledMcAfee, Remove-McAfeeInstallation, Invoke-McAfeeRemovalTool
