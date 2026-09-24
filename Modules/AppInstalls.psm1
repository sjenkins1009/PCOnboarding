#Requires -Version 5.1
<#
    AppInstalls.psm1
    Silent, unattended downloads/installs for standard onboarding software:
    Google Chrome Enterprise and Adobe Acrobat Reader.
#>

function Install-ChromeEnterprise {
    <#
        Downloads the official Google Chrome Enterprise standalone MSI and
        installs it silently.
    #>
    [CmdletBinding()]
    param(
        [string]$DownloadDir = (Join-Path $env:ProgramData 'PCOnboarding\Downloads'),
        [string]$Url = 'https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi',
        [string]$LogPath
    )

    if (-not (Test-Path $DownloadDir)) { New-Item -Path $DownloadDir -ItemType Directory -Force | Out-Null }
    $msiPath = Join-Path $DownloadDir 'GoogleChromeStandaloneEnterprise64.msi'

    try {
        Write-Verbose "Downloading Chrome Enterprise from $Url"
        $prevProgressPref = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $Url -OutFile $msiPath -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = $prevProgressPref
    }
    catch {
        Write-Warning "Failed to download Google Chrome Enterprise: $_"
        return $false
    }

    if (-not (Test-Path $msiPath) -or (Get-Item $msiPath).Length -eq 0) {
        Write-Warning 'Chrome Enterprise download appears empty or missing.'
        return $false
    }

    $logArg = if ($LogPath) { "/l*v `"$LogPath`"" } else { '' }
    $arguments = "/i `"$msiPath`" /qn /norestart $logArg"

    Write-Verbose "Running: msiexec.exe $arguments"
    $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru
    return ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010)
}

function Install-AdobeReader {
    <#
        Downloads Adobe's official Acrobat Reader DC web-installer package
        (a small, permanently-static bootstrapper Adobe hosts precisely so
        scripts don't need to chase version-specific filenames) and installs
        it silently.
    #>
    [CmdletBinding()]
    param(
        [string]$DownloadDir = (Join-Path $env:ProgramData 'PCOnboarding\Downloads'),
        [string]$Url = 'https://trials.adobe.com/AdobeProducts/APRO/Acrobat_HelpX/win32/Acrobat_DC_Web_WWMUI.zip'
    )

    if (-not (Test-Path $DownloadDir)) { New-Item -Path $DownloadDir -ItemType Directory -Force | Out-Null }
    $zipPath = Join-Path $DownloadDir 'Acrobat_DC_Web_WWMUI.zip'
    $extractDir = Join-Path $DownloadDir 'AdobeReaderInstall'

    try {
        Write-Verbose "Downloading Adobe Reader from $Url"
        $prevProgressPref = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $Url -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = $prevProgressPref
    }
    catch {
        Write-Warning "Failed to download Adobe Reader: $_"
        return $false
    }

    if (-not (Test-Path $zipPath) -or (Get-Item $zipPath).Length -eq 0) {
        Write-Warning 'Adobe Reader download appears empty or missing.'
        return $false
    }

    try {
        if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
        Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to extract Adobe Reader installer: $_"
        return $false
    }

    $setupExe = Get-ChildItem -Path $extractDir -Filter 'setup.exe' -Recurse | Select-Object -First 1
    if (-not $setupExe) {
        Write-Warning 'setup.exe not found in extracted Adobe Reader package.'
        return $false
    }

    # /sAll = suppress all UI, /rs = suppress any reboot, /msi ... = passed
    # through to the underlying msiexec (EULA_ACCEPT avoids the EULA prompt,
    # /qn keeps the MSI phase silent too).
    $arguments = '/sAll /rs /msi EULA_ACCEPT=YES /qn'

    Write-Verbose "Running: $($setupExe.FullName) $arguments"
    $proc = Start-Process -FilePath $setupExe.FullName -ArgumentList $arguments -Wait -PassThru
    return ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010)
}

Export-ModuleMember -Function Install-ChromeEnterprise, Install-AdobeReader
