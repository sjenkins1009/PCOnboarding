#Requires -Version 5.1
<#
    OptionalAppInstalls.psm1
    Silent, unattended installers for apps that are NOT installed by
    default during onboarding - only run when explicitly requested:
    Dropbox, Slack, Google Drive, and Cisco Secure Client.
#>

function Install-Dropbox {
    <#
        Downloads Dropbox's official offline installer (the "?full=1"
        query on their download endpoint always redirects to the current
        full standalone .exe, so this URL doesn't need version-chasing)
        and installs it silently without launching the app afterward.
    #>
    [CmdletBinding()]
    param(
        [string]$DownloadDir = (Join-Path $env:ProgramData 'PCOnboarding\Downloads'),
        [string]$Url = 'https://www.dropbox.com/download?full=1&os=win'
    )

    if (-not (Test-Path $DownloadDir)) { New-Item -Path $DownloadDir -ItemType Directory -Force | Out-Null }
    $exePath = Join-Path $DownloadDir 'DropboxOfflineInstaller.exe'

    try {
        Write-Verbose "Downloading Dropbox from $Url"
        $prevProgressPref = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $Url -OutFile $exePath -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = $prevProgressPref
    }
    catch {
        Write-Warning "Failed to download Dropbox: $_"
        return $false
    }

    if (-not (Test-Path $exePath) -or (Get-Item $exePath).Length -eq 0) {
        Write-Warning 'Dropbox download appears empty or missing.'
        return $false
    }

    Write-Verbose "Running: $exePath /NOLAUNCH"
    $proc = Start-Process -FilePath $exePath -ArgumentList '/NOLAUNCH' -Wait -PassThru
    return ($proc.ExitCode -eq 0)
}

function Install-Slack {
    <#
        Downloads Slack's official current Windows MSIX package (Slack's
        own deployment docs point at this API redirect rather than a
        version-pinned filename) and installs it via Add-AppxPackage -
        Slack has moved off a traditional MSI to MSIX for Windows.
    #>
    [CmdletBinding()]
    param(
        [string]$DownloadDir = (Join-Path $env:ProgramData 'PCOnboarding\Downloads'),
        [string]$Url = 'https://slack.com/api/desktop.latestRelease?arch=x64&variant=msix&redirect=true'
    )

    if (-not (Test-Path $DownloadDir)) { New-Item -Path $DownloadDir -ItemType Directory -Force | Out-Null }
    $msixPath = Join-Path $DownloadDir 'Slack.msix'

    try {
        Write-Verbose "Downloading Slack from $Url"
        $prevProgressPref = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $Url -OutFile $msixPath -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = $prevProgressPref
    }
    catch {
        Write-Warning "Failed to download Slack: $_"
        return $false
    }

    if (-not (Test-Path $msixPath) -or (Get-Item $msixPath).Length -eq 0) {
        Write-Warning 'Slack download appears empty or missing.'
        return $false
    }

    try {
        Write-Verbose "Running: Add-AppxPackage -Path $msixPath"
        Add-AppxPackage -Path $msixPath -ErrorAction Stop
        return $true
    }
    catch {
        Write-Warning "Failed to install Slack: $_"
        return $false
    }
}

function Install-GoogleDrive {
    <#
        Downloads Google's official "Drive for desktop" installer and
        installs it silently without a post-install launch.
    #>
    [CmdletBinding()]
    param(
        [string]$DownloadDir = (Join-Path $env:ProgramData 'PCOnboarding\Downloads'),
        [string]$Url = 'https://dl.google.com/drive-file-stream/GoogleDriveFSSetup.exe'
    )

    if (-not (Test-Path $DownloadDir)) { New-Item -Path $DownloadDir -ItemType Directory -Force | Out-Null }
    $exePath = Join-Path $DownloadDir 'GoogleDriveFSSetup.exe'

    try {
        Write-Verbose "Downloading Google Drive from $Url"
        $prevProgressPref = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $Url -OutFile $exePath -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = $prevProgressPref
    }
    catch {
        Write-Warning "Failed to download Google Drive: $_"
        return $false
    }

    if (-not (Test-Path $exePath) -or (Get-Item $exePath).Length -eq 0) {
        Write-Warning 'Google Drive download appears empty or missing.'
        return $false
    }

    $arguments = '--silent --desktop_shortcut --skip_launch_new'
    Write-Verbose "Running: $exePath $arguments"
    $proc = Start-Process -FilePath $exePath -ArgumentList $arguments -Wait -PassThru
    return ($proc.ExitCode -eq 0)
}

function Install-CiscoSecureClient {
    <#
        Cisco Secure Client has no public download - Cisco gates it behind
        a Cisco.com account and your org's own entitlement/VPN headend, so
        there is no generic URL to fetch it from. This installs a locally
        supplied MSI instead of downloading one.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$InstallerPath,

        [string]$LogPath
    )

    if (-not (Test-Path $InstallerPath)) {
        Write-Warning "Cisco Secure Client installer not found at $InstallerPath."
        return $false
    }

    $logArg = if ($LogPath) { "/l*v `"$LogPath`"" } else { '' }
    $arguments = "/i `"$InstallerPath`" /qn /norestart $logArg"

    Write-Verbose "Running: msiexec.exe $arguments"
    $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru
    return ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010)
}

Export-ModuleMember -Function Install-Dropbox, Install-Slack, Install-GoogleDrive, Install-CiscoSecureClient
