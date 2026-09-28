#Requires -Version 5.1
<#
    DriverUpdates.psm1
    Manufacturer driver updates: Dell Command | Update on Dell PCs, Lenovo
    System Update on Lenovo PCs.

    If the vendor tool isn't already on the PC it's installed through winget.
    Neither vendor publishes a stable "latest" download link (both installer
    URLs change every release); winget resolves the current installer and
    verifies its hash. Both tools are then run limited to drivers, with
    reboots suppressed so nothing restarts the PC in the middle of onboarding.

    Update-* functions return Status (Updated, UpToDate, Ran, NotSupported,
    Failed), Detail, RebootRequired, and ToolInstalled (installed this run).
#>

function Get-DeviceManufacturer {
    [CmdletBinding()]
    param()

    $computer = Get-CimInstance -ClassName Win32_ComputerSystem
    $product = Get-CimInstance -ClassName Win32_ComputerSystemProduct

    $vendor = switch -Regex ($computer.Manufacturer) {
        '^Dell'  { 'Dell'; break }
        'Lenovo' { 'Lenovo'; break }
        default  { 'Other' }
    }

    [pscustomobject]@{
        Vendor       = $vendor
        Manufacturer = $computer.Manufacturer
        # Lenovo keeps the friendly name ("ThinkPad T14 Gen 4") in Version; its
        # Model field is the machine-type code. Everyone else uses Model.
        Model        = if ($vendor -eq 'Lenovo' -and $product.Version) { $product.Version } else { $computer.Model }
    }
}

function Install-WingetPackage {
    # Not exported.
    param([Parameter(Mandatory)] [string]$Id)

    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) {
        Write-Warning 'winget is not available for this account (App Installer missing or not registered yet).'
        return $false
    }

    $wingetArgs = @('install', '--id', $Id, '--exact', '--silent',
        '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity')
    $proc = Start-Process -FilePath $winget.Source -ArgumentList $wingetArgs -Wait -PassThru -WindowStyle Hidden
    return $proc.ExitCode -eq 0
}

function New-DriverResult {
    # Not exported.
    param([string]$Status, [string]$Detail, [bool]$RebootRequired = $false, [bool]$ToolInstalled = $false)
    [pscustomobject]@{ Status = $Status; Detail = $Detail; RebootRequired = $RebootRequired; ToolInstalled = $ToolInstalled }
}

function ConvertFrom-DcuExitCode {
    # Not exported. Codes from the Dell Command | Update 5.x reference guide.
    param([int]$ExitCode, [bool]$ToolInstalled)

    switch ($ExitCode) {
        0       { New-DriverResult 'Updated' 'Driver updates installed.' $false $ToolInstalled }
        1       { New-DriverResult 'Updated' 'Driver updates installed; a restart is required.' $true $ToolInstalled }
        500     { New-DriverResult 'UpToDate' 'No driver updates were needed.' $false $ToolInstalled }
        7       { New-DriverResult 'NotSupported' 'Dell Command | Update does not support this model (consumer models use Dell SupportAssist).' $false $ToolInstalled }
        5       { New-DriverResult 'Failed' 'A restart is pending from an earlier update; restart and run Dell Command | Update again.' $false $ToolInstalled }
        default { New-DriverResult 'Failed' "Dell Command | Update exited with code $ExitCode." $false $ToolInstalled }
    }
}

function Update-DellDrivers {
    [CmdletBinding()]
    param(
        # Must end in .log - Dell Command | Update rejects anything else.
        [string]$LogPath
    )

    $findCli = {
        @("$env:ProgramFiles\Dell\CommandUpdate\dcu-cli.exe",
          "${env:ProgramFiles(x86)}\Dell\CommandUpdate\dcu-cli.exe") |
            Where-Object { Test-Path $_ } | Select-Object -First 1
    }

    $toolInstalled = $false
    $cli = & $findCli
    if (-not $cli) {
        $null = Install-WingetPackage -Id 'Dell.CommandUpdate.Universal'
        $cli = & $findCli
        if (-not $cli) {
            return New-DriverResult 'Failed' 'Could not install Dell Command | Update (via winget).'
        }
        $toolInstalled = $true
    }

    $dcuArgs = @('/applyUpdates', '-updateType=driver', '-reboot=disable', '-silent')
    if ($LogPath) { $dcuArgs += "-outputLog=`"$LogPath`"" }

    Write-Verbose "Running: $cli $dcuArgs"
    $proc = Start-Process -FilePath $cli -ArgumentList $dcuArgs -Wait -PassThru -WindowStyle Hidden
    return ConvertFrom-DcuExitCode -ExitCode $proc.ExitCode -ToolInstalled $toolInstalled
}

function Update-LenovoDrivers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Model
    )

    # System Update only covers Lenovo's commercial "Think" lines; IdeaPad,
    # Yoga, Legion etc. get drivers through Lenovo Vantage instead.
    if ($Model -notmatch 'Think') {
        return New-DriverResult 'NotSupported' "Lenovo System Update doesn't support $Model (consumer Lenovo models use Lenovo Vantage)."
    }

    $tvsu = "${env:ProgramFiles(x86)}\Lenovo\System Update\tvsu.exe"
    $toolInstalled = $false
    if (-not (Test-Path $tvsu)) {
        $null = Install-WingetPackage -Id 'Lenovo.SystemUpdate'
        if (-not (Test-Path $tvsu)) {
            return New-DriverResult 'Failed' 'Could not install Lenovo System Update (via winget).'
        }
        $toolInstalled = $true
    }

    # -packagetypes 2 = drivers only. -includerebootpackages 3 adds drivers that
    # need a normal restart (suppressed by -noreboot) and leaves out types 1/4/5,
    # which force a reboot or shutdown themselves and would kill this run.
    $suArgs = '/CM -search A -action INSTALL -packagetypes 2 -includerebootpackages 3 -noreboot -noicon -nolicense'

    Write-Verbose "Running: $tvsu $suArgs"
    Start-Process -FilePath $tvsu -ArgumentList $suArgs -Wait | Out-Null

    # System Update has no documented exit codes, so success can't be confirmed
    # here. Drivers it installs may need a restart, so recommend one.
    return New-DriverResult 'Ran' 'Lenovo System Update checked for and installed driver updates.' $true $toolInstalled
}

Export-ModuleMember -Function Get-DeviceManufacturer, Update-DellDrivers, Update-LenovoDrivers
