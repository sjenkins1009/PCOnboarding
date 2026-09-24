#Requires -Version 5.1
<#
    DeviceJoin.psm1
    Renames the PC and joins it to an Active Directory domain, Microsoft
    Entra ID, or both (hybrid) at the start of onboarding.

    Entra join has no command-line path for a signing-in user: dsregcmd /join
    only performs *hybrid* join (it proves identity with the AD computer
    account), and a fully unattended Entra join needs a pre-built
    provisioning package or Autopilot. So Entra join opens the Settings join
    screen for the tech to sign in, and the caller confirms the result after.
#>

function Get-JoinStatus {
    [CmdletBinding()]
    param()

    $computer = Get-CimInstance -ClassName Win32_ComputerSystem
    $dsreg = & dsregcmd.exe /status 2>$null
    $tenantLine = $dsreg | Where-Object { $_ -match '^\s*TenantName\s*:' } | Select-Object -First 1

    [pscustomobject]@{
        DomainJoined = [bool]$computer.PartOfDomain
        Domain       = if ($computer.PartOfDomain) { $computer.Domain } else { $null }
        EntraJoined  = [bool]($dsreg | Where-Object { $_ -match '^\s*AzureAdJoined\s*:\s*YES' })
        TenantName   = if ($tenantLine) { ($tenantLine -split ':', 2)[1].Trim() } else { $null }
    }
}

function Test-JoinSupportedEdition {
    # Windows Home can't join a domain or Microsoft Entra ID. That's common on
    # retail consumer PCs, which have to be upgraded to Pro first.
    return (Get-CimInstance -ClassName Win32_OperatingSystem).Caption -notmatch 'Home'
}

function Join-ADDomain {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DomainName,

        [Parameter(Mandatory)]
        [pscredential]$Credential,

        # Renaming separately and then joining before a restart would join
        # under the old name, so a rename rides along in the same call.
        [string]$NewName
    )

    $params = @{
        DomainName  = $DomainName
        Credential  = $Credential
        Force       = $true
        ErrorAction = 'Stop'
    }
    if ($NewName) { $params.NewName = $NewName }

    try {
        # No -Restart: the rest of onboarding runs first, and the join takes
        # effect on the next restart.
        Add-Computer @params
        return $true
    }
    catch {
        Write-Warning "Domain join failed: $($_.Exception.Message)"
        return $false
    }
}

function Rename-Device {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$NewName,

        # Required when the PC is already domain joined: the rename also
        # renames its computer account in Active Directory.
        [pscredential]$DomainCredential
    )

    $params = @{
        NewName     = $NewName
        Force       = $true
        ErrorAction = 'Stop'
    }
    if ($DomainCredential) { $params.DomainCredential = $DomainCredential }

    try {
        Rename-Computer @params
        return $true
    }
    catch {
        Write-Warning "Rename failed: $($_.Exception.Message)"
        return $false
    }
}

function Start-EntraJoin {
    # Settings > Accounts > Access work or school, where "Connect" offers
    # "Join this device to Microsoft Entra ID".
    Start-Process 'ms-settings:workplace'
}

Export-ModuleMember -Function Get-JoinStatus, Test-JoinSupportedEdition, Join-ADDomain, Rename-Device, Start-EntraJoin
