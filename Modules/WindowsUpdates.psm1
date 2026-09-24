#Requires -Version 5.1
<#
    WindowsUpdates.psm1
    Windows Update at the end of onboarding, two ways:
      - Get-PendingWindowsUpdate / Install-WindowsUpdateItem install updates
        inside the run through the built-in Windows Update Agent COM API (no
        add-on modules), so the time and results land in the ticket summary.
      - Start-WindowsUpdateScan just triggers "Check for updates" and lets
        Windows finish in the background after the run ends.
#>

# Updates must be downloaded/installed through the same session that found them.
$script:UpdateSession = $null

function Get-PendingWindowsUpdate {
    <#
        Returns the updates Windows would install automatically: not
        installed, not hidden, and not optional (BrowseOnly=0). Feature
        upgrades (e.g. 24H2 -> 25H2) are skipped - they can take over an hour
        and are a separate decision from routine patching.
    #>
    [CmdletBinding()]
    param()

    $script:UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $searcher = $script:UpdateSession.CreateUpdateSearcher()
    $result = $searcher.Search('IsInstalled=0 and IsHidden=0 and BrowseOnly=0')

    $updates = foreach ($update in $result.Updates) {
        $isUpgrade = $false
        foreach ($category in $update.Categories) {
            if ($category.Name -eq 'Upgrades') { $isUpgrade = $true }
        }
        if (-not $isUpgrade) { $update }
    }
    return @($updates)
}

function Install-WindowsUpdateItem {
    <#
        Downloads (if needed) and installs one update from
        Get-PendingWindowsUpdate. Returns Succeeded / RebootRequired.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Update
    )

    try {
        if (-not $Update.EulaAccepted) { $Update.AcceptEula() }

        $collection = New-Object -ComObject Microsoft.Update.UpdateColl
        [void]$collection.Add($Update)

        if (-not $Update.IsDownloaded) {
            $downloader = $script:UpdateSession.CreateUpdateDownloader()
            $downloader.Updates = $collection
            $download = $downloader.Download()
            # ResultCode 2 = succeeded, 3 = succeeded with errors.
            if ($download.ResultCode -notin 2, 3) {
                Write-Warning "Download failed for $($Update.Title) (result code $($download.ResultCode))."
                return [pscustomobject]@{ Succeeded = $false; RebootRequired = $false }
            }
        }

        $installer = $script:UpdateSession.CreateUpdateInstaller()
        $installer.ForceQuiet = $true
        $installer.Updates = $collection
        $install = $installer.Install()

        return [pscustomobject]@{
            Succeeded      = $install.ResultCode -in 2, 3
            RebootRequired = [bool]$install.RebootRequired
        }
    }
    catch {
        Write-Warning "Update failed for $($Update.Title): $($_.Exception.Message)"
        return [pscustomobject]@{ Succeeded = $false; RebootRequired = $false }
    }
}

function Start-WindowsUpdateScan {
    # Same as clicking "Check for updates" in Settings; Windows then downloads
    # and installs what it finds on its own, after the run ends.
    Start-Process 'ms-settings:windowsupdate-action'
}

Export-ModuleMember -Function Get-PendingWindowsUpdate, Install-WindowsUpdateItem, Start-WindowsUpdateScan
