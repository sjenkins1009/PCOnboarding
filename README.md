# PC Onboarding Application

PowerShell-based onboarding tool. Asks a few yes/no questions up front, then
runs unattended:

1. Detect and remove all Microsoft Office installations.
2. Detect and remove preloaded/consumer Teams and the new Outlook app.
   OneDrive is intentionally left in place.
3. Detect and remove preloaded McAfee products, then run McAfee's own MCPR
   removal tool to clear what their uninstallers leave behind.
4. Download and silently install Google Chrome Enterprise.
5. Download and silently install Adobe Acrobat Reader.
6. Optional apps — **off by default**, only installed if you say yes at the
   startup prompts: Dropbox, Slack, Google Drive, Cisco Secure Client.

## Quick start on a new PC

1. In a browser, sign in to GitHub (the repo is private) and go to
   **https://github.com/sjenkins1009/PCOnboarding/releases/latest**.
2. Download **PCOnboarding.zip** and extract it (right-click → Extract All).
3. Open the extracted `PCOnboarding` folder and double-click
   **`Run-Onboarding.cmd`**. Approve the admin (UAC) prompt.

Windows may show a "Do you want to run this file?" warning for a downloaded
file. Click **Run**. The window stays open after the run so you can read the
summary.

## Structure

- `Run-Onboarding.cmd` — double-click launcher for `Start-Onboarding.ps1` (no typed commands needed).
- `Start-Onboarding.ps1` — entry point, self-elevates, runs the steps, logs progress.
- `Modules/OfficeRemoval.psm1` — `Get-InstalledOffice` (scan) and `Remove-OfficeInstallation` (uninstall).
- `Modules/BundledAppRemoval.psm1` — `Get-InstalledBundledApps` and `Remove-BundledAppInstallation` (Teams, new Outlook).
- `Modules/McAfeeRemoval.psm1` — `Get-InstalledMcAfee`, `Remove-McAfeeInstallation`, and `Invoke-McAfeeRemovalTool` (downloads/runs MCPR).
- `Modules/AppInstalls.psm1` — `Install-ChromeEnterprise` and `Install-AdobeReader`.
- `Modules/OptionalAppInstalls.psm1` — `Install-Dropbox`, `Install-Slack`, `Install-GoogleDrive`, `Install-CiscoSecureClient`.
- `Logs/` — transcript + per-product/app logs, one run per timestamp.

## Usage

Run on the target Windows PC (requires admin — the script will self-elevate
via UAC if not already running elevated):

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Start-Onboarding.ps1
```

Dry run (detect only, no removal):

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Start-Onboarding.ps1 -WhatIf
```

Fully unattended (skips the one interactive step — MCPR; the silent McAfee
uninstalls still run):

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Start-Onboarding.ps1 -SkipMcprCleanup
```

At startup it asks:

```
Also install any optional apps (Dropbox, Slack, Google Drive, Cisco Secure Client)? (y/N)
```

Answer `N` (or just press Enter) to run only the default set. Answer `y` and
it asks about each app in turn:

```
  Install Dropbox? (y/N)
  Install Slack? (y/N)
  Install Google Drive? (y/N)
  Install Cisco Secure Client? (y/N)
    Path to the Cisco Secure Client installer (.msi)
```

The Cisco path prompt only appears if you said yes to Cisco Secure Client,
and re-asks until you give it a file that actually exists.

Aside from that startup Q&A, the only step that can prompt is the MCPR
cleanup in step 3, and only if McAfee was actually found — MCPR is a
CAPTCHA-gated wizard, so it needs someone at the keyboard (pass
`-SkipMcprCleanup` to skip it). Everything else runs unattended, with
progress shown via a progress bar and console output, and a full transcript
written to `Logs\`.

## How detection/removal works

- **Click-to-Run** (Microsoft 365, Office 2019/2021/2024): read installed
  product IDs from `HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration`,
  then remove via `OfficeClickToRun.exe ... DisplayLevel=False` — the same
  silent path Programs & Features uses under the hood.
- **MSI-based** (Office 2007/2010/2013/2016 non-C2R): scan the registry
  Uninstall keys (64-bit, 32-bit-on-64-bit, and per-user) for Office/Visio/
  Project/OneNote entries, then run `msiexec /x {GUID} /qn /norestart`.
- **Per-language Click-to-Run entries**: both the main suite and standalone
  free OneNote register one Programs & Features entry per installed
  language — e.g. `Microsoft 365 - en-us`, `Microsoft 365 - es-es` for the
  suite, `Microsoft OneNote - es-es`, `Microsoft OneNote - pt-br` for
  OneNote. These match the registry scan above (the exact-shape
  `Microsoft 365 - <locale>` pattern is matched separately from the general
  Office/Visio/Project/OneNote pattern, specifically so it doesn't also
  catch unrelated "365"-branded products like Microsoft 365 Copilot), but
  their `UninstallString` already invokes `OfficeClickToRun.exe` directly
  rather than `msiexec`, so they're detected as a distinct type and run
  as-is with `DisplayLevel=False` forced, instead of being (incorrectly)
  treated as an MSI product.
- **OneNote for Windows 10** (the UWP/Store app): this one isn't a registry
  Uninstall entry, it's an AppX package — and it's commonly what people mean
  by "multiple OneNote installs," since each user profile on the machine can
  have its own copy, plus a machine-wide *provisioned* copy that silently
  reinstalls it for any new profile. The scan checks both:
  - `Get-AppxPackage -AllUsers` → one hit per user profile that has it, removed with `Remove-AppxPackage -AllUsers`.
  - `Get-AppxProvisionedPackage -Online` → the "will install for new users" copy, removed with `Remove-AppxProvisionedPackage`.
  Classic desktop OneNote (bundled in an Office MSI/C2R suite, or the
  standalone free Click-to-Run OneNote) is already covered by the two
  detections above.

A post-removal re-scan confirms nothing Office-related remains (some leftovers
may require a reboot to fully clear, which is called out if seen).

## Teams / new Outlook removal (step 2)

These come preloaded on new machines and get replaced once the licensed
Microsoft 365 deployment runs, so step 2 clears them out first. OneDrive is
deliberately not touched by this step.

- **Teams (classic)**: detected two ways — the `Teams Machine-Wide Installer`
  MSI (removed via `msiexec /x`), and a per-user install at
  `%LocalAppData%\Microsoft\Teams\Update.exe` (removed via `Update.exe
  --uninstall -s`).
- **Teams (new)** and **Outlook (new)**: both ship as AppX packages
  (`MSTeams`, `MicrosoftTeams` for the Windows 11 chat icon, and
  `Microsoft.OutlookForWindows`), so they use the same
  `Get-AppxPackage -AllUsers` / `Get-AppxProvisionedPackage -Online` pattern
  used for OneNote for Windows 10 above.

**Will this stop these from reappearing on user profiles created later?**
Yes for both. What we remove *is* the machine-wide provisioning mechanism
itself — the Machine-Wide Installer MSI is exactly what installs classic
Teams for every new login, and the AppX *provisioned* package is exactly
what installs the new AppX-based Teams/Outlook for every new login. Removing
them removes that trigger, so future profiles won't get it either.

## McAfee removal (step 3)

McAfee ships preloaded on most consumer/OEM machines as a trial, usually as
several separate entries rather than one. Removal is two stages, because in
practice one isn't enough.

**Stage 1 — silent uninstall of each detected product.** Detection scans the
same registry Uninstall keys used elsewhere in this project, matching on
*either* display name (`McAfee`, `WebAdvisor`, `Safe Connect`, `LiveSafe`,
`Total Protection`) **or** publisher (`McAfee`). The publisher check matters:
some OEM preloads register components like WebAdvisor and Safe Connect
without the McAfee brand anywhere in the display name, so a name-only scan
misses them. The Store-delivered "McAfee Security" app is picked up
separately via `Get-AppxPackage -AllUsers` / `Get-AppxProvisionedPackage
-Online`, same pattern as OneNote and Teams above — the provisioned copy is
what reinstalls it for every new user profile, so it has to go too.

Products are removed in dependency order: the bolt-on components first, the
main suite (Total Protection / LiveSafe / Internet Security / Antivirus /
Personal Security) last. Taking the suite out first leaves its bolt-ons
holding a dead uninstall string and they become unremovable without manual
cleanup. Entries with a GUID-shaped key go to `msiexec /x ... /qn
/norestart`; the rest run their own recorded uninstall string, preferring
`QuietUninstallString` when the product registered one, and otherwise having
`/quiet` appended if no silence flag is already present.

**Stage 2 — MCPR.** McAfee's own uninstallers routinely strand services,
drivers and registry keys that block a later clean reinstall, so MCPR
(McAfee Consumer Product Removal) runs as a cleanup pass whenever McAfee was
detected — even if every uninstall in stage 1 reported success.

> **MCPR is interactive.** It's a wizard and it gates itself behind a
> CAPTCHA, so it cannot be silenced and it will stall an unattended run
> waiting for clicks. That's why it's a cleanup pass rather than the primary
> removal path, and why `-SkipMcprCleanup` exists. A reboot is required
> afterwards to finish clearing McAfee's drivers and services.

### Where MCPR is downloaded from

Straight from McAfee's CDN:

```
https://download.mcafee.com/products/licensed/cust_support_patches/MCPR.exe
```

Older removal guides circulate a shortened alias, `mcpr.notlong.com`. That
alias is **deliberately not used here**, even though it currently 302s to
exactly the URL above and serves a byte-identical file (12,647,224 bytes,
verified). The reason is control: `mcpr.notlong.com` and bare `notlong.com`
resolve to the same third-party host, so it's a wildcard redirector someone
else owns and can repoint at any binary at any time — and this script
executes whatever comes back, elevated. Fetching from McAfee directly gets
the same file with nobody in the middle.

As a second layer, `Invoke-McAfeeRemovalTool` checks the downloaded file's
Authenticode signature before running it, and refuses to execute unless the
signature is valid *and* the signing subject is McAfee. If the URL ever
starts serving something else, the script fails there instead of running it.

## App installs (steps 4 & 5)

Both apps are downloaded to `%ProgramData%\PCOnboarding\Downloads` and
installed silently — no prompts, no bundled-offer opt-outs to click through.

- **Google Chrome Enterprise**: downloads Google's official standalone MSI
  (`dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi` —
  a static URL Google maintains that always resolves to current stable),
  then `msiexec /i ... /qn /norestart`.
- **Adobe Acrobat Reader**: downloads Adobe's official "web installer"
  package (`trials.adobe.com/.../Acrobat_DC_Web_WWMUI.zip` — likewise a
  permanently-static URL Adobe hosts so scripts don't need to chase
  version-specific filenames), extracts it, then runs the bundled
  `setup.exe /sAll /rs /msi EULA_ACCEPT=YES /qn`.

Both installer functions download with `Invoke-WebRequest`, verify a non-empty
file landed before proceeding, and treat MSI exit codes 0 and 3010 (success,
reboot required) as success.

## Optional apps (step 6) — off by default

None of these run unless you answer yes to them at the startup prompts (see
Usage above). All three downloadable ones land in
`%ProgramData%\PCOnboarding\Downloads`, same as Chrome/Adobe.

- **Dropbox**: downloads the official offline installer
  (`dropbox.com/download?full=1&os=win` — a stable redirect to the current
  version, confirmed live), installs via `DropboxOfflineInstaller.exe
  /NOLAUNCH` (silent, no post-install launch).
- **Slack**: downloads the official current MSIX package
  (`slack.com/api/desktop.latestRelease?arch=x64&variant=msix&redirect=true`
  — Slack's own deployment docs point at this endpoint; Slack has moved off
  a traditional MSI to MSIX on Windows), installs via `Add-AppxPackage`.
- **Google Drive**: downloads Google's official "Drive for desktop"
  installer (`dl.google.com/drive-file-stream/GoogleDriveFSSetup.exe`,
  confirmed live), installs via `GoogleDriveFSSetup.exe --silent
  --desktop_shortcut --skip_launch_new`.
- **Cisco Secure Client**: **cannot be auto-downloaded — you must supply the
  installer file.** Cisco gates it behind a Cisco.com account and your
  organization's own entitlement/VPN headend package; there's no generic
  public URL for it. If you say yes, the script prompts for a local path to
  the MSI (re-asking until the file actually exists) and runs
  `msiexec /i ... /qn /norestart` against it, same as the other MSI-based
  installs in this project.

## Notes / next steps

- Office removal (step 1) only targets Office suites/apps (Word, Excel,
  Outlook, Visio, Project, Microsoft 365 Apps, OneNote in all its forms). It
  deliberately excludes Edge, Visual Studio, .NET, etc. — and Teams/the new
  Outlook, which step 2 handles separately since they're not part of the
  Office suite and use different removal mechanisms.
- OneDrive is intentionally left alone by this script (both steps 1 and 2
  exclude it) — it's not being removed as part of this onboarding flow.
- If Google or Adobe change these download URLs in the future, only
  `Modules/AppInstalls.psm1` needs updating — nothing else references them.
  Likewise, the MCPR URL lives in one place, `$script:McprUrl` at the top of
  `Modules/McAfeeRemoval.psm1`.
- A GUI/frontend can wrap this later — the module functions are already
  decoupled from console output, so they can be called from anything that
  can invoke PowerShell.
