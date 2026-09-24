# PC Onboarding Application

PowerShell-based onboarding tool. Asks a few questions up front, then runs:

1. Optionally rename the PC, and join it to Microsoft Entra ID, an Active
   Directory domain, or both (chosen at startup — see
   [Rename and join](#rename-and-join-step-1)).
2. Detect and remove all Microsoft Office installations.
3. Detect and remove preloaded/consumer Teams and the new Outlook app.
   OneDrive is intentionally left in place.
4. Detect and remove preloaded McAfee products, then run McAfee's own MCPR
   removal tool to clear what their uninstallers leave behind.
5. Download and silently install Google Chrome Enterprise.
6. Download and silently install Adobe Acrobat Reader.
7. Optional apps — **off by default**, only installed if you say yes at the
   startup prompts: Dropbox, Slack, Google Drive, Firefox, Zoom, Cisco Secure Client.
8. Windows Update — chosen at startup: install updates as part of the run
   (time counted, updates listed in the summary), or start them in the
   background and finish right away (see [Windows Update](#windows-update-step-8)).

## Quick start on a new PC

In PowerShell:

```powershell
iwr https://git-pc.skjenkins.com/PCOnboarding.zip -OutFile $env:TEMP\PCOnboarding.zip
Expand-Archive $env:TEMP\PCOnboarding.zip -DestinationPath C:\PCOnboarding -Force
C:\PCOnboarding\Run-Onboarding.cmd
```

Approve the admin (UAC) prompt if one appears. The window stays open after the
run so you can read the summary.

`PCOnboarding.zip` is rebuilt automatically from `main` on every push (see
`.github/workflows/pages.yml`), so the link always has the latest version.

## Ticket summary

When the run finishes, the script prints a summary ready to paste into a
ticket, **copies it to the clipboard**, and saves it as
`Logs\Summary_<timestamp>.txt`:

```
PC Onboarding Summary - CONTOSO-LT01 (renamed from DESKTOP-ABC123)
Start Time: 9/24/2026 2:05 PM
End Time:   9/24/2026 2:41 PM
Runtime:    45 min (actual 36 min, rounded up to 15-min increments)

Joined:
  - Domain: contoso.local
  - Microsoft Entra ID (hybrid): completes after restart
Removed (5):
  - Microsoft 365 - en-us
  - Microsoft Teams (new)
  - ...
Installed (2):
  - Google Chrome Enterprise
  - Adobe Acrobat Reader
Windows Updates installed (2):
  - 2026-09 Cumulative Update for Windows 11 Version 25H2 for x64-based Systems (KB5066001)
  - Security Intelligence Update for Microsoft Defender Antivirus (KB2267602)
Failed (1):
  - Remove Microsoft OneNote - pt-br
```

If you chose to run Windows Update in the background instead, the updates
section is replaced by: `Windows Update: started at the end of the run;
updates finish installing in the background.`

- The clock starts before the startup questions, so they count toward runtime.
- Runtime is rounded **up** to the next 15 minutes (a 16-minute run bills as
  30). The actual minutes are shown alongside it.
- Sections with nothing in them are left out, and apps you didn't choose
  aren't listed.
- If the run stops partway because of an error, you still get a summary of
  what it finished, with a note that it stopped early.
- Windows Update time is only in the runtime if you chose to install updates
  as part of the run. In background mode, if a restart is also needed for a
  rename/join, let the updates finish first so one restart covers both.

## Structure

- `Run-Onboarding.cmd` — double-click launcher for `Start-Onboarding.ps1` (no typed commands needed).
- `Start-Onboarding.ps1` — entry point, self-elevates, runs the steps, logs progress.
- `Modules/DeviceJoin.psm1` — `Get-JoinStatus`, `Join-ADDomain`, `Rename-Device`, `Start-EntraJoin`, `Test-JoinSupportedEdition`.
- `Modules/OfficeRemoval.psm1` — `Get-InstalledOffice` (scan) and `Remove-OfficeInstallation` (uninstall).
- `Modules/BundledAppRemoval.psm1` — `Get-InstalledBundledApps` and `Remove-BundledAppInstallation` (Teams, new Outlook).
- `Modules/McAfeeRemoval.psm1` — `Get-InstalledMcAfee`, `Remove-McAfeeInstallation`, and `Invoke-McAfeeRemovalTool` (downloads/runs MCPR).
- `Modules/AppInstalls.psm1` — `Install-ChromeEnterprise` and `Install-AdobeReader`.
- `Modules/OptionalAppInstalls.psm1` — `Install-Dropbox`, `Install-Slack`, `Install-GoogleDrive`, `Install-Firefox`, `Install-Zoom`, `Install-CiscoSecureClient`.
- `Modules/WindowsUpdates.psm1` — `Get-PendingWindowsUpdate`, `Install-WindowsUpdateItem`, `Start-WindowsUpdateScan`.
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

Skip the interactive MCPR cleanup (the silent McAfee uninstalls still run):

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Start-Onboarding.ps1 -SkipMcprCleanup
```

At startup it first asks whether to rename the PC, then how to join it:

```
Rename this PC? Current name is DESKTOP-ABC123. Enter a new name, or press Enter to keep it

How should this PC be joined?
  1) Microsoft Entra ID (Entra joined)
  2) Active Directory domain (domain joined)
  3) Both - domain join, then hybrid Entra join
Enter 1, 2, or 3 (or just press Enter to skip joining)
```

Options 2 and 3 then ask for the domain name and pop up a sign-in box for an
account that's allowed to join computers to it. Then it asks:

```
Also install any optional apps (Dropbox, Slack, Google Drive, Cisco Secure Client, Firefox, Zoom)? (y/N)
```

Answer `N` (or just press Enter) to run only the default set. Answer `y` and
it asks about each app in turn:

```
  Install Dropbox? (y/N)
  Install Slack? (y/N)
  Install Google Drive? (y/N)
  Install Firefox? (y/N)
  Install Zoom? (y/N)
  Install Cisco Secure Client? (y/N)
    Path to the Cisco Secure Client installer (.msi)
```

The Cisco path prompt only appears if you said yes to Cisco Secure Client,
and re-asks until you give it a file that actually exists.

Last, it asks how to handle Windows Update:

```
Windows Update at the end of the run:
  1) Install updates as part of the run - update time counts toward runtime, updates listed in the summary
  2) Start updates in the background - summary is ready right away, updates finish afterward
Enter 1 or 2 (or just press Enter for 2)
```

This is asked at startup (not when the run reaches step 8) so the run never
sits waiting on a prompt, with the clock running, after you've walked away.

Aside from that startup Q&A, only two things can need someone at the
keyboard: an **Entra join** (option 1), which happens in step 1 right after
the questions, and the **MCPR cleanup** in step 4, only if McAfee was actually
found (MCPR is a CAPTCHA-gated wizard; pass `-SkipMcprCleanup` to skip it). A
failed domain join also asks whether to retry. Everything else runs
unattended, with progress shown via a progress bar and console output, and a
full transcript written to `Logs\`.

## Rename and join (step 1)

Runs before anything is uninstalled.

**Rename:** new names must be 1–15 letters, numbers, or hyphens, not all
numbers, and can't start or end with a hyphen; the prompt re-asks until the
name is valid. The new name takes effect after restart, and the ticket
summary shows it as `NEW-NAME (renamed from OLD-NAME)`.

- When a domain join happens in the same run, the rename is done *inside* the
  join (`Add-Computer -NewName`). Renaming first and then joining before a
  restart would put the PC in the domain under its **old** name.
- If the PC is already in a domain, renaming it also renames its account in
  Active Directory, so the script asks for a domain account at startup.
- If a domain join fails, the PC is still renamed on its own.
- With an Entra join (option 1), Entra ID may list the PC under its old name
  until it restarts.

**Join:** Windows **Home** can't be joined to a domain or Entra ID, so the
script checks the edition first and reports it as failed on Home PCs (upgrade
to Pro, then join); a rename still works on Home. It also skips any join the
PC already has.

- **1 – Microsoft Entra ID:** Windows has no command line that signs a user
  in and Entra-joins the PC, so the script opens **Settings → Access work or
  school** and prints the clicks: Connect → "Join this device to Microsoft
  Entra ID" → sign in with the user's work account → Done. Press Enter in the
  script when finished; it checks `dsregcmd /status` to confirm the join and
  offers to reopen the screen if it didn't work. Don't restart when Windows
  offers — the script finishes first and tells you when.
- **2 – Active Directory domain:** fully command line (`Add-Computer`). The PC
  must be able to reach a domain controller, so it has to be on the client's
  network or VPN. If the join fails (wrong password, typo, can't reach the
  domain), it offers to retry with a corrected domain name or account. The
  join takes effect after restart, so the rest of onboarding runs first.
- **3 – Both:** Windows won't Entra-join a PC that's in a domain directly; the
  "both" state is **hybrid join**. The script does the domain join, and after
  the restart Windows registers the PC with Entra ID by itself — **only if the
  client's Microsoft Entra Connect is set up for hybrid join**. Check afterward
  with `dsregcmd /status` (look for `AzureAdJoined : YES` and
  `DomainJoined : YES`).

The domain account password goes only to `Add-Computer`; it isn't written to
the log or the ticket summary. When a rename or join was done, the run ends
with **RESTART REQUIRED**.

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

## Teams / new Outlook removal (step 3)

These come preloaded on new machines and get replaced once the licensed
Microsoft 365 deployment runs, so step 3 clears them out first. OneDrive is
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

## McAfee removal (step 4)

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

## App installs (steps 5 & 6)

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

## Optional apps (step 7) — off by default

None of these run unless you answer yes to them at the startup prompts (see
Usage above). All the downloadable ones land in
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
- **Firefox**: downloads Mozilla's official MSI
  (`download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=en-US`
  — always redirects to the current 64-bit release, confirmed live), installs
  via `msiexec /i ... /qn /norestart`.
- **Zoom**: downloads Zoom's official IT-deployment MSI
  (`zoom.us/client/latest/ZoomInstallerFull.msi?archType=x64` — always
  redirects to the current 64-bit Zoom Workplace, confirmed live), installs
  via `msiexec /i ... /qn /norestart`.
- **Cisco Secure Client**: **cannot be auto-downloaded — you must supply the
  installer file.** Cisco gates it behind a Cisco.com account and your
  organization's own entitlement/VPN headend package; there's no generic
  public URL for it. If you say yes, the script prompts for a local path to
  the MSI (re-asking until the file actually exists) and runs
  `msiexec /i ... /qn /norestart` against it, same as the other MSI-based
  installs in this project.

## Windows Update (step 8)

Which mode runs is chosen at startup.

- **1 – Install as part of the run:** uses the Windows Update service built
  into Windows (no add-on modules). It checks for updates, then downloads and
  installs them one at a time with a progress bar. The time counts toward the
  runtime, each installed update is listed in the ticket summary, and any that
  fail go under **Failed**.
  - It installs what Windows would install automatically: security and
    cumulative updates, Defender updates, and drivers. Optional/preview
    updates are skipped, and so are **feature upgrades** (e.g. 24H2 → 25H2),
    which can take over an hour.
  - It runs one pass. Some updates only show up after the restart that the
    first batch needs, so a later check may find a few more.
  - If an update needs a restart, the run ends with **RESTART REQUIRED**.
- **2 – Background (default):** opens Settings → Windows Update and triggers
  **Check for updates**. Windows downloads and installs on its own after the
  run ends, so none of that time is in the runtime, and the summary just notes
  that updates were started.

## Notes / next steps

- Office removal (step 2) only targets Office suites/apps (Word, Excel,
  Outlook, Visio, Project, Microsoft 365 Apps, OneNote in all its forms). It
  deliberately excludes Edge, Visual Studio, .NET, etc. — and Teams/the new
  Outlook, which step 3 handles separately since they're not part of the
  Office suite and use different removal mechanisms.
- OneDrive is intentionally left alone by this script (both steps 2 and 3
  exclude it) — it's not being removed as part of this onboarding flow.
- If Google or Adobe change these download URLs in the future, only
  `Modules/AppInstalls.psm1` needs updating — nothing else references them.
  Likewise, the MCPR URL lives in one place, `$script:McprUrl` at the top of
  `Modules/McAfeeRemoval.psm1`.
- A GUI/frontend can wrap this later — the module functions are already
  decoupled from console output, so they can be called from anything that
  can invoke PowerShell.
