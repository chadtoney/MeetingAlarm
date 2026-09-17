# Meeting Alarm

A standalone Windows meeting-reminder prototype with always-on-top alarms,
a repeating chime, snooze, and meeting join links. It reads your default calendar
from classic Outlook or Microsoft Graph; local alarms work without either.

Available under the [MIT license](LICENSE). This is a prototype, not an official
Microsoft product or PowerToy. It does not guarantee that you will notice every
meeting.

## Quick start

Requires Windows 10 version 2004 or later, or Windows 11. The app is built for
x64; Windows 11 on Arm64 can run it using Windows' x64 emulation.

### Install (recommended; no developer tools needed)

1. Open the [Releases page](https://github.com/chadtoney/MeetingAlarm/releases)
   and download `MeetingAlarm-<version>-win-x64-Setup.exe` from **Assets**.
2. Double-click the downloaded file and follow the setup wizard. You can keep the
   default installation folder and optionally create a desktop shortcut.
3. Open **Meeting Alarm** from the Windows Start menu, then click **Test in 10
   seconds** to check the popup and sound.
4. Connect your calendar using the steps below. Optionally enable **Start at
   Windows sign-in** in the app.

Setup installs for your Windows user without requesting administrator rights.
It includes .NET and the Windows App SDK, adds a Start menu shortcut, and registers
an uninstaller in Windows Settings. You do not need Git, Visual Studio, or the
.NET SDK. Calendar connection is still a separate step; setup does not sign in or
read your mailbox.

Builds are unsigned. If Windows or organizational policy blocks execution, follow
the applicable approval process rather than disabling protections. If an installer
has not been published yet, a maintainer must build and upload one using the
instructions below.

### Portable ZIP (alternative)

If a release includes `MeetingAlarm-win-x64.zip`, use File Explorer's **Extract
All** to unpack it to a permanent folder. Open that folder and run
`MeetingAlarm.App.exe`. Do not run inside the ZIP or move just the executable:
keep all extracted files together. Portable copies do not add Start menu shortcuts
or a Windows uninstall entry.

### Build and install from source

Install [Git](https://git-scm.com/downloads/win) and the
[.NET 10 SDK for Windows x64](https://dotnet.microsoft.com/en-us/download/dotnet/10.0).
Clone the repository and open its folder in PowerShell:

```powershell
git clone https://github.com/chadtoney/MeetingAlarm.git
Set-Location MeetingAlarm
```

If you already have the source folder, skip cloning. Build and run:

```powershell
.\build.ps1
.\artifacts\app\MeetingAlarm.App.exe
```

The build restores pinned public NuGet dependencies and publishes a self-contained
app, including .NET and the Windows App SDK. Internet access is required for the
initial restore. Keep the **entire** `artifacts\app` folder, not just the executable.
You can copy that folder's contents to a permanent location such as
`%LOCALAPPDATA%\Programs\MeetingAlarm`, then run `MeetingAlarm.App.exe` there.
Before rebuilding, quit any copy running from that folder using the tray icon's
**Quit** command; closing its window only hides it and leaves build files locked.

### Update or uninstall

**Installed with setup:** Right-click the tray icon and choose **Quit**, then run
the newer installer. Keep the same installation folder to preserve the startup
path. To uninstall, quit from the tray, open **Windows Settings > Apps**, find
**Meeting Alarm**, and choose **Uninstall**. The uninstaller removes the startup
entry if it still points to that installation.

**Portable copy:** Quit from the tray before replacing application files with a
fresh extraction. To uninstall, disable **Start at Windows sign-in**, quit, and
delete the application folder. Disable startup before moving the app or switching
to the setup-installed version; re-enable it in the new copy if desired.

Both methods retain settings and calendar state in `%LOCALAPPDATA%\MeetingAlarm`.
Optionally delete that folder to remove saved settings, calendar data, and the
local token cache. Deleting the local cache does not revoke previously issued
tokens or sign you out of Outlook or your browser.

### Connect your calendar

1. Click **Test in 10 seconds** to check the popup and sound.
2. Expand **Calendar connection**, select **Classic Outlook (on this PC)**, and
   click **Connect**. Classic Outlook must be installed with a configured default
   mail profile. New Outlook alone is not sufficient. No separate Entra app
   registration is needed for this source.
3. Alternatively, select **Microsoft Graph** and supply an approved Entra
   application/client ID and tenant ID. It uses delegated `Calendars.Read`,
   a desktop public client, and `http://localhost` as the redirect URI. No shared
   client ID or client secret is included; organizational consent may be required.

You can use **Schedule local alarm** without connecting a calendar. Closing the
main window leaves the app in the tray; right-click the tray icon to quit.
Starting at Windows sign-in is optional and disabled by default.

See [GettingStarted.txt](GettingStarted.txt) for complete setup, alarm rules,
privacy controls, troubleshooting, and uninstall instructions. Connector details:
[classic Outlook](src/MeetingAlarm.ClassicOutlook/SETUP.txt) and
[Microsoft Graph](src/MeetingAlarm.Outlook/SETUP.txt).

## Privacy and limitations

Calendar access is read-only: the app does not send messages, edit meetings, or
change responses. Calendar snapshots and settings are stored under
`%LOCALAPPDATA%\MeetingAlarm` using Windows current-user DPAPI encryption; the
Graph token cache also uses Windows-protected persistence.

Calendar data refreshes every two minutes and may be stale while offline.
Sleep, lock screen, muted audio, and full-screen apps can prevent an effective
reminder. The app does not wake the computer, unmute audio, or detect Teams calls.
Alarm titles are hidden by default, but titles remain visible in the main window.

Do not attach real meeting subjects, attendee details, join URLs, tokens, or local
state files to issues or pull requests. Use synthetic examples and redact
screenshots and diagnostics before sharing.

## Development

Run the tests on Windows with the .NET 10 SDK:

```powershell
dotnet test MeetingAlarm.slnx -c Release
```

Tests use synthetic data and fake calendar readers; they do not sign in or read
your mailbox. Live calendar connections require a separate, explicitly authorized
manual check. Run `.\build.ps1` after app changes to verify the published resources.

The solution separates the scheduling engine (`MeetingAlarm.Core`), calendar
connectors (`MeetingAlarm.ClassicOutlook` and `MeetingAlarm.Outlook`), and WinUI 3
desktop host (`MeetingAlarm.App`). Keep changes focused and add tests for behavior
changes. Report reproducible issues in this repository without including private
calendar data.

## Publishing and sharing (maintainers)

Review the files staged for each commit. `.gitignore` excludes build output
and common local state/credential files, but is not a substitute for reviewing
content. Do not include `%LOCALAPPDATA%\MeetingAlarm` or another user's cache.
Contributors must have permission to share and license their contributions.

Install [Inno Setup 6.3 or later](https://jrsoftware.org/isdl.php) on the build
machine (not on users' computers), then run:

```powershell
.\build.ps1 -Installer
```

The installer is written to
`artifacts\installer\MeetingAlarm-<version>-win-x64-Setup.exe`. Its version comes
from the published application's file version; increment the application version
before publishing a new release. The compiler is discovered on PATH or in the
standard per-user and Program Files locations. The build includes the MIT license,
user guide, and third-party notices collected from restored packages. App settings
and token caches are not part of the installer.

On a clean Windows test account with no installed Meeting Alarm, shortcuts, or
startup entry, exercise installation, update, shortcuts, running-app guards, and
uninstallation with:

```powershell
.\tools\Test-Installer.ps1 -InstallerPath .\artifacts\installer\MeetingAlarm-<version>-win-x64-Setup.exe
```

This temporarily installs the app and creates test shortcuts/startup entries,
then uninstalls and removes those entries. It does not launch the app or connect
to a calendar. Diagnostic logs are kept in the temporary folder printed by the test.

Attach the setup executable to a GitHub release so users can install without
building. A private repository still requires GitHub access to download releases;
MIT licensing does not change repository visibility. The build does not code-sign
the application or installer; follow applicable signing and software-distribution
policy before distributing.

For an optional portable archive after building (use a new archive name if the
destination already exists):

```powershell
Compress-Archive -Path .\artifacts\app\* -DestinationPath .\artifacts\MeetingAlarm-win-x64.zip
```

## License

Meeting Alarm is licensed under the [MIT license](LICENSE), which permits use,
modification, and redistribution, including commercial use, subject to its terms.
Include the copyright and permission notice when redistributing the software.
The software is provided without warranty.

Third-party dependencies remain subject to their respective licenses; the
Meeting Alarm license does not replace those terms. Published builds include
`THIRD-PARTY-NOTICES.txt` with package attribution, supplied license/notice text,
and license references.
