# Meeting Alarm

A standalone Windows meeting-reminder prototype with always-on-top alarms,
a repeating chime, snooze, and meeting join links. It reads your default calendar
from classic Outlook or Microsoft Graph; local alarms work without either.

**Intended for sharing with Microsoft employees in an access-restricted GitHub
repository.** This is a prototype, not an official Microsoft product or PowerToy.
It does not guarantee that you will notice every meeting.

## Quick start

Requires Windows 10 version 2004 or later, or Windows 11, on x64.

### Build from source

Install the .NET 10 SDK for Windows x64, then open PowerShell in this folder:

```powershell
.\build.ps1
.\artifacts\app\MeetingAlarm.App.exe
```

The build restores pinned public NuGet dependencies and publishes a self-contained
app, including .NET and the Windows App SDK. Internet access is required for the
initial restore. Keep the **entire** `artifacts\app` folder, not just the executable.
Before rebuilding, quit any copy running from that folder using the tray icon's
**Quit** command; closing its window only hides it and leaves build files locked.

### Run a shared build

If a maintainer provides a build in the employee-only repository, extract the
entire archive to a local folder and run `MeetingAlarm.App.exe`. The .NET SDK is
not required to run the published app. Use only an approved distribution source;
if Windows or organizational policy blocks execution, follow your organization's
approval process rather than disabling protections.

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
screenshots and diagnostics before sharing, even in an employee-only repository.

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
changes. Report reproducible issues in the access-restricted repository.

## Publishing and sharing (maintainers)

Before the first push, create an organization-approved repository whose effective
access is limited to the intended Microsoft employees. A repository visibility
label alone does not establish employee-only access; check organization membership,
outside collaborators, and enterprise access policies.

Review the files staged for the initial commit. `.gitignore` excludes build output
and common local state/credential files, but is not a substitute for reviewing
content. Do not include `%LOCALAPPDATA%\MeetingAlarm` or another user's cache.
No open-source license or public redistribution permission is established here;
obtain the appropriate ownership and release approvals before any public release.

To prepare a runnable archive after a successful build:

```powershell
Compress-Archive -Path .\artifacts\app\* -DestinationPath .\artifacts\MeetingAlarm-win-x64.zip
```

Share the archive only through the approved employee-only repository or another
approved internal channel. This build script does not sign the executable or
produce an installer; follow applicable signing and software-distribution policy.
