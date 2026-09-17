#define AppSourceDir SourcePath + "..\artifacts\app"
#define AppVersion GetFileVersion(AppSourceDir + "\MeetingAlarm.App.exe")

[Setup]
AppId={{19B63BD1-4E9D-47CC-A502-FA617D0F6D3F}
AppName=Meeting Alarm
AppVerName=Meeting Alarm
AppVersion={#AppVersion}
AppPublisher=chadtoney
AppPublisherURL=https://github.com/chadtoney/MeetingAlarm
AppSupportURL=https://github.com/chadtoney/MeetingAlarm/issues
AppUpdatesURL=https://github.com/chadtoney/MeetingAlarm/releases
SourceDir=..
DefaultDirName={localappdata}\Programs\MeetingAlarm
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.19041
LicenseFile=LICENSE
SetupIconFile=src\MeetingAlarm.App\Assets\MeetingAlarm.ico
UninstallDisplayIcon={app}\Assets\MeetingAlarm.ico
OutputDir=artifacts\installer
OutputBaseFilename=MeetingAlarm-{#AppVersion}-win-x64-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
AppMutex=Local\MeetingAlarm
CloseApplications=no
RestartApplications=no

[Messages]
SetupAppRunningError=Meeting Alarm is running.%n%nRight-click its icon in the Windows notification area and choose Quit. Closing the window only hides it.%n%nClick OK after quitting, or Cancel to exit setup.
UninstallAppRunningError=Meeting Alarm is running.%n%nRight-click its icon in the Windows notification area and choose Quit, then click OK to uninstall.

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked

[Files]
Source: "artifacts\app\*"; DestDir: "{app}"; Excludes: "*.pdb,*.msalcache*,state.bin,state.bin.*"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{userprograms}\Meeting Alarm"; Filename: "{app}\MeetingAlarm.App.exe"; WorkingDir: "{app}"
Name: "{userdesktop}\Meeting Alarm"; Filename: "{app}\MeetingAlarm.App.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\MeetingAlarm.App.exe"; Description: "Launch Meeting Alarm"; WorkingDir: "{app}"; Flags: nowait postinstall skipifsilent

[Code]
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Command: String;
begin
  if CurUninstallStep = usUninstall then
  begin
    { Leave startup entries for other portable copies alone. }
    if RegQueryStringValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Run',
      'MeetingAlarm', Command) then
    begin
      if CompareText(Command,
        '"' + ExpandConstant('{app}\MeetingAlarm.App.exe') + '" --background') = 0 then
      begin
        if not RegDeleteValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Run',
          'MeetingAlarm') then
          RaiseException('Could not remove Meeting Alarm from Windows startup. Disable Start at Windows sign-in in the app before uninstalling.');
      end;
    end;
  end;
end;
