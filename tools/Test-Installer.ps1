param(
    [Parameter(Mandatory)]
    [string]$InstallerPath
)
$ErrorActionPreference = 'Stop'
$InstallerPath = (Resolve-Path -LiteralPath $InstallerPath).Path
$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{19B63BD1-4E9D-47CC-A502-FA617D0F6D3F}_is1'
$startupKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$menuShortcut = Join-Path ([Environment]::GetFolderPath('Programs')) 'Meeting Alarm.lnk'
$desktopShortcut = Join-Path ([Environment]::GetFolderPath('DesktopDirectory')) 'Meeting Alarm.lnk'

function Get-StartupCommand {
    if (-not (Test-Path -LiteralPath $startupKey)) { return $null }
    $key = Get-Item -LiteralPath $startupKey
    try { $key.GetValue('MeetingAlarm') }
    finally { $key.Close() }
}

if ((Test-Path -LiteralPath $uninstallKey) -or
    (Test-Path -LiteralPath $menuShortcut) -or
    (Test-Path -LiteralPath $desktopShortcut) -or
    ($null -ne (Get-StartupCommand))) {
    throw 'Run this test on a Windows account without an installed Meeting Alarm, its shortcuts, or its startup entry.'
}
$existing = $null
if ([Threading.Mutex]::TryOpenExisting('Local\MeetingAlarm', [ref]$existing)) {
    $existing.Dispose()
    throw 'Quit Meeting Alarm before running installer tests.'
}

$testDirectory = Join-Path $env:TEMP ("MeetingAlarm-Installer-Test-" + [Guid]::NewGuid().ToString('N'))
$installDirectory = Join-Path $testDirectory 'app'
$uninstaller = Join-Path $installDirectory 'unins000.exe'
$ownStartup = '"' + (Join-Path $installDirectory 'MeetingAlarm.App.exe') + '" --background'
$otherStartup = '"' + (Join-Path $testDirectory 'portable\MeetingAlarm.App.exe') + '" --background'
New-Item -ItemType Directory -Path $testDirectory | Out-Null

function Invoke-InstallerProcess($Executable, $Arguments, $LogName, [switch]$ExpectFailure) {
    $log = Join-Path $testDirectory $LogName
    $process = Start-Process -FilePath $Executable -ArgumentList (
        @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', "/LOG=`"$log`"") + $Arguments
    ) -Wait -PassThru
    if ($ExpectFailure) {
        if ($process.ExitCode -eq 0) { throw "Expected the running-app check to block installation/uninstallation. Log: $log" }
        if ((Get-Content -LiteralPath $log -Raw) -notmatch 'Meeting Alarm is running\.') {
            throw "Installer failed for a reason other than the running-app check. Log: $log"
        }
    }
    elseif ($process.ExitCode -ne 0) { throw "Installer exited with $($process.ExitCode). Log: $log" }
}

function Assert-Condition($Condition, $Message) {
    if (-not $Condition) { throw $Message }
}

$installArguments = @("/DIR=`"$installDirectory`"", '/SP-')
try {
    $mutex = [Threading.Mutex]::new($false, 'Local\MeetingAlarm')
    try {
        Invoke-InstallerProcess $InstallerPath $installArguments 'blocked-install.log' -ExpectFailure
        Assert-Condition (-not (Test-Path -LiteralPath $uninstallKey)) 'Blocked install created an uninstall entry.'
    }
    finally { $mutex.Dispose() }

    Invoke-InstallerProcess $InstallerPath $installArguments 'install.log'
    foreach ($name in @('MeetingAlarm.App.exe', 'MeetingAlarm.App.pri', 'App.xbf',
        'coreclr.dll', 'hostfxr.dll', 'Microsoft.ui.xaml.dll', 'Assets\MeetingAlarm.ico',
        'LICENSE', 'THIRD-PARTY-NOTICES.txt', 'GettingStarted.txt')) {
        Assert-Condition (Test-Path -LiteralPath (Join-Path $installDirectory $name)) "Missing installed file: $name"
    }
    Assert-Condition (Test-Path -LiteralPath $menuShortcut) 'Start menu shortcut is missing.'
    Assert-Condition (-not (Test-Path -LiteralPath $desktopShortcut)) 'Desktop shortcut should be opt-in.'
    Assert-Condition ($null -eq (Get-StartupCommand)) 'Install enabled startup without consent.'
    $registration = Get-ItemProperty -LiteralPath $uninstallKey
    Assert-Condition ($registration.DisplayName -eq 'Meeting Alarm') 'Incorrect Windows uninstall display name.'
    Assert-Condition ($registration.InstallLocation.TrimEnd('\') -eq $installDirectory) 'Incorrect uninstall install location.'

    $marker = Join-Path $installDirectory 'installer-test-user-file.txt'
    Set-Content -LiteralPath $marker -Value 'Keep user-created files during updates and uninstall.'
    if (-not (Test-Path -LiteralPath $startupKey)) {
        New-Item -Path $startupKey -Force | Out-Null
    }
    New-ItemProperty -LiteralPath $startupKey -Name MeetingAlarm -Value $ownStartup -PropertyType String | Out-Null
    Invoke-InstallerProcess $InstallerPath ($installArguments + '/TASKS=desktopicon') 'update.log'
    Assert-Condition (Test-Path -LiteralPath $desktopShortcut) 'Selected desktop shortcut is missing.'
    Assert-Condition (Test-Path -LiteralPath $marker) 'Update removed a user-created file.'
    Assert-Condition ((Get-StartupCommand) -eq $ownStartup) 'Update changed the startup preference.'

    $mutex = [Threading.Mutex]::new($false, 'Local\MeetingAlarm')
    try {
        Invoke-InstallerProcess $uninstaller @() 'blocked-uninstall.log' -ExpectFailure
        Assert-Condition (Test-Path -LiteralPath $uninstallKey) 'Blocked uninstall removed the installation.'
    }
    finally { $mutex.Dispose() }

    Invoke-InstallerProcess $uninstaller @() 'uninstall.log'
    Assert-Condition ($null -eq (Get-StartupCommand)) 'Uninstall left its own startup entry behind.'
    Assert-Condition (-not (Test-Path -LiteralPath $menuShortcut)) 'Uninstall left the Start menu shortcut behind.'
    Assert-Condition (-not (Test-Path -LiteralPath $desktopShortcut)) 'Uninstall left the desktop shortcut behind.'
    Assert-Condition (-not (Test-Path -LiteralPath $uninstallKey)) 'Uninstall left its Windows Settings entry behind.'
    Assert-Condition (-not (Test-Path -LiteralPath (Join-Path $installDirectory 'MeetingAlarm.App.exe'))) 'Uninstall left the application behind.'
    Assert-Condition (Test-Path -LiteralPath $marker) 'Uninstall deleted a user-created file.'

    Invoke-InstallerProcess $InstallerPath $installArguments 'reinstall.log'
    New-ItemProperty -LiteralPath $startupKey -Name MeetingAlarm -Value $otherStartup -PropertyType String | Out-Null
    Invoke-InstallerProcess $uninstaller @() 'uninstall-other-startup.log'
    Assert-Condition ((Get-StartupCommand) -eq $otherStartup) 'Uninstall removed another copy''s startup entry.'
    Write-Host "Installer lifecycle checks passed. Logs: $testDirectory"
}
finally {
    if (Test-Path -LiteralPath $uninstaller) {
        Invoke-InstallerProcess $uninstaller @() 'cleanup.log'
    }
    if ((Get-StartupCommand) -in @($ownStartup, $otherStartup)) {
        Remove-ItemProperty -LiteralPath $startupKey -Name MeetingAlarm
    }
}
