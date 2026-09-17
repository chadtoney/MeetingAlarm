param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    [switch]$Launch,
    [switch]$Installer
)
$ErrorActionPreference = 'Stop'
$compiler = $null
if ($Installer) {
    $command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($command) { $compiler = $command.Source }
    else {
        foreach ($directory in @(
            (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6'),
            (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6'),
            (Join-Path $env:ProgramFiles 'Inno Setup 6')
        )) {
            $candidate = Join-Path $directory 'ISCC.exe'
            if (Test-Path -LiteralPath $candidate) {
                $compiler = $candidate
                break
            }
        }
    }
    if (-not $compiler) {
        throw 'Inno Setup 6.3 or later is required for -Installer. Install it from https://jrsoftware.org/isdl.php or put ISCC.exe on PATH.'
    }
}
$project = Join-Path $PSScriptRoot 'src\MeetingAlarm.App\MeetingAlarm.App.csproj'
$output = Join-Path $PSScriptRoot 'artifacts\app'
& dotnet publish $project -c $Configuration -r win-x64 --self-contained true -o $output --nologo
if ($LASTEXITCODE -ne 0) { throw 'Meeting Alarm publish failed.' }
$build = Join-Path $PSScriptRoot "src\MeetingAlarm.App\bin\$Configuration\net10.0-windows10.0.19041.0\win-x64"
foreach ($name in @('MeetingAlarm.App.pri', 'App.xbf', 'hostfxr.dll', 'coreclr.dll', 'Microsoft.ui.xaml.dll', 'Assets\MeetingAlarm.ico', 'Assets\MeetingAlarm.png')) {
    $source = Get-FileHash (Join-Path $build $name)
    $published = Get-FileHash (Join-Path $output $name)
    if ($source.Hash -ne $published.Hash) { throw "Published file does not match the build: $name" }
}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'LICENSE'), (Join-Path $PSScriptRoot 'GettingStarted.txt') -Destination $output
& (Join-Path $PSScriptRoot 'tools\Write-ThirdPartyNotices.ps1') -PublishDirectory $output
Write-Host "Application: $output\MeetingAlarm.App.exe"
if ($Installer) {
    & $compiler /Qp (Join-Path $PSScriptRoot 'installer\MeetingAlarm.iss')
    if ($LASTEXITCODE -ne 0) { throw 'Meeting Alarm installer build failed.' }
    $version = (Get-Item -LiteralPath (Join-Path $output 'MeetingAlarm.App.exe')).VersionInfo.FileVersion
    $setup = Join-Path $PSScriptRoot "artifacts\installer\MeetingAlarm-$version-win-x64-Setup.exe"
    if (-not (Test-Path -LiteralPath $setup)) { throw "Installer output is missing: $setup" }
    Write-Host "Installer: $setup"
}
if ($Launch) { Start-Process (Join-Path $output 'MeetingAlarm.App.exe') }
