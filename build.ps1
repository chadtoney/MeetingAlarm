param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    [switch]$Launch
)
$ErrorActionPreference = 'Stop'
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
Write-Host "Application: $output\MeetingAlarm.App.exe"
if ($Launch) { Start-Process (Join-Path $output 'MeetingAlarm.App.exe') }
