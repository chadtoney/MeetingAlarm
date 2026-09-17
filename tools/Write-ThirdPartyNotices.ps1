param(
    [Parameter(Mandatory)]
    [string]$PublishDirectory
)
$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot
$assets = Get-Content -LiteralPath (Join-Path $repository 'src\MeetingAlarm.App\obj\project.assets.json') -Raw | ConvertFrom-Json
$dependencies = Get-Content -LiteralPath (Join-Path $PublishDirectory 'MeetingAlarm.App.deps.json') -Raw | ConvertFrom-Json
$packages = @($assets.libraries.PSObject.Properties |
    Where-Object { $_.Value.type -eq 'package' } | ForEach-Object Name)
$packages += @($dependencies.libraries.PSObject.Properties |
    Where-Object { $_.Value.type -eq 'runtimepack' } |
    ForEach-Object { $_.Name.Substring('runtimepack.'.Length) })
$mitLicense = Get-Content -LiteralPath (Join-Path $repository 'LICENSE') -Raw
$permissionStart = $mitLicense.IndexOf('Permission is hereby granted', [StringComparison]::Ordinal)
if (-not $mitLicense.StartsWith('MIT License') -or $permissionStart -lt 0) {
    throw 'License collection requires the MIT permission text. Review this script if the project license changes.'
}
$mitTerms = $mitLicense.Substring($permissionStart)
$notices = [Text.StringBuilder]::new()
[void]$notices.AppendLine('THIRD-PARTY NOTICES')
[void]$notices.AppendLine('Licenses and notices for packages used to build or run Meeting Alarm.')
[void]$notices.AppendLine('Some packages contain build-time tools that are not installed with the app.')

foreach ($package in ($packages | Sort-Object -Unique)) {
    $packageDirectory = $null
    foreach ($root in $assets.packageFolders.PSObject.Properties.Name) {
        $candidate = Join-Path $root $package.ToLowerInvariant().Replace('/', '\')
        if (Test-Path -LiteralPath $candidate) {
            $packageDirectory = $candidate
            break
        }
    }
    if (-not $packageDirectory) { throw "Cannot locate the restored package for license collection: $package" }
    $nuspec = @(Get-ChildItem -LiteralPath $packageDirectory -Filter '*.nuspec')
    if ($nuspec.Count -ne 1) { throw "Expected one package manifest for $package." }
    [xml]$manifest = Get-Content -LiteralPath $nuspec[0].FullName -Raw
    $metadata = $manifest.package.metadata
    [void]$notices.AppendLine()
    [void]$notices.AppendLine("===== $package =====")
    [void]$notices.AppendLine([string]$metadata.copyright)
    $files = @(Get-ChildItem -LiteralPath $packageDirectory -File |
        Where-Object { $_.Name -match '(?i)license|notice' } | ForEach-Object FullName)
    if ($metadata.license -and $metadata.license.type -eq 'file') {
        $files += Join-Path $packageDirectory ([string]$metadata.license.InnerText).Replace('/', '\')
    }
    elseif ($metadata.license -and $metadata.license.InnerText -eq 'MIT') {
        if (-not $metadata.copyright) { throw "MIT attribution is missing for $package." }
        $license = "MIT License`n`n$($metadata.copyright)`n`n$mitTerms"
        [void]$notices.AppendLine($license)
    }
    elseif ($metadata.licenseUrl) {
        [void]$notices.AppendLine("License terms: $($metadata.licenseUrl)")
    }
    else {
        throw "Review the unsupported or missing license metadata for $package before distributing."
    }
    foreach ($file in ($files | Sort-Object -Unique)) {
        [void]$notices.AppendLine("--- $(Split-Path $file -Leaf) ---")
        [void]$notices.AppendLine((Get-Content -LiteralPath $file -Raw))
    }
}
$noticePath = Join-Path $PublishDirectory 'THIRD-PARTY-NOTICES.txt'
[IO.File]::WriteAllText($noticePath, $notices.ToString(), [Text.UTF8Encoding]::new($false))
