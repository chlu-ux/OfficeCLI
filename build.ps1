# Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
# SPDX-License-Identifier: Apache-2.0

[CmdletBinding(DefaultParameterSetName = 'Single')]
param(
    [Parameter(ParameterSetName = 'Single')]
    [ValidateSet('osx-arm64', 'osx-x64', 'linux-x64', 'linux-arm64',
        'linux-musl-x64', 'linux-musl-arm64', 'win-x64', 'win-arm64')]
    [string] $Runtime,

    [Parameter(ParameterSetName = 'All', Mandatory = $true)]
    [switch] $All,

    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release',

    [Nullable[bool]] $Trimmed
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$project = Join-Path $PSScriptRoot 'src\officecli\officecli.csproj'
$targets = [ordered]@{
    'osx-arm64'         = 'officecli-mac-arm64'
    'osx-x64'           = 'officecli-mac-x64'
    'linux-x64'         = 'officecli-linux-x64'
    'linux-arm64'       = 'officecli-linux-arm64'
    'linux-musl-x64'    = 'officecli-linux-alpine-x64'
    'linux-musl-arm64'  = 'officecli-linux-alpine-arm64'
    'win-x64'           = 'officecli-win-x64.exe'
    'win-arm64'         = 'officecli-win-arm64.exe'
}

if (-not $All -and [string]::IsNullOrWhiteSpace($Runtime)) {
    throw 'Specify -Runtime <RID> or -All.'
}

$effectiveTrimmed = if ($null -ne $Trimmed) {
    $Trimmed.Value
} else {
    # A 32-bit Windows host can run out of address space in the linker. Keep
    # local cross-RID diagnostics useful there; release artifacts remain CI-built.
    [Environment]::Is64BitProcess
}

$selected = if ($All) { @($targets.Keys) } else { @($Runtime) }
$outputDir = Join-Path $PSScriptRoot (Join-Path 'bin' $Configuration.ToLowerInvariant())
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$sdk = (& dotnet --version).Trim()
$commit = try { (& git -C $PSScriptRoot rev-parse --short HEAD).Trim() } catch { 'unknown' }
Write-Host "SDK: $sdk"
Write-Host "Commit: $commit"
Write-Host "Configuration: $Configuration"
Write-Host "PublishTrimmed: $($effectiveTrimmed.ToString().ToLowerInvariant())"

foreach ($rid in $selected) {
    $assetName = $targets[$rid]
    $stageDir = Join-Path $outputDir ('.stage-' + [Guid]::NewGuid().ToString('N'))
    $stagedAsset = Join-Path $outputDir ($assetName + '.new')
    $destination = Join-Path $outputDir $assetName
    $stagedPdb = Join-Path $outputDir (($assetName -replace '\.exe$', '') + '.pdb.new')
    $destinationPdb = $stagedPdb.Substring(0, $stagedPdb.Length - 4)

    try {
        New-Item -ItemType Directory -Path $stageDir | Out-Null
        Write-Host "[$Configuration] Building $rid -> $assetName"
        & dotnet publish $project -c $Configuration -r $rid -o $stageDir `
            --self-contained true --nologo -p:PublishAot=false `
            "-p:PublishTrimmed=$($effectiveTrimmed.ToString().ToLowerInvariant())"
        if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed for $rid (exit $LASTEXITCODE)." }

        $publishedBinary = Join-Path $stageDir $(if ($rid.StartsWith('win-')) { 'officecli.exe' } else { 'officecli' })
        if (-not (Test-Path -LiteralPath $publishedBinary -PathType Leaf)) {
            throw "Published binary was not found: $publishedBinary"
        }

        Copy-Item -LiteralPath $publishedBinary -Destination $stagedAsset
        Move-Item -LiteralPath $stagedAsset -Destination $destination -Force

        $publishedPdb = Join-Path $stageDir 'officecli.pdb'
        if (Test-Path -LiteralPath $publishedPdb -PathType Leaf) {
            Copy-Item -LiteralPath $publishedPdb -Destination $stagedPdb
            Move-Item -LiteralPath $stagedPdb -Destination $destinationPdb -Force
        }

        $hash = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash.ToLowerInvariant()
        Write-Host "RID: $rid"
        Write-Host "Artifact: $destination"
        Write-Host "SHA-256: $hash"
    }
    finally {
        if (Test-Path -LiteralPath $stageDir) { Remove-Item -LiteralPath $stageDir -Recurse -Force }
        if (Test-Path -LiteralPath $stagedAsset) { Remove-Item -LiteralPath $stagedAsset -Force }
        if (Test-Path -LiteralPath $stagedPdb) { Remove-Item -LiteralPath $stagedPdb -Force }
    }
}
