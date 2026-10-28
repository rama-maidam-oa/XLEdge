# post_build.ps1 - XLEdge.Addin.Core
#
# Publishes the freshly built core as a release: SetupFiles\<Configuration>\Manifest\ gets
#   v<version>_<releaseDate>.zip  - every file of the build output except *.pdb / *.xml
#   manifest.json                 - a one-entry array the XLEdge loader reads (version, releaseDate,
#                                   folderName, fileName, checksum = SHA-256 of the zip, contractVersion, ...)
# The shell's own post-build copies this folder into its bin\<Configuration>\AddinCore\Manifest\ so the next
# Excel start (or Reload Add-in) installs it.
param(
    [Parameter(Mandatory)][string]$TargetDir,
    [Parameter(Mandatory)][string]$Configuration
)
$ErrorActionPreference = 'Stop'
try {
    # MSBuild passes "$(TargetDir) " (a trailing space keeps the closing quote from being escaped by the trailing '\').
    $TargetDir = $TargetDir.Trim().TrimEnd('\', '"').Trim()
    $projectDir = $PSScriptRoot
    $outDir = Join-Path $projectDir "SetupFiles\$Configuration\Manifest"
    $coreDll = Join-Path $TargetDir 'XLEdge.Addin.Core.dll'
    $contractsDll = Join-Path $TargetDir 'XLEdge.Contracts.dll'

    # Assembly version without loading (no file lock).
    $version = [Reflection.AssemblyName]::GetAssemblyName($coreDll).Version.ToString()

    # ContractInfo.Version is a constant: read it in a child process so this one never locks the DLL.
    $contractVersion = [int](& powershell -NoProfile -Command "([Reflection.Assembly]::LoadFile('$contractsDll')).GetType('XLEdge.Contracts.ContractInfo').GetField('Version').GetRawConstantValue()")
    if ($contractVersion -lt 1) { throw "Could not read XLEdge.Contracts.ContractInfo.Version from $contractsDll" }

    $releaseDate = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
    # Same rule as ReleaseHistoryStore.BuildFolderName: characters invalid in file names become '-'.
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $safeDate = -join ($releaseDate.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '-' } else { $_ } })
    $folderName = "v${version}_$safeDate"
    $fileName = "$folderName.zip"

    $notesFile = Join-Path $projectDir 'ReleaseNotes.txt'
    $notes = 'XLEdge core build'
    if (Test-Path $notesFile) {
        $first = Get-Content $notesFile -TotalCount 1
        if ($first -and $first.Trim()) { $notes = $first.Trim() }
    }

    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    Get-ChildItem $outDir -File | Where-Object { $_.Extension -eq '.zip' -or $_.Name -eq 'manifest.json' } | Remove-Item -Force

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zipPath = Join-Path $outDir $fileName
    $archive = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $root = (Resolve-Path $TargetDir).Path.TrimEnd('\') + '\'
        Get-ChildItem $TargetDir -Recurse -File |
            Where-Object { $_.Extension -notin '.pdb', '.xml', '.zip' -and $_.FullName -notlike "$($root)AddinCore\*" } |
            ForEach-Object {
                $entry = $_.FullName.Substring($root.Length).Replace('\', '/')
                [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $_.FullName, $entry, [IO.Compression.CompressionLevel]::Optimal) | Out-Null
            }
    }
    finally {
        $archive.Dispose()
    }

    $sha = [Security.Cryptography.SHA256]::Create()
    try { $checksum = ([BitConverter]::ToString($sha.ComputeHash([IO.File]::ReadAllBytes($zipPath)))).Replace('-', '') }
    finally { $sha.Dispose() }

    $entryObject = [ordered]@{
        version         = $version
        releaseDate     = $releaseDate
        folderName      = $folderName
        fileName        = $fileName
        checksum        = $checksum
        notes           = $notes
        mandatory       = $false
        contractVersion = $contractVersion
    }
    $json = ConvertTo-Json -InputObject @([pscustomobject]$entryObject) -Depth 3
    [IO.File]::WriteAllText((Join-Path $outDir 'manifest.json'), $json, (New-Object Text.UTF8Encoding $false))

    Write-Output "XLEdge.Addin.Core post-build: published $fileName (contract $contractVersion) to $outDir"
    exit 0
}
catch {
    Write-Output "error XLEDGE0001: XLEdge.Addin.Core post-build failed: $($_.Exception.Message)"
    exit 1
}
