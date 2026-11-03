# verify-signed.ps1 - the setup project's last PreBuild step: every file of ours the MSI ships must carry a valid Authenticode
# signature - the shell (the obj\Release copy the Primary Output packages, and the bin copy), both loaders, Contracts, Shared,
# Loader.Core, and the core DLL + shared DLLs inside the installer payload zip. A build that reused unsigned outputs (an earlier
# XLEDGE_SKIP_SIGNING=1 build, or a Build that skipped up-to-date projects) fails here instead of shipping them.
# XLEDGE_SKIP_SIGNING=1: an unsigned MSI is allowed, with a warning (Claude's content checks).
param(
    [Parameter(Mandatory)][string]$ShellBin,
    [Parameter(Mandatory)][string]$ShellObj,
    [Parameter(Mandatory)][string]$Payload
)
if ($env:XLEDGE_SKIP_SIGNING -eq '1') {
    Write-Output "warning XLEDGE0004: building an UNSIGNED MSI (XLEDGE_SKIP_SIGNING=1) - not for release."
    exit 0
}
$files = @()
foreach ($name in 'XLEdge.dll', 'adxloader.dll', 'adxloader64.dll', 'XLEdge.Contracts.dll', 'XLEdge.Shared.dll', 'XLEdge.Loader.Core.dll') {
    $files += Join-Path $ShellBin $name
}
$files += Join-Path $ShellObj 'XLEdge.dll'
$unsigned = New-Object System.Collections.Generic.List[string]
foreach ($file in $files) {
    if (-not (Test-Path -LiteralPath $file)) { $unsigned.Add("$file (missing)"); continue }
    if ((Get-AuthenticodeSignature -LiteralPath $file).Status -ne 'Valid') { $unsigned.Add($file) }
}
$temp = Join-Path ([IO.Path]::GetTempPath()) ("XLEdgeVerify_" + [Guid]::NewGuid().ToString('N'))
try {
    if (-not (Test-Path -LiteralPath $Payload)) { $unsigned.Add("$Payload (missing)") }
    else {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [IO.Compression.ZipFile]::ExtractToDirectory($Payload, $temp)
        foreach ($name in 'XLEdge.Addin.Core.dll', 'XLEdge.Contracts.dll', 'XLEdge.Shared.dll') {
            $inside = Join-Path $temp $name
            if (-not (Test-Path -LiteralPath $inside)) { $unsigned.Add("$Payload!$name (missing)"); continue }
            if ((Get-AuthenticodeSignature -LiteralPath $inside).Status -ne 'Valid') { $unsigned.Add("$Payload!$name") }
        }
    }
}
finally {
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }
}
if ($unsigned.Count -gt 0) {
    Write-Output ("error XLEDGE0005: not signed - rebuild Release (/Rebuild) with sign_file.cmd in place: " + ($unsigned -join '; '))
    exit 1
}
Write-Output "XLEdge: every shipped file of ours is signed."
exit 0
