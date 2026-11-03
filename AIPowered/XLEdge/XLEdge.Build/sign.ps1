# sign.ps1 - signs one build output with the developer's sign_file.cmd (GLSense's pattern: "<file>" "<Configuration>").
#
# Release only. XLEDGE_SKIP_SIGNING=1 skips it (Claude's build and test tools set it - DigiCert signings are metered). A Release
# build without sign_file.cmd and without the skip variable FAILS, so an unsigned release is never built by accident.
# sign_file.cmd lives at the twin root (AIPowered\XLEdge\sign_file.cmd), is copied there by the developer, and is git-ignored.
param(
    [Parameter(Mandatory)][string]$File,
    [Parameter(Mandatory)][string]$Configuration,
    [string]$SignScript
)
$Configuration = $Configuration.Trim()
if ($Configuration -ne 'Release') { exit 0 }
if ($env:XLEDGE_SKIP_SIGNING -eq '1') {
    Write-Output "XLEdge: signing skipped (XLEDGE_SKIP_SIGNING=1): $File"
    exit 0
}
if (-not $SignScript) { $SignScript = Join-Path (Split-Path -Parent $PSScriptRoot) 'sign_file.cmd' }
if (-not (Test-Path -LiteralPath $SignScript)) {
    Write-Output "error XLEDGE0002: sign_file.cmd not found at $SignScript - copy it there, or set XLEDGE_SKIP_SIGNING=1 for an unsigned build"
    exit 1
}
& $SignScript $File $Configuration
if ($LASTEXITCODE -ne 0) {
    Write-Output "error XLEDGE0003: signing $File failed (sign_file.cmd exit $LASTEXITCODE)"
    exit 1
}
Write-Output "XLEdge: signed $File"
exit 0
