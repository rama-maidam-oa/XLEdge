<#
.SYNOPSIS
    XLEdge AIPowered health check for QA machines: what is installed, registered and loaded, read-only.

.DESCRIPTION
    Reads what the add-in itself uses and reports one line per check (PASS / WARN / FAIL), or JSON with -Json:
      Registration     HKCU Excel add-in entry XLEdge.AddinModule (LoadBehavior 3 = loaded when Excel starts)
      Loader           the registered adxloader (COM InprocServer32) - its folder is the install root
      Shell            XLEdge.dll and its file version;  Signature: its Authenticode signature
      SharedAssemblies XLEdge.Contracts / Shared / Loader.Core next to it (same file version as the shell)
      AddinCore        <install root>\AddinCore: the catalogued core releases (ReleaseHistory.json)
      SelectedRelease  AddinCore\SelectedRelease.txt - the release Excel loads next
      StagedRelease    AddinCore\Manifest\ - a release installed at the next Excel start
      Log              today's XLEdge log: ERROR lines (count and the last ones)
      GLSense          the GLSense add-in entry (side by side)
    Nothing is written. Exit code 1 when any check FAILs, else 0.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File XLEdgeHealthCheck.ps1
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File XLEdgeHealthCheck.ps1 -Json > health.json
#>
param(
    # The install root; default: the folder of the registered adxloader.
    [string]$InstallRoot,
    # %LOCALAPPDATA%\ORBIT\Excel_Logs (logs under XLEdge_Logs\Logs).
    [string]$DataRoot = (Join-Path $env:LOCALAPPDATA 'ORBIT\Excel_Logs'),
    # HKCU:\Software; tests pass a temporary key.
    [string]$RegistryRoot = 'HKCU:\Software',
    # How many of today's ERROR lines to show.
    [int]$LogErrors = 10,
    [switch]$Json
)

$ErrorActionPreference = 'Continue'
$Clsid = '{80B0FB76-A5F4-41E0-B283-C175B7374B60}'
$results = New-Object System.Collections.Generic.List[object]
function Report([string]$check, [string]$status, [string]$detail) {
    $results.Add([pscustomobject][ordered]@{ check = $check; status = $status; detail = $detail })
}
function RegValue([string]$path, [string]$name) {
    try { (Get-ItemProperty -LiteralPath $path -Name $name -ErrorAction Stop).$name } catch { $null }
}
function FileVersion([string]$path) {
    try { [Diagnostics.FileVersionInfo]::GetVersionInfo($path).FileVersion } catch { $null }
}

# ---- Registration ------------------------------------------------------------------------------------------------------
$addinKey = Join-Path $RegistryRoot 'Microsoft\Office\Excel\Addins\XLEdge.AddinModule'
if (-not (Test-Path -LiteralPath $addinKey)) {
    Report 'Registration' 'FAIL' "XLEdge is not registered as an Excel add-in ($addinKey missing) - install the MSI, or register the build."
}
else {
    $loadBehavior = RegValue $addinKey 'LoadBehavior'
    if ($loadBehavior -eq 3) { Report 'Registration' 'PASS' 'XLEdge.AddinModule registered, LoadBehavior 3 (loaded when Excel starts).' }
    else { Report 'Registration' 'WARN' "LoadBehavior $loadBehavior`: Excel does not load XLEdge at start (3 expected; Excel sets 2 after a crash at start - re-enable it in COM Add-ins)." }
}

# ---- Loader / install root -----------------------------------------------------------------------------------------------
$serverKey = Join-Path $RegistryRoot "Classes\CLSID\$Clsid\InprocServer32"
$loader = RegValue $serverKey '(default)'
if (-not $loader) {
    Report 'Loader' 'FAIL' "No COM server registered for XLEdge ($serverKey)."
}
elseif (-not (Test-Path -LiteralPath $loader)) {
    Report 'Loader' 'FAIL' "The registered loader is missing: $loader"
}
else {
    Report 'Loader' 'PASS' "Registered loader: $loader"
    if (-not $InstallRoot) { $InstallRoot = Split-Path -Parent $loader }
}

if (-not $InstallRoot) {
    Report 'Shell' 'FAIL' 'No install root: XLEdge is not registered and -InstallRoot was not given.'
}
else {
    # ---- Shell + signature ---------------------------------------------------------------------------------------------
    $shell = Join-Path $InstallRoot 'XLEdge.dll'
    $shellVersion = $null
    if (-not (Test-Path -LiteralPath $shell)) {
        Report 'Shell' 'FAIL' "XLEdge.dll is missing in $InstallRoot"
    }
    else {
        $shellVersion = FileVersion $shell
        Report 'Shell' 'PASS' "XLEdge.dll $shellVersion in $InstallRoot"
        $signature = (Get-AuthenticodeSignature -LiteralPath $shell).Status
        if ("$signature" -eq 'Valid') { Report 'Signature' 'PASS' 'XLEdge.dll is signed (Authenticode Valid).' }
        else { Report 'Signature' 'WARN' "XLEdge.dll is not signed ($signature) - a development build?" }
    }

    # ---- Shared assemblies ---------------------------------------------------------------------------------------------
    $missing = @(); $mismatched = @()
    foreach ($name in 'XLEdge.Contracts.dll', 'XLEdge.Shared.dll', 'XLEdge.Loader.Core.dll') {
        $path = Join-Path $InstallRoot $name
        if (-not (Test-Path -LiteralPath $path)) { $missing += $name; continue }
        $version = FileVersion $path
        if ($shellVersion -and $version -ne $shellVersion) { $mismatched += "$name $version" }
    }
    $singleDll = ($missing.Count -eq 3) -and -not (Test-Path -LiteralPath (Join-Path $InstallRoot 'AddinCore'))
    if ($singleDll) { Report 'SharedAssemblies' 'WARN' 'The single-DLL XLEdge (11.1.2 or older): no Contracts/Shared/Loader.Core - none expected there.' }
    elseif ($missing.Count -gt 0) { Report 'SharedAssemblies' 'FAIL' ("Missing next to the shell: " + ($missing -join ', ')) }
    elseif ($mismatched.Count -gt 0) { Report 'SharedAssemblies' 'WARN' ("Not the shell's version ($shellVersion): " + ($mismatched -join ', ')) }
    else { Report 'SharedAssemblies' 'PASS' "XLEdge.Contracts, Shared, Loader.Core present ($shellVersion)." }

    # ---- AddinCore -----------------------------------------------------------------------------------------------------
    $addinCore = Join-Path $InstallRoot 'AddinCore'
    if (-not (Test-Path -LiteralPath $addinCore)) {
        Report 'AddinCore' 'WARN' "AddinCore is not set up in $InstallRoot - the old single-DLL XLEdge, or Excel has not started XLEdge yet."
    }
    else {
        $historyFile = Join-Path $addinCore 'ReleaseHistory.json'
        $versions = Join-Path $addinCore 'Versions'
        if (-not (Test-Path -LiteralPath $historyFile)) {
            Report 'AddinCore' 'WARN' 'No ReleaseHistory.json yet (no core installed so far).'
        }
        else {
            try {
                # Windows PowerShell's ConvertFrom-Json hands a JSON array over as ONE object: enumerate it.
                $history = @(Get-Content -LiteralPath $historyFile -Raw | ConvertFrom-Json | ForEach-Object { $_ })
                $list = ($history | ForEach-Object { "$($_.Version) ($($_.FolderName), $($_.Source))" }) -join '; '
                Report 'AddinCore' 'PASS' ("$($history.Count) catalogued release(s): " + $list)
            }
            catch { Report 'AddinCore' 'FAIL' "ReleaseHistory.json cannot be read: $($_.Exception.Message)" }
        }

        $selectedFile = Join-Path $addinCore 'SelectedRelease.txt'
        # An empty file is "no choice" to the loader (RememberLoaded writes ""); Get-Content -Raw gives $null for it.
        $selected = if (Test-Path -LiteralPath $selectedFile) { "$(Get-Content -LiteralPath $selectedFile -Raw)".Trim() } else { '' }
        if (-not $selected) {
            Report 'SelectedRelease' 'PASS' 'No release chosen yet: the newest catalogued release loads.'
        }
        else {
            if (Test-Path -LiteralPath (Join-Path $versions $selected)) { Report 'SelectedRelease' 'PASS' "Excel loads $selected next." }
            else { Report 'SelectedRelease' 'FAIL' "SelectedRelease.txt names '$selected', which is gone: the next start falls back to the newest catalogued release." }
        }

        $manifestDir = Join-Path $addinCore 'Manifest'
        $staged = @(Get-ChildItem -LiteralPath $manifestDir -Filter *.zip -File -ErrorAction SilentlyContinue)
        if ($staged.Count -gt 0) { Report 'StagedRelease' 'PASS' ("Staged: " + ($staged.Name -join ', ') + " - installed at the next Excel start.") }
        else { Report 'StagedRelease' 'PASS' 'Nothing staged.' }
    }
}

# ---- Log -----------------------------------------------------------------------------------------------------------------
$log = Join-Path $DataRoot ('XLEdge_Logs\Logs\XLEdge_Logs_' + (Get-Date).ToString('dd-MMM-yyyy', [Globalization.CultureInfo]::InvariantCulture) + '.log')
if (-not (Test-Path -LiteralPath $log)) {
    Report 'Log' 'WARN' "No XLEdge log for today ($log) - XLEdge has not run today."
}
else {
    try {
        # Shared read: Excel may have the log open.
        $stream = New-Object IO.FileStream($log, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try { $text = (New-Object IO.StreamReader($stream)).ReadToEnd() } finally { $stream.Dispose() }
        # Logger.cs: NLog's "${longdate}|" then "{indent}ERROR | HH:mm:ss | message".
        $errors = @($text -split "`r?`n" | Where-Object { $_ -match '^(\d{4}-\d{2}-\d{2} [0-9:.]+\|)?\s*ERROR\s*\|' })
        $last = @($errors | Select-Object -Last $LogErrors | ForEach-Object { $_.Trim() })
        $detail = "$($errors.Count) ERROR line(s) today in $log"
        if ($last.Count -gt 0) { $detail += ': ' + ($last -join ' || ') }
        Report 'Log' 'PASS' $detail
    }
    catch { Report 'Log' 'WARN' "The log could not be read: $($_.Exception.Message)" }
}

# ---- GLSense (side by side) ----------------------------------------------------------------------------------------------
$glsenseKey = Join-Path $RegistryRoot 'Microsoft\Office\Excel\Addins\GLSense.AddinModule'
if (Test-Path -LiteralPath $glsenseKey) { Report 'GLSense' 'PASS' ("GLSense registered, LoadBehavior " + (RegValue $glsenseKey 'LoadBehavior') + ' (side by side with XLEdge).') }
else { Report 'GLSense' 'PASS' 'GLSense not installed.' }

# ---- Output --------------------------------------------------------------------------------------------------------------
if ($Json) {
    # A plain array (Windows PowerShell 5.1 cannot serialize the generic list itself).
    ConvertTo-Json -InputObject ([object[]]$results.ToArray()) -Depth 3
}
else {
    foreach ($r in $results) { Write-Output ("[{0}] {1}: {2}" -f $r.status, $r.check, $r.detail) }
    $fails = @($results | Where-Object status -eq 'FAIL').Count
    $warns = @($results | Where-Object status -eq 'WARN').Count
    Write-Output ("XLEdge health check: {0} FAIL, {1} WARN." -f $fails, $warns)
}
if (@($results | Where-Object status -eq 'FAIL').Count -gt 0) { exit 1 }
exit 0
