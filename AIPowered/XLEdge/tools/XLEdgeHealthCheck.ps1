<#
.SYNOPSIS
    XLEdge AIPowered health check for QA machines: what is installed, registered and loaded, read-only.

.DESCRIPTION
    Reads what the add-in itself uses and reports one line per check (PASS / WARN / FAIL), or JSON with -Json:
      Registration     HKCU Excel add-in entry XLEdge.AddinModule (LoadBehavior 3 = loaded when Excel starts)
      Loader           the adxloader registered for Excel's bitness (COM InprocServer32: adxloader64.dll in the 64-bit view,
                       adxloader.dll in the WOW6432Node view for 32-bit Excel) - its folder is the install root
      OtherLoader      the loader of the other bitness (information: adxregistrator registers both)
      Shell            XLEdge.dll and its file version;  Signature: its Authenticode signature
      SharedAssemblies XLEdge.Contracts / Shared / Loader.Core next to it (same file version as the shell)
      AddinCore        <install root>\AddinCore: the catalogued core releases (ReleaseHistory.json)
      StagedRelease    AddinCore\Manifest\ - a release installed at the next Excel start (zip + manifest.json), refused ones
      SelectedRelease  AddinCore\SelectedRelease.txt - the release Excel loads next (unless one is staged)
      Log              today's XLEdge log files (incl. rollovers): ERROR lines; XLEdge_Error_*.log (the logger failed to start)
      GLSense          the GLSense add-in entry (side by side)
    The rules follow UpdateBootstrapper (XLEdge.Loader.Core). Nothing is written. Exit code 1 when any check FAILs, else 0.

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
    # 32 or 64; default: read from Office's Click-to-Run configuration (0 = unknown).
    [ValidateSet(0, 32, 64)][int]$ExcelBitness = 0,
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
function HasDlls([string]$folder) {
    (Test-Path -LiteralPath $folder -PathType Container) -and @(Get-ChildItem -LiteralPath $folder -Filter *.dll -File -ErrorAction SilentlyContinue).Count -gt 0
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

# ---- Loaders / install root ----------------------------------------------------------------------------------------------
if ($ExcelBitness -eq 0) {
    $platform = RegValue 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' 'Platform'
    if ($platform -eq 'x64') { $ExcelBitness = 64 } elseif ($platform -eq 'x86') { $ExcelBitness = 32 }
}
$views = @(
    [pscustomobject]@{ Bits = 64; Key = (Join-Path $RegistryRoot "Classes\CLSID\$Clsid\InprocServer32") },
    [pscustomobject]@{ Bits = 32; Key = (Join-Path $RegistryRoot "Classes\WOW6432Node\CLSID\$Clsid\InprocServer32") }
)
foreach ($v in $views) {
    $v | Add-Member NoteProperty Path (RegValue $v.Key '(default)')
    $v | Add-Member NoteProperty Exists ([bool]($v.Path -and (Test-Path -LiteralPath $v.Path)))
}
$used = if ($ExcelBitness -eq 32) { $views[1] } else { $views[0] }
$other = if ($ExcelBitness -eq 32) { $views[0] } else { $views[1] }
$excelText = if ($ExcelBitness -eq 0) { 'Excel bitness unknown - assuming 64-bit' } else { "$ExcelBitness-bit Excel" }

if (-not $used.Path) { Report 'Loader' 'FAIL' "No COM server registered for XLEdge for $($used.Bits)-bit Excel ($($used.Key)) [$excelText]." }
elseif (-not $used.Exists) { Report 'Loader' 'FAIL' "The registered $($used.Bits)-bit loader is missing: $($used.Path) [$excelText]" }
else { Report 'Loader' 'PASS' "Registered $($used.Bits)-bit loader: $($used.Path) [$excelText]" }

if (-not $other.Path) { Report 'OtherLoader' 'WARN' "No $($other.Bits)-bit loader registered ($($other.Key)) - only matters for $($other.Bits)-bit Excel." }
elseif (-not $other.Exists) { Report 'OtherLoader' 'WARN' "The registered $($other.Bits)-bit loader is missing: $($other.Path) - only matters for $($other.Bits)-bit Excel." }
else { Report 'OtherLoader' 'PASS' "Registered $($other.Bits)-bit loader: $($other.Path)" }

# The install root: Excel's loader's folder, else the other loader's; a registered path whose file is gone still names it.
$rootFrom = @($used, $other) | Where-Object { $_.Path } | Select-Object -First 1
if (-not $InstallRoot -and $rootFrom) { $InstallRoot = Split-Path -Parent $rootFrom.Path }

if (-not $InstallRoot) {
    Report 'Shell' 'FAIL' 'No install root: no XLEdge loader is registered and -InstallRoot was not given.'
}
elseif (-not (Test-Path -LiteralPath $InstallRoot -PathType Container)) {
    Report 'Shell' 'FAIL' "The install folder of the registered loader is gone: $InstallRoot - reinstall XLEdge."
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
        $rebuilt = 'the next Excel start rebuilds it from the release folders'' manifest.json'
        if (-not (Test-Path -LiteralPath $historyFile)) {
            Report 'AddinCore' 'WARN' "No ReleaseHistory.json yet (no core installed so far, or it was deleted: $rebuilt)."
        }
        else {
            $history = $null
            try {
                # Windows PowerShell's ConvertFrom-Json hands a JSON array over as ONE object: enumerate it.
                $history = @(Get-Content -LiteralPath $historyFile -Raw | ConvertFrom-Json -ErrorAction Stop | ForEach-Object { $_ } | Where-Object { $_ })
            }
            catch { Report 'AddinCore' 'WARN' "ReleaseHistory.json is unreadable ($($_.Exception.Message)) - $rebuilt." }
            if ($null -ne $history) {
                if ($history.Count -eq 0) { Report 'AddinCore' 'WARN' "ReleaseHistory.json is empty - $rebuilt." }
                else {
                    $list = ($history | ForEach-Object { "$($_.Version) ($($_.FolderName), $($_.Source))" }) -join '; '
                    Report 'AddinCore' 'PASS' ("$($history.Count) catalogued release(s): " + $list)
                }
            }
        }

        # ---- Staged release (UpdateBootstrapper.FindStagedZip: none without manifest.json) ----------------------------
        $manifestDir = Join-Path $addinCore 'Manifest'
        $zips = @(Get-ChildItem -LiteralPath $manifestDir -Filter *.zip -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -eq '.zip' })
        $rejected = @(Get-ChildItem -LiteralPath $manifestDir -Filter '*.zip.rejected-*' -File -ErrorAction SilentlyContinue)
        $hasManifest = Test-Path -LiteralPath (Join-Path $manifestDir 'manifest.json')
        $staged = $zips.Count -gt 0 -and $hasManifest
        $notes = @()
        if ($rejected.Count -gt 0) { $notes += 'refused and parked (not retried): ' + ($rejected.Name -join ', ') + ' - see the log of that start' }
        if ($zips.Count -gt 0 -and -not $hasManifest) {
            Report 'StagedRelease' 'WARN' ("Zip without Manifest\manifest.json - NOT installed: " + ($zips.Name -join ', ') + ($(if ($notes) { '; ' + ($notes -join '; ') } else { '' })))
        }
        elseif ($staged) {
            $detail = "Staged: " + ($zips.Name -join ', ') + " - installed and loaded at the next Excel start"
            if ($notes) { Report 'StagedRelease' 'WARN' ($detail + '; ' + ($notes -join '; ')) } else { Report 'StagedRelease' 'PASS' ($detail + '.') }
        }
        elseif ($notes) { Report 'StagedRelease' 'WARN' ('Nothing staged; ' + ($notes -join '; ')) }
        else { Report 'StagedRelease' 'PASS' 'Nothing staged.' }

        # ---- Selected release (UpdateBootstrapper: a remembered folder loads only with DLLs in it) -------------------
        $selectedFile = Join-Path $addinCore 'SelectedRelease.txt'
        # An empty file is "no choice" to the loader (RememberLoaded writes ""); Get-Content -Raw gives $null for it.
        $selected = if (Test-Path -LiteralPath $selectedFile) { "$(Get-Content -LiteralPath $selectedFile -Raw)".Trim() } else { '' }
        $stagedNote = if ($staged) { ' (but the staged release is installed and loaded first)' } else { '' }
        $selectedFolder = if ($selected) { Join-Path $versions $selected } else { $null }
        if (-not $selected) {
            Report 'SelectedRelease' 'PASS' "No release chosen yet: the newest catalogued release loads$stagedNote."
        }
        elseif (-not (Test-Path -LiteralPath $selectedFolder -PathType Container)) {
            Report 'SelectedRelease' 'FAIL' "SelectedRelease.txt names '$selected', which is gone: the next start falls back to the newest catalogued release$stagedNote."
        }
        elseif (-not (HasDlls $selectedFolder)) {
            Report 'SelectedRelease' 'FAIL' "SelectedRelease.txt names '$selected', whose folder has no DLLs: the next start falls back to the newest catalogued release$stagedNote."
        }
        else { Report 'SelectedRelease' 'PASS' "Excel loads $selected next$stagedNote." }
    }
}

# ---- Log -----------------------------------------------------------------------------------------------------------------
$today = (Get-Date).ToString('dd-MMM-yyyy', [Globalization.CultureInfo]::InvariantCulture)
$logsFolder = Join-Path $DataRoot 'XLEdge_Logs\Logs'
# Logger.cs: XLEdge_Logs_<date>.log, rolled over above 20 MB into XLEdge_Logs_<date>(1).log, (2)...
$logFiles = @(Get-ChildItem -LiteralPath $logsFolder -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -eq "XLEdge_Logs_$today.log" -or $_.Name -match ('^XLEdge_Logs_' + [regex]::Escape($today) + '\(\d+\)\.log$') } |
    Sort-Object LastWriteTime)
# Logger.cs writes XLEdge_Error_<date>.log when the logger itself cannot start.
$errorFile = Join-Path $logsFolder "XLEdge_Error_$today.log"
function ReadShared([string]$path) {
    # Shared read: Excel may have the log open.
    $stream = New-Object IO.FileStream($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try { (New-Object IO.StreamReader($stream)).ReadToEnd() } finally { $stream.Dispose() }
}
$logStatus = 'PASS'; $logDetail = $null
if ($logFiles.Count -eq 0) {
    $logStatus = 'WARN'; $logDetail = "No XLEdge log for today ($(Join-Path $logsFolder "XLEdge_Logs_$today.log")) - XLEdge has not run today."
}
else {
    try {
        $errors = @()
        foreach ($f in $logFiles) {
            # Logger.cs: NLog's "${longdate}|" then "{indent}ERROR | HH:mm:ss | message".
            $errors += @((ReadShared $f.FullName) -split "`r?`n" | Where-Object { $_ -match '^(\d{4}-\d{2}-\d{2} [0-9:.]+\|)?\s*ERROR\s*\|' })
        }
        $last = @($errors | Select-Object -Last $LogErrors | ForEach-Object { $_.Trim() })
        $logDetail = "$($errors.Count) ERROR line(s) today in " + ($logFiles.Name -join ', ')
        if ($last.Count -gt 0) { $logDetail += ': ' + ($last -join ' || ') }
    }
    catch { $logStatus = 'WARN'; $logDetail = "The log could not be read: $($_.Exception.Message)" }
}
if (Test-Path -LiteralPath $errorFile) {
    $text = try { ((ReadShared $errorFile) -replace '\s+', ' ').Trim() } catch { "(unreadable: $($_.Exception.Message))" }
    if ($text.Length -gt 300) { $text = $text.Substring(0, 300) + '...' }
    $logStatus = 'WARN'; $logDetail = "The logger failed to start today ($(Split-Path -Leaf $errorFile)): $text; " + $logDetail
}
Report 'Log' $logStatus $logDetail

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
