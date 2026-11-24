# This script will scan the registry for all keys associated with install,
# create a log file, and optionally delete them (not recommended - uninstall via Inno Setup unins000.exe)

$ProgID = "SeleniumVBA"
$GuidPrefix = "38ED0FFA-E3F3-41C4-B601-"
$LogPath = ".\COM_Registry_Log.txt"
$Delete = $false

$ProgIDPattern = "$ProgID*"

# Initialize log file (clear any existing content from previous runs)
if (Test-Path $LogPath) { Remove-Item $LogPath -Force }

$log = New-Object System.Collections.Generic.List[string]
function Log($msg) {
    $timestamped = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    $log.Add($timestamped)
    Write-Host $timestamped
}

# Tracks how many keys matched at each scanned location, for the summary at end.
# Key: "$basePath -> $pattern", Value: match count.
$summary = New-Object System.Collections.Specialized.OrderedDictionary

function MatchKeys($basePath, $pattern) {
    Log "Scanning: $basePath with pattern: $pattern"
    try {
        $items = Get-ChildItem -Path "Registry::$basePath" -ErrorAction SilentlyContinue
        $matchedKeys = @($items | Where-Object {
            $_.PSChildName -like $pattern -or $_.PSChildName -like "{$pattern"
        } | ForEach-Object { $_.Name })
        Log "Matched $($matchedKeys.Count) keys"
        $summary["$basePath -> $pattern"] = $matchedKeys.Count
        return $matchedKeys
    } catch {
        Log "Failed to scan: $basePath - $_"
        $summary["$basePath -> $pattern"] = "ERROR"
        return @()
    }
}

function DeleteKey($keyPath) {
    try {
        Remove-Item -Path "Registry::$keyPath" -Recurse -Force -ErrorAction Stop
        Log "Deleted: $keyPath"
    } catch {
        Log "Failed to delete: $keyPath - $_"
    }
}

# Registry hives to scan
$hives = @(
    "HKCR",
    "HKCR\Wow6432Node",
    "HKCU\Software\Classes",
    "HKCU\Software\Classes\Wow6432Node",
    "HKCU\Software\Wow6432Node",
    "HKLM\Software\Classes",
    "HKLM\Software\Classes\Wow6432Node",
    "HKLM\Software\Wow6432Node\Classes"
)

# Targets to match (Path is relative to each hive; empty means scan the hive root)
$targets = @(
    @{ Path = "";          Pattern = $ProgIDPattern },
    @{ Path = "CLSID";     Pattern = "$GuidPrefix*" },
    @{ Path = "Interface"; Pattern = "$GuidPrefix*" },
    @{ Path = "TypeLib";   Pattern = "$GuidPrefix*" }
)

foreach ($hive in $hives) {
    foreach ($target in $targets) {
        $base = if ($target.Path -eq "") { $hive } else { "$hive\$($target.Path)" }
        $matchedKeys = MatchKeys $base $target.Pattern
        foreach ($match in $matchedKeys) {
            Log "Found: $match"
            if ($Delete) { DeleteKey $match }
        }
    }
}

# Trusted Location tracking for summary
$trustedLocationCount = 0

# Detect Office version
function Get-OfficeVersion($app) {
    $roots = @(
        "HKCU:\Software\Microsoft\Office",
        "HKLM:\Software\Microsoft\Office",
        "HKLM:\Software\Wow6432Node\Microsoft\Office"  # 32-bit Office on 64-bit Windows
    )
    $latest = $null
    $latestRoot = $null
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        $versions = Get-ChildItem -Path $root -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
            Sort-Object -Property {[version]$_.PSChildName} -Descending

        foreach ($version in $versions) {
            $strictPath = "$root\$($version.PSChildName)\$app\Security\Trusted Locations"
            $broadPath  = "$root\$($version.PSChildName)\$app"

            if (Test-Path $strictPath) {
                Log "Found Office version: $($version.PSChildName) (strict match)"
                Log "Office installed on: $root"
                return $version.PSChildName
            }
            elseif (-not $latest -and (Test-Path $broadPath)) {
                $latest = $version.PSChildName  # fallback candidate
                $latestRoot = $root
            }
        }
    }
    if ($latest) {
        Log "Found Office version: $latest (broad-path fallback)"
        Log "Office installed on: $latestRoot"
    } else {
        Log "Office install not found"
    }
    return $latest
}

# Find Trusted Location key
function FindTrustedLocation($app) {
    $version = Get-OfficeVersion $app
    if ($version) {
        $keyPath = "HKCU:\Software\Microsoft\Office\$version\$app\Security\Trusted Locations\$ProgID"
        if (Test-Path -Path $keyPath) {
            Log "Found Trusted Location: $keyPath"
            $script:trustedLocationCount++
            if ($Delete) {
                try {
                    Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                    Log "Deleted Trusted Location: $keyPath"
                } catch {
                    Log "Failed to delete $keyPath - $_"
                }
            }
        } else {
            Log "No Trusted Location found for $app ($version): $keyPath"
        }
    } else {
        Log "No Trusted Location or Office version found for $app"
    }
}

# Find trusted locations for supported apps
FindTrustedLocation "Excel"
FindTrustedLocation "Word"
FindTrustedLocation "PowerPoint"
FindTrustedLocation "Access"

# ---- Summary ----
Log ""
Log "=== Summary ==="

# Group totals by target type across all hives
$totalsByTarget = @{
    "ProgIDs"    = 0
    "CLSIDs"     = 0
    "Interfaces" = 0
    "TypeLibs"   = 0
}

foreach ($entry in $summary.GetEnumerator()) {
    $key = $entry.Key
    $count = $entry.Value
    if ($count -isnot [int]) { continue }  # Skip error entries

    if ($key -match '\\CLSID -> ') {
        $totalsByTarget["CLSIDs"] += $count
    } elseif ($key -match '\\Interface -> ') {
        $totalsByTarget["Interfaces"] += $count
    } elseif ($key -match '\\TypeLib -> ') {
        $totalsByTarget["TypeLibs"] += $count
    } elseif ($key -match " -> $([regex]::Escape($ProgIDPattern))$") {
        $totalsByTarget["ProgIDs"] += $count
    }
}

Log "ProgIDs:           $($totalsByTarget['ProgIDs'])"
Log "CLSIDs:            $($totalsByTarget['CLSIDs'])"
Log "Interfaces:        $($totalsByTarget['Interfaces'])"
Log "TypeLibs:          $($totalsByTarget['TypeLibs'])"
Log "Trusted Locations: $trustedLocationCount"

Log ""
Log "--- Per-location detail ---"
foreach ($entry in $summary.GetEnumerator()) {
    if ($entry.Value -is [int] -and $entry.Value -gt 0) {
        Log "  $($entry.Key): $($entry.Value)"
    } elseif ($entry.Value -isnot [int]) {
        Log "  $($entry.Key): $($entry.Value)"
    }
}

# Save log to file (atomic write at the end of the run)
$log | Out-File -FilePath $LogPath -Encoding UTF8
Write-Host "Log saved to: $LogPath"