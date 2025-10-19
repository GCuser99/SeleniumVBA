# This script will scan the registry for all keys associated with SeleniumVBA install,
# create a log file, and optionally delete them (not recommended - uninstall via Inno Setup unins000.exe)

$ProgID = "SeleniumVBA"
$GuidPrefix = "38ED0FFA-E3F3-41C4-B601-"
$LogPath = ".\COM_Registry_Log.txt"
$Delete = $false

$ProgIDPattern = "$ProgID.*"

$log = New-Object System.Collections.Generic.List[string]
function Log($msg) {
    $timestamped = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    $log.Add($timestamped)
    Add-Content -Path $LogPath -Value $timestamped
    Write-Host $timestamped
}

function MatchKeys($basePath, $pattern) {
    Log "Scanning: $basePath with pattern: $pattern"
    try {
        $items = Get-ChildItem -Path "Registry::$basePath" -ErrorAction SilentlyContinue
        $matches = $items | Where-Object {
            $_.PSChildName -match "^$pattern" -or $_.PSChildName -match "^\{$pattern"
            } | ForEach-Object { $_.Name }
        Log "Matched $($matches.Count) keys"
        return $matches
    } catch {
        Log "Failed to scan: $basePath - $_"
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
"HKCR\WOW6432Node",
"HKCU\Software\Classes",
"HKCU\Software\Classes\Wow6432Node",
"HKCU\Software\Wow6432Node",
"HKLM\Software\Classes",
"HKLM\Software\Classes\WOW6432Node",
"HKLM\Software\WOW6432Node\Classes"
)

#"Software\Classes\TypeLib"

# Targets to match
$targets = @(
@{ Path = ""; Pattern = $ProgIDPattern },
@{ Path = "CLSID"; Pattern = "$GuidPrefix*" },
@{ Path = "Interface"; Pattern = "$GuidPrefix*" },
@{ Path = "TypeLib"; Pattern = "$GuidPrefix*" }
)

foreach ($hive in $hives) {
    foreach ($target in $targets) {
        $base = if ($target.Path -eq "") { $hive } else { "$hive\$($target.Path)" }
        $matches = MatchKeys $base $target.Pattern
        foreach ($match in $matches) {
            Log "Found: $match"
            if ($Delete) { DeleteKey $match }
        }
    }
}

# Detect Office version
function Get-OfficeVersion($app) {
    $roots = @(
        "HKCU:\Software\Microsoft\Office",
        "HKLM:\Software\Microsoft\Office",
        "HKLM:\Software\WOW6432Node\Microsoft\Office"  # 32‑bit Office on 64‑bit Windows
    )
    $latest = $null
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        $versions = Get-ChildItem -Path $root -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
            Sort-Object -Property {[version]$_.PSChildName} -Descending

        foreach ($version in $versions) {
            $strictPath = "$root\$($version.PSChildName)\$app\Security\Trusted Locations"
            $broadPath  = "$root\$($version.PSChildName)\$app"

            if (Test-Path $strictPath) {
		$latest = $version.PSChildName  # confirmed initialized
		Log "Found Office version: $latest" 
		Log "Office installed on: $root" 
                return $latest  
            }
            elseif (-not $latest -and (Test-Path $broadPath)) {
                $latest = $version.PSChildName  # fallback candidate
            }
        }
    }
    if ($latest) {
        Log "Found Office version: $version.PSChildName" 
        Log "Office installed on: $root"
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
            if ($Delete) {
                try {
                    Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                    Log "Deleted Trusted Location: $keyPath"
                } catch {
                    Log "Failed to delete $keyPath - $_"
                }
            }
        } else {
            Log "No Trusted Location found for $app {$version}: $keyPath"
        }

    } else {
        Log "No Trusted Location or Office version found for $app"
    }
}

# Find trusted locations for Excel and Access
FindTrustedLocation "Excel"
FindTrustedLocation "Access"

# Save log to file
$log | Out-File -FilePath $LogPath -Encoding UTF8
Log "Log saved to: $LogPath"