# This script launches browser in debugger mode on port 9222
# Sets the required temporary profile directory location (configurable)
# Optionally kills any browser processes before launch

param (
    [string]$userDataDir = $env:TMP + "\msedgeDebugProfile",
    [switch]$killBrowsersFirst
)

if ($killBrowsersFirst) {
    # Kill all Edge browser processes and child processes
    Get-Process msedge -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            if (Get-Process -Id $_.Id -ErrorAction SilentlyContinue) {
                Stop-Process -Id $_.Id -Force
                Write-Output "Terminated Edge process ID: $($_.Id)"
            }
        } catch {
            Write-Warning "Failed to terminate process ID: $($_.Id). Error: $_"
        }
    }
}

# Get path to executable
$msedgePaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"
)

foreach ($path in $msedgePaths) {
    try {
        $msedgeExe = (Get-ItemProperty -Path $path).'(default)'
        if (Test-Path $msedgeExe) {
            $msedgeExe
            break
        }
    } catch {
        # Ignore any paths that don't exist
    }
}

# Launch Edge in Debugger mode
Start-Process -FilePath $msedgeExe -ArgumentList "--remote-debugging-port=9222 --user-data-dir=`"$userDataDir`""