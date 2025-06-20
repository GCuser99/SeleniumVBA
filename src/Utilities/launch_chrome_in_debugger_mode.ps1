# This script launches browser in debugger mode on port 9222
# Sets the required temporary profile directory location (configurable)
# Optionally kills any browser processes before launch

param (
    [string]$userDataDir = $env:TMP + "\chromeDebugProfile",
    [switch]$killBrowsersFirst
)

if ($killBrowsersFirst) {
    # Kill all Chrome browser processes and child processes
    Get-Process chrome -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            if (Get-Process -Id $_.Id -ErrorAction SilentlyContinue) {
                Stop-Process -Id $_.Id -Force
                Write-Output "Terminated chrome process ID: $($_.Id)"
            }
        } catch {
            Write-Warning "Failed to terminate process ID: $($_.Id). Error: $_"
        }
    }
}

# Get path to executable
$chromePaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"
)

foreach ($path in $chromePaths) {
    try {
        $chromeExe = (Get-ItemProperty -Path $path).'(default)'
        if (Test-Path $chromeExe) {
            $chromeExe
            break
        }
    } catch {
        # Ignore any paths that don't exist
    }
}

# Launch Chrome in Debugger mode
Start-Process -FilePath $chromeExe -ArgumentList "--remote-debugging-port=9222 --user-data-dir=`"$userDataDir`""