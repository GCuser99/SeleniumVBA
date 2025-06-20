# This script terminates any "stranded" WebDriver and associated child processes

$webDrivers = @("chromedriver", "geckodriver", "msedgedriver")
$killCount = 0

foreach ($driver in $webDrivers) {
    $processes = Get-Process -Name $driver -ErrorAction SilentlyContinue
    foreach ($proc in $processes) {
        try {
            Stop-Process -Id $proc.Id -Force
            $killCount++
        } catch {
            Write-Warning "Could not stop $($proc.Name) with ID $($proc.Id): $_"
        }
    }
}

Write-Output "Total WebDriver processes terminated: $killCount"