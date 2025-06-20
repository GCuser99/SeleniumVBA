# This script is used by Inno DLL Setup program to create an updated SeleniumVBA.ini file in the setup directory

param (
    [string]$iniFilePath = "..\SeleniumVBA.ini",
    [switch]$keepExistingValues
)

# Function to read INI values
Function Read-IniEntry {
    param (
        [string]$FilePath,
        [string]$Section,
        [string]$Key,
        [string]$DefaultValue,
        [bool]$keepExistingValues = $true
    )

    if (-Not (Test-Path $FilePath)) { return $DefaultValue }
    if (-Not $keepExistingValues) { return $DefaultValue }

    $iniContent = Get-Content $FilePath
    $sectionFound = $false

    foreach ($line in $iniContent) {
        $line = $line.Trim()  # Remove any leading/trailing whitespace

        # Detect section header
        if ($line -match "^\s*\[\s*$Section\s*\]\s*$") { 
            $sectionFound = $true
            continue
        }

        # Stop when reaching the next section
        if ($sectionFound -and $line -match "^\s*\[\s*$Section\s*\]\s*$") { break }

        # Extract key-value pair
        if ($sectionFound -and $line -match "^\s*$Key\s*=\s*(.+)$") { 
            return $matches[1].Trim()
        }
    }

    return $DefaultValue
}


$entries = @(
    "# This settings file is completely optional. For it to have effect,",
    "# it must be located in the same folder as the SeleniumVBA code",
    "# library, and be named SeleniumVBA.ini.",
    "",
    "# If a value for an entry is not specified, then the system",
    "# default value will be used.",
    "# Note that all path-type entry values recognize the %[Environ]% syntax.",
    "",
    "# A few useful Environ values for reference:",
    "", 
    "# %USERPROFILE%=C:\Users\[user name]",
    "# %APPDATA%=C:\Users\[user name]\AppData\Roaming",
    "# %LOCALAPPDATA%=C:\Users\[user name]\AppData\Local",
    "# %TEMP%=C:\Users\[user name]\AppData\Local\Temp",
    "",
    "[GENERAL]",
    "",
    "# The driver_location_folder system defaults to Downloads folder.",
    "# The default_io_folder system defaults to the active vba project's",
    "# document location - leave this blank to use default.",
    "# Valid values for command_window_style are vbHide (default),", 
    "# vbNormalFocus, vbMinimizedFocus, vbMaximizedFocus, vbNormalNoFocus,",
    "# and vbMinimizedNoFocus.",
    "# The system default values for implicit_wait, page_load_timeout, and",
    "# script_timeout are 0, 300000, and 30000 ms respectively.",
    "",
    "driver_location_folder=$((Read-IniEntry $iniFilePath 'GENERAL' 'driver_location_folder' '%USERPROFILE%\Downloads' $keepExistingValues))",
    "default_io_folder=$((Read-IniEntry $iniFilePath 'GENERAL' 'default_io_folder' '' $keepExistingValues))",
    "command_window_style=$((Read-IniEntry $iniFilePath 'GENERAL' 'command_window_style' 'vbHide' $keepExistingValues))",
    "",
    "implicit_wait=$((Read-IniEntry $iniFilePath 'GENERAL' 'implicit_wait' '0' $keepExistingValues))",
    "page_load_timeout=$((Read-IniEntry $iniFilePath 'GENERAL' 'page_load_timeout' '300000' $keepExistingValues))",
    "script_timeout=$((Read-IniEntry $iniFilePath 'GENERAL' 'script_timeout' '30000' $keepExistingValues))",
    "",
    "[AUTO-DRIVER-UPDATE]",
    "",
    "# If auto_detect_and_update=True (system default) then everytime",
    "# the WebDriver's Start* method is called, the Selenium driver's",
    "# version is checked against the corresponding browser version.",
    "# If the driver is not compatible with browser, it will be updated.",
    "# min_compatibility_level determines trigger for updating an",
    "# an out-of-date driver. System default is svbaBuildMajor.",
    "# Use svbaMinor for less frequent updating, and svbaExactMatch",
    "# for more frequent updating.",
    "",
    "auto_detect_and_update=$((Read-IniEntry $iniFilePath 'AUTO-DRIVER-UPDATE' 'auto_detect_and_update' 'True' $keepExistingValues))",
    "min_compatibility_level=$((Read-IniEntry $iniFilePath 'AUTO-DRIVER-UPDATE' 'min_compatibility_level' 'svbaBuildMajor' $keepExistingValues))",
    "",
    "# Below are browser-specific initializations.",
    "# To automatically initialize a set of capabilities each time the",
    "# OpenBrowser method of WebDriver class is invoked, set the",
    "# preload_capabilities_file_path entry to the path of a valid json",
    "# capabilities file. Note that if preload_capabilities_file_path is",
    "# set to a blank value, or the entry is missing or commented out,",
    "# then this option is ignored. Use the SaveToFile method of the",
    "# WebCapabilities class to save a default set of capabilities",
    "# for pre-loading.",
    "# The system defaults for local_host_port:",
    "# Chrome - 9515, Edge - 9516, Firefox - 4444",
    "",
    "[CHROME]",
    "",
    "preload_capabilities_file_path=$((Read-IniEntry $iniFilePath 'CHROME' 'preload_capabilities_file_path' '' $keepExistingValues))",
    "local_host_port=$((Read-IniEntry $iniFilePath 'CHROME' 'local_host_port' '9515' $keepExistingValues))",
    "",
    "[EDGE]",
    "",
    "preload_capabilities_file_path=$((Read-IniEntry $iniFilePath 'EDGE' 'preload_capabilities_file_path' '' $keepExistingValues))",
    "local_host_port=$((Read-IniEntry $iniFilePath 'EDGE' 'local_host_port' '9516' $keepExistingValues))",
    "",
    "[FIREFOX]",
    "",
    "preload_capabilities_file_path=$((Read-IniEntry $iniFilePath 'FIREFOX' 'preload_capabilities_file_path' '' $keepExistingValues))",
    "local_host_port=$((Read-IniEntry $iniFilePath 'FIREFOX' 'local_host_port' '4444' $keepExistingValues))"
    "",
    "[PDF_DEFAULT_PRINT_SETTINGS]",
    "",
    "# Valid units values are svbaInches (default) or svbaCentimeters.",
    "# Valid orientation values are svbaPortrait (default) or svbaLandscape.",
    "", 
    "# Common Metric print settings:",
    "# units=svbaCentimeters",
    "# page_height=27.94",
    "# page_width=21.59",
    "# margin_bottom=1",
    "# margin_top=1",
    "# margin_right=1",
    "# margin_left=1",
    "", 
    "# Common Imperial print settings:",
    "# units=svbaInches",
    "# page_height=11",
    "# page_width=8.5",
    "# margin_bottom=.393701",
    "# margin_top=.393701",
    "# margin_right=.393701",
    "# margin_left=.393701",
    "",
    "units=$((Read-IniEntry $iniFilePath 'PDF_DEFAULT_PRINT_SETTINGS' 'units' 'svbaInches' $keepExistingValues))",
    "page_height=$((Read-IniEntry $iniFilePath 'PDF_DEFAULT_PRINT_SETTINGS' 'page_height' '11' $keepExistingValues))",
    "page_width=$((Read-IniEntry $iniFilePath 'PDF_DEFAULT_PRINT_SETTINGS' 'page_width' '8.5' $keepExistingValues))",
    "margin_bottom=$((Read-IniEntry $iniFilePath 'PDF_DEFAULT_PRINT_SETTINGS' 'margin_bottom' '.393701' $keepExistingValues))",
    "margin_top=$((Read-IniEntry $iniFilePath 'PDF_DEFAULT_PRINT_SETTINGS' 'margin_top' '.393701' $keepExistingValues))",
    "margin_right=$((Read-IniEntry $iniFilePath 'PDF_DEFAULT_PRINT_SETTINGS' 'margin_right' '.393701' $keepExistingValues))",
    "margin_left=$((Read-IniEntry $iniFilePath 'PDF_DEFAULT_PRINT_SETTINGS' 'margin_left' '.393701' $keepExistingValues))",
    "background=$((Read-IniEntry $iniFilePath 'PDF_DEFAULT_PRINT_SETTINGS' 'background' 'False' $keepExistingValues))",
    "orientation=$((Read-IniEntry $iniFilePath 'PDF_DEFAULT_PRINT_SETTINGS' 'orientation' 'svbaPortrait' $keepExistingValues))",
    "print_scale=$((Read-IniEntry $iniFilePath 'PDF_DEFAULT_PRINT_SETTINGS' 'print_scale' '1.0' $keepExistingValues))",
    "shrink_to_fit=$((Read-IniEntry $iniFilePath 'PDF_DEFAULT_PRINT_SETTINGS' 'shrink_to_fit' 'True' $keepExistingValues))"
)

# Write the entries to the INI file
$entries | Set-Content -Path $iniFilePath
#$entries | Out-File -FilePath $iniFilePath -Encoding utf8
    