# Change Log

[v7.4, 24 Nov 2026]

- ActiveX DLL (twinBASIC Project)
    - Now references new tbRegExp and tbFileSysTools from tb's Package Server
    - Most problems with misreported error descriptions solved with latest tb releases
    - Significant refactor and code comments for Inno setup and analize_registry.ps1 utility scripts
    - All of the same test files from VBA ported to twinBASIC project for running in the IDE
    - Source files relocated to twinBASIC Project folder
- WebDriver Class
    - Renamed FileClosed procedure of WebDriver class to IsFileLocked to better reflect what the procedure does and 
        align with tb Project
    - Fixed SendKeys bug concerning sending long numeric-like strings [@alfreramos76, @hanamichi77777]
    - Fixed TableToArray edge-case bugs - thanks to [@hanamichi77777] for reporting and proposing changes.
    - Added TableToArray edge-case tests [@hanamichi77777]
    - Added AllowGeolocation method to turn off recent browser geolocation permission alerts (Edge/Chrome only)
    - WaitForIdleJQuery and WaitForIdleDOM refactored for robustness
    - Added InstallNetworkAndDOMHooks and WaitForIdleNetworkAndDOM method pairs for improved Network/DOM activity
        idle state tracking - use instead of legacy WaitForIdleNetwork which is now marked as DEPRECATED [@hanamichi77777]
    - Added tests for WaitForIdleJQuery, WaitForIdleDOM, and InstallNetworkAndDOMHooks/WaitForIdleNetworkAndDOM
    - Removed unused code from class
- WebShared Module
    - Refactored downloadToFile procedure - proxy setup moved to applySystemProxy, which frees WinHTTP memory on every path.
- WebDriverManger Class
    - Dramatically sped up FireFox Check-Driver-Browser-Alignment process (now almost as fast as Edge/Chrome)
    - Improved robustness for the Chrome/Edge Check-Driver-Browser-Alignment process
    - Improved speed and robustness of DownloadAndInstallDriver by switching to Tar.exe zip extraction method with 
        Shell32.CopyHere as a fallback [@hanamichi77777]
    - Added optional browserPath argument to AlignDriverWithBrowser to support portable browsers
- WebCapabilities Class
    - Marked SetGeolocationAware as DEPRECATED - use AllowGeolocation of WebDriver class in
    - EnableBiDiMode no longer sets OptionsKeepStringifiedNumbersAsStrings of WebJsonConverter class - that option
        is now enabled by default
