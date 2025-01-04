VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used to maintain Selenium driver/browser version alignment"
'@ModuleDescription "This class is used to maintain Selenium driver/browser version alignment"
'@Exposed
'@folder("SeleniumVBA.Source")
' ==========================================================================
' SeleniumVBA v6.3
'
' A Selenium wrapper for browser automation developed for MS Office VBA
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' Contact Info:
'
' https://github.com/6DiegoDiego9
' https://github.com/GCUser99
' ==========================================================================
' MIT License
'
' Copyright (c) 2023-2024, GCUser99 and 6DiegoDiego9 (https://github.com/GCuser99/SeleniumVBA)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ==========================================================================
'
' DESCRIPTION:
'
'    This class was written for use in Selenium VBA projects to manage webdriver-browser
'    version compatibility.
'
' USAGE:
'
'    Dim mngr As New WebDriverManager
'
'    Debug.Print mngr.AlignEdgeDriverWithBrowser()
'    Debug.Print mngr.AlignChromeDriverWithBrowser()
'    Debug.Print mngr.AlignFirefoxDriverWithBrowser()
'
'    Alternatively, the update process can be controlled with more granularity, as in the following perturbation...
'
'    Dim mngr As New WebDriverManager
'    driverPath = ".\chromedriver.exe"
'    bverInstalled = mngr.GetInstalledBrowserVersion(svbaBrowser.Chrome)
'    dverInstalled = mngr.GetInstalledDriverVersion(svbaBrowser.Chrome, driverPath)
'    If mngr.CheckCompatibilityLevel(bverInstalled, dverInstalled) < svbaExactMatch Then
'        resp = MsgBox("There is a more compatible driver available - would you like to install?", vbYesNo)
'        If resp = vbYes Then
'            dverCompat = mngr.GetCompatibleDriverVersion(svbaBrowser.Chrome, bverInstalled)
'            mngr.DownloadAndInstallDriver svbaBrowser.Chrome, dverCompat, driverPath
'        End If
'    End If
'
' NOTES:
'
'   If the driver has not yet been installed, the Align* methods will detect that and install
'   The Align* methods have a runSilent input parameter that if set to True, will not throw any of the expected errors -
'   this is used with the auto-check/install of the WebDriver Start* methods to make this feature as innocuous/stealth as possible.
'   Relative driver paths are supported (e.g. ".\msedgedriver.exe").
'   Use the DefaultDriverFolder property or the SeleniumVBA.ini file to specify a driver location other than the Downloads dir
'   A different level of minimum version compatibility can be specified [svbaMajor, svbaMinor, svbaBuildMajor, svbaExactMatch].
'   The default level is set for the Align* methods to svbaExactMatch, meaning that the most recent compatible release
'   will be installed if it is not already. However, as an example, the setting for the auto-check in the Start method of
'   the WebDriver class is to a level of svbaBuildMajor, meaning that only if there is a newer compatible driver with a different
'   build major version will the install occur, with the intent to reduce the number of requests to the driver servers.
'
' REFERENCES:
'
'   Windows Script Host Model
'   Microsoft Shell Controls and Automation
'   Microsoft XML, V6.0
'   Microsoft WMI Scripting V1.2 Library
'
Option Explicit

Private mDefaultDriverFolder As String

Private Const verBrowserNotFound As String = "-99.0.000.00"
Private Const verDriverNotFound As String = "000.0.000.00"

Public Enum svbaCompatibility
    svbaNotCompatible = 0
    svbaMajor = 1
    svbaMinor = 2
    svbaBuildMajor = 3
    svbaExactMatch = 4
End Enum

Public Enum svbaBrowser
    Chrome = 1
    Edge = 2
    Firefox = 3
End Enum

Public Enum svbaPlatform
    svbaSystem = 0
    svbaWin32 = 1
    svbaWin64 = 2
End Enum

Public Enum svbaChannel
    svbaStable = 0
    svbaBeta = 1
    svbaDev = 2
    svbaCanary = 3
End Enum

Private Sub Class_Initialize()
    Dim oShell As New Shell32.Shell
    Dim iniFilePath As String
    
    'attempt to read global variable values from optional settings file - if not successful,
    'then assign a system default value (last input parameter of ReadIniFileEntry)
    iniFilePath = getFullLocalPath(thisLibFolderPath) & "\SeleniumVBA.ini"
    Me.DefaultDriverFolder = readIniFileEntry(iniFilePath, "GENERAL", "driver_location_folder", oShell.Namespace("shell:Downloads").Self.Path)
End Sub

'@Description("Sets the default driver folder path")
Public Property Let DefaultDriverFolder(ByVal folderPath As String)
Attribute DefaultDriverFolder.VB_Description = "Sets the default driver folder path"
    mDefaultDriverFolder = WebShared.getFullLocalPath(folderPath)
End Property

Public Property Get DefaultDriverFolder() As String
    DefaultDriverFolder = mDefaultDriverFolder
End Property

Friend Function AlignDriverWithBrowser(ByVal browser As svbaBrowser, ByVal driverPath As String, Optional ByVal minCompatibilityLevel As svbaCompatibility = svbaExactMatch, Optional ByVal runSilent As Boolean = False, Optional ByVal platform As svbaPlatform = svbaSystem) As String
    Dim bverInstalled As String
    Dim dverCompat As String
    Dim respStr As String
    Dim dverInstalled As String
    
    'get the installed versions of driver and browser
    bverInstalled = GetInstalledBrowserVersion(browser)
    
    'user might be using a portable browser that will be set in capabilities
    'so if browser is not found in registry, then set browser version to a
    'dummy value and then install latest driver, if not already installed
    If Left$(bverInstalled, 5) = "Error" Then bverInstalled = verBrowserNotFound
    
    dverInstalled = GetInstalledDriverVersion(browser, driverPath)
    
    Select Case browser
    Case svbaBrowser.Edge, svbaBrowser.Chrome
        
        'because of the way Edge and Chrome maintain synchrony between driver and browser version numbers
        'we can compare installed driver and installed browser version strings to determine compatibility
        'without having to make a url request - only if they are incompatible do we need to make the request
        If CheckCompatibilityLevel(bverInstalled, dverInstalled) < minCompatibilityLevel Then
            'the installed driver version is sufficiently different than the installed browser version
            'so find the most compatible driver version for installed browser
            If bverInstalled = verBrowserNotFound Then
                dverCompat = GetLatestDriverVersion(browser)
            Else
                dverCompat = GetCompatibleDriverVersion(browser, bverInstalled)
            End If
        
            'check if we have an error and then take appropriate action
            If dverCompat Like "Error*" Then
                If runSilent Then 'don't throw the error - just report back to caller
                    AlignDriverWithBrowser = dverCompat
                    Exit Function
                Else
                    Err.Raise 1, "WebDriverManager", dverCompat
                End If
            End If
            
            'It's possible that the browser version is a dev version, and there is not
            'an updated driver version. It's also possible that the browser was not found and
            'that dverCompat is latest version. Check to make sure dverCompat and dverInstalled
            'are sufficiently different, then if so, install the "compatible" driver
            If CheckCompatibilityLevel(dverCompat, dverInstalled) < minCompatibilityLevel Then
                DownloadAndInstallDriver browser, dverCompat, driverPath, platform
                dverInstalled = dverCompat
                respStr = "The latest " & getBrowserString(browser) & " Webdriver was installed."
            Else
                respStr = getBrowserString(browser) & " Webdriver and Browser are compatible."
            End If
        Else
            respStr = getBrowserString(browser) & " Webdriver and Browser are compatible."
        End If
    Case svbaBrowser.Firefox
        'get the compatible driver version for the installed browser
        If bverInstalled = verBrowserNotFound Then
            dverCompat = GetLatestDriverVersion(browser)
        Else
            dverCompat = GetCompatibleDriverVersion(browser, bverInstalled)
        End If
                 
        'check if we have an error and then take appropriate action
        If dverCompat Like "Error*" Then
            If runSilent Then 'don't throw the error - just report back to caller
                AlignDriverWithBrowser = dverCompat
                Exit Function
            Else
                Err.Raise 1, "WebDriverManager", dverCompat
            End If
        End If
        
        If CheckCompatibilityLevel(dverCompat, dverInstalled) < minCompatibilityLevel Then
            'the installed driver version is sufficiently different than the installed browser version
            'so install it
            DownloadAndInstallDriver browser, dverCompat, driverPath, platform
            dverInstalled = dverCompat
            respStr = "The latest " & getBrowserString(browser) & " Webdriver was installed."
        Else
            respStr = getBrowserString(browser) & " Webdriver and Browser are compatible."
        End If
    End Select
    
    If bverInstalled = verBrowserNotFound Then bverInstalled = "Not found"
    
    respStr = respStr & vbNewLine & vbNewLine & "Browser version: " & bverInstalled & vbNewLine & "Driver version:    " & dverInstalled
    
    AlignDriverWithBrowser = respStr
End Function

'@Description("Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed")
Public Function AlignEdgeDriverWithBrowser(Optional ByVal driverPath As String = "msedgedriver.exe", Optional ByVal minCompatibilityLevel As svbaCompatibility = svbaExactMatch, Optional ByVal runSilent As Boolean = False, Optional ByVal platform As svbaPlatform = svbaSystem) As String
Attribute AlignEdgeDriverWithBrowser.VB_Description = "Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed"
    AlignEdgeDriverWithBrowser = AlignDriverWithBrowser(svbaBrowser.Edge, driverPath, minCompatibilityLevel, runSilent, platform)
End Function

'@Description("Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed")
Public Function AlignChromeDriverWithBrowser(Optional ByVal driverPath As String = "chromedriver.exe", Optional ByVal minCompatibilityLevel As svbaCompatibility = svbaExactMatch, Optional ByVal runSilent As Boolean = False, Optional ByVal platform As svbaPlatform = svbaSystem) As String
Attribute AlignChromeDriverWithBrowser.VB_Description = "Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed"
    AlignChromeDriverWithBrowser = AlignDriverWithBrowser(svbaBrowser.Chrome, driverPath, minCompatibilityLevel, runSilent, platform)
End Function

'@Description("Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed")
Public Function AlignFirefoxDriverWithBrowser(Optional ByVal driverPath As String = "geckodriver.exe", Optional ByVal minCompatibilityLevel As svbaCompatibility = svbaExactMatch, Optional ByVal runSilent As Boolean = False, Optional ByVal platform As svbaPlatform = svbaSystem) As String
Attribute AlignFirefoxDriverWithBrowser.VB_Description = "Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed"
    AlignFirefoxDriverWithBrowser = AlignDriverWithBrowser(svbaBrowser.Firefox, driverPath, minCompatibilityLevel, runSilent, platform)
End Function

'@Description("Returns the installed Driver version")
Public Function GetInstalledDriverVersion(ByVal browser As svbaBrowser, Optional ByVal driverPath As Variant) As String
Attribute GetInstalledDriverVersion.VB_Description = "Returns the installed Driver version"
    Dim wsh As IWshRuntimeLibrary.WshShell 'CreateObject("WScript.Shell")
    Dim fso As FileSystemObject
    Dim resp As Variant
    Dim i As Long
    Dim ver As String

    If IsMissing(driverPath) Then
        Select Case browser
        Case svbaBrowser.Chrome
            driverPath = ".\chromedriver.exe"
        Case svbaBrowser.Edge
            driverPath = ".\msedgedriver.exe"
        Case svbaBrowser.Firefox
            driverPath = ".\geckodriver.exe"
        End Select
    End If

    driverPath = WebShared.getFullLocalPath(driverPath, mDefaultDriverFolder)
    
    Set fso = New FileSystemObject
    If Not fso.FileExists(driverPath) Then
        GetInstalledDriverVersion = verDriverNotFound
        Exit Function
    End If
    
    If browser = svbaBrowser.Edge Then
        'see if it can be found from embedded extended file properties first - much faster and less disruptive
        resp = findVersionInString(fso.GetFileVersion(driverPath), browser)
        If resp <> vbNullString Then
            GetInstalledDriverVersion = resp
            Exit Function
        End If
    End If
    
    Set wsh = New IWshRuntimeLibrary.WshShell
    resp = wsh.Exec(Chr$(34) & driverPath & Chr$(34) & " --version").StdOut.ReadAll
    GetInstalledDriverVersion = findVersionInString(resp, browser)
End Function

'@Description("Returns the compatibility level between a pair of version strings")
Public Function CheckCompatibilityLevel(ByVal versionString1 As String, ByVal versionString2 As String) As svbaCompatibility
Attribute CheckCompatibilityLevel.VB_Description = "Returns the compatibility level between a pair of version strings"
    Dim vermajor1 As Variant, verminor1 As Variant, verbuildmajor1 As Variant, verbuildminor1 As Variant
    Dim vermajor2 As Variant, verminor2 As Variant, verbuildmajor2 As Variant, verbuildminor2 As Variant
    
    If versionString1 = "driver not found" Then CheckCompatibilityLevel = -1: Exit Function
    vermajor1 = selectVersionPart(versionString1, 1)
    verminor1 = selectVersionPart(versionString1, 2)
    verbuildmajor1 = selectVersionPart(versionString1, 3)
    If countVersionParts(versionString1) > 3 Then verbuildminor1 = selectVersionPart(versionString1, 4) Else verbuildminor1 = "0"
    
    If versionString2 = "driver not found" Then CheckCompatibilityLevel = -1: Exit Function
    vermajor2 = selectVersionPart(versionString2, 1)
    verminor2 = selectVersionPart(versionString2, 2)
    verbuildmajor2 = selectVersionPart(versionString2, 3)
    If countVersionParts(versionString2) > 3 Then verbuildminor2 = selectVersionPart(versionString2, 4) Else verbuildminor2 = "0"
    
    If vermajor1 <> vermajor2 Then
        CheckCompatibilityLevel = svbaNotCompatible
        Exit Function
    End If
    
    If verminor1 <> verminor2 Then
        CheckCompatibilityLevel = svbaMajor
        Exit Function
    End If
    
    If verbuildmajor1 <> verbuildmajor2 Then
        CheckCompatibilityLevel = svbaMinor
        Exit Function
    End If
    
    If verbuildminor1 <> verbuildminor2 Then
        CheckCompatibilityLevel = svbaBuildMajor
        Exit Function
    End If
    
    CheckCompatibilityLevel = svbaExactMatch
End Function

'@Description("Returns the latest Driver version")
Public Function GetLatestDriverVersion(ByVal browser As svbaBrowser, Optional ByVal channel As svbaChannel = svbaStable) As String
Attribute GetLatestDriverVersion.VB_Description = "Returns the latest Driver version"
    Dim Url As String
    Dim respTxt As Variant
    Dim json As Dictionary
    
    Select Case browser
    Case svbaBrowser.Edge
        Url = "https://msedgedriver.azureedge.net/LATEST_" & UCase$(getChannelString(channel)) 'STABLE, BETA, DEV, CANARY
        respTxt = requestData(Url)
        If respTxt = "Release Not Found" Then GetLatestDriverVersion = "Error: Release Not Found": Exit Function
        GetLatestDriverVersion = Replace$(respTxt, vbNewLine, vbNullString)
    Case svbaBrowser.Chrome
        Url = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_" & UCase$(getChannelString(channel)) 'STABLE, BETA, DEV, CANARY
        respTxt = requestData(Url)
        If respTxt = "Release Not Found" Then GetLatestDriverVersion = "Error: Release Not Found": Exit Function
        GetLatestDriverVersion = respTxt
    Case svbaBrowser.Firefox
        Url = "https://api.github.com/repos/mozilla/geckodriver/releases/latest"
        respTxt = requestData(Url)
        If respTxt = "Release Not Found" Then GetLatestDriverVersion = "Error: Release Not Found": Exit Function
        GetLatestDriverVersion = findLatestFirefoxVersion(respTxt)
    End Select
End Function

'@Description("Returns the most recent compatible Driver version given the installed Browser version")
Public Function GetCompatibleDriverVersion(ByVal browser As svbaBrowser, ByVal browserVersion As String) As String
Attribute GetCompatibleDriverVersion.VB_Description = "Returns the most recent compatible Driver version given the installed Browser version"
    Dim htmlDoc As New HTMLDocument
    Dim table As HTMLTable
    Dim row As HTMLTableRow
    Dim verMajor As String
    Dim browserBuild  As String
    Dim respTxt As Variant
    Dim Url As String
    
    'This function, if called from WebDriver class as part of auto-updater, needs to execution as silently
    'as possible - we don't want the non-mission-critical error in driver check/updating to get in the way of user
    'running WebDriver. For example, user may be running on local urls, and so in no internet connection, user can still work
    
    'for chrome info:
    'https://chromedriver.chromium.org/downloads/version-selection
    'https://groups.google.com/g/chromedriver-users/c/clpipqvOGjE
    'https://github.com/GoogleChromeLabs/chrome-for-testing#json-api-endpoints
    
    'for edge info:
    'https://blogs.windows.com/msedgedev/2019/11/04/edge-chromium-release-candidate-get-ready/
    'https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
    
    Select Case browser
    Case svbaBrowser.Edge
        
        verMajor = selectVersionPart(browserVersion, 1)

        'search for major version, decrementing by one until success
        Do While verMajor >= 75
            Url = "https://msedgedriver.azureedge.net/LATEST_RELEASE_" & verMajor & "_WINDOWS"
            
            respTxt = requestData(Url)
            
            If respTxt Like "Error*" Then GetCompatibleDriverVersion = respTxt: Exit Function
            
            If respTxt <> "Release Not Found" Then Exit Do
            
            verMajor = verMajor - 1
        Loop
        
        GetCompatibleDriverVersion = Replace$(respTxt, vbNewLine, vbNullString)
        Exit Function
    Case svbaBrowser.Chrome
        browserBuild = filterVersionParts(browserVersion, 3)
        verMajor = selectVersionPart(browserVersion, 1)
        
        If compareVersions(browserBuild, "115.0.5763") >= 0 Then
            'use new api...
            'as of 8/22/2023, this special end-point was added that
            'shortcuts using latest-patch-versions-per-build end-point!!
            Url = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_" & browserBuild
            
            respTxt = requestData(Url)
            
            If respTxt Like "Error*" Then GetCompatibleDriverVersion = respTxt: Exit Function
            
            If respTxt = "Release Not Found" Then
                'in theory, we should never get here, but just in case...
                'if an exact match to the browser build was not found above, then
                'search for the maximum version that is less than the specified browser build
                Url = "https://googlechromelabs.github.io/chrome-for-testing/latest-patch-versions-per-build.json"
                respTxt = requestData(Url)
                
                If respTxt Like "Error*" Then GetCompatibleDriverVersion = respTxt: Exit Function
                
                GetCompatibleDriverVersion = findLatestChromeVersionFromBuild(respTxt, browserBuild)
                Exit Function
            End If
            
            GetCompatibleDriverVersion = respTxt
            Exit Function
        ElseIf verMajor = 115 Then
            'we are in the "gap" between the two api's
            'builds 115.0.5739 to 115.0.5762 do not have a corresponding ChromeDriver
            'so get the last known major version from 114...
            GetCompatibleDriverVersion = "114.0.5735.90"
            Exit Function
        Else
            'use old api
            'first see if we can find an exact build match
            Url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & browserBuild
            
            respTxt = requestData(Url)
            
            If respTxt Like "Error*" Then GetCompatibleDriverVersion = respTxt: Exit Function
            
            If respTxt = "Release Not Found" Then
                'if exact build match failed, then maybe we have a Dev or Canary channel
                'so search for major version match, decrementing by one until success
                Do While verMajor >= 70
                    Url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & verMajor
                    respTxt = requestData(Url)
                    
                    If respTxt Like "Error*" Then GetCompatibleDriverVersion = respTxt: Exit Function
                    
                    If respTxt <> "Release Not Found" Then Exit Do
                    verMajor = verMajor - 1
                Loop
            End If
            GetCompatibleDriverVersion = respTxt
            Exit Function
        End If
    Case svbaBrowser.Firefox
        'Firefox browser version has this pattern: 117.0 (x64 en-US)
        verMajor = selectVersionPart(browserVersion, 1)
        
        'the rules are different for Firefox driver - we must lookup the browser version in a compatibility matrix
        'to get the most recent driver version compatible with the installed browser
        Url = "https://firefox-source-docs.mozilla.org/testing/geckodriver/Support.html"
        
        respTxt = requestData(Url)
        
        If respTxt = "Release Not Found" Then respTxt = "Error - Release Not Found"
        
        If respTxt Like "Error*" Then GetCompatibleDriverVersion = respTxt: Exit Function
        
        'search the table for the first driver version entry (most recent) for a given browser version
        htmlDoc.body.innerHTML = respTxt
        Set table = htmlDoc.getElementsByTagName("table")(0)
        
        table.deleteTHead
        
        For Each row In table.rows
            If CInt(verMajor) >= CInt(Split(row.cells(2).innerText)(0)) Then
                GetCompatibleDriverVersion = Trim$(row.cells(0).innerText)
                Exit Function
            End If
        Next row
        
        GetCompatibleDriverVersion = "Error: Release Not Found"
    End Select
End Function

'@Description("Returns the installed Browser version")
Public Function GetInstalledBrowserVersion(ByVal browser As svbaBrowser) As String
Attribute GetInstalledBrowserVersion.VB_Description = "Returns the installed Browser version"
    Dim wsh As New IWshRuntimeLibrary.WshShell 'CreateObject("WScript.Shell")
    Dim regkey As String
    Dim regKeyAlt As String
    Dim resp As String
    
    'Notes for FF - when first installed using the FF installer only the HKLM key exists - but a subsequent update from the browser creates the HKCU version as well
    'HKEY_CURRENT_USER\SOFTWARE\Mozilla\Mozilla Firefox\CurrentVersion
    'also, if running in 32-bit tB IDE the HKLM version redirects to potentially non-existent:
    'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Mozilla\Mozilla Firefox\CurrentVersion
    'but the HKCU version works

    Select Case browser
    Case svbaBrowser.Chrome
        regkey = "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon\version"
        regKeyAlt = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Google\Update\Clients\{8A69D345-D564-463c-AFF1-A69D9E530F96}\pv"
    Case svbaBrowser.Edge
        regkey = "HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon\version"
        regKeyAlt = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}\pv"
    Case svbaBrowser.Firefox
        regkey = "HKEY_LOCAL_MACHINE\SOFTWARE\Mozilla\Mozilla Firefox\CurrentVersion"
        regKeyAlt = "HKEY_CURRENT_USER\SOFTWARE\Mozilla\Mozilla Firefox\CurrentVersion"
    End Select

    'this throws error if key doesnt exist - report back to caller if browser isn't installed
    On Error Resume Next
    resp = wsh.RegRead(regkey)
    If Err.Number <> 0 Then
        'try again with the alternate key
        Err.Clear
        resp = wsh.RegRead(regKeyAlt)
        If Err.Number <> 0 Then resp = "Error: " & getBrowserString(browser) & " Browser Not Installed"
    End If

    GetInstalledBrowserVersion = resp
End Function

'@Description("Downloads and installs the specified Driver version")
Public Sub DownloadAndInstallDriver(ByVal browser As svbaBrowser, ByVal driverVersion As String, Optional ByVal driverPath As Variant, Optional ByVal platform As svbaPlatform = svbaSystem)
Attribute DownloadAndInstallDriver.VB_Description = "Downloads and installs the specified Driver version"
    Dim Url As String
    Dim fso As New IWshRuntimeLibrary.FileSystemObject
    Dim shll As New Shell32.Shell
    Dim driverFolder As String
    Dim driverFileName As String
    Dim tmpInstallFolder As String
    Dim createdTmpFolder As Boolean
    Dim tmpInstallZipPath As String
    Dim tmpZipDriverFilename As String
    Dim tmpInstallDriverpath As String
    Dim driverBuild As String
    Dim respTxt As String
    Dim platformString As String
    
    If Left$(driverVersion, 5) = "Error" Then Err.Raise 1, , "Error in DownloadAndInstallDriver: driverVersion not valid"
    
    If IsMissing(driverPath) Then
        Select Case browser
        Case svbaBrowser.Chrome
            driverPath = ".\chromedriver.exe"
        Case svbaBrowser.Edge
            driverPath = ".\msedgedriver.exe"
        Case svbaBrowser.Firefox
            driverPath = ".\geckodriver.exe"
        End Select
    End If
    
    platformString = getPlatformString(platform)
    
    driverPath = WebShared.getFullLocalPath(driverPath, mDefaultDriverFolder)
    
    driverFolder = fso.GetParentFolderName(driverPath)
    driverFileName = fso.GetFileName(driverPath)
    
    tmpInstallFolder = driverFolder & "\tmp\"
    
    Select Case browser
    Case svbaBrowser.Chrome
        driverBuild = filterVersionParts(driverVersion, 3)
        If compareVersions(driverBuild, "115.0.5763") >= 0 Then
            'use the new download api
            Url = "https://storage.googleapis.com/chrome-for-testing-public/" & driverVersion & "/" & platformString & "/chromedriver-" & platformString & ".zip"
            tmpInstallZipPath = tmpInstallFolder & "chromedriver-" & platformString & ".zip"
            tmpZipDriverFilename = "chromedriver-" & platformString & "\chromedriver.exe"
            tmpInstallDriverpath = tmpInstallFolder & "chromedriver.exe"
        Else
            'use the old download api
            'no 64-bit version available for drivers before 115.0.5763
            Url = "https://chromedriver.storage.googleapis.com/" & driverVersion & "/chromedriver_win32.zip"
            tmpInstallZipPath = tmpInstallFolder & "chromedriver_win32.zip"
            tmpZipDriverFilename = "chromedriver.exe"
            tmpInstallDriverpath = tmpInstallFolder & tmpZipDriverFilename
        End If
    Case svbaBrowser.Edge
        Url = "https://msedgedriver.azureedge.net/" & driverVersion & "/edgedriver_" & platformString & ".zip"
        tmpInstallZipPath = tmpInstallFolder & "edgedriver_" & platformString & ".zip"
        tmpZipDriverFilename = "msedgedriver.exe"
        tmpInstallDriverpath = tmpInstallFolder & tmpZipDriverFilename
    Case svbaBrowser.Firefox
        Url = "https://github.com/mozilla/geckodriver/releases/download/v" & driverVersion & "/geckodriver-v" & driverVersion & "-" & platformString & ".zip"
        tmpInstallZipPath = tmpInstallFolder & "geckodriver-v" & driverVersion & "-" & platformString & ".zip"
        tmpZipDriverFilename = "geckodriver.exe"
        tmpInstallDriverpath = tmpInstallFolder & tmpZipDriverFilename
    End Select
    
    'if tmp folder doesn't exist then create it - remember for clean up
    createdTmpFolder = False
    If Not fso.FolderExists(tmpInstallFolder) Then
        createdTmpFolder = True
        fso.CreateFolder tmpInstallFolder
    End If
    
    UrlDownloadToFile 0&, StrPtr(Url), StrPtr(tmpInstallZipPath), 0&, 0&
    
    'make sure driver zip file was created successfully
    If Not fso.FileExists(tmpInstallZipPath) Then
        'clean up first
        If createdTmpFolder Then fso.DeleteFolder Left$(tmpInstallFolder, Len(tmpInstallFolder) - 1), True
        Err.Raise 1, "WebDriverManager", "Error: Driver Download Not Successful"
    End If
    
    If fso.FileExists(tmpInstallDriverpath) Then fso.DeleteFile tmpInstallDriverpath, True

    shll.Namespace(tmpInstallFolder).CopyHere shll.Namespace(tmpInstallZipPath).Items.Item(tmpZipDriverFilename)

    'check if driver exists and if it does, make sure it isn't in use before overwrite
    If fso.FileExists(driverPath) Then
        If isProcessRunning(driverFileName) Then taskKillByImage (driverFileName): WebShared.SleepWinAPI 100
    End If

    'copy file from temp folder to driver path
    fso.CopyFile tmpInstallDriverpath, driverPath, True
    
    'if created tmp folder, then delete it
    If createdTmpFolder Then fso.DeleteFolder Left$(tmpInstallFolder, Len(tmpInstallFolder) - 1), True
End Sub

'@Description("Returns the folder path to Florent Breheret's SeleniumBasic application in local AppData, assuming that it is installed")
Public Function GetSeleniumBasicFolderPath() As String
Attribute GetSeleniumBasicFolderPath.VB_Description = "Returns the folder path to Florent Breheret's SeleniumBasic application in local AppData, assuming that it is installed"
    'this may be of use for updating drivers for SeleniumBasic
    Dim fso As New FileSystemObject
    GetSeleniumBasicFolderPath = Environ$("LocalAppData") & "\SeleniumBasic\"
    If Not fso.FolderExists(GetSeleniumBasicFolderPath) Then Err.Raise 1, , "Error in GetSeleniumBasicFolderPath: SeleniumBasic not installed in the expected location"
End Function

'@Description("Returns the folder path to the Temp folder in local AppData")
Public Function GetUserTempFolderPath() As String
Attribute GetUserTempFolderPath.VB_Description = "Returns the folder path to the Temp folder in local AppData"
    GetUserTempFolderPath = Environ$("LocalAppData") & "\Temp\"
End Function

'@Description("Returns the folder path to the local Downloads folder")
Public Function GetDownloadsFolderPath() As String
Attribute GetDownloadsFolderPath.VB_Description = "Returns the folder path to the local Downloads folder"
    Dim oShell As New Shell32.Shell
    GetDownloadsFolderPath = oShell.Namespace("shell:Downloads").Self.Path
End Function

'@Description("Deletes the Edge Driver from the specified path")
Public Sub DeleteEdgeDriver(Optional ByVal driverPath As String = "msedgedriver.exe")
Attribute DeleteEdgeDriver.VB_Description = "Deletes the Edge Driver from the specified path"
    deleteDriver driverPath
End Sub

'@Description("Deletes the Chrome Driver from the specified path")
Public Sub DeleteChromeDriver(Optional ByVal driverPath As String = "chromedriver.exe")
Attribute DeleteChromeDriver.VB_Description = "Deletes the Chrome Driver from the specified path"
    deleteDriver driverPath
End Sub

'@Description("Deletes the Firefox Driver from the specified path")
Public Sub DeleteFirefoxDriver(Optional ByVal driverPath As String = "geckodriver.exe")
Attribute DeleteFirefoxDriver.VB_Description = "Deletes the Firefox Driver from the specified path"
    deleteDriver driverPath
End Sub

'---------------------- Private Helper Functions ----------------------

Private Sub deleteDriver(ByVal driverPath As String)
    Dim fso As New FileSystemObject
    Dim driverFileName As String
    'check if driver exists and if it does, make sure it isn't in use before deleting
    driverPath = WebShared.getFullLocalPath(driverPath, mDefaultDriverFolder)
    driverFileName = fso.GetFileName(driverPath)
    If fso.FileExists(driverPath) Then
        If isProcessRunning(driverFileName) Then taskKillByImage (driverFileName)
        WebShared.SleepWinAPI 100
        fso.DeleteFile driverPath, True
    End If
End Sub

Private Function isProcessRunning(ByVal process As String) As Boolean
    Dim objServices As WbemScripting.SWbemServices
    Dim objProcessList As WbemScripting.SWbemObjectSet
    
    Set objServices = GetObject("winmgmts:")
    Set objProcessList = objServices.ExecQuery("select * from win32_process where name='" & process & "'")
    
    isProcessRunning = (objProcessList.Count > 0)
End Function

Private Function requestData(ByVal Url As String, Optional maxTries As Long = 20) As Variant
    'see https://github.com/GCuser99/SeleniumVBA/discussions/29#discussioncomment-3673682
    Dim client As New MSXML2.XMLHTTP60
    Dim try As Long
        
    'try until a connection to server is established, up to a maximum # of tries
    For try = 1 To maxTries
    
        On Error Resume Next
        client.Open "GET", Url, False
        
        client.send
        
        While client.readyState <> 4
            DoEvents
        Wend
        
        On Error GoTo 0
        
        Select Case client.Status
        Case 200
            requestData = client.ResponseText
            Exit Function
        Case 404
            requestData = "Release Not Found"
            Exit Function
        Case 12007
            requestData = "Error: No Internet Connection"
            Exit Function
        Case 12029, 12031
            'Don't exit loop - make another try to see if we can connect
            'RequestData = "Error: Cannot Connect to Server"
        Case Else 'an unexpected error
            Debug.Print "Please report this error to developers: RequestData Status= " & client.Status & ", " & "Response= " & client.ResponseText
            requestData = "Error: Status= " & client.Status & ", " & "Response= " & client.ResponseText
            Exit Function
        End Select
        
        Debug.Print "try: ", try, client.Status, client.ResponseText, Url
    Next try
    'if here then had too many 12029/12031's
    requestData = "Error: Cannot Connect to Server"
End Function

Private Function getBrowserString(ByVal val As svbaBrowser) As String
    Select Case val
    Case svbaBrowser.Edge
        getBrowserString = "Edge"
    Case svbaBrowser.Chrome
        getBrowserString = "Chrome"
    Case svbaBrowser.Firefox
        getBrowserString = "Firefox"
    End Select
End Function

Private Function taskKillByImage(ByVal taskName As String) As Variant
    Dim wsh As New IWshRuntimeLibrary.WshShell
    taskKillByImage = wsh.Run("taskkill /f /t /im " & taskName, 0, True)
End Function

Private Function getPlatformString(ByVal platform As svbaPlatform) As String
    Select Case platform
    Case svbaPlatform.svbaSystem
        If InStr(GetObject("winmgmts:Win32_OperatingSystem=@").OSArchitecture, "64") Then
            getPlatformString = "win64"
        Else
            getPlatformString = "win32"
        End If
    Case svbaPlatform.svbaWin32
        getPlatformString = "win32"
    Case svbaPlatform.svbaWin64
        getPlatformString = "win64"
    End Select
End Function

Private Function getChannelString(ByVal val As svbaChannel) As String
    Select Case val
    Case svbaChannel.svbaStable
        getChannelString = "Stable"
    Case svbaChannel.svbaDev
        getChannelString = "Dev"
    Case svbaChannel.svbaBeta
        getChannelString = "Beta"
    Case svbaChannel.svbaCanary
        getChannelString = "Canary"
    End Select
End Function

'---------------------- RegExp Json Parser Functions ----------------------

Private Function findLatestChromeVersionFromBuild(ByVal jsonString As String, ByVal browserBuild As String) As String
    'this function assumes that input json is ordered from oldest to newest build/version numbers
    'it finds the MOST RECENT version of chrome driver that is compatible with the input browser build
    Dim oRegExp As New VBScript_RegExp_55.RegExp
    Dim matches As VBScript_RegExp_55.MatchCollection
    Dim match As VBScript_RegExp_55.match
    Dim lastVersion As String

    oRegExp.Global = True
    oRegExp.Pattern = """(\d+\.\d+\.\d+)"":{""version"":""(\d+\.\d+\.\d+\.\d+)"""
    Set matches = oRegExp.execute(jsonString)
    'find the maximum build that is less than or equal to installed browser build
    For Each match In matches
        If compareVersions(match.SubMatches(0), browserBuild) = 1 Then
            findLatestChromeVersionFromBuild = lastVersion
            Exit Function
        End If
        lastVersion = match.SubMatches(1)
    Next match
    'if here then return last version in the file
    findLatestChromeVersionFromBuild = lastVersion
End Function

Private Function findLatestFirefoxVersion(ByVal jsonString As String) As String
    'this function finds the latest Firefox driver version number
    Dim oRegExp As New VBScript_RegExp_55.RegExp
    Dim matches As VBScript_RegExp_55.MatchCollection
    
    oRegExp.Global = False
    oRegExp.Pattern = "geckodriver-v(\d+\.\d+\.\d+)-win32\.zip"

    Set matches = oRegExp.execute(jsonString)
    
    If matches.Count = 0 Then findLatestFirefoxVersion = "Error: Release Not Found": Exit Function
    
    findLatestFirefoxVersion = matches(0).SubMatches(0)
End Function

Private Function findVersionInString(ByVal stringToSearch As String, ByVal browser As svbaBrowser) As String
    'this function parses the input string for the version number substring
    Dim oRegExp As VBScript_RegExp_55.RegExp
    Dim matches As VBScript_RegExp_55.MatchCollection
    Dim searchPattern As String

    Select Case browser
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        searchPattern = "\d+\.\d+\.\d+\.\d+"
    Case svbaBrowser.Firefox
        searchPattern = "\d+\.\d+\.\d+"
    End Select
    
    On Error GoTo EH
    Set oRegExp = New VBScript_RegExp_55.RegExp
    oRegExp.Global = False
    oRegExp.Pattern = searchPattern

    Set matches = oRegExp.execute(stringToSearch)
    If matches.Count > 0 Then findVersionInString = matches(0)
    Exit Function
EH:
    Err.Raise 1, , "Error in findVersionInString of WebDriverManager: VBScript_RegExp failed - make sure that it is installed on your Windows system and is not being blocked by AV software."
End Function

'---------------------- Version String Math Functions ----------------------

Private Function compareVersions(ByVal ver1 As String, ByVal ver2 As String) As Long
    'returns 1 if ver1 > ver2
    'returns 0 if ver1 = ver2
    'returns -1 if ver1 < ver2
    Dim ver1PartsCount  As Long
    Dim ver2PartsCount  As Long
    Dim ver1parts As Variant
    Dim ver2parts As Variant
    Dim i As Long
    
    If ver1 = ver2 Then
        compareVersions = 0
        Exit Function
    End If
        
    'normalize version strings before compare
    
    'append 0 to string if ends in "."
    If Right$(ver1, 1) = "." Then ver1 = ver1 & "0"
    If Right$(ver2, 1) = "." Then ver2 = ver2 & "0"
    
    'make sure the version strings have same number of parts
    ver1PartsCount = countVersionParts(ver1)
    ver2PartsCount = countVersionParts(ver2)
    If ver1PartsCount <> ver2PartsCount Then
        If ver1PartsCount > ver2PartsCount Then
            For i = 1 To ver1PartsCount - ver2PartsCount
                ver2 = ver2 & ".0"
            Next i
        Else
            For i = 1 To ver2PartsCount - ver1PartsCount
                ver1 = ver1 & ".0"
            Next i
        End If
    End If
    
    'after normalization, check again if strings are equal
    If ver1 = ver2 Then
        compareVersions = 0
        Exit Function
    End If
    
    'split the strings into part arrays
    ver1parts = Split(ver1, ".")
    ver2parts = Split(ver2, ".")
    
    'compare each part until unequal, return 1 if ver1 > ver2, -1 otherwise
    For i = 0 To UBound(ver1parts)
        If CInt(ver1parts(i)) <> CInt(ver2parts(i)) Then
            compareVersions = Sgn(CInt(ver1parts(i)) - CInt(ver2parts(i)))
            Exit Function
        End If
    Next i
    
    'if we get here, then must be equal
    compareVersions = 0
End Function

Private Function filterVersionParts(ByVal versionString As String, ByVal numberOfPartsToKeep As Long) As String
    Dim verParts As Variant
    Dim i As Long
    verParts = Split(versionString, ".")
    If numberOfPartsToKeep >= UBound(verParts) + 1 Then
        filterVersionParts = versionString
        Exit Function
    End If
    versionString = verParts(0)
    For i = 1 To numberOfPartsToKeep - 1
        versionString = versionString & "." & verParts(i)
    Next i
    filterVersionParts = versionString
End Function

Private Function selectVersionPart(ByVal versionString As String, ByVal partNumber As Long) As String
    Dim verParts As Variant
    verParts = Split(versionString, ".")
    If partNumber > UBound(verParts) + 1 Then
        selectVersionPart = vbNullString
        Exit Function
    End If
    selectVersionPart = verParts(partNumber - 1)
End Function

Private Function countVersionParts(ByVal versionString As String) As String
    countVersionParts = UBound(Split(versionString, ".")) + 1
End Function

