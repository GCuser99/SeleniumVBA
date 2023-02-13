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
' SeleniumVBA v3.5
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
' Copyright (c) 2023, GCUser99 and 6DiegoDiego9 (https://github.com/GCuser99/SeleniumVBA)
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
'   If the driver has not yet been installed, the Align* methods with detect that and install
'   The Align* methods have a runSilent input parameter that if set to True, will not throw any of the expected errors -
'   this is used with the auto-check/install of the WebDriver Start* methods to make this feature as innocuous/stealth as possible.
'   Relative driver paths are supported (e.g. ".\msedgedriver.exe").
'   Use the DefaultDriverFolder property to specify a driver location other than the Downloads dir
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
    IE = 4
End Enum

Private Sub Class_Initialize()
    Dim oShell As New Shell32.Shell
    Dim iniFilePath As String
    
    'attempt to read global variable values from optional settings file - if not successful,
    'then assign a system default value (last input parameter of ReadIniFileEntry)
    iniFilePath = GetFullLocalPath(ThisLibFolderPath) & "\SeleniumVBA.ini"
    Me.DefaultDriverFolder = ReadIniFileEntry(iniFilePath, "GENERAL", "driver_location_folder", oShell.Namespace("shell:Downloads").Self.Path)
End Sub

'@Description("Sets the default driver folder path")
Public Property Let DefaultDriverFolder(ByVal folderPath As String)
Attribute DefaultDriverFolder.VB_Description = "Sets the default driver folder path"
    mDefaultDriverFolder = WebShared.GetFullLocalPath(folderPath)
End Property

Public Property Get DefaultDriverFolder() As String
    DefaultDriverFolder = mDefaultDriverFolder
End Property

Friend Function AlignDriverWithBrowser(ByVal browser As svbaBrowser, ByVal driverPath As String, Optional ByVal minCompatibilityLevel As svbaCompatibility = svbaExactMatch, Optional ByVal runSilent As Boolean = False) As String
    Dim bverInstalled As String
    Dim dverCompat As String
    Dim respStr As String
    Dim dverInstalled As String
    
    bverInstalled = GetInstalledBrowserVersion(browser)
    'user might be using a portable browser that will be set in capabilities
    'so if browser is not found in registry, then set browser version to a
    'dummy value and then install latest driver, if not already installed
    If Left$(bverInstalled, 5) = "Error" Then bverInstalled = verBrowserNotFound
    
    dverInstalled = GetInstalledDriverVersion(browser, driverPath)
    
    Select Case browser
    Case svbaBrowser.Edge, svbaBrowser.Chrome
        
        'because of the way Edge and Chrome maintain synchrony between diver and browser version numbers
        'we can compare driver and browser version strings to determine compatibility
        If CheckCompatibilityLevel(bverInstalled, dverInstalled) < minCompatibilityLevel Then
        
            'the installed driver version is sufficiently different than the installed browser version
            'so find the most compatible driver version for installed browser
            If bverInstalled = verBrowserNotFound Then
                dverCompat = GetLatestDriverVersion(browser)
            Else
                dverCompat = GetCompatibleDriverVersion(browser, bverInstalled)
            End If
        
            'check if we have an error and then take appropriate action
            If Left$(dverCompat, 5) = "Error" Then
                If runSilent Then 'don't throw the error - just report back to caller
                    AlignDriverWithBrowser = dverCompat
                    Exit Function
                Else
                    Err.raise 1, "WebDriverManager", dverCompat
                End If
            End If
            
            'It's possible that the browser version is a dev version, and there is not
            'an updated driver version. It's also possible that the browser was not found and
            'that dverCompat is latest version. Check to make sure dverCompat and dverInstalled
            'are sufficiently different, then if so, install the "compatible" driver
            If CheckCompatibilityLevel(dverCompat, dverInstalled) < minCompatibilityLevel Then
                DownloadAndInstallDriver browser, dverCompat, driverPath
                dverInstalled = dverCompat
                respStr = "The latest " & BrowserNameStr(browser) & " Webdriver was installed."
            Else
                respStr = BrowserNameStr(browser) & " Webdriver and Browser are compatible."
            End If
        Else
            respStr = BrowserNameStr(browser) & " Webdriver and Browser are compatible."
        End If
    Case svbaBrowser.Firefox
                 
        'the installed driver version is sufficiently different than the installed browser version
        'so find the most compatible driver version for installed browser
        If bverInstalled = verBrowserNotFound Then
            dverCompat = GetLatestDriverVersion(browser)
        Else
            dverCompat = GetCompatibleDriverVersion(browser, bverInstalled)
        End If
                 
        'check if we have an error and then take appropriate action
        If Left$(dverCompat, 5) = "Error" Then
            If runSilent Then 'don't throw the error - just report back to caller
                AlignDriverWithBrowser = dverCompat
                Exit Function
            Else
                Err.raise 1, "WebDriverManager", dverCompat
            End If
        End If
        
        If CheckCompatibilityLevel(dverCompat, dverInstalled) < minCompatibilityLevel Then
            'the installed driver version is sufficiently different than the installed browser version
            'so install it
            DownloadAndInstallDriver browser, dverCompat, driverPath
            dverInstalled = dverCompat
            respStr = "The latest " & BrowserNameStr(browser) & " Webdriver was installed."
        Else
            respStr = BrowserNameStr(browser) & " Webdriver and Browser are compatible."
        End If
    Case svbaBrowser.IE
    
        'the installed driver version is sufficiently different than the installed browser version
        'so find the most compatible driver version for installed browser
        If bverInstalled = verBrowserNotFound Then
            dverCompat = GetLatestDriverVersion(browser)
        Else
            dverCompat = GetCompatibleDriverVersion(browser, bverInstalled)
        End If
                 
        'check if we have an error and then take appropriate action
        If Left$(dverCompat, 5) = "Error" Then
            If runSilent Then 'don't throw the error - just report back to caller
                AlignDriverWithBrowser = dverCompat
                Exit Function
            Else
                Err.raise 1, "WebDriverManager", dverCompat
            End If
        End If
        
        If CheckCompatibilityLevel(dverCompat, dverInstalled) < minCompatibilityLevel Then
            'the installed driver version is sufficiently different than the installed browser version
            'so install it
            DownloadAndInstallDriver browser, dverCompat, driverPath
            dverInstalled = dverCompat
            respStr = "The latest " & BrowserNameStr(browser) & " Webdriver was installed."
        Else
            respStr = BrowserNameStr(browser) & " Webdriver and Browser are compatible."
        End If
    
    End Select
    
    If bverInstalled = verBrowserNotFound Then bverInstalled = "Not found"
    
    respStr = respStr & vbNewLine & vbNewLine & "Browser version: " & bverInstalled & vbNewLine & "Driver version:    " & dverInstalled
    
    AlignDriverWithBrowser = respStr
End Function

'@Description("Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed")
Public Function AlignEdgeDriverWithBrowser(Optional ByVal driverPath As String = "msedgedriver.exe", Optional ByVal minCompatibilityLevel As svbaCompatibility = svbaExactMatch, Optional ByVal runSilent As Boolean = False) As String
Attribute AlignEdgeDriverWithBrowser.VB_Description = "Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed"
    AlignEdgeDriverWithBrowser = AlignDriverWithBrowser(svbaBrowser.Edge, driverPath, minCompatibilityLevel, runSilent)
End Function

'@Description("Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed")
Public Function AlignChromeDriverWithBrowser(Optional ByVal driverPath As String = "chromedriver.exe", Optional ByVal minCompatibilityLevel As svbaCompatibility = svbaExactMatch, Optional ByVal runSilent As Boolean = False) As String
Attribute AlignChromeDriverWithBrowser.VB_Description = "Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed"
    AlignChromeDriverWithBrowser = AlignDriverWithBrowser(svbaBrowser.Chrome, driverPath, minCompatibilityLevel, runSilent)
End Function

'@Description("Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed")
Public Function AlignFirefoxDriverWithBrowser(Optional ByVal driverPath As String = "geckodriver.exe", Optional ByVal minCompatibilityLevel As svbaCompatibility = svbaExactMatch, Optional ByVal runSilent As Boolean = False) As String
Attribute AlignFirefoxDriverWithBrowser.VB_Description = "Checks if driver is installed, or if installed driver is compatibile with installed browser, and then installs an updated driver if needed"
    AlignFirefoxDriverWithBrowser = AlignDriverWithBrowser(svbaBrowser.Firefox, driverPath, minCompatibilityLevel, runSilent)
End Function

'@Description("Checks if driver is installed, or if installed driver is compatibile with installed browser, and then if needed, installs an updated driver")
Public Function AlignIEDriverWithBrowser(Optional ByVal driverPath As String = "IEdriverServer.exe", Optional ByVal minCompatibilityLevel As svbaCompatibility = svbaExactMatch, Optional ByVal runSilent As Boolean = False) As String
Attribute AlignIEDriverWithBrowser.VB_Description = "Checks if driver is installed, or if installed driver is compatibile with installed browser, and then if needed, installs an updated driver"
    AlignIEDriverWithBrowser = AlignDriverWithBrowser(svbaBrowser.IE, driverPath, minCompatibilityLevel, runSilent)
End Function

'@Description("Returns the installed Driver version")
Public Function GetInstalledDriverVersion(ByVal browser As svbaBrowser, Optional ByVal driverPath As Variant) As String
Attribute GetInstalledDriverVersion.VB_Description = "Returns the installed Driver version"
    Dim wsh As New IWshRuntimeLibrary.WshShell 'CreateObject("WScript.Shell")
    Dim fso As New IWshRuntimeLibrary.FileSystemObject
    Dim resp As Variant
    Dim i As Integer
    Dim ver As String

    If IsMissing(driverPath) Then
        Select Case browser
        Case svbaBrowser.Chrome
            driverPath = ".\chromedriver.exe"
        Case svbaBrowser.Edge
            driverPath = ".\msedgedriver.exe"
        Case svbaBrowser.Firefox
            driverPath = ".\geckodriver.exe"
        Case svbaBrowser.IE
            driverPath = ".\IEdriverServer.exe"
        End Select
    End If
    
    driverPath = WebShared.GetFullLocalPath(driverPath, mDefaultDriverFolder)
    
    If Not fso.FileExists(driverPath) Then
        GetInstalledDriverVersion = verDriverNotFound
        Exit Function
    End If
    
    resp = wsh.Exec(Chr$(34) & driverPath & Chr$(34) & " --version").StdOut.ReadAll
    
    'assume only numbers and periods are allowed in version numbers(?)
    Dim vparts() As String
    vparts = Split(resp, " ")
    For i = 0 To UBound(vparts)
        If IsNumeric(Replace$(vparts(i), ".", vbNullString)) Then
            ver = vparts(i)
            Exit For
        End If
    Next i
    
    GetInstalledDriverVersion = ver
End Function

'@Description("Returns the compatibility level between a pair of version strings")
Public Function CheckCompatibilityLevel(ByVal versionString1 As String, ByVal versionString2 As String) As svbaCompatibility
Attribute CheckCompatibilityLevel.VB_Description = "Returns the compatibility level between a pair of version strings"
    Dim vermajor1 As Variant, verminor1 As Variant, verbuildmajor1 As Variant, verbuildminor1 As Variant
    Dim vermajor2 As Variant, verminor2 As Variant, verbuildmajor2 As Variant, verbuildminor2 As Variant
    
    If versionString1 = "driver not found" Then CheckCompatibilityLevel = -1: Exit Function
    vermajor1 = Split(versionString1, ".")(0)
    verminor1 = Split(versionString1, ".")(1)
    verbuildmajor1 = Split(versionString1, ".")(2)
    If UBound(Split(versionString1, ".")) > 2 Then verbuildminor1 = Split(versionString1, ".")(3) Else verbuildminor1 = "0"
    
    vermajor2 = Split(versionString2, ".")(0)
    verminor2 = Split(versionString2, ".")(1)
    verbuildmajor2 = Split(versionString2, ".")(2)
    If UBound(Split(versionString2, ".")) > 2 Then verbuildminor2 = Split(versionString2, ".")(3) Else verbuildminor2 = "0"
    
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
Public Function GetLatestDriverVersion(ByVal browser As svbaBrowser) As String
Attribute GetLatestDriverVersion.VB_Description = "Returns the latest Driver version"
    Dim jc As New WebJsonConverter
    Dim resp As New Dictionary
    Dim url As String
    Dim responseText As Variant
    Dim try As Long
    Dim searchOS As String
    Dim i As Long, j As Long
    Dim durl As String
    Dim col As Collection
    Dim json As Dictionary
    
    Select Case browser
    Case svbaBrowser.Edge
        url = "https://msedgedriver.azureedge.net/LATEST_STABLE"
        responseText = RequestData(url)
        If responseText = "Release Not Found" Then GetLatestDriverVersion = "Error: " & responseText: Exit Function
        GetLatestDriverVersion = Replace$(responseText, vbNewLine, vbNullString)
    Case svbaBrowser.Chrome
        url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
        responseText = RequestData(url)
        If responseText = "Release Not Found" Then GetLatestDriverVersion = "Error: " & responseText: Exit Function
        GetLatestDriverVersion = responseText
    Case svbaBrowser.Firefox
        url = "https://api.github.com/repos/mozilla/geckodriver/releases/latest"
        responseText = RequestData(url)
        If responseText = "Release Not Found" Then GetLatestDriverVersion = "Error: " & responseText: Exit Function
        Set resp = jc.ParseJson(responseText)
        'now search all
        'if os=64bit then search for win64 otherwise win32
        'below returns download url
        If InStr(GetObject("winmgmts:Win32_OperatingSystem=@").OSArchitecture, "64") Then
            searchOS = "win64"
        Else
            searchOS = "win32"
        End If

        For i = 1 To resp("assets").Count
            durl = resp("assets")(i)("browser_download_url")
            If InStr(durl, searchOS) Then Exit For
        Next i
        GetLatestDriverVersion = resp("name")
    Case svbaBrowser.IE
        url = "https://api.github.com/repos/SeleniumHQ/selenium/releases"
        responseText = RequestData(url)
        If responseText = "Release Not Found" Then GetLatestDriverVersion = "Error: " & responseText: Exit Function
        Set col = jc.ParseJson(responseText)
        'search the release log until an IEDriveServer is found
        For j = 1 To col.Count
            Set json = col.Item(j)
            For i = 1 To json("assets").Count
                If json("assets")(i)("name") Like "IEDriverServer_Win32*" Then
                    GetLatestDriverVersion = Mid$(json("tag_name"), 10)
                    Exit Function
                End If
            Next i
        Next j
    End Select
End Function

'@Description("Returns the most recent compatible Driver version")
Public Function GetCompatibleDriverVersion(ByVal browser As svbaBrowser, ByVal browserVersion As String) As String
Attribute GetCompatibleDriverVersion.VB_Description = "Returns the most recent compatible Driver version"
    Dim htmlDoc As New HTMLDocument
    Dim table As htmlTable
    Dim row As HTMLTableRow
    Dim verMajor As String
    Dim verMajorSearch As String
    Dim verMinor As String
    Dim verBuild As String
    Dim verSearch As String
    Dim responseText As Variant
    Dim url As String
    Dim try As Integer
    
    'This function, if called from WebDriver class as part of auto-updater, needs to execution as silently
    'as possible - we don't want the non-mission-critical error in driver check/updating to get in the way of user
    'running WebDriver. For example, user may be running on local urls, and so in no internet connection, user can still work
    
    'for chrome info:
    'https://chromedriver.chromium.org/downloads/version-selection
    
    'for edge info:
    'https://blogs.windows.com/msedgedev/2019/11/04/edge-chromium-release-candidate-get-ready/
    'https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
    
    Select Case browser
    Case svbaBrowser.Edge
        
        verMajor = Split(browserVersion, ".")(0)

        'search for major version, decrementing by one until success
        Do While verMajor >= 75
            url = "https://msedgedriver.azureedge.net/LATEST_RELEASE_" & verMajor & "_WINDOWS"
            
            responseText = RequestData(url)
            
            If Left$(responseText, 5) = "Error" Then
                GetCompatibleDriverVersion = responseText
                Exit Function
            End If
            
            If responseText <> "Release Not Found" Then Exit Do
            
            verMajor = verMajor - 1
        Loop
        
        GetCompatibleDriverVersion = Replace$(responseText, vbNewLine, vbNullString)

    Case svbaBrowser.Chrome

        verMajor = Split(browserVersion, ".")(0)
        verMinor = Split(browserVersion, ".")(1)
        verBuild = Split(browserVersion, ".")(2)
        
        'first see if we can find an exact match
        url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & verMajor & "." & verMinor & "." & verBuild
        
        responseText = RequestData(url)
        
        If Left$(responseText, 5) = "Error" Then
            GetCompatibleDriverVersion = responseText
            Exit Function
        End If
        
        If responseText = "Release Not Found" Then
            'if exact match failed, then maybe we have a Dev or Canary channel
            'so search for major version match, decrementing by one until success

            Do While verMajor >= 70
                url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & verMajor
                responseText = RequestData(url)
                
                If Left$(responseText, 5) = "Error" Then
                    GetCompatibleDriverVersion = responseText
                    Exit Function
                End If
                
                If responseText <> "Release Not Found" Then Exit Do
                verMajor = verMajor - 1
            Loop
        End If
        
        GetCompatibleDriverVersion = responseText
        
    Case svbaBrowser.Firefox
        
        verMajor = Split(browserVersion, ".")(0)
        
        'the rules are different for Firefox driver - we must lookup the browser version in a compatibility matrix
        'to get the most recent driver version compatible with the installed browser
        url = "https://firefox-source-docs.mozilla.org/testing/geckodriver/Support.html"
        
        responseText = RequestData(url)
        
        If Left$(responseText, 5) = "Error" Then
            GetCompatibleDriverVersion = responseText
            Exit Function
        End If
        
        'search the table for the first driver version entry (most recent) for a given browser version
        htmlDoc.body.innerHTML = responseText
        Set table = htmlDoc.getElementsByTagName("table")(0)
        
        table.deleteTHead
        
        For Each row In table.Rows
            If CInt(verMajor) >= CInt(Split(row.Cells(2).innerText)(0)) Then
                GetCompatibleDriverVersion = Trim$(row.Cells(0).innerText)
                Exit For
            End If
        Next row
    Case svbaBrowser.IE
        'just get the latest release, independent of Edge browser version
        GetCompatibleDriverVersion = GetLatestDriverVersion(svbaBrowser.IE)
    End Select
End Function

'@Description("Returns the installed Browser version")
Public Function GetInstalledBrowserVersion(ByVal browser As svbaBrowser) As String
Attribute GetInstalledBrowserVersion.VB_Description = "Returns the installed Browser version"
    Dim wsh As New IWshRuntimeLibrary.WshShell 'CreateObject("WScript.Shell")
    Dim regkey As String
    Dim resp As String
    
    Select Case browser
    Case svbaBrowser.Chrome
        regkey = "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon\version"
    Case svbaBrowser.Edge
        regkey = "HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon\version"
    Case svbaBrowser.Firefox
        regkey = "HKEY_LOCAL_MACHINE\SOFTWARE\Mozilla\Mozilla Firefox\CurrentVersion"
    Case svbaBrowser.IE
        regkey = "HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon\version"
    End Select
    
    'this throws error if key doesnt exist - report back to caller if browser isn't installed
    On Error Resume Next
    resp = wsh.RegRead(regkey)
    If Err.Number <> 0 Then
        resp = "Error: " & BrowserNameStr(browser) & " Browser Not Installed"
    End If

    GetInstalledBrowserVersion = resp
End Function

'@Description("Downloads and installs the specified Driver version")
Public Sub DownloadAndInstallDriver(ByVal browser As svbaBrowser, ByVal driverVersion As String, Optional ByVal driverPath As Variant)
Attribute DownloadAndInstallDriver.VB_Description = "Downloads and installs the specified Driver version"
    Dim url As String
    Dim rurl As String
    Dim fso As New IWshRuntimeLibrary.FileSystemObject
    Dim shll As New Shell32.Shell
    Dim driverFolder As String
    Dim driverFilename As String
    Dim tmpInstallFolder As String
    Dim createdTmpFolder As Boolean
    Dim tmpInstallZipPath As String
    Dim tmpZipDriverFilename As String
    Dim tmpInstallDriverpath As String
    Dim i As Long, j As Long
    Dim jc As New WebJsonConverter
    Dim json As Dictionary
    Dim col As Collection
    
    If IsMissing(driverPath) Then
        Select Case browser
        Case svbaBrowser.Chrome
            driverPath = ".\chromedriver.exe"
        Case svbaBrowser.Edge
            driverPath = ".\msedgedriver.exe"
        Case svbaBrowser.Firefox
            driverPath = ".\geckodriver.exe"
        Case Else
            driverPath = ".\IEDriverServer.exe"
        End Select
    End If
    
    driverPath = WebShared.GetFullLocalPath(driverPath, mDefaultDriverFolder)
    
    driverFolder = fso.GetParentFolderName(driverPath)
    driverFilename = fso.GetFileName(driverPath)
    
    tmpInstallFolder = driverFolder & "\tmp\"
    
    Select Case browser
    Case svbaBrowser.Chrome
        url = "https://chromedriver.storage.googleapis.com/" & driverVersion & "/chromedriver_win32.zip"
        tmpInstallZipPath = tmpInstallFolder & "chromedriver_win32.zip"
        tmpZipDriverFilename = "chromedriver.exe"
        tmpInstallDriverpath = tmpInstallFolder & tmpZipDriverFilename
    Case svbaBrowser.Edge
        'test for 64 bit vs 32 bit
        If InStr(GetObject("winmgmts:Win32_OperatingSystem=@").OSArchitecture, "64") Then
            url = "https://msedgedriver.azureedge.net/" & driverVersion & "/edgedriver_win64.zip"
            tmpInstallZipPath = tmpInstallFolder & "edgedriver_win64.zip"
        Else
            url = "https://msedgedriver.azureedge.net/" & driverVersion & "/edgedriver_win32.zip"
            tmpInstallZipPath = tmpInstallFolder & "edgedriver_win32.zip"
        End If
        tmpZipDriverFilename = "msedgedriver.exe"
        tmpInstallDriverpath = tmpInstallFolder & tmpZipDriverFilename
    Case svbaBrowser.Firefox
        'test for 64 bit vs 32 bit
        If InStr(GetObject("winmgmts:Win32_OperatingSystem=@").OSArchitecture, "64") Then
            url = "https://github.com/mozilla/geckodriver/releases/download/v" & driverVersion & "/geckodriver-v" & driverVersion & "-win64.zip"
            tmpInstallZipPath = tmpInstallFolder & "geckodriver-v" & driverVersion & "-win64.zip"
        Else
            url = "https://github.com/mozilla/geckodriver/releases/download/v" & driverVersion & "/geckodriver-v" & driverVersion & "-win32.zip"
            tmpInstallZipPath = tmpInstallFolder & "geckodriver-v" & driverVersion & "-win32.zip"
        End If
        tmpZipDriverFilename = "geckodriver.exe"
        tmpInstallDriverpath = tmpInstallFolder & tmpZipDriverFilename
    Case svbaBrowser.IE
        'just get the latest release, independent of Edge browser version
        rurl = "https://api.github.com/repos/SeleniumHQ/selenium/releases"
        Set col = jc.ParseJson(RequestData(rurl))
        'search the release log until an IEDriveServer is found
        For j = 1 To col.Count
            Set json = col.Item(j)
            'this finds the executable download link
            For i = 1 To json("assets").Count
                If json("assets")(i)("name") Like "IEDriverServer_Win32*" Then
                    url = json("assets")(i)("browser_download_url")
                    tmpInstallZipPath = tmpInstallFolder & json("assets")(i)("name")
                    Exit For
                End If
            Next i
            If url <> vbNullString Then Exit For
        Next j
        If url = vbNullString Then Err.raise 1, "WebDriverManager", "Error: Could not locate latest version of IEDriverServer"
        tmpZipDriverFilename = "IEDriverServer.exe"
        tmpInstallDriverpath = tmpInstallFolder & tmpZipDriverFilename
    End Select
    
    'if tmp folder doesn't exist then create it - remember for clean up
    createdTmpFolder = False
    If Not fso.FolderExists(tmpInstallFolder) Then
        createdTmpFolder = True
        fso.CreateFolder tmpInstallFolder
    End If
    
    URLDownloadToFile 0, url, tmpInstallZipPath, 0, 0
    
    'make sure driver zip file was created successfully
    If Not fso.FileExists(tmpInstallZipPath) Then
        'clean up first
        If createdTmpFolder Then fso.DeleteFolder Left$(tmpInstallFolder, Len(tmpInstallFolder) - 1), True
        Err.raise 1, "WebDriverManager", "Error: Driver Download Not Successful"
    End If
    
    If fso.FileExists(tmpInstallDriverpath) Then fso.DeleteFile tmpInstallDriverpath, True

    shll.Namespace(tmpInstallFolder).CopyHere shll.Namespace(tmpInstallZipPath).Items.Item(tmpZipDriverFilename)

    'check if driver exists and if it does, make sure it isn't in use before overwrite
    If fso.FileExists(driverPath) Then
        If IsProcessRunning(driverFilename) Then TaskKillByImage (driverFilename)
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
    GetSeleniumBasicFolderPath = Environ$("LocalAppData") & "\SeleniumBasic\"
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

Private Function IsProcessRunning(ByVal process As String) As Boolean
    Dim objServices As WbemScripting.SWbemServices
    Dim objProcessList As WbemScripting.SWbemObjectSet
    
    Set objServices = GetObject("winmgmts:")
    Set objProcessList = objServices.ExecQuery("select * from win32_process where name='" & process & "'")
    
    If objProcessList.Count > 0 Then
        IsProcessRunning = True
    Else
        IsProcessRunning = False
    End If
End Function

Private Function RequestData(ByVal url As String, Optional maxTries As Integer = 20) As Variant
    'see https://github.com/GCuser99/SeleniumVBA/discussions/29#discussioncomment-3673682
    Dim client As New MSXML2.XMLHTTP60
    Dim try As Integer
        
    'try until a connection to server is established, up to a maximum # of tries
    For try = 1 To maxTries
    
        On Error Resume Next
        client.Open "GET", url, False
        
        client.send
        
        While client.readyState <> 4
            DoEvents
        Wend
        
        On Error GoTo 0
        
        Select Case client.status
        Case 200
            RequestData = client.responseText
            Exit Function
        Case 404
            RequestData = "Release Not Found"
            Exit Function
        Case 12007
            RequestData = "Error: No Internet Connection"
            Exit Function
        Case 12029, 12031
            'Don't exit loop - make another try to see if we can connect
            'RequestData = "Error: Cannot Connect to Server"
        Case Else 'an unexpected error
            Debug.Print "Please report this error to developers: RequestData Status= " & client.status & ", " & "Response= " & client.responseText
            RequestData = "Error: Status= " & client.status & ", " & "Response= " & client.responseText
            Exit Function
        End Select
        
        Debug.Print "try: ", try, client.status, client.responseText, url
    Next try
    'if here then had too many 12029/12031's
    RequestData = "Error: Cannot Connect to Server"
End Function

Private Function BrowserNameStr(ByVal val As svbaBrowser) As String
    Select Case val
    Case svbaBrowser.Edge
        BrowserNameStr = "Edge"
    Case svbaBrowser.Chrome
        BrowserNameStr = "Chrome"
    Case svbaBrowser.Firefox
        BrowserNameStr = "Firefox"
    Case svbaBrowser.IE
        BrowserNameStr = "IE"
    End Select
End Function

Private Function TaskKillByImage(ByVal taskName As String) As Variant
    Dim wsh As New IWshRuntimeLibrary.WshShell
    TaskKillByImage = wsh.Run("taskkill /f /t /im " & taskName, 0, True)
End Function
