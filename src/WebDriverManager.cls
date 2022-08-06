VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================
' SeleniumVBA v1.6
' A Selenium wrapper for Edge and Chrome written in Windows VBA based on JSon
' wire protocol.
'
' (c) GCUser99
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' ==========================================================================
' MIT License
'
' Copyright (c) 2022 GCUser99
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
'    mngr.AlignEdgeDriverWithBrowser ".\msedgedriver.exe"
'    mngr.AlignChromeDriverWithBrowser ".\chromedriver.exe"
'    mngr.AlignFirefoxDriverWithBrowser ".\geckodriver.exe"
'
'    The Methods above will also return an informative status string, as in:
'
'    msg = mngr.AlignEdgeDriverWithBrowser()
'
'    Alternatively, the update process can be controlled with more granularity, as in the following perturbation...
'
'    Dim mngr As New WebDriverManager
'    driverPath = ".\chromedriver.exe"
'    If Not mngr.IsInstalledDriverCompatible(svbaBrowser.Chrome, driverPath) Then
'        resp = MsgBox("WebDriver is not compatible with installed browser - would you like to install the compatible WebDriver?", vbYesNo)
'        If resp = vbYes Then
'            bverInstalled = mngr.GetInstalledBrowserVersion(svbaBrowser.Chrome)
'            dverCompat = mngr.GetCompatibleDriverVersion(svbaBrowser.Chrome, bverInstalled)
'            mngr.DownloadAndInstallDriver svbaBrowser.Chrome, dverCompat, driverPath
'            MsgBox "Chrome" & " " & "Webdriver and Browser are compatible!" & vbCrLf & vbCrLf & "Browser version: " & bverInstalled & vbCrLf & "Driver version:    " & dverCompat, , "SeleniumVBA"
'        End If
'    Else
'        MsgBox "Chrome " & "Webdriver and Browser are compatible!" & vbCrLf & vbCrLf & "Browser version: " & mngr.GetInstalledBrowserVersion(svbaBrowser.Chrome) & vbCrLf & "Driver version:    " & mngr.GetInstalledDriverVersion(svbaBrowser.Chrome, driverPath), , "SeleniumVBA"
'    End If
'
' NOTES:
'
'   Relative driver paths are supported (e.g. ".\msedgedriver.exe").
'   A different level of version compatibility can be specified [svbaMajor, svbaMinor, svbaBuildMajor, svbaBuildMinor].
'   For example the default is set to svbaMajor, meaning that only the major version numbers to the left of the first
'   period need to match to deem the browser and driver are compatible. In fact this seems to be the consensus on
'   coding sites but just in case that is not always the case, this class will accommodate a more stringent
'   test if needed. To change the desired minimum level of compatibility, modify the value assigned to
'   mMinCompatibilityLevel in the Class_Initialize sub of this class.
'
' REFERENCES:
'
'   Windows Script Host Model
'   Microsoft ActiveX Data Objects 6.1 Library
'   Microsoft Shell Controls and Automation
'   Microsoft XML, V6.0
'   Microsoft WMI Scripting V1.2 Library
'

Option Explicit

Private mMinCompatibilityLevel As svbaCompatibility

Public Enum svbaCompatibility
    svbaNotCompatible = 0
    svbaMajor = 1
    svbaMinor = 2
    svbaBuildMajor = 3
    svbaBuildMinor = 4
End Enum

Public Enum svbaBrowser
    Chrome = 1
    Edge = 2
    Firefox = 3
End Enum

Private Function GetBrowserName(ByVal browser As svbaBrowser) As String
    Dim browserName As String
    Select Case browser
    Case svbaBrowser.Chrome
        browserName = "chrome"
    Case svbaBrowser.Edge
        browserName = "msedge"
    Case svbaBrowser.Firefox
        browserName = "firefox"
    End Select
    GetBrowserName = browserName
End Function

Property Let MinCompatibilityLevel(ByVal val As svbaCompatibility)
    mMinCompatibilityLevel = val
End Property

Property Get MinCompatibilityLevel() As svbaCompatibility
    MinCompatibilityLevel = mMinCompatibilityLevel
End Property

'this one does it all
Private Function AlignDriverWithBrowser(ByVal browser As svbaBrowser, ByVal driverPath As String) As String
    Dim bverInstalled As String, dverCompat As String, respStr As String
    
    bverInstalled = GetInstalledBrowserVersion(browser)
    
    If bverInstalled = "browser not installed" Then
        respStr = "Could not find the " & BrowserNameStr(browser) & " browser - install before using the WebDriver"
        AlignDriverWithBrowser = respStr
        Exit Function
    End If
    
    dverCompat = GetCompatibleDriverVersion(browser, bverInstalled)
    
    If Not IsInstalledDriverCompatible(browser, driverPath, bverInstalled, dverCompat) Then
        DownloadAndInstallDriver browser, dverCompat, driverPath
        respStr = "The latest " & BrowserNameStr(browser) & " Webdriver was installed."
    Else
        respStr = BrowserNameStr(browser) & " Webdriver and Browser are compatible."
    End If
    
    respStr = respStr & vbCrLf & vbCrLf & "Browser version: " & bverInstalled & vbCrLf & "Driver version:    " & dverCompat
    
    AlignDriverWithBrowser = respStr
End Function

Public Function AlignEdgeDriverWithBrowser(Optional ByVal driverPath As String = ".\msedgedriver.exe") As String
    AlignEdgeDriverWithBrowser = AlignDriverWithBrowser(svbaBrowser.Edge, driverPath)
End Function

Public Function AlignChromeDriverWithBrowser(Optional ByVal driverPath As String = ".\chromedriver.exe") As String
    AlignChromeDriverWithBrowser = AlignDriverWithBrowser(svbaBrowser.Chrome, driverPath)
End Function

Public Function AlignFirefoxDriverWithBrowser(Optional ByVal driverPath As String = ".\geckodriver.exe") As String
    AlignFirefoxDriverWithBrowser = AlignDriverWithBrowser(svbaBrowser.Firefox, driverPath)
End Function

Public Function IsInstalledDriverCompatible(ByVal browser As svbaBrowser, Optional ByVal driverPath, _
                                                        Optional bverInstalled, Optional dverCompat) As Boolean
    Dim fso As New IWshRuntimeLibrary.FileSystemObject, clevel As svbaCompatibility
    Dim dverinstalled As String
    
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
    
    driverPath = GetAbsolutePath(driverPath)
    
    If Not fso.FileExists(driverPath) Then
        IsInstalledDriverCompatible = False
        Exit Function
    End If
    
    If IsMissing(bverInstalled) Then bverInstalled = Me.GetInstalledBrowserVersion(browser)
    If bverInstalled = "browser not installed" Then
        IsInstalledDriverCompatible = False
        Exit Function
    End If
    
    dverinstalled = Me.GetInstalledDriverVersion(browser, driverPath)
    If IsMissing(dverCompat) Then dverCompat = Me.GetCompatibleDriverVersion(browser, bverInstalled)
    clevel = Me.CheckCompatibilityLevel(dverinstalled, dverCompat)
    
    If clevel >= mMinCompatibilityLevel Then
        IsInstalledDriverCompatible = True
    Else
        IsInstalledDriverCompatible = False
    End If
End Function

Public Function GetInstalledDriverVersion(ByVal browser As svbaBrowser, Optional ByVal driverPath) As String
    Dim wsh As New IWshRuntimeLibrary.wshShell 'CreateObject("WScript.Shell")
    Dim fso As New IWshRuntimeLibrary.FileSystemObject
    Dim resp As Variant, i As Integer, ver As String

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
    
    driverPath = GetAbsolutePath(driverPath)
    
    If Not fso.FileExists(driverPath) Then
        GetInstalledDriverVersion = "driver not found"
        Exit Function
    End If
    
    resp = wsh.Exec(Chr(34) & driverPath & Chr(34) & " --version").StdOut.ReadAll
    
    'ipos = InStr(resp, ".")
    'iposStart = InStrRev(resp, " ", ipos) + 1
    'iposEnd = InStr(iposStart, resp, " ") - 1
    'ver = Mid(resp, iposStart, iposEnd - iposStart + 1)
    
    'assume only numbers and periods are allowed in version numbers(?)
    Dim vparts() As String
    vparts = Split(resp, " ")
    For i = 0 To UBound(vparts)
        If IsNumeric(Replace(vparts(i), ".", "")) Then
            ver = vparts(i)
            Exit For
        End If
    Next i
    
    GetInstalledDriverVersion = ver
End Function

Public Function CheckCompatibilityLevel(ByVal installedVersion As String, ByVal compatibleVersion As String) As svbaCompatibility
    Dim vermajor1 As Variant, verminor1 As Variant, verbuildmajor1 As Variant, verbuildminor1 As Variant
    Dim vermajor2 As Variant, verminor2 As Variant, verbuildmajor2 As Variant, verbuildminor2 As Variant
    
    If installedVersion = "driver not found" Then CheckCompatibilityLevel = -1: Exit Function
    vermajor1 = Split(installedVersion, ".")(0)
    verminor1 = Split(installedVersion, ".")(1)
    verbuildmajor1 = Split(installedVersion, ".")(2)
    If UBound(Split(installedVersion, ".")) > 2 Then verbuildminor1 = Split(installedVersion, ".")(3) Else verbuildminor1 = ""
    
    vermajor2 = Split(compatibleVersion, ".")(0)
    verminor2 = Split(compatibleVersion, ".")(1)
    verbuildmajor2 = Split(compatibleVersion, ".")(2)
    If UBound(Split(compatibleVersion, ".")) > 2 Then verbuildminor2 = Split(compatibleVersion, ".")(3) Else verbuildminor2 = ""
    
    If val(Replace(installedVersion, ".", "")) > val(Replace(compatibleVersion, ".", "")) Then
        Stop 'installed is greater than compatible
    End If
    
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
    
    CheckCompatibilityLevel = svbaBuildMinor
End Function

Public Function GetLatestDriverVersion(ByVal browser As svbaBrowser) As String
    'this function is here for completeness, although not used
    Dim jc As New WebJSonConverter, resp As New Dictionary, url As String
    Dim responseText As Variant, try As Integer, searchOS As String, i As Integer, durl As String
    
    Select Case browser
    Case svbaBrowser.Edge
        url = "https://msedgedriver.azureedge.net/LATEST_STABLE"
        'the following loop is to deal with "-2147012739 - An error occurred in the secure channel support" for msedge driver LATEST_RELEASE_XX api
        'this usually gets a connection in less than 4 tries
        Do
            responseText = RequestData(url)
            try = try + 1
            'Debug.Print "try: " & try
            DoEvents
            If try > 10 Then
                MsgBox "could not connect to msedgedriver server"
                GetLatestDriverVersion = False
                Exit Function
            End If
        Loop Until responseText <> False
        GetLatestDriverVersion = Replace(responseText, vbCrLf, "")
    Case svbaBrowser.Chrome
        url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
        responseText = RequestData(url)
        GetLatestDriverVersion = responseText
    Case svbaBrowser.Firefox
        url = "https://api.github.com/repos/mozilla/geckodriver/releases/latest"
        responseText = RequestData(url)
        
        Set resp = jc.ParseJSON(responseText)
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
    End Select
End Function

Public Function GetCompatibleDriverVersion(ByVal browser As svbaBrowser, ByVal browserVersion) As String
    Dim htmldoc As New HTMLDocument, table As IHTMLTable, row As IHTMLTableRow, cells As IHTMLElementCollection
    Dim vermajor As String, vermajorsearch As String, verminor As String, verbuild As String, versearch As String
    Dim responseText As Variant, url As String, try As Integer
    'for chrome info:
    'https://chromedriver.chromium.org/downloads/version-selection
    
    'for edge info:
    'https://blogs.windows.com/msedgedev/2019/11/04/edge-chromium-release-candidate-get-ready/
    'https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
    
    Select Case browser
    Case svbaBrowser.Edge
        
        vermajor = Split(browserVersion, ".")(0)

        'search for major version, decrementing by one until success
        vermajorsearch = vermajor
        Do While vermajorsearch >= 75
            url = "https://msedgedriver.azureedge.net/LATEST_RELEASE_" & vermajorsearch
            
            'the following loop is to deal with "-2147012739 - An error occurred in the secure channel support" for msedge driver LATEST_RELEASE_XX api
            'this usually gets a connection in less than 4 tries
            Do
                responseText = RequestData(url)
                try = try + 1
                'Debug.Print "try: " & try
                DoEvents
                If try > 10 Then
                    MsgBox "could not connect to msedgedriver server"
                    GetCompatibleDriverVersion = False
                    Exit Function
                End If
            Loop Until responseText <> False
            
            GetCompatibleDriverVersion = Replace(responseText, vbCrLf, "")
            
            If InStr(GetCompatibleDriverVersion, "Error") = 0 Then Exit Do
            vermajorsearch = vermajorsearch - 1
        Loop

    Case svbaBrowser.Chrome
    
        vermajor = Split(browserVersion, ".")(0)
        verminor = Split(browserVersion, ".")(1)
        verbuild = Split(browserVersion, ".")(2)
    
        versearch = vermajor & "." & verminor & "." & verbuild
        
        'first see if we can find an exact match
        url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & versearch
        
        responseText = RequestData(url)
        
        GetCompatibleDriverVersion = responseText
        
        If InStr(GetCompatibleDriverVersion, "Error") = 0 Then Exit Function
        
        'if exact match failed, then maybe we have a Dev or Canary channel
        'so search for major version match, decrementing by one until success
        vermajorsearch = vermajor
        Do While vermajorsearch >= 70
            url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & vermajorsearch
            responseText = RequestData(url)
            
            GetCompatibleDriverVersion = responseText
            
            If InStr(GetCompatibleDriverVersion, "Error") = 0 Then Exit Do
            vermajorsearch = vermajorsearch - 1
        Loop
        
    Case svbaBrowser.Firefox
        
        vermajor = Split(browserVersion, ".")(0)
        verminor = Split(browserVersion, ".")(1)
        
        'the rules are different for Firefox driver - we must lookup the browser version in a compatibility matrix
        'to get the most recent driver version compatible with the installed browser
        url = "https://firefox-source-docs.mozilla.org/testing/geckodriver/Support.html"
        responseText = RequestData(url)
        If responseText = False Then
            MsgBox "Could not access Firefox compatibility table from Mozilla", , "SeleniumVBA"
            Exit Function
        End If
        
        'search the table for the first driver version entry (most recent) for a given browser version
        htmldoc.body.innerHTML = responseText
        Set table = htmldoc.getElementsByTagName("table")(0)
        For Each row In table.Rows
            If row.getElementsByTagName("td").length > 0 Then
                Set cells = row.getElementsByTagName("td")
                If CInt(vermajor) >= CInt(Split(cells(2).outerText)(0)) Then
                    GetCompatibleDriverVersion = Trim(cells(0).outerText)
                    Exit For
                End If
            End If
        Next row

    End Select
End Function

Public Function GetInstalledBrowserVersion(ByVal browser As svbaBrowser) As String
    Dim wsh As New IWshRuntimeLibrary.wshShell 'CreateObject("WScript.Shell")
    Dim regkey As String, resp As String
    
    Select Case browser
    Case svbaBrowser.Chrome
        regkey = "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon\version"
    Case svbaBrowser.Edge
        regkey = "HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon\version"
    Case svbaBrowser.Firefox
        regkey = "HKEY_LOCAL_MACHINE\SOFTWARE\Mozilla\Mozilla Firefox\CurrentVersion"
    End Select
    
    'this throws error if key doesnt exist - report back to user if browser isn't installed
    On Error Resume Next
    resp = wsh.RegRead(regkey)
    If Err.Number <> 0 Then
        resp = "browser not installed"
    End If

    GetInstalledBrowserVersion = resp
End Function

Public Sub DownloadAndInstallDriver(ByVal browser As svbaBrowser, ByVal driverVersion As String, Optional ByVal driverPath)
    Dim url As String
    Dim BinaryStream As New ADODB.stream 'deals with binary - FileSystemObject does not
    Dim fso As New IWshRuntimeLibrary.FileSystemObject
    Dim shll As New Shell32.Shell 'Microsoft Shell Controls and Automation (Shell32) 'CreateObject("Shell.Application")
    Dim driverFolder As String, driverFilename As String, tmpInstallFolder As String, createdTmpFolder As Boolean
    Dim tmpInstallZipPath As String, tmpZipDriverFilename As String, tmpInstallDriverpath As String, responseBody As Variant
    
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
    
    driverPath = GetAbsolutePath(driverPath)
    
    driverFolder = fso.GetParentFolderName(driverPath)
    driverFilename = fso.GetFileName(driverPath)
    
    tmpInstallFolder = driverFolder & "\tmp\"
    
    'if tmp folder doesn't exist then create it - remember for clean up
    createdTmpFolder = False
    If Not fso.FolderExists(tmpInstallFolder) Then
        createdTmpFolder = True
        fso.CreateFolder tmpInstallFolder
    End If
    
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
    End Select
    
    responseBody = RequestData(url, "responseBody")
        
    With BinaryStream
        .Open
        .Type = adTypeBinary
        .Write responseBody
        .SaveToFile tmpInstallZipPath, adSaveCreateOverWrite
        .Close
    End With
    
    If fso.FileExists(tmpInstallDriverpath) Then fso.DeleteFile tmpInstallDriverpath, True

    shll.Namespace(tmpInstallFolder).CopyHere shll.Namespace(tmpInstallZipPath).Items.Item(tmpZipDriverFilename)

    'check if driver exists and if it does, make sure it isn't in use before overwrite
    If fso.FileExists(driverPath) Then
        If IsProcessRunning(driverFilename) Then TaskKillByImage (driverFilename)
    End If

    fso.CopyFile tmpInstallDriverpath, driverPath, True
    
    'if created tmp folder, then delete it
    If createdTmpFolder Then
        fso.DeleteFolder Left(tmpInstallFolder, Len(tmpInstallFolder) - 1), True
    End If
End Sub

Public Function GetSeleniumBasicFolder()
    'this may be of use for updating drivers for SeleniumBasic
    GetSeleniumBasicFolder = Environ$("LocalAppData") & "\SeleniumBasic\"
End Function

Public Function GetUserTempFolder()
    GetUserTempFolder = Environ$("LocalAppData") & "\Temp\"
End Function

Private Function IsProcessRunning(ByVal process As String)
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

Private Function RequestData(ByVal url, Optional ByVal returnType = "responseText") As Variant
    Dim client As New MSXML2.ServerXMLHTTP60
    
    On Error GoTo HaveError
    client.Open "GET", url, False
    
    client.send
    
    While client.readyState <> 4
        DoEvents
    Wend

    If client.Status = 200 Then
        Select Case LCase(returnType)
        Case "responsetext"
            RequestData = client.responseText
        Case "responsebody"
            RequestData = client.responseBody
        End Select
    Else
        Debug.Print client.Status
    End If
    
    Exit Function
    
HaveError:
    'Debug.Print Err.Number & " - " & Err.Description, vbCritical
    RequestData = False
End Function

Private Function BrowserNameStr(ByVal val As svbaBrowser)
    Select Case val
    Case svbaBrowser.Edge
        BrowserNameStr = "Edge"
    Case svbaBrowser.Chrome
        BrowserNameStr = "Chrome"
    Case svbaBrowser.Firefox
        BrowserNameStr = "Firefox"
    End Select
End Function

Private Sub Class_Initialize()
    mMinCompatibilityLevel = svbaMajor
End Sub

Private Function TaskKillByImage(ByVal taskName As String)
    Dim wsh As New IWshRuntimeLibrary.wshShell
    TaskKillByImage = wsh.Run("taskkill /f /t /im " & taskName, 0, True)
End Function
