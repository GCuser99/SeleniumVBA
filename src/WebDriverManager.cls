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
' SeleniumVBA v1.1
' A Selenium wrapper for Edge and Chrome written in Windows VBA based on JSon wire protocol.
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


'This class was written by GCUser99 for use in Selenium VBA projects to manage webdriver-browser version compatibility

'References:
'Windows Script Host Model
'Microsoft ActiveX Data Objects 6.1 Library
'Microsoft Shell Controls and Automation
'Microsoft XML, V6.0
'Microsoft WMI Scripting V1.2 Library

Private mSupressWarnings As Boolean

Public Enum svbaCompatibility
    svbaNotCompatible = 0
    svbaMajor = 1
    svbaMinor = 2
    svbaBuildMajor = 3
    svbaBuildMinor = 4
End Enum

Property Let supressWarnings(ByVal val As Boolean)
    mSupressWarnings = val
End Property

Property Get supressWarnings() As Boolean
    supressWarnings = mSupressWarnings
End Property

Public Function IsInstalledDriverCompatible(ByVal browserName As String, Optional ByVal driverPath, Optional ByVal compatibilityLevel As svbaCompatibility = svbaMajor) As Boolean
    Dim fso As New IWshRuntimeLibrary.FileSystemObject, clevel As svbaCompatibility
    
    browserName = LCase(browserName)
    If IsMissing(driverPath) Then
        driverPath = ThisWorkbook.Path & "\" & browserName & "driver.exe"
    Else
        driverPath = GetAbsolutePath(driverPath)
    End If
    
    If Not fso.FileExists(driverPath) Then
        If Not mSupressWarnings Then MsgBox "driver not installed"
        IsInstalledDriverCompatible = False
        Exit Function
    End If
    
    bverInstalled = Me.GetInstalledBrowserVersion(browserName)
    If bverInstalled = "browser not installed" Then
        If Not mSupressWarnings Then MsgBox "browser not installed"
        IsInstalledDriverCompatible = False
        Exit Function
    End If
    
    dverinstalled = Me.GetInstalledDriverVersion(browserName, driverPath)
    dverCompat = Me.GetCompatibleDriverVersion(browserName, bverInstalled)
    clevel = Me.CheckCompatibilityLevel(dverinstalled, dverCompat)
    
    If clevel < compatibilityLevel Then
        IsInstalledDriverCompatible = False
    Else
        IsInstalledDriverCompatible = True
    End If

End Function

Public Function GetInstalledDriverVersion(ByVal browserName As String, Optional ByVal driverPath) As String
    Dim wsh As New IWshRuntimeLibrary.WshShell 'CreateObject("WScript.Shell")
    Dim fso As New IWshRuntimeLibrary.FileSystemObject

    browserName = LCase(browserName)

    If IsMissing(driverPath) Then
        driverPath = ThisWorkbook.Path & "\" & browserName & "driver.exe"
    Else
        driverPath = GetAbsolutePath(driverPath)
    End If
    
    If Not fso.FileExists(driverPath) Then
        If Not mSupressWarnings Then MsgBox "driver not installed"
        GetInstalledDriverVersion = "driver not found"
        Exit Function
    End If
    
    resp = wsh.Exec(Chr(34) & driverPath & Chr(34) & " --version").StdOut.ReadAll
    GetInstalledDriverVersion = Split(resp, " ")(1)

End Function

Public Function CheckCompatibilityLevel(ByVal installedVersion As String, ByVal compatibleVersion As String) As svbaCompatibility
    
    If installedVersion = "driver not found" Then CheckComptibilityLevel = -1: Exit Function
    vermajor1 = Split(installedVersion, ".")(0)
    verminor1 = Split(installedVersion, ".")(1)
    verbuildmajor1 = Split(installedVersion, ".")(2)
    verbuildminor1 = Split(installedVersion, ".")(3)
    
    vermajor2 = Split(compatibleVersion, ".")(0)
    verminor2 = Split(compatibleVersion, ".")(1)
    verbuildmajor2 = Split(compatibleVersion, ".")(2)
    verbuildminor2 = Split(compatibleVersion, ".")(3)
    
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

Public Function GetLatestDriverVersion(ByVal browserName As String) As String
    'this function is here for completeness, although not used
    
    browserName = LCase(browserName)
    
    Select Case browserName
    Case "msedge"
        url = "https://msedgedriver.azureedge.net/LATEST_STABLE"
        'the following loop is to deal with "-2147012739 - An error occurred in the secure channel support" for msedge driver LATEST_RELEASE_XX api
        'this usually gets a connection in less than 4 tries
        Do
            responseText = RequestData(url)
            try = try + 1
            'Debug.Print "try: " & try
            DoEvents
            If try > 10 Then
                If Not mSupressWarnings Then MsgBox "could not connect to msedgedriver server"
                GetLatestDriverVersion = False
                Exit Function
            End If
        Loop Until responseText <> False
        GetLatestDriverVersion = Replace(responseText, vbCrLf, "")
    Case "chrome"
        url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
        responseText = RequestData(url)
        GetLatestDriverVersion = responseText
    End Select

End Function

Public Function GetCompatibleDriverVersion(ByVal browserName As String, ByVal browserVersion) As String

    'for chrome info:
    'https://chromedriver.chromium.org/downloads/version-selection
    
    'for edge info:
    'https://blogs.windows.com/msedgedev/2019/11/04/edge-chromium-release-candidate-get-ready/
    'https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
    
    browserName = LCase(browserName)
    
    Select Case browserName
    Case "msedge"
        
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
                    If Not mSupressWarnings Then MsgBox "could not connect to msedgedriver server"
                    GetCompatibleDriverVersion = False
                    Exit Function
                End If
            Loop Until responseText <> False
            
            GetCompatibleDriverVersion = Replace(responseText, vbCrLf, "")
            
            If InStr(GetCompatibleDriverVersion, "Error") = 0 Then Exit Do
            vermajorsearch = vermajorsearch - 1
        Loop

    Case "chrome"
    
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

    End Select

End Function

Public Function GetInstalledBrowserVersion(ByVal browserName As String) As String
    Dim wsh As New IWshRuntimeLibrary.WshShell 'CreateObject("WScript.Shell")
    
    browserName = LCase(browserName)
    
    Select Case browserName
    Case "chrome"
        regkey = "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon\version"
    Case "msedge"
        regkey = "HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon\version"
    End Select
    
    'this throws error if key doesnt exist - report back to user if browser isn't installed
    On Error Resume Next
    resp = wsh.RegRead(regkey)
    If Err.Number <> 0 Then
        resp = "browser not installed"
    End If

    GetInstalledBrowserVersion = resp
    
End Function

Public Sub DownloadAndInstallDriver(ByVal browserName As String, ByVal driverVersion As String, Optional ByVal driverPath)
    Dim url As String
    Dim BinaryStream As New ADODB.stream 'deals with binary - FileSystemObject does not
    Dim fso As New IWshRuntimeLibrary.FileSystemObject
    Dim shll As New Shell32.shell 'Microsoft Shell Controls and Automation (Shell32) 'CreateObject("Shell.Application")
    
    browserName = LCase(browserName)
    
    If IsMissing(driverPath) Then
        driverPath = ThisWorkbook.Path & "\" & browserName & "driver.exe"
    Else
        driverPath = GetAbsolutePath(driverPath)
    End If
    
    driverFolder = fso.GetParentFolderName(driverPath)
    driverFileName = fso.GetFileName(driverPath)
    
    tmpInstallFolder = driverFolder & "\tmp\"
    
    'if tmp folder doesn't exist then create it - remember for clean up
    createdTmpFolder = False
    If Not fso.FolderExists(tmpInstallFolder) Then
        createdTmpFolder = True
        fso.CreateFolder tmpInstallFolder
    End If
    
    Select Case browserName
    Case "chrome"
        url = "https://chromedriver.storage.googleapis.com/" & driverVersion & "/chromedriver_win32.zip"
        tmpInstallZipPath = tmpInstallFolder & "chromedriver_win32.zip"
        tmpZipDriverFilename = "chromedriver.exe"
        tmpInstallDriverpath = tmpInstallFolder & tmpZipDriverFilename
    Case "msedge"
        'test for 64 bit vs 32 bit
        If InStr(Application.OperatingSystem, "64-bit") Then
            url = "https://msedgedriver.azureedge.net/" & driverVersion & "/edgedriver_win64.zip"
            tmpInstallZipPath = tmpInstallFolder & "edgedriver_win64.zip"
        Else
            url = "https://msedgedriver.azureedge.net/" & driverVersion & "/edgedriver_win32.zip"
            tmpInstallZipPath = tmpInstallFolder & "edgedriver_win32.zip"
        End If
        tmpZipDriverFilename = "msedgedriver.exe"
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
        If IsProcessRunning(driverFileName) Then TaskKill driverFileName
    End If

    fso.CopyFile tmpInstallDriverpath, driverPath, True
    
    'if created tmp folder, then delete it
    If createdTmpFolder Then
        fso.DeleteFolder Left(tmpInstallFolder, Len(tmpInstallFolder) - 1), True
    End If
    
End Sub

Public Sub AlignDriverAndBrowser(ByVal browserName As String, ByVal driverPath As String, Optional ByVal supressWarnings As Boolean = False, Optional ByVal compatibilityLevel As svbaCompatibility = svbaMajor)
    mSupressWarnings = supressWarnings
    If Not IsInstalledDriverCompatible(browserName, driverPath, compatibilityLevel) Then
        bverInstalled = GetInstalledBrowserVersion(browserName)
        dverCompat = GetCompatibleDriverVersion(browserName, bverInstalled)
        DownloadAndInstallDriver browserName, dverCompat, driverPath
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

Private Function TaskKill(ByVal taskName As String)
    Dim wsh As New IWshRuntimeLibrary.WshShell 'CreateObject("WScript.Shell")
    TaskKill = wsh.Run("taskkill /f /im " & taskName, 0, True)
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
        Case Else
            'Stop
        End Select
    Else
        'Debug.Print client.Status
    End If
    
    Exit Function
    
HaveError:
    'Debug.Print Err.Number & " - " & Err.Description, vbCritical
    RequestData = False
End Function

Private Function GetAbsolutePath(ByVal strPath As String) As String
    Dim fso As New IWshRuntimeLibrary.FileSystemObject
    saveppath = CurDir()
    ChDrive ThisWorkbook.Path
    ChDir ThisWorkbook.Path
    GetAbsolutePath = fso.GetAbsolutePathName(strPath)
    ChDrive saveppath
    ChDir saveppath
End Function

Private Sub Class_Initialize()
    mSupressWarnings = False
End Sub
