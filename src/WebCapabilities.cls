VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebCapabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================
' SeleniumVBA v2.3.1
' A Selenium wrapper for Edge and Chrome written in Windows VBA based on JSon wire protocol.
'
' (c) GCUser99
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' ==========================================================================
'
' MODIFIED/EXTENDED BY GCUser99 FROM:
'
' TinySeleniumVBA v0.1.3
' A tiny Selenium wrapper written in pure VBA
'
' (c)2021 uezo
'
' Mail: uezo@uezo.net
' Twitter: @uezochan
' https://github.com/uezo/TinySeleniumVBA
'
' ==========================================================================
' MIT License
'
' Copyright (c) 2021 uezo
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

' Specs for Capabilities:
'
' - Chrome: https://chromedriver.chromium.org/capabilities
' - Edge: https://docs.microsoft.com/en-us/microsoft-edge/webdriver-chromium/capabilities-edge-options
'
' Notes:
'
' - WebCapabilities object should be instantiated using WebDriver.CreateCapabilities
' - It should be instantiated after the StartEdge, StartChrome, or StartFirefox methods,
' - and passed to OpenBrowser method, as in following example...
'
' Example:
'
'   Dim caps as WebCapabilities, driver as new WebDriver
'
'   driver.StartEdge
'   Set caps = driver.MakeCapabilities
'   caps.RunInvisble
'   driver.OpenBrowser caps
'
' The OpenBrowser method of WebDriver class has optional input parameters that can be used to avoid directly
'   instantiating the Capabilities object:
'
'   invisible (Boolean): If set to True (default=False), adds the --headless argument to the default Capabilities object
'   incognito (Boolean): If set to True (default=False), adds the --Incognito argument to the default Capabilities object
'   capabilitiesFilePath (String): If specified, initializes the default Capabilities object with a previously saved json file
'   (see: SaveToFile method of the WebCapabilities class, and example in test_Capabilities module)
'

Option Explicit

Private data_ As New Dictionary
Private browserOptionKey As String
Private browser_ As svbaBrowser
Private driver_ As WebDriver

Public Enum svbaAlertHandling
    svbaAccept
    svbaDismiss
    svbaIgnore
    svbaDismissAndNotify
    svbaAcceptAndNotify
End Enum

'-------------------------------------------------------------------------
' Capability Settings
'-------------------------------------------------------------------------

Public Sub SetCapability(ByVal key As String, val As Variant)
    'set a generic capability
    If IsObject(val) Then
        Set data_("alwaysMatch")(key) = val
    Else
        If Left(val, 1) = "[" Or Left(val, 1) = "{" Then 'its a json string - try to parse
            Dim jc As New WebJSonConverter
            Set data_("alwaysMatch")(key) = jc.ParseJSON(val)
        Else 'its a string parameter
            data_("alwaysMatch")(key) = val
        End If
    End If
End Sub

Public Sub SetAcceptInsecureCerts(ByVal val As Boolean)
    Me.SetCapability "acceptInsecureCerts", val
End Sub

Public Sub SetUnhandledPromptBehavior(ByVal val As svbaAlertHandling)
    Me.SetCapability "unhandledPromptBehavior", GetAlertHandlingString(val)
End Sub

'-------------------------------------------------------------------------
' Option Settings
'-------------------------------------------------------------------------

Public Sub SetOption(ByVal key As String, val As Variant)
    'set an option
    'input val can be a string, a json string, array, or an object (dictionary or collection)
    If IsObject(val) Then 'collection or dictionary
        Set data_("alwaysMatch")(browserOptionKey)(key) = val
    Else
        If IsArray(val) Then
            data_("alwaysMatch")(browserOptionKey)(key) = val
        Else 'string
            If Left(val, 1) = "[" Or Left(val, 1) = "{" Then 'its a json string - try to parse
                Dim jc As New WebJSonConverter
                Set data_("alwaysMatch")(browserOptionKey)(key) = jc.ParseJSON(val)
            Else 'its a string parameter
                data_("alwaysMatch")(browserOptionKey)(key) = val
            End If
        End If
    End If
End Sub

Public Sub AddExcludeSwitches(ParamArray addList() As Variant)
    Dim varry() As Variant
    varry = addList
    AddToOptionList "excludeSwitches", ParamListToStringArray(varry)
End Sub

Public Sub AddWindowTypes(ParamArray addList() As Variant)
    Dim varry() As Variant
    varry = addList
    AddToOptionList "windowTypes", ParamListToStringArray(varry)
End Sub

Public Sub AddExtensions(ParamArray addList() As Variant)
    Dim varry() As Variant
    varry = addList
    AddToOptionList "extensions", ParamListToStringArray(varry)
End Sub

Public Sub SetBrowserBinary(ByVal browserPath As String)
    'set the path to the browser executable to use
    Me.SetOption "binary", browserPath
End Sub

Public Sub RemoveControlNotification()
    'this removes the "Chrome/Edge is being controlled by automated test software" browser notification
    Select Case browser_
    Case svbaBrowser.Chrome
        Me.AddExcludeSwitches "enable-automation"
        Me.SetOption "useAutomationExtension", False
    Case svbaBrowser.Edge
        Me.AddExcludeSwitches "enable-automation"
        Me.SetOption "useAutomationExtension", False
        'for msedge, you may be greeted with an unwanted "Sync your profile" dialog
        'to get rid of that, run in incognito mode
    Case svbaBrowser.Firefox
        'no need to do anything
    End Select
End Sub

Public Sub SetDetachBrowser(ByVal val As Boolean)
    'this sets whether browser is closed (false) or left open (true) when the driver is sent the shutdown command
    If browser_ = svbaBrowser.Firefox Then
        Exit Sub 'as yet have not found a working equivalent - SetPreference("detach", True) does not work
    End If
    Me.SetOption "detach", val
End Sub

'-------------------------------------------------------------------------
' Option Argument Settings
'-------------------------------------------------------------------------

Public Sub AddArguments(ParamArray addList() As Variant)
    Dim varry() As Variant
    varry = addList
    AddToOptionList "args", ParamListToStringArray(varry)
End Sub

Public Sub RunInvisible()
    Me.AddArguments "--headless"
End Sub

Public Sub RunIncognito()
    Select Case browser_
    Case svbaBrowser.Chrome
        Me.AddArguments "--incognito"
    Case svbaBrowser.Edge
        Me.AddArguments "--inprivate"
    Case svbaBrowser.Firefox
        Me.AddArguments "-private"
    End Select
End Sub

Public Sub SetProfile(ByVal profilePath As String)
    'this sets a user-defined profile - it will create and populate the profile directory even if it does not exist
    'it is recommended not to make profiles for SeleniumVBA in the default C:\Users\[user]\AppData\Local\Google\Chrome\User Data
    'Instead copy profiles needed from the default location into a different location to access with SeleniumVBA
    Dim fso As New FileSystemObject
    Dim userDir As String, profileName As String
    
    profilePath = WebShared.GetFullLocalPath(profilePath)
    
    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        userDir = fso.GetParentFolderName(profilePath)
        profileName = fso.GetBaseName(profilePath)
        Me.AddArguments "--user-data-dir=" & userDir
        Me.AddArguments "--profile-directory=" & profileName
    Case svbaBrowser.Firefox
        'unlike edge and chrome, firefox requires that the folder to exist
        'hence check and if doesn't exist, then create it
        If Not fso.FolderExists(profilePath) Then CreateFolderPath profilePath
        Me.AddArguments "-profile", profilePath
    End Select
End Sub

Public Sub SetUserAgent(ByVal val As String)
    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        Me.AddArguments "--user-agent=" & val
    Case svbaBrowser.Firefox
        Me.SetPreference "general.useragent.override", val
    End Select
End Sub

'-------------------------------------------------------------------------
' Option Preference Settings
'-------------------------------------------------------------------------

Public Sub SetPreference(ByVal key As String, ByVal val As Variant)
    'set a generic preference
    'input val can be a string, a json string, or an object (dictionary)
    If Not data_("alwaysMatch")(browserOptionKey).Exists("prefs") Then
        SetOption "prefs", New Dictionary
    End If
    
    'need to come up with a more generalized approach to paths - this is just a band-aid
    If key = "download.default_directory" Then val = WebShared.GetFullLocalPath(val, driver_.DefaultIOFolder)
    If key = "browser.download.dir" Then val = WebShared.GetFullLocalPath(val, driver_.DefaultIOFolder)
    If key = "savefile.default_directory" Then val = WebShared.GetFullLocalPath(val, driver_.DefaultIOFolder)
    
    If IsObject(val) Then
        Set data_("alwaysMatch")(browserOptionKey)("prefs")(key) = val
    Else
        If Left(val, 1) = "[" Or Left(val, 1) = "{" Then 'it's json string - try to parse
            'convert to dictionary or collection
            Dim jc As New WebJSonConverter
            Set data_("alwaysMatch")(browserOptionKey)("prefs")(key) = jc.ParseJSON(val)
        Else
            data_("alwaysMatch")(browserOptionKey)("prefs")(key) = val
        End If
    End If
End Sub

Public Sub SetDownloadPrefs(Optional ByVal downloadDirectoryPath As String = ".\")
    downloadDirectoryPath = WebShared.GetFullLocalPath(downloadDirectoryPath, driver_.DefaultIOFolder)

    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        Me.SetPreference "download.default_directory", downloadDirectoryPath
        Me.SetPreference "download.directory_upgrade", True
        Me.SetPreference "download.prompt_for_download", False 'false disables prompting user for download permission
        Me.SetPreference "safebrowsing.enabled", True
        Me.SetPreference "plugins.always_open_pdf_externally", True 'instead of using the pdf viewer
        'Me.SetPreference "safebrowsing.disable_download_protection", True
        'Me.SetPreference "profile.default_content_settings.popups", False '0
        'Me.SetPreference "profile.default_content_settings.javascript", True '1
        'Me.SetPreference "profile.default_content_setting_values.notifications", 2
        'Me.SetPreference "profile.content_settings.exceptions.automatic_downloads.*.setting", True '1
        'Me.SetPreference "profile.default_content_setting_values.automatic_downloads", True '1
    Case svbaBrowser.Firefox
        Me.SetPreference "browser.download.folderList", 2 '0 is Desktop, 1 is Downloads
        Me.SetPreference "browser.download.dir", downloadDirectoryPath
        Me.SetPreference "pdfjs.disabled", True  'disable the built-in PDF viewer
        'Me.SetPreference "browser.download.manager.showWhenStarting", False
        'Me.SetPreference "browser.helperApps.neverAsk.saveToDisk", "application/pdf;text/plain;text/csv;application/text;text/xml;application/xml;application/csv;pplication/octet-stream doc xls pdf txt"
        'Me.SetPreference "browser.download.useDownloadDir", True
        'Me.SetPreference "browser.download.viewableInternally.enabledTypes", ""
    End Select
End Sub

'-------------------------------------------------------------------------
' Public helper methods
'-------------------------------------------------------------------------

'convert to Dictionary (used in WebDriver class)
Public Function Dictionary() As Dictionary
Attribute Dictionary.VB_UserMemId = 0
    Set Dictionary = data_
End Function

'convert to JSON string for debugging, also used in SaveToFile method
Public Function ToJson() As String
    Dim jc As New WebJSonConverter
    ToJson = jc.ConvertToJson(data_, 4)
End Function

Public Sub SaveToFile(Optional ByVal filePath As String = ".\capabilities.json")
    Dim fso As New FileSystemObject, ts As textstream
    Dim jc As New WebJSonConverter
    filePath = WebShared.GetFullLocalPath(filePath, driver_.DefaultIOFolder)
    Set ts = fso.CreateTextFile(filePath, True, True)
    ts.Write Me.ToJson
    ts.Close
End Sub

Public Sub LoadFromFile(Optional ByVal filePath As String = ".\capabilities.json")
    Dim fso As New FileSystemObject, ts As textstream
    Dim jc As New WebJSonConverter
        
    filePath = WebShared.GetFullLocalPath(filePath, driver_.DefaultIOFolder)
    
    If Not fso.FileExists(filePath) Then Err.raise 1, "WebCapabilities", "Capabilities Json file does not exist!"
    'note that this will fail if input file format is utf-8 with non-ascii chars - see https://github.com/GCuser99/SeleniumVBA/discussions/32
    Set ts = fso.OpenTextFile(filePath, , , TristateUseDefault)
    Set data_ = jc.ParseJSON(ts.ReadAll)
    ts.Close
    
    If data_("browserName") <> WebShared.GetBrowserName(browser_) Then Err.raise 1, "WebCapabilities", "Browser name from file does not match browser name assigned to driver instance."
End Sub

Public Sub InitializeForChrome()
    InitializeFor svbaBrowser.Chrome
End Sub

Public Sub InitializeForEdge()
    InitializeFor svbaBrowser.Edge
End Sub

Public Sub InitializeForFirefox()
    InitializeFor svbaBrowser.Firefox
End Sub

'-------------------------------------------------------------------------
' Private helper methods
'-------------------------------------------------------------------------
Friend Sub InitializeFor(ByVal browser As svbaBrowser)
    Set data_ = New Dictionary
    data_.Add "browserName", WebShared.GetBrowserName(browser)
    data_.Add "alwaysMatch", New Dictionary
    browser_ = browser
    
    Dim browserOptions As New Dictionary
    
    Select Case browser
        Case svbaBrowser.Chrome
            browserOptionKey = "goog:chromeOptions"
        Case svbaBrowser.Edge
            browserOptionKey = "ms:edgeOptions"
        Case svbaBrowser.Firefox
            browserOptionKey = "moz:firefoxOptions"
    End Select
    
    data_("alwaysMatch").Add browserOptionKey, browserOptions
End Sub

Friend Property Set WebDriver(ByRef obj As WebDriver)
    Set driver_ = obj
End Property

Private Function GetAlertHandlingString(ByVal val As svbaAlertHandling) As String
    Select Case val
    Case svbaAlertHandling.svbaAccept
        GetAlertHandlingString = "accept"
    Case svbaAlertHandling.svbaAcceptAndNotify
        GetAlertHandlingString = "accept and notify"
    Case svbaAlertHandling.svbaDismiss
        GetAlertHandlingString = "dismiss"
    Case svbaAlertHandling.svbaDismissAndNotify
        GetAlertHandlingString = "dismiss and notify"
    Case svbaAlertHandling.svbaIgnore
        GetAlertHandlingString = "ignore"
    End Select
End Function

Private Sub AddToOptionList(ByVal optionName As String, addList() As String)
    'add to list of option parameters, while checking for duplicates
    Dim parameters() As String, found As Boolean, j As Integer, i As Integer
    Dim lboundParms As Integer, uboundParms As Integer

    If IsArray(data_("alwaysMatch")(browserOptionKey)(optionName)) Then
        parameters = data_("alwaysMatch")(browserOptionKey)(optionName)
        lboundParms = LBound(parameters)
        uboundParms = UBound(parameters)
    Else 'starting from scratch
        lboundParms = 1 '1 for 1-based, 0 for 0-based
        uboundParms = lboundParms - 1
    End If

    For i = LBound(addList) To UBound(addList)
        found = False
        For j = lboundParms To uboundParms
            If addList(i) = parameters(j) Then found = True: Exit For
        Next j
        If Not found Then
            uboundParms = uboundParms + 1
            ReDim Preserve parameters(lboundParms To uboundParms)
            parameters(uboundParms) = addList(i)
        End If
    Next i
    'put the expanded list back
    data_("alwaysMatch")(browserOptionKey)(optionName) = parameters
End Sub

Private Function ParamListToStringArray(parmList() As Variant) As String()
    'a parameter array can include arrays as list elements - expand into a string array
    Dim i As Integer, j As Integer, k As Integer
    Dim strlist() As String
    Dim temp As Variant
    
    'count number of parameters in list
    k = 0
    For i = LBound(parmList) To UBound(parmList)
        temp = parmList(i)
        If IsArray(parmList(i)) Then
            k = k + UBound(parmList(i)) - LBound(parmList(i)) + 1
        Else
            If Left(parmList(i), 1) = "[" Then  'its a json string - try to parse
                Dim jc As New WebJSonConverter
                k = k + jc.ParseJSON(parmList(i)).Count
            Else 'its a string parameter
                k = k + 1
            End If
        End If
    Next i
    
    'size the output array
    ReDim strlist(1 To k)
    
    'now store the parameters from list into string array
    k = 0
    For i = LBound(parmList) To UBound(parmList)
        If IsArray(parmList(i)) Then
            For j = LBound(parmList(i)) To UBound(parmList(i))
                k = k + 1
                strlist(k) = parmList(i)(j)
            Next j
        Else
            If Left(parmList(i), 1) = "[" Then  'its a json string
                Dim col As Collection
                Set col = jc.ParseJSON(parmList(i))
                For j = 1 To col.Count
                    k = k + 1
                    strlist(k) = col(j)
                Next j
            Else 'its a string parameter
                k = k + 1
                strlist(k) = parmList(i)
            End If
        End If
    Next i
    
    ParamListToStringArray = strlist
End Function

'https://stackoverflow.com/questions/31033820/filesystemobject-createfolder-to-create-directory-and-subdirectories
Private Function CreateFolderPath(ByVal folderPath As String) As Boolean
    Static fso As FileSystemObject
 
    'Initialize fso variable if not already setup
    If fso Is Nothing Then Set fso = New FileSystemObject

    'Is the folderPath parameter populated
    If Len(folderPath) = 0 Then
      CreateFolderPath = False
      Exit Function
    End If

    'If the folderPath exists as a file, the function fails.
    If fso.FileExists(folderPath) Then
        CreateFolderPath = False
        Exit Function
    End If
 
    'If the folderPath already exists as a folder, don't do anything and return success.
    If fso.FolderExists(folderPath) Then
        CreateFolderPath = True
        Exit Function
    End If
 
    'recursively create the parent folder, then if successful create the top folder.
    If CreateFolderPath(fso.GetParentFolderName(folderPath)) Then
        If fso.CreateFolder(folderPath) Is Nothing Then
            CreateFolderPath = False
        Else
           CreateFolderPath = True
        End If
    Else
        CreateFolderPath = False
    End If
End Function
