VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebCapabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used to manage/set Selenium optional Capabilities - must be instantiated using WebDriver's CreateCapabilities method"
'@ModuleDescription "This class is used to manage/set Selenium optional Capabilities - must be instantiated using WebDriver's CreateCapabilities method"
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
'
' Based originally on uezo's TinySeleniumVBA with extensive changes/additions
' ==========================================================================
' MIT License
'
' Copyright (c) 2023-2024, GCUser99 and 6DiegoDiego9 (https://github.com/GCuser99/SeleniumVBA)
' Copyright (c) 2021, uezo (https://github.com/uezo/TinySeleniumVBA)
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
    svbaAccept = 0
    svbaDismiss = 1
    svbaIgnore = 2
    svbaDismissAndNotify = 3
    svbaAcceptAndNotify = 4
End Enum

Public Enum svbaPageLoadStrategy
    svbaNormal = 0
    svbaEager = 1
    svbaNone = 2
End Enum

'-------------------------------------------------------------------------
' Capability Settings
'-------------------------------------------------------------------------

'@Description("Sets a generic capability")
Public Sub SetCapability(ByVal key As String, val As Variant)
Attribute SetCapability.VB_Description = "Sets a generic capability"
    If IsObject(val) Then
        Set data_.Item("alwaysMatch")(key) = val
    Else
        If Left$(val, 1) = "[" Or Left$(val, 1) = "{" Then 'its a json string - try to parse
            Set data_.Item("alwaysMatch")(key) = WebJsonConverter.ParseJson(val)
        Else 'its a string parameter
            data_.Item("alwaysMatch")(key) = val
        End If
    End If
End Sub

'@Description("Sets the acceptInsecureCerts capability")
Public Sub SetAcceptInsecureCerts(ByVal val As Boolean)
Attribute SetAcceptInsecureCerts.VB_Description = "Sets the acceptInsecureCerts capability"
    Me.SetCapability "acceptInsecureCerts", val
End Sub

'@Description("Sets the pageLoadStrategy capability - defaults to normal")
Public Sub SetPageLoadStrategy(ByVal val As svbaPageLoadStrategy)
Attribute SetPageLoadStrategy.VB_Description = "Sets the pageLoadStrategy capability - defaults to normal"
    Select Case val
    Case svbaPageLoadStrategy.svbaNormal
        Me.SetCapability "pageLoadStrategy", "normal"
    Case svbaPageLoadStrategy.svbaEager
        Me.SetCapability "pageLoadStrategy", "eager"
    Case svbaPageLoadStrategy.svbaNone
        Me.SetCapability "pageLoadStrategy", "none"
    End Select
End Sub

'@Description("Sets the unhandledPromptBehavior capability - defaults to dismiss and notify")
Public Sub SetUnhandledPromptBehavior(ByVal val As svbaAlertHandling)
Attribute SetUnhandledPromptBehavior.VB_Description = "Sets the unhandledPromptBehavior capability - defaults to dismiss and notify"
    Me.SetCapability "unhandledPromptBehavior", getAlertHandlingString(val)
End Sub

'-------------------------------------------------------------------------
' Option Settings
'-------------------------------------------------------------------------

'@Description("Sets an option key/value")
Public Sub SetOption(ByVal key As String, val As Variant)
Attribute SetOption.VB_Description = "Sets an option key/value"
    'input val can be a string, a json string, array, or an object (dictionary or collection)
    If IsObject(val) Then 'collection or dictionary
        Set data_.Item("alwaysMatch")(browserOptionKey)(key) = val
    Else
        If IsArray(val) Then
            data_.Item("alwaysMatch")(browserOptionKey)(key) = val
        Else 'string
            If Left$(val, 1) = "[" Or Left$(val, 1) = "{" Then 'its a json string - try to parse
                Set data_.Item("alwaysMatch")(browserOptionKey)(key) = WebJsonConverter.ParseJson(val)
            Else 'its a string parameter
                data_.Item("alwaysMatch")(browserOptionKey)(key) = val
            End If
        End If
    End If
End Sub

'@Description("Adds a list of option excludeSwitches")
Public Sub AddExcludeSwitches(ParamArray addList() As Variant)
Attribute AddExcludeSwitches.VB_Description = "Adds a list of option excludeSwitches"
    Dim varry() As Variant
    varry = addList
    addToOptionList "excludeSwitches", paramListToStringArray(varry)
End Sub

'@Description("Adds a list of option windowTypes")
Public Sub AddWindowTypes(ParamArray addList() As Variant)
Attribute AddWindowTypes.VB_Description = "Adds a list of option windowTypes"
    Dim varry() As Variant
    varry = addList
    addToOptionList "windowTypes", paramListToStringArray(varry)
End Sub

'@Description("Adds a list of option extensions")
Public Sub AddExtensions(ParamArray addList() As Variant)
Attribute AddExtensions.VB_Description = "Adds a list of option extensions"
    Dim varry() As Variant
    Dim pathList() As String
    Dim extList() As String
    Dim i As Long
    Dim fso As New FileSystemObject
    Dim filePath As String
    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        varry = addList
        pathList = paramListToStringArray(varry)
        ReDim extList(1 To UBound(pathList))
        For i = 1 To UBound(pathList)
            filePath = driver_.ResolvePath(pathList(i))
            If Not fso.FileExists(filePath) Then Err.Raise 1, "WebCapabilities", "The extension file path :" & vbNewLine & vbNewLine & filePath & vbNewLine & vbNewLine & "does not exist - please correct the path provided."
            extList(i) = encodeBase64(readByteArrayFromFile(filePath), , False)
        Next i
        addToOptionList "extensions", extList
    Case svbaBrowser.Firefox
        Err.Raise 1, "WebCapabilities", "The AddExtensions method is not supported by the Firefox browser - use InstallAddon method of Webdriver class instead"
    End Select
End Sub

'@Description("Sets the path to the browser executable in a non-standard location")
Public Sub SetBrowserBinary(ByVal browserPath As String)
Attribute SetBrowserBinary.VB_Description = "Sets the path to the browser executable in a non-standard location"
    'set the path to the browser executable to use
    Me.SetOption "binary", driver_.ResolvePath(browserPath)
End Sub

'@Description("Removes control notification for Edge/Chrome")
Public Sub RemoveControlNotification()
Attribute RemoveControlNotification.VB_Description = "Removes control notification for Edge/Chrome"
    'this removes the "Chrome/Edge is being controlled by automated test software" browser notification
    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        Me.AddExcludeSwitches "enable-automation"
        Me.SetOption "useAutomationExtension", False
        'for msedge, you may be greeted with an unwanted "Sync your profile" dialog
        'to get rid of that, run in incognito mode
    Case svbaBrowser.Firefox
        Err.Raise 1, "WebCapabilities", "The RemoveControlNotification method is not supported by the Firefox browser."
    End Select
End Sub

'@Description("Sets the detach option for Edge/Chrome")
Public Sub SetDetachBrowser(ByVal val As Boolean)
Attribute SetDetachBrowser.VB_Description = "Sets the detach option for Edge/Chrome"
    'this sets whether browser is closed (false) or left open (true) when the driver is sent the shutdown command
    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        If Not (driver_.CommandWindowStyle = vbHide Or driver_.CommandWindowStyle = vbMinimizedNoFocus) Then _
            Err.Raise 1, "WebCapabilities", "The SetDetachBrowser method is not compatible with command window style - use vbHide or vbMinimizeNoFocus."
        Me.SetOption "detach", val
    Case svbaBrowser.Firefox
        Err.Raise 1, "WebCapabilities", "The SetDetachBrowser method is not supported by the Firefox browser."
    End Select
End Sub

'@Description("Sets the Debugger Address option for Edge/Chrome needed to connect with an already existing browser")
Public Sub SetDebuggerAddress(Optional ByVal val As String = "localhost:9222")
Attribute SetDebuggerAddress.VB_Description = "Sets the Debugger Address option for Edge/Chrome needed to connect with an already existing browser"
    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        Me.SetOption "debuggerAddress", val
    Case svbaBrowser.Firefox
        Err.Raise 1, "WebCapabilities", "The SetDebuggerAddress method is not supported by the Firefox browser."
    End Select
End Sub

'@Description("Adds a list of experimental browser flags (edge://flags, chrome://flags). Flags must end in @0 (default), @1 (enable), or @2 (disable)")
Public Sub AddLocalStateFlags(ParamArray flagsList() As Variant)
Attribute AddLocalStateFlags.VB_Description = "Adds a list of experimental browser flags (edge://flags, chrome://flags). Flags must end in @0 (default), @1 (enable), or @2 (disable)"
    Dim flags As Collection
    Dim localState As New Dictionary
    Dim browser As New Dictionary
    Dim varry() As Variant
    Dim addList() As String
    Dim flag As String
    varry = flagsList
    addList = paramListToStringArray(varry)
    
    If browser_ <> Chrome And browser_ <> Edge Then Err.Raise 1, , "Error in SetLocalStateFlag - only valid for Edge and Chrome browsers."
    
    'build the enabled_labs_experiments structure if needed
    If Not data_.Item("alwaysMatch")(browserOptionKey).Exists("localState") Then
        browser.Add "enabled_labs_experiments", New Collection
        localState.Add "browser", browser
        SetOption "localState", localState
    Else
        If Not data_.Item("alwaysMatch")(browserOptionKey)("localState").Exists("browser") Then
            browser.Add "enabled_labs_experiments", New Collection
            data_.Item("alwaysMatch")(browserOptionKey)("localState").Add "browser", browser
        Else
            If Not data_.Item("alwaysMatch")(browserOptionKey)("localState")("browser").Exists("enabled_labs_experiments") Then
                data_.Item("alwaysMatch")(browserOptionKey)("localState")("browser").Add "enabled_labs_experiments", New Collection
            End If
        End If
    End If
    
    'get the current flags collection
    Set flags = data_.Item("alwaysMatch")(browserOptionKey)("localState")("browser")("enabled_labs_experiments")
    
    'check if the flag already exists in the collection, and if so, remove it
    Dim i As Long
    Dim j As Long
    For j = 1 To UBound(addList)
        flag = addList(j)
        For i = 1 To flags.Count
            If Split(flags.Item(i), "@")(0) = Split(flag, "@")(0) Then
                flags.Remove i
                Exit For
            End If
        Next i
    Next j
    
    'add the flags
    'this does not guard against user specifying same flag twice
    For j = 1 To UBound(addList)
        flags.Add addList(j)
    Next j

    'reset the modified flags collection
    Set data_.Item("alwaysMatch")(browserOptionKey)("localState")("browser")("enabled_labs_experiments") = flags
End Sub

'-------------------------------------------------------------------------
' Option Argument Settings
'-------------------------------------------------------------------------

'@Description("Adds a list of option arguments")
Public Sub AddArguments(ParamArray addList() As Variant)
Attribute AddArguments.VB_Description = "Adds a list of option arguments"
    Dim varry() As Variant
    varry = addList
    addToOptionList "args", paramListToStringArray(varry)
End Sub

'@Description("Runs headless argument mode")
Public Sub RunInvisible()
Attribute RunInvisible.VB_Description = "Runs headless argument mode"
    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        Me.AddArguments "--headless=new"
    Case svbaBrowser.Firefox
        Me.AddArguments "--headless"
    End Select
End Sub

'@Description("Runs incognito/private argument mode")
Public Sub RunIncognito()
Attribute RunIncognito.VB_Description = "Runs incognito/private argument mode"
    Select Case browser_
    Case svbaBrowser.Chrome
        Me.AddArguments "--incognito"
    Case svbaBrowser.Edge
        Me.AddArguments "--inprivate"
    Case svbaBrowser.Firefox
        Me.AddArguments "-private"
    End Select
End Sub

'@Description("Sets a user-defined profile path")
Public Sub SetProfile(ByVal profilePath As String)
Attribute SetProfile.VB_Description = "Sets a user-defined profile path"
    'this sets a user-defined profile - it will create and populate the profile directory even if it does not exist
    'it is recommended not to make profiles for SeleniumVBA in the default C:\Users\[user]\AppData\Local\Google\Chrome\User Data
    'Instead copy profiles needed from the default location into a different location to access with SeleniumVBA
    Dim fso As New FileSystemObject
    Dim userDir As String
    Dim profileName As String
    
    profilePath = driver_.ResolvePath(profilePath)
    
    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        userDir = fso.GetParentFolderName(profilePath)
        profileName = fso.GetBaseName(profilePath)
        Me.AddArguments "--user-data-dir=" & userDir
        Me.AddArguments "--profile-directory=" & profileName
    Case svbaBrowser.Firefox
        'unlike edge and chrome, firefox requires that the folder to exist
        'hence check and if doesn't exist, then create it
        If Not fso.FolderExists(profilePath) Then createFolderPath profilePath
        Me.AddArguments "-profile", profilePath
    End Select
End Sub

'@Description("Sets the user agent string")
Public Sub SetUserAgent(ByVal val As String)
Attribute SetUserAgent.VB_Description = "Sets the user agent string"
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

'@Description("Sets an option preference key/value")
Public Sub SetPreference(ByVal key As String, ByVal val As Variant)
Attribute SetPreference.VB_Description = "Sets an option preference key/value"
    'input val can be a string, a json string, or an object (dictionary)
    If Not data_.Item("alwaysMatch")(browserOptionKey).Exists("prefs") Then
        SetOption "prefs", New Dictionary
    End If
    
    'need to come up with a more generalized approach to paths - this is just a band-aid
    Select Case key
    Case "download.default_directory", "browser.download.dir", "savefile.default_directory"
        val = driver_.ResolvePath(val)
    Case Else
    End Select
    
    If IsObject(val) Then
        Set data_.Item("alwaysMatch")(browserOptionKey)("prefs")(key) = val
    Else
        If Left$(val, 1) = "[" Or Left$(val, 1) = "{" Then 'it's json string - try to parse
            'convert to dictionary or collection
            Set data_.Item("alwaysMatch")(browserOptionKey)("prefs")(key) = WebJsonConverter.ParseJson(val)
        Else
            data_.Item("alwaysMatch")(browserOptionKey)("prefs")(key) = val
        End If
    End If
End Sub

'@Description("Sets auto-download preferences and directory path")
Public Sub SetDownloadPrefs(Optional ByVal downloadFolderPath As String = ".\", _
                            Optional ByVal promptForDownload As Boolean = False, _
                            Optional ByVal disablePDFViewer As Boolean = True)
Attribute SetDownloadPrefs.VB_Description = "Sets auto-download preferences and directory path"
                            
    downloadFolderPath = driver_.ResolvePath(downloadFolderPath)

    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        Me.SetPreference "download.default_directory", downloadFolderPath
        Me.SetPreference "download.directory_upgrade", True
        Me.SetPreference "safebrowsing.enabled", True
        Me.SetPreference "download.prompt_for_download", promptForDownload 'false disables prompting user for download permission
        Me.SetPreference "plugins.always_open_pdf_externally", disablePDFViewer 'instead of using the pdf viewer
        'Me.SetPreference "browser.set_download_behavior", "allow"
        'Me.SetPreference "safebrowsing.disable_download_protection", True
        'Me.SetPreference "profile.default_content_settings.popups", False '0
        'Me.SetPreference "profile.default_content_settings.javascript", True '1
        'Me.SetPreference "profile.default_content_setting_values.notifications", 2
        'Me.SetPreference "profile.content_settings.exceptions.automatic_downloads.*.setting", True '1
        'Me.SetPreference "profile.default_content_setting_values.automatic_downloads", True '1
    Case svbaBrowser.Firefox
        Me.SetPreference "browser.download.folderList", 2 '0 is Desktop, 1 is Downloads
        Me.SetPreference "browser.download.dir", downloadFolderPath
        Me.SetPreference "pdfjs.disabled", disablePDFViewer  'disable the built-in PDF viewer
        
        'Me.SetPreference "browser.download.manager.showWhenStarting", False
        'Me.SetPreference "browser.helperApps.neverAsk.saveToDisk", "application/pdf;text/plain;text/csv;application/text;text/xml;application/xml;application/csv;pplication/octet-stream doc xls pdf txt"
        'Me.SetPreference "browser.download.useDownloadDir", True
        'Me.SetPreference "browser.download.viewableInternally.enabledTypes", ""
    End Select
End Sub

'@Description("Sets whether or not to allow geolocation awareness")
Public Sub SetGeolocationAware(Optional ByVal aware As Boolean = True)
Attribute SetGeolocationAware.VB_Description = "Sets whether or not to allow geolocation awareness"
    'this may be required, for example, if using incognito mode and
    'wanting navigated url to be aware of geolocation
    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge
        Me.SetPreference "profile.managed_default_content_settings.geolocation", IIf(aware, 1, 2)
        'Me.SetPreference "googlegeolocationaccess.enabled", aware
        'Me.SetPreference "profile.default_content_setting_values.notifications", IIf(aware, 1, 2)
        'Me.SetPreference "profile.default_content_setting_values.geolocation", IIf(aware, 1, 2)
    Case svbaBrowser.Firefox
        Err.Raise 1, "WebCapabilities", "The SetGeolocationAware method is not supported by the Firefox browser."
    End Select
End Sub

'-------------------------------------------------------------------------
' Public helper methods
'-------------------------------------------------------------------------

'@DefaultMember
'@Description("Returns a capabilities Dictionary object")
Public Function Dictionary() As Dictionary
Attribute Dictionary.VB_Description = "Returns a capabilities Dictionary object"
Attribute Dictionary.VB_UserMemId = 0
    Set Dictionary = data_
End Function

'@Description("Returns Json string for qc")
Public Function ToJson() As String
Attribute ToJson.VB_Description = "Returns Json string for qc"
    ToJson = WebJsonConverter.ConvertToJson(data_, 4)
End Function

'@Description("Saves the capabilities to file path")
Public Sub SaveToFile(Optional ByVal filePath As String = ".\capabilities.json")
Attribute SaveToFile.VB_Description = "Saves the capabilities to file path"
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    filePath = driver_.ResolvePath(filePath)
    Set ts = fso.CreateTextFile(filePath, True, True)
    ts.Write Me.ToJson
    ts.Close
End Sub

'@Description("Loads the capabilities from file path")
Public Sub LoadFromFile(Optional ByVal filePath As String = ".\capabilities.json")
Attribute LoadFromFile.VB_Description = "Loads the capabilities from file path"
    Dim fso As New FileSystemObject
    Dim ts As TextStream
        
    filePath = driver_.ResolvePath(filePath)
    
    If Not fso.FileExists(filePath) Then Err.Raise 53, "WebCapabilities", "Capabilities Json file does not exist:" & vbNewLine & vbNewLine & filePath
    'note that this will fail if input file format is utf-8 with non-ascii chars - see https://github.com/GCuser99/SeleniumVBA/discussions/32
    Set ts = fso.OpenTextFile(filePath, , , TristateUseDefault)
    Set data_ = WebJsonConverter.ParseJson(ts.ReadAll)
    ts.Close
    
    If data_.Item("browserName") <> WebShared.getBrowserNameString(browser_) Then Err.Raise 1, "WebCapabilities", "Browser name from file does not match browser name assigned to driver instance."
End Sub

'-------------------------------------------------------------------------
' Private helper methods
'-------------------------------------------------------------------------

Friend Sub initializeFor(ByVal browser As svbaBrowser)
    Set data_ = New Dictionary
    data_.Add "browserName", WebShared.getBrowserNameString(browser)
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
    
    data_.Item("alwaysMatch").Add browserOptionKey, browserOptions
End Sub

Friend Property Set WebDriver(ByRef obj As WebDriver)
    Set driver_ = obj
End Property

Private Function getAlertHandlingString(ByVal val As svbaAlertHandling) As String
    Select Case val
    Case svbaAlertHandling.svbaAccept
        getAlertHandlingString = "accept"
    Case svbaAlertHandling.svbaAcceptAndNotify
        getAlertHandlingString = "accept and notify"
    Case svbaAlertHandling.svbaDismiss
        getAlertHandlingString = "dismiss"
    Case svbaAlertHandling.svbaDismissAndNotify
        getAlertHandlingString = "dismiss and notify"
    Case svbaAlertHandling.svbaIgnore
        getAlertHandlingString = "ignore"
    End Select
End Function

Private Sub addToOptionList(ByVal optionName As String, addList() As String)
    'add to list of option parameters, while checking for duplicates
    Dim parameters As Collection
    Dim found As Boolean
    Dim j As Long
    Dim i As Long

    'lists are represented as collections for caps initialized from LoadFromFile (via WebJsonConverter.ParseJson)
    If IsObject(data_.Item("alwaysMatch")(browserOptionKey)(optionName)) Then
        Set parameters = data_.Item("alwaysMatch")(browserOptionKey)(optionName)
    Else 'starting from scratch
        Set parameters = New Collection
    End If

    For i = LBound(addList) To UBound(addList)
        found = False
        For j = 1 To parameters.Count
            If addList(i) = parameters(j) Then found = True: Exit For
        Next j
        If Not found Then
            parameters.Add addList(i)
        End If
    Next i
    'put the expanded list back
    Set data_.Item("alwaysMatch")(browserOptionKey)(optionName) = parameters
End Sub

Private Function paramListToStringArray(parmList() As Variant) As String()
    'a parameter array can include arrays as list elements - expand into a string array
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strlist() As String
    
    'count number of parameters in list
    k = 0
    For i = LBound(parmList) To UBound(parmList)
        If IsArray(parmList(i)) Then
            k = k + UBound(parmList(i)) - LBound(parmList(i)) + 1
        Else
            If Left$(parmList(i), 1) = "[" Then  'its a json string - try to parse
                k = k + WebJsonConverter.ParseJson(parmList(i)).Count
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
            If Left$(parmList(i), 1) = "[" Then  'its a json string
                Dim col As Collection
                Set col = WebJsonConverter.ParseJson(parmList(i))
                For j = 1 To col.Count
                    k = k + 1
                    strlist(k) = col.Item(j)
                Next j
            Else 'its a string parameter
                k = k + 1
                strlist(k) = parmList(i)
            End If
        End If
    Next i
    
    paramListToStringArray = strlist
End Function

'https://stackoverflow.com/questions/31033820/filesystemobject-createfolder-to-create-directory-and-subdirectories
Private Function createFolderPath(ByVal folderPath As String) As Boolean
    Static fso As FileSystemObject
 
    'Initialize fso variable if not already setup
    If fso Is Nothing Then Set fso = New FileSystemObject

    'Is the folderPath parameter populated
    If Len(folderPath) = 0 Then
        createFolderPath = False
        Exit Function
    End If

    'If the folderPath exists as a file, the function fails.
    If fso.FileExists(folderPath) Then
        createFolderPath = False
        Exit Function
    End If
 
    'If the folderPath already exists as a folder, don't do anything and return success.
    If fso.FolderExists(folderPath) Then
        createFolderPath = True
        Exit Function
    End If
 
    'recursively create the parent folder, then if successful create the top folder.
    If createFolderPath(fso.GetParentFolderName(folderPath)) Then
        createFolderPath = Not (fso.CreateFolder(folderPath) Is Nothing)
    Else
        createFolderPath = False
    End If
End Function
