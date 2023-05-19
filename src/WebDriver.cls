VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class provides the main mechanism to control/automate a browser application"
'@ModuleDescription "This class provides the main mechanism to control/automate a browser application"
'@Exposed
'@folder("SeleniumVBA.Source")
' ==========================================================================
' SeleniumVBA v4.1
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
' Copyright (c) 2023, GCUser99 and 6DiegoDiego9 (https://github.com/GCuser99/SeleniumVBA)
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

Option Explicit

Private checkDriverBrowserVersionAlignment As Boolean

Private minDriverBrowserCompatibilityLevel As Long

'desired lower bound of returned arrays
Private Const baseArrayIdx As Integer = 1

'change baseDomIdx to 0 if zero-based Dom (de)selection is desired
'with baseDomIdx=1: SelectByIndex 1 gets first element
'with baseDomIdx=0: SelectByIndex 0 gets first element
Private Const baseDomIdx As Integer = 1

Private sessionId_ As String
Private driverUrl_ As String
Private browser_ As svbaBrowser
Private appWinStyle As VbAppWinStyle
Private defaultIOFolder_ As String
Private defaultDriverFolder_ As String

Private preloadChromeCapabilitiesFilePath As String
Private preloadEdgeCapabilitiesFilePath As String
Private preloadFirefoxCapabilitiesFilePath As String
Private preloadIECapabilitiesFilePath As String

Private driverDefaultPortChrome As Long
Private driverDefaultPortEdge As Long
Private driverDefaultPortFirefox As Long
Private driverDefaultPortIE As Long

'store original style of Highlighted elements to be able to UnHighlight them at a future time
Private SetHighlight_ElemsOrigStyle As Scripting.Dictionary

Public Enum VBAcolors
    Unchanged = -1
    Black = VBA.ColorConstants.vbBlack
    Blue = VBA.ColorConstants.vbBlue
    Cyan = VBA.ColorConstants.vbCyan
    Green = VBA.ColorConstants.vbGreen
    Magenta = VBA.ColorConstants.vbMagenta
    Red = VBA.ColorConstants.vbRed
    Yellow = VBA.ColorConstants.vbYellow
    'other custom fine-tuned colors can be added, for example: "Brown = VBA.RGB(128, 0, 0)"
End Enum

'Highlighting globals:
Private setHightlightFoundElems_Settings As Collection
Private lastElems As WebElements

Private driverFilename As String
Private processID As String

Private Const ELEMENT_KEY As String = "element-6066-11e4-a52e-4f735466cecf"
Private Const SHADOWROOT_KEY As String = "shadow-6066-11e4-a52e-4f735466cecf"

Private Type typCMD
    CMD_STATUS As Variant
    CMD_NEW_SESSION As Variant
    CMD_GET_ALL_SESSIONS As Variant
    CMD_QUIT As Variant
    CMD_GET_CURRENT_WINDOW_HANDLE As Variant
    CMD_GET_WINDOW_HANDLES As Variant
    CMD_GET As Variant
    CMD_GO_FORWARD As Variant
    CMD_GO_BACK As Variant
    CMD_REFRESH As Variant
    CMD_EXECUTE_SCRIPT As Variant
    CMD_EXECUTE_SCRIPT_ASYNC As Variant
    CMD_GET_CURRENT_URL As Variant
    CMD_GET_TITLE As Variant
    CMD_GET_PAGE_SOURCE As Variant
    CMD_SCREENSHOT As Variant
    CMD_ELEMENT_SCREENSHOT As Variant
    CMD_FIND_ELEMENT As Variant
    CMD_FIND_ELEMENTS As Variant
    CMD_GET_ACTIVE_ELEMENT As Variant
    CMD_FIND_ELEMENT_FROM_ELEMENT As Variant
    CMD_FIND_ELEMENTS_FROM_ELEMENT As Variant
    CMD_CLICK_ELEMENT As Variant
    CMD_CLEAR_ELEMENT As Variant
    CMD_GET_ELEMENT_TEXT As Variant
    CMD_SEND_KEYS_TO_ELEMENT As Variant
    CMD_GET_ELEMENT_TAG_NAME As Variant
    CMD_IS_ELEMENT_SELECTED As Variant
    CMD_IS_ELEMENT_ENABLED As Variant
    CMD_IS_ELEMENT_DISPLAYED As Variant
    CMD_GET_ELEMENT_RECT As Variant
    CMD_GET_ELEMENT_ATTRIBUTE As Variant
    CMD_GET_ELEMENT_PROPERTY As Variant
    CMD_GET_ALL_COOKIES As Variant
    CMD_ADD_COOKIE As Variant
    CMD_GET_COOKIE As Variant
    CMD_DELETE_ALL_COOKIES As Variant
    CMD_DELETE_COOKIE As Variant
    CMD_SWITCH_TO_FRAME As Variant
    CMD_SWITCH_TO_PARENT_FRAME As Variant
    CMD_SWITCH_TO_WINDOW As Variant
    CMD_NEW_WINDOW As Variant
    CMD_CLOSE As Variant
    CMD_GET_ELEMENT_VALUE_OF_CSS_PROPERTY As Variant
    CMD_SET_TIMEOUTS As Variant
    CMD_GET_TIMEOUTS As Variant
    CMD_DISMISS_ALERT As Variant
    CMD_ACCEPT_ALERT As Variant
    CMD_SET_ALERT_TEXT As Variant
    CMD_GET_ALERT_TEXT As Variant
    CMD_CLICK As Variant
    CMD_SET_WINDOW_RECT As Variant
    CMD_GET_WINDOW_RECT As Variant
    CMD_MAXIMIZE_WINDOW As Variant
    CMD_GET_LOCATION As Variant
    CMD_SET_LOCATION As Variant
    CMD_FULLSCREEN_WINDOW As Variant
    CMD_MINIMIZE_WINDOW As Variant
    CMD_SHUTDOWN As Variant
    CMD_PRINT_PAGE As Variant
    CMD_GET_ELEMENT_SHADOW_ROOT As Variant
    CMD_FIND_ELEMENT_FROM_SHADOW_ROOT As Variant
    CMD_FIND_ELEMENTS_FROM_SHADOW_ROOT As Variant
    CMD_GET_ELEMENT_ARIA_ROLE As Variant
    CMD_GET_ELEMENT_ARIA_LABEL As Variant
End Type

Private tCMD As typCMD

Public Enum By
    ID = 0
    tagName = 1
    className = 2
    Name = 3
    cssSelector = 4
    XPath = 5
    linkText = 6
    PartialLinkText = 7
End Enum

Public Enum ScrollIntoViewOptions_speed
    jump_smooth = 1
    jump_instant = 2
    jump_auto = 3
End Enum

Public Enum ScrollIntoViewOptions_alignment
    align_start = 1
    align_center = 2
    align_end = 3
    align_nearest = 4
End Enum

Public Enum svbaWindowType
    svbaWindow = 0
    svbaTab = 1
End Enum

Private Sub Class_Initialize()
    Dim oShell As New Shell32.Shell
    Dim iniFilePath As String
    initCommands

    'attempt to read global variable values from optional settings file - if not successful,
    'then assign a system default value (last input parameter of ReadIniFileEntry)
    iniFilePath = getFullLocalPath(thisLibFolderPath) & "\SeleniumVBA.ini"
    
    Me.CommandWindowStyle = enumTextToValue(readIniFileEntry(iniFilePath, "GENERAL", "command_window_style", vbHide))
    Me.DefaultDriverFolder = readIniFileEntry(iniFilePath, "GENERAL", "driver_location_folder", oShell.Namespace("shell:Downloads").Self.Path)
    
    defaultIOFolder_ = readIniFileEntry(iniFilePath, "GENERAL", "default_io_folder", vbNullString)
    
    checkDriverBrowserVersionAlignment = readIniFileEntry(iniFilePath, "AUTO-DRIVER-UPDATE", "auto_detect_and_update", True)
    minDriverBrowserCompatibilityLevel = enumTextToValue(readIniFileEntry(iniFilePath, "AUTO-DRIVER-UPDATE", "min_compat_level", svbaBuildMajor))

    preloadChromeCapabilitiesFilePath = readIniFileEntry(iniFilePath, "CHROME", "preload_capabilities_file_path", vbNullString)
    preloadEdgeCapabilitiesFilePath = readIniFileEntry(iniFilePath, "EDGE", "preload_capabilities_file_path", vbNullString)
    preloadFirefoxCapabilitiesFilePath = readIniFileEntry(iniFilePath, "FIREFOX", "preload_capabilities_file_path", vbNullString)
    preloadIECapabilitiesFilePath = readIniFileEntry(iniFilePath, "INTERNET EXPLORER", "preload_capabilities_file_path", vbNullString)
    
    driverDefaultPortChrome = CLng(readIniFileEntry(iniFilePath, "CHROME", "local_host_port", 9515))
    driverDefaultPortEdge = CLng(readIniFileEntry(iniFilePath, "EDGE", "local_host_port", 9516))
    driverDefaultPortFirefox = CLng(readIniFileEntry(iniFilePath, "FIREFOX", "local_host_port", 4444))
    driverDefaultPortIE = CLng(readIniFileEntry(iniFilePath, "INTERNET EXPLORER", "local_host_port", 5555))
End Sub

Private Sub initCommands()
    With tCMD
        .CMD_STATUS = Array("GET", "/status")
        .CMD_NEW_SESSION = Array("POST", "/session")
        .CMD_GET_ALL_SESSIONS = Array("GET", "/sessions")
        .CMD_QUIT = Array("DELETE", "/session/$sessionId")
        .CMD_GET_CURRENT_WINDOW_HANDLE = Array("GET", "/session/$sessionId/window")
        .CMD_GET_WINDOW_HANDLES = Array("GET", "/session/$sessionId/window/handles")
        .CMD_GET = Array("POST", "/session/$sessionId/url")
        .CMD_GO_FORWARD = Array("POST", "/session/$sessionId/forward")
        .CMD_GO_BACK = Array("POST", "/session/$sessionId/back")
        .CMD_REFRESH = Array("POST", "/session/$sessionId/refresh")
        .CMD_EXECUTE_SCRIPT = Array("POST", "/session/$sessionId/execute/sync")
        .CMD_EXECUTE_SCRIPT_ASYNC = Array("POST", "/session/$sessionId/execute/async")
        .CMD_GET_CURRENT_URL = Array("GET", "/session/$sessionId/url")
        .CMD_GET_TITLE = Array("GET", "/session/$sessionId/title")
        .CMD_GET_PAGE_SOURCE = Array("GET", "/session/$sessionId/source")
        .CMD_SCREENSHOT = Array("GET", "/session/$sessionId/screenshot")
        .CMD_ELEMENT_SCREENSHOT = Array("GET", "/session/$sessionId/element/$id/screenshot")
        .CMD_FIND_ELEMENT = Array("POST", "/session/$sessionId/element")
        .CMD_FIND_ELEMENTS = Array("POST", "/session/$sessionId/elements")
        .CMD_GET_ACTIVE_ELEMENT = Array("GET", "/session/$sessionId/element/active")
        .CMD_FIND_ELEMENT_FROM_ELEMENT = Array("POST", "/session/$sessionId/element/$id/element")
        .CMD_FIND_ELEMENTS_FROM_ELEMENT = Array("POST", "/session/$sessionId/element/$id/elements")
        .CMD_CLICK_ELEMENT = Array("POST", "/session/$sessionId/element/$id/click")
        .CMD_CLEAR_ELEMENT = Array("POST", "/session/$sessionId/element/$id/clear")
        .CMD_GET_ELEMENT_TEXT = Array("GET", "/session/$sessionId/element/$id/text")
        .CMD_SEND_KEYS_TO_ELEMENT = Array("POST", "/session/$sessionId/element/$id/value")
        .CMD_GET_ELEMENT_TAG_NAME = Array("GET", "/session/$sessionId/element/$id/name")
        .CMD_IS_ELEMENT_SELECTED = Array("GET", "/session/$sessionId/element/$id/selected")
        .CMD_IS_ELEMENT_ENABLED = Array("GET", "/session/$sessionId/element/$id/enabled")
        .CMD_IS_ELEMENT_DISPLAYED = Array("GET", "/session/$sessionId/element/$id/displayed")
        .CMD_GET_ELEMENT_RECT = Array("GET", "/session/$sessionId/element/$id/rect")
        .CMD_GET_ELEMENT_ATTRIBUTE = Array("GET", "/session/$sessionId/element/$id/attribute/$name")
        .CMD_GET_ELEMENT_PROPERTY = Array("GET", "/session/$sessionId/element/$id/property/$name")
        .CMD_GET_ALL_COOKIES = Array("GET", "/session/$sessionId/cookie")
        .CMD_ADD_COOKIE = Array("POST", "/session/$sessionId/cookie")
        .CMD_GET_COOKIE = Array("GET", "/session/$sessionId/cookie/$name")
        .CMD_DELETE_ALL_COOKIES = Array("DELETE", "/session/$sessionId/cookie")
        .CMD_DELETE_COOKIE = Array("DELETE", "/session/$sessionId/cookie/$name")
        .CMD_SWITCH_TO_FRAME = Array("POST", "/session/$sessionId/frame")
        .CMD_SWITCH_TO_PARENT_FRAME = Array("POST", "/session/$sessionId/frame/parent")
        .CMD_SWITCH_TO_WINDOW = Array("POST", "/session/$sessionId/window")
        .CMD_NEW_WINDOW = Array("POST", "/session/$sessionId/window/new")
        .CMD_CLOSE = Array("DELETE", "/session/$sessionId/window")
        .CMD_GET_ELEMENT_VALUE_OF_CSS_PROPERTY = Array("GET", "/session/$sessionId/element/$id/css/$name")
        .CMD_SET_TIMEOUTS = Array("POST", "/session/$sessionId/timeouts")
        .CMD_GET_TIMEOUTS = Array("GET", "/session/$sessionId/timeouts")
        .CMD_DISMISS_ALERT = Array("POST", "/session/$sessionId/alert/dismiss")
        .CMD_ACCEPT_ALERT = Array("POST", "/session/$sessionId/alert/accept")
        .CMD_SET_ALERT_TEXT = Array("POST", "/session/$sessionId/alert/text")
        .CMD_GET_ALERT_TEXT = Array("GET", "/session/$sessionId/alert/text")
        .CMD_CLICK = Array("POST", "/session/$sessionId/click")
        .CMD_SET_WINDOW_RECT = Array("POST", "/session/$sessionId/window/rect")
        .CMD_GET_WINDOW_RECT = Array("GET", "/session/$sessionId/window/rect")
        .CMD_MAXIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/maximize")
        .CMD_GET_LOCATION = Array("GET", "/session/$sessionId/location")
        .CMD_SET_LOCATION = Array("POST", "/session/$sessionId/location")
        .CMD_FULLSCREEN_WINDOW = Array("POST", "/session/$sessionId/window/fullscreen")
        .CMD_MINIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/minimize")
        .CMD_SHUTDOWN = Array("GET", "/shutdown")
        .CMD_PRINT_PAGE = Array("POST", "/session/$sessionId/print")
        .CMD_GET_ELEMENT_SHADOW_ROOT = Array("GET", "/session/$sessionId/element/$id/shadow")
        .CMD_FIND_ELEMENT_FROM_SHADOW_ROOT = Array("POST", "/session/$sessionId/shadow/$sid/element")
        .CMD_FIND_ELEMENTS_FROM_SHADOW_ROOT = Array("POST", "/session/$sessionId/shadow/$sid/elements")
        .CMD_GET_ELEMENT_ARIA_ROLE = Array("GET", "/session/$sessionId/element/$id/computedrole")
        .CMD_GET_ELEMENT_ARIA_LABEL = Array("GET", "/session/$sessionId/element/$id/computedlabel")
    End With
End Sub

'@Description("Set the driver command window display mode - defaults to vbHide")
Public Property Let CommandWindowStyle(ByVal style As VbAppWinStyle)
Attribute CommandWindowStyle.VB_Description = "Set the driver command window display mode - defaults to vbHide"
    appWinStyle = style
End Property

'@Description("Start an Edge driver")
Public Sub StartEdge(Optional ByVal driverPath As String = "msedgedriver.exe", Optional ByVal localHostPort As Long, Optional ByVal enableLogging As Boolean = False, Optional ByVal logFilePath As String = ".\edge.log")
Attribute StartEdge.VB_Description = "Start an Edge driver"
    If localHostPort = 0 Then localHostPort = driverDefaultPortEdge
    start svbaBrowser.Edge, driverPath, localHostPort, enableLogging, logFilePath
End Sub

'@Description("Start a Chrome driver")
Public Sub StartChrome(Optional ByVal driverPath As String = "chromedriver.exe", Optional ByVal localHostPort As Long, Optional ByVal enableLogging As Boolean = False, Optional ByVal logFilePath As String = ".\chrome.log")
Attribute StartChrome.VB_Description = "Start a Chrome driver"
    If localHostPort = 0 Then localHostPort = driverDefaultPortChrome
    start svbaBrowser.Chrome, driverPath, localHostPort, enableLogging, logFilePath
End Sub

'@Description("Start a Firefox driver")
Public Sub StartFirefox(Optional ByVal driverPath As String = "geckodriver.exe", Optional ByVal localHostPort As Long, Optional ByVal enableLogging As Boolean = False, Optional ByVal logFilePath As String = ".\firefox.log")
Attribute StartFirefox.VB_Description = "Start a Firefox driver"
    If localHostPort = 0 Then localHostPort = driverDefaultPortFirefox
    start svbaBrowser.Firefox, driverPath, localHostPort, enableLogging, logFilePath
End Sub

'@Description("Start an IE mode driver")
Public Sub StartIE(Optional ByVal driverPath As String = "IEDriverServer.exe", Optional ByVal localHostPort As Long, Optional ByVal enableLogging As Boolean = False, Optional ByVal logFilePath As String = ".\ie.log")
Attribute StartIE.VB_Description = "Start an IE mode driver"
    If localHostPort = 0 Then localHostPort = driverDefaultPortIE
    start svbaBrowser.IE, driverPath, localHostPort, enableLogging, logFilePath
End Sub

Private Sub start(ByVal browser As svbaBrowser, ByVal driverPath As String, ByVal localHostPort As Long, ByVal enableLogging As Boolean, ByVal logFilePath As String)
    Dim fso As New Scripting.FileSystemObject, serviceArgs As String, commandStr As String
    
    browser_ = browser
    
    If browser_ = svbaBrowser.Firefox Or browser_ = svbaBrowser.IE Then
        'for firefox, if previous session is still open on same port, then will get "Session is already started"
        'on OpenBrowser so kill all driver instances on the assigned port
        'could use KillDriver on Chrome/Edge too, but since they support multi-sessions on same port, this will eliminate
        'that functionality if important
        killDriver browser_, localHostPort
    End If
    
    driverPath = WebShared.getFullLocalPath(driverPath, defaultDriverFolder_)

    If checkDriverBrowserVersionAlignment Then
        Dim mngr As New WebDriverManager
        'it's important to run these in silent mode so that if there is a problem, we don't get in the way of automation activity
        mngr.AlignDriverWithBrowser browser_, driverPath, minDriverBrowserCompatibilityLevel, True
    End If
    
    If Not fso.FileExists(driverPath) Then
        Err.Raise 1, "WebDriver", "Could not find Selenium WebDriver at :" & vbNewLine & vbNewLine & driverPath & vbNewLine & vbNewLine & "Either manually install the driver or let WebDriverManager class install"
    End If
    
    commandStr = vbNullString: serviceArgs = vbNullString

    If enableLogging Then
        logFilePath = WebShared.getFullLocalPath(logFilePath, defaultIOFolder_)
        Select Case browser_
        Case svbaBrowser.Edge, svbaBrowser.Chrome
            serviceArgs = " --verbose --log-path=" & Chr$(34) & logFilePath & Chr$(34)
        Case svbaBrowser.Firefox
            'https://firefox-source-docs.mozilla.org/testing/geckodriver/TraceLogs.html
            'apparently firefox does not support command line argument for specifying log file path
            'so using cmd.exe to redirect log from console window
            serviceArgs = " -vv"" > " & Replace$(logFilePath, " ", "^ ") 'need to escape spaces in log path
            commandStr = Chr$(34) & Environ("comspec") & Chr$(34) & " /c " & """"
        Case svbaBrowser.IE
            serviceArgs = " --log-level=TRACE --log-file=" & Chr$(34) & logFilePath & Chr$(34)
        End Select
    End If
    
    commandStr = commandStr & Chr$(34) & driverPath & Chr$(34) & " --port=" & localHostPort & serviceArgs
    
    processID = Shell(commandStr, appWinStyle)

    If processID = 0 Then
        Err.Raise 1, "WebDriver", "Failed in starting Selenium WebDriver." & vbNewLine & "WebDriver path: " & commandStr
    End If
    
    'if we've got this far, then store some global variable values for later use
    driverUrl_ = "http://localhost:" & localHostPort
    driverFilename = fso.GetFileName(driverPath)
End Sub

'@Description("Shuts down all driver instances attached to this driver's port")
Public Sub Shutdown()
Attribute Shutdown.VB_Description = "Shuts down all driver instances attached to this driver's port"
    'this will shut down all WebDriver instances using this driver's assigned port
    'which can be a good thing if there are multiple "rogue" instances lurking
    
    Select Case browser_
    Case svbaBrowser.Edge, svbaBrowser.Chrome
        'this will shutdown driver and session - equivalent to CloseBrowser and kill driver process
        'see "detach" option for chrome/edge to affect this behavior
        execute tCMD.CMD_SHUTDOWN
    Case svbaBrowser.Firefox
        'firefox does not recognize shutdown command, so need to roll our own...
        'exit as gracefully as possible - first be sure user closed the browser, and if not, then let's do it here
        If sessionId_ <> vbNullString Then CloseBrowser
        taskKillByPid processID
    Case svbaBrowser.IE
        'unfortunately, CMD_SHUTDOWN in IE mode does not kill all driver sub-processes such as the broswer
        'so similar to Firefox, must check if browser is still open, and close if needed
        If sessionId_ <> vbNullString Then CloseBrowser
        execute tCMD.CMD_SHUTDOWN
    End Select
End Sub

'@Description("Opens browser instance with specified capabilities")
Public Sub OpenBrowser(Optional caps As WebCapabilities, Optional ByVal invisible As Boolean = False, Optional ByVal incognito As Boolean = False, Optional capabilitiesFilePath As String = vbNullString)
Attribute OpenBrowser.VB_Description = "Opens browser instance with specified capabilities"
    Dim resp As Dictionary
    
    If sessionId_ <> vbNullString Then
        Err.Raise 1, "WebDriver", "Only one OpenBrowser command per WebDriver instance is allowed! Use CloseBrowser before another OpenBrowser to fix this."
    End If
    
    If caps Is Nothing Then
        'need to initialize one
        If capabilitiesFilePath <> vbNullString Then
            'user created capabilities from file
            Set caps = CreateCapabilities(initializeFromSettingsFile:=False)
            caps.LoadFromFile capabilitiesFilePath
        Else
            'create from scratch, but then initialize from settings file if applicable
            Set caps = CreateCapabilities(initializeFromSettingsFile:=True)
        End If
    End If
    
    If invisible Then
        caps.RunInvisible
    End If
    
    If incognito Then
        caps.RunIncognito
    End If
    
    DoEvents 'this is needed to prevent intermittent error if command window is not hidden or minimized
    
    Set resp = execute(tCMD.CMD_NEW_SESSION, params("capabilities", caps.Dictionary))("value")
    
    sessionId_ = resp("sessionId")
End Sub

'@Description("Closes browser without terminating driver instance")
Public Sub CloseBrowser()
Attribute CloseBrowser.VB_Description = "Closes browser without terminating driver instance"
    execute tCMD.CMD_QUIT, New Dictionary
    sessionId_ = vbNullString
End Sub

'@Description("Navigates browser to a specified webpage and returns true if no timeout or other error")
Public Function NavigateTo(ByVal url As String, Optional ByVal timeOutms As Variant, Optional ByVal raiseError As Boolean = True) As Boolean
Attribute NavigateTo.VB_Description = "Navigates browser to a specified webpage and returns true if no timeout or other error"
    Dim data As New Dictionary
    Dim resp As Dictionary
    Dim savTimeOutms As Long
    
    data.Add "url", url
    If Not IsMissing(timeOutms) Then
        savTimeOutms = PageLoadTimeout()
        If savTimeOutms <> timeOutms Then PageLoadTimeout = timeOutms
    End If
    
    Set resp = execute(tCMD.CMD_GET, data, raiseError)

    If Not IsMissing(timeOutms) Then
        If savTimeOutms <> timeOutms Then PageLoadTimeout = savTimeOutms
    End If
    
    If isResponseError(resp) Then
        NavigateTo = False
        Debug.Print getResponseErrorMessage(resp)
    Else
        NavigateTo = True
    End If
End Function

'@Description("Navigates browser to an html file of the specified file path")
Public Sub NavigateToFile(ByVal htmlFilePath As String)
Attribute NavigateToFile.VB_Description = "Navigates browser to an html file of the specified file path"
    'input file must have html extension to be interpreted as a webpage
    'NavigateTo (tCMD.CMD_GET) will automatically discern whether the input is a file or url, even without the 'file:///'
    'So the only functional value of this method is to handle relative paths!
    NavigateTo "file:///" & WebShared.getFullLocalPath(htmlFilePath, defaultIOFolder_)
End Sub

'@Description("Finds the first element for a given selector and value")
Public Function FindElement(by_ As By, ByVal val As String, Optional parentElement As WebElement) As WebElement
Attribute FindElement.VB_Description = "Finds the first element for a given selector and value"
    Dim data As Dictionary
    Dim vCMD As Variant
    Dim elem As New WebElement
    Set data = toSelector(by_, val)
    If Not parentElement Is Nothing Then
        data.Add "id", parentElement.elementId
        vCMD = tCMD.CMD_FIND_ELEMENT_FROM_ELEMENT
    Else
        vCMD = tCMD.CMD_FIND_ELEMENT
    End If

    Set elem = toWebElement(execute(vCMD, data)("value")(ELEMENT_KEY))
    
    If Not setHightlightFoundElems_Settings Is Nothing Then setHightlightFoundElems_exec elem
    
    Set FindElement = elem
End Function

'@Description("Finds all elements that satisfy the specified selector and value")
Public Function FindElements(by_ As By, ByVal val As String, Optional parentElement As WebElement) As WebElements
Attribute FindElements.VB_Description = "Finds all elements that satisfy the specified selector and value"
    Dim data As Dictionary
    Dim i As Integer
    Dim vCMD As Variant
    Dim elements As Collection
    Dim elems As New WebElements
    
    Set data = toSelector(by_, val)
    
    If Not parentElement Is Nothing Then
        data.Add "id", parentElement.elementId
        vCMD = tCMD.CMD_FIND_ELEMENTS_FROM_ELEMENT
    Else
        vCMD = tCMD.CMD_FIND_ELEMENTS
    End If
    
    Set elements = execute(vCMD, data)("value")
    
    For i = 1 To elements.Count
        elems.Add toWebElement(elements(i)(ELEMENT_KEY))
    Next

    If Not setHightlightFoundElems_Settings Is Nothing Then setHightlightFoundElems_exec elems
    
    Set elems.WebDriver = Me
    
    Set FindElements = elems
End Function

Private Sub setHightlightFoundElems_exec(elementOrElements As Object)
    With setHightlightFoundElems_Settings
        Select Case TypeName(elementOrElements)
            Case "WebElement"
                SetHighlight True, elementOrElements, .Item(1), .Item(2), .Item(3)
            Case "WebElements"
                SetHighlight True, elementOrElements, .Item(4), .Item(5), .Item(6)
        End Select
    End With
End Sub

Private Function toSelector(by_ As By, ByVal val As String) As Dictionary
    Dim data As New Dictionary
    
    Select Case by_
    Case By.ID
        data.Add "using", "css selector"
        data.Add "value", "[id=""" & val & """]"
    Case By.tagName
        data.Add "using", "css selector"
        data.Add "value", val
    Case By.className
        data.Add "using", "css selector"
        data.Add "value", "." & val
    Case By.Name
        data.Add "using", "css selector"
        data.Add "value", "[name=""" & val & """]"
    Case By.cssSelector
        data.Add "using", "css selector"
        data.Add "value", val
    Case By.XPath
        data.Add "using", "xpath"
        data.Add "value", val
    Case By.linkText
        data.Add "using", "link text"
        data.Add "value", val
    Case By.PartialLinkText
        data.Add "using", "partial link text"
        data.Add "value", val
    Case Else
        data.Add "using", "css selector"
        data.Add "value", val
    End Select
    
    Set toSelector = data
End Function

'@Description("Finds the first element by the specified id value")
Public Function FindElementByID(ByVal val As String, Optional parentElement As WebElement) As WebElement
Attribute FindElementByID.VB_Description = "Finds the first element by the specified id value"
    Set FindElementByID = FindElement(By.ID, val, parentElement)
End Function

'@Description("Finds the first element by the specified tag name value")
Public Function FindElementByTagName(ByVal val As String, Optional parentElement As WebElement) As WebElement
Attribute FindElementByTagName.VB_Description = "Finds the first element by the specified tag name value"
    Set FindElementByTagName = FindElement(By.tagName, val, parentElement)
End Function

'@Description("Finds the first element by the specified class name value")
Public Function FindElementByClassName(ByVal val As String, Optional parentElement As WebElement) As WebElement
Attribute FindElementByClassName.VB_Description = "Finds the first element by the specified class name value"
    Set FindElementByClassName = FindElement(By.className, val, parentElement)
End Function

'@Description("Finds the first element by the specified css selector value")
Public Function FindElementByCssSelector(ByVal val As String, Optional parentElement As WebElement) As WebElement
Attribute FindElementByCssSelector.VB_Description = "Finds the first element by the specified css selector value"
    Set FindElementByCssSelector = FindElement(By.cssSelector, val, parentElement)
End Function

'@Description("Finds the first element by the specified link text value")
Public Function FindElementByLinkText(ByVal val As String, Optional parentElement As WebElement) As WebElement
Attribute FindElementByLinkText.VB_Description = "Finds the first element by the specified link text value"
    Set FindElementByLinkText = FindElement(By.linkText, val, parentElement)
End Function

'@Description("Finds the first element by the specified name value")
Public Function FindElementByName(ByVal val As String, Optional parentElement As WebElement) As WebElement
Attribute FindElementByName.VB_Description = "Finds the first element by the specified name value"
    Set FindElementByName = FindElement(By.Name, val, parentElement)
End Function

'@Description("Finds the first element by the specified partial link text value")
Public Function FindElementByPartialLinkText(ByVal val As String, Optional parentElement As WebElement) As WebElement
Attribute FindElementByPartialLinkText.VB_Description = "Finds the first element by the specified partial link text value"
    Set FindElementByPartialLinkText = FindElement(By.PartialLinkText, val, parentElement)
End Function

'@Description("Finds the first element by the specified xpath query value")
Public Function FindElementByXPath(ByVal val As String, Optional parentElement As WebElement) As WebElement
Attribute FindElementByXPath.VB_Description = "Finds the first element by the specified xpath query value"
    Set FindElementByXPath = FindElement(By.XPath, val, parentElement)
End Function

'@Description("Finds all elements that satisfy the specified xpath query value")
Public Function FindElementsByXPath(ByVal val As String, Optional parentElement As WebElement) As WebElements
Attribute FindElementsByXPath.VB_Description = "Finds all elements that satisfy the specified xpath query value"
    Set FindElementsByXPath = FindElements(By.XPath, val, parentElement)
End Function

'@Description("Finds all elements that satisfy the specified id value")
Public Function FindElementsByID(ByVal val As String, Optional parentElement As WebElement) As WebElements
Attribute FindElementsByID.VB_Description = "Finds all elements that satisfy the specified id value"
    Set FindElementsByID = FindElements(By.ID, val, parentElement)
End Function

'@Description("Finds all elements that satisfy the specified tag name value")
Public Function FindElementsByTagName(ByVal val As String, Optional parentElement As WebElement) As WebElements
Attribute FindElementsByTagName.VB_Description = "Finds all elements that satisfy the specified tag name value"
    Set FindElementsByTagName = FindElements(By.tagName, val, parentElement)
End Function

'@Description("Finds all elements that satisfy the specified class name value")
Public Function FindElementsByClassName(ByVal val As String, Optional parentElement As WebElement) As WebElements
Attribute FindElementsByClassName.VB_Description = "Finds all elements that satisfy the specified class name value"
    Set FindElementsByClassName = FindElements(By.className, val, parentElement)
End Function

'@Description("Finds all elements that satisfy the specified css selector value")
Public Function FindElementsByCssSelector(ByVal val As String, Optional parentElement As WebElement) As WebElements
Attribute FindElementsByCssSelector.VB_Description = "Finds all elements that satisfy the specified css selector value"
    Set FindElementsByCssSelector = FindElements(By.cssSelector, val, parentElement)
End Function

'@Description("Finds all elements that satisfy the specified link text value")
Public Function FindElementsByLinkText(ByVal val As String, Optional parentElement As WebElement) As WebElements
Attribute FindElementsByLinkText.VB_Description = "Finds all elements that satisfy the specified link text value"
    Set FindElementsByLinkText = FindElements(By.linkText, val, parentElement)
End Function

'@Description("Finds all elements that satisfy the specified name value")
Public Function FindElementsByName(ByVal val As String, Optional parentElement As WebElement) As WebElements
Attribute FindElementsByName.VB_Description = "Finds all elements that satisfy the specified name value"
    Set FindElementsByName = FindElements(By.Name, val, parentElement)
End Function

'@Description("Finds all elements that satisfy the specified partial link text value")
Public Function FindElementsByPartialLinkText(ByVal val As String, Optional parentElement As WebElement) As WebElements
Attribute FindElementsByPartialLinkText.VB_Description = "Finds all elements that satisfy the specified partial link text value"
    Set FindElementsByPartialLinkText = FindElements(By.PartialLinkText, val, parentElement)
End Function

'@Description("Sends a key sequence to the specified element - see WebKeyboard class for how to send special keys")
Public Sub SendKeys(element As WebElement, ByVal keys As String, Optional ByVal clearBeforeTyping As Boolean = True)
Attribute SendKeys.VB_Description = "Sends a key sequence to the specified element - see WebKeyboard class for how to send special keys"
    Dim data As New Dictionary
    Dim chunk() As String
    Dim chars() As String
    
    data.Add "id", element.elementId
    If clearBeforeTyping = True Then execute tCMD.CMD_CLEAR_ELEMENT, data
    
    'split key sequence into chunks with special keys (pattern "\ue0*") by themselves
    ReDim chunk(0 To Len(keys) - 1)
    Dim i As Long
    i = -1
    Do Until keys = vbNullString
        Dim ipos As String
        ipos = InStr(keys, "\ue0")
        If ipos > 0 Then
            If ipos > 1 Then 'text to left is chunk
                i = i + 1
                chunk(i) = Left$(keys, ipos - 1)
            End If
            'now process special key
            i = i + 1
            chunk(i) = Mid$(keys, ipos, 6)
            keys = Right$(keys, Len(keys) - (ipos + 5))
        Else
            i = i + 1
            chunk(i) = keys
            keys = vbNullString
        End If
    Loop
    
    ReDim Preserve chunk(0 To i)
    
    'initialize param keys
    data.Add "text", vbNullString
    data.Add "value", vbNullString
    
    'loop through and send text chunks to target
    For i = 0 To UBound(chunk)
        If Left$(chunk(i), 4) = "\ue0" Then
            ReDim chars(0 To 0)
            chars(0) = chunk(i)
        Else
            ReDim chars(0 To Len(chunk(i)) - 1)
            Dim j As Long
            For j = 0 To Len(chunk(i)) - 1
                chars(j) = Mid$(chunk(i), j + 1, 1)
            Next j
        End If
        data("text") = chunk(i)
        data("value") = chars
        execute tCMD.CMD_SEND_KEYS_TO_ELEMENT, data
    Next i
End Sub

'@Description("Clicks the element")
Public Sub Click(Optional element As WebElement)
Attribute Click.VB_Description = "Clicks the element"
    Dim data As New Dictionary
    If Not element Is Nothing Then
        data.Add "id", element.elementId
        execute tCMD.CMD_CLICK_ELEMENT, data
    Else
        execute tCMD.CMD_CLICK, data
    End If
End Sub

'@Description("Submits a form element")
Public Sub Submit(element As WebElement)
Attribute Submit.VB_Description = "Submits a form element"
    Click element
End Sub

'@Description("Gets the text of the element - set optional visibleOnly to False to return text whether hidden or not")
Public Function GetText(element As WebElement, Optional ByVal visibleOnly As Boolean = True) As String
Attribute GetText.VB_Description = "Gets the text of the element - set optional visibleOnly to False to return text whether hidden or not"
    If visibleOnly Then 'default behavior of W3C GetText
        Dim data As New Dictionary
        data.Add "id", element.elementId
        GetText = execute(tCMD.CMD_GET_ELEMENT_TEXT, data)("value")
    Else 'get the underlying text whether visible or not
        GetText = ExecuteScript("return arguments[0].textContent;", element)
    End If
End Function

'@Description("Simulates clicking the back arrow on the browser window")
Public Sub GoBack()
Attribute GoBack.VB_Description = "Simulates clicking the back arrow on the browser window"
    execute tCMD.CMD_GO_BACK, New Dictionary
End Sub

'@Description("Simulates clicking the forward arrow on the browser window")
Public Sub GoForward()
Attribute GoForward.VB_Description = "Simulates clicking the forward arrow on the browser window"
    execute tCMD.CMD_GO_FORWARD, New Dictionary
End Sub

'@Description("Refreshes the current url")
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refreshes the current url"
    execute tCMD.CMD_REFRESH, New Dictionary
End Sub

'@Description("Returns the current url string")
Public Function GetCurrentURL() As String
Attribute GetCurrentURL.VB_Description = "Returns the current url string"
    Dim encodeURL As String
    encodeURL = execute(tCMD.CMD_GET_CURRENT_URL, New Dictionary)("value")
    GetCurrentURL = ExecuteScript("return decodeURI('" & encodeURL & "')")
End Function

'@Description("Returns the active element")
Public Function GetActiveElement() As WebElement
Attribute GetActiveElement.VB_Description = "Returns the active element"
    Dim elem As WebElement
    
    Set elem = toWebElement(execute(tCMD.CMD_GET_ACTIVE_ELEMENT, New Dictionary)("value")(ELEMENT_KEY))
    
    If Not setHightlightFoundElems_Settings Is Nothing Then setHightlightFoundElems_exec elem
    
    Set GetActiveElement = elem
End Function

'@Description("Gets the property state of the element, such as the value of an input box - if the property does not exist, then it returns vbNullString")
Public Function GetProperty(ByVal propName As String, element As WebElement) As String
Attribute GetProperty.VB_Description = "Gets the property state of the element, such as the value of an input box - if the property does not exist, then it returns vbNullString"
    'This returns the current state of a DOM node object, such as the value of an input box.
    'If the property does not exist, then this returns vbNullString
    Dim data As New Dictionary, resp As Dictionary
    data.Add "name", propName
    data.Add "id", element.elementId
    
    Set resp = execute(tCMD.CMD_GET_ELEMENT_PROPERTY, data)
    
    If IsNull(resp("value")) Then
        GetProperty = vbNullString
    Else
        GetProperty = resp("value")
    End If
End Function

'@Description("Highlights/unhighlights specified element(s) with color")
Public Sub SetHighlight(ByVal onOffSwitch As Boolean, elementOrElements As Object, _
                Optional ByVal borderSizePx As Byte = 4, Optional ByVal borderColor As VBAcolors = Blue, _
                Optional ByVal backgroundColor As VBAcolors = VBAcolors.Unchanged, _
                Optional ByVal bScrollIntoView As Boolean = True, Optional ByVal unHighlightLast As Boolean = True, _
                Optional ByVal bForgiving_UnHighlight As Boolean = False)
Attribute SetHighlight.VB_Description = "Highlights/unhighlights specified element(s) with color"

    Dim elements As WebElements
    Dim element As WebElement
    
    Set elements = objToWebElements(elementOrElements)

    If onOffSwitch = True Then
        If lastElems Is Nothing Then Set lastElems = New WebElements
        If SetHighlight_ElemsOrigStyle Is Nothing Then Set SetHighlight_ElemsOrigStyle = New Scripting.Dictionary

        'handling of last highlighted elements
        If unHighlightLast = True And lastElems.Count > 0 Then
            SetHighlight False, lastElems, bForgiving_UnHighlight:=True
            lastElems.RemoveAll
        End If
        
        For Each element In elements
            If Not lastElems.Exists(element) Then lastElems.Add element
        Next element

        'Create the JS string for the highlighting SetAttribute
        Dim borderColorString As String, backgroundColorString As String
        If (borderSizePx = 0 Or borderColor = VBAcolors.Unchanged) And backgroundColor = VBAcolors.Unchanged Then
            Err.Raise 1, "WebDriver", "Error: Highlight function used with no border and no background specified. Choose at least one or leave defaults."
        End If
        If borderSizePx > 0 And borderColor <> VBAcolors.Unchanged Then
            borderColorString = "border:" & CStr(borderSizePx) & "px solid " & rgbLongToString(borderColor)
        End If
        If backgroundColor <> VBAcolors.Unchanged Then
            backgroundColorString = "background:" & rgbLongToString(backgroundColor)
        End If
        
        Dim finalString As String
        finalString = IIf(borderColorString <> vbNullString, borderColorString & "; ", vbNullString) & backgroundColorString

        For Each element In elements
            If bScrollIntoView And (element Is elements.Item(1)) Then element.ScrollIntoView
            'Dim maxYcoord As Long
            'maxYcoord = Application.WorksheetFunction.Max(maxYcoord, element.GetRect("y"))

            'store original style of element for future SetUnHighlight
            If Not SetHighlight_ElemsOrigStyle.Exists(element.elementId) Then
                SetHighlight_ElemsOrigStyle.Add element.elementId, element.GetAttribute("style")
            End If

            'highlight element
            element.SetAttribute "style", finalString
        Next
        'draft string (incomplete code) to highlight WebElements by single command:
        'ExecuteScript "arguments[0].setAttribute('style', arguments[1])", elements, finalString
        'MsgBox maxYcoord
        'Stop
    Else 'if UnHighlight
        For Each element In elements
            If SetHighlight_ElemsOrigStyle.Exists(element.elementId) Then
                If bForgiving_UnHighlight = True Then On Error GoTo ErrHnd
                element.SetAttribute "style", SetHighlight_ElemsOrigStyle.Item(element.elementId)
                If bForgiving_UnHighlight = True Then On Error GoTo 0
                SetHighlight_ElemsOrigStyle.Remove element.elementId
            End If
        Next
    End If
    Exit Sub
ErrHnd:
If Err.Number = 404 Then 'Forgiving only with 404 (Not Found) error types
    Resume Next
Else
    With Err: Err.Raise .Number, .Source, .Description, .HelpFile, .HelpContext: End With
End If
End Sub

'@Description("Highlights/unhighlights future found element(s) with color")
Public Sub SetHightlightFoundElems(ByVal onOffSwitch As Boolean, Optional ByVal elemBorderSizePx As Byte = 4, _
                Optional ByVal elemBorderColor As VBAcolors = Blue, Optional ByVal elemBackgroundColor As VBAcolors = VBAcolors.Unchanged, _
                Optional ByVal elemsBorderSizePx As Byte = 4, _
                Optional ByVal elemsBorderColor As VBAcolors = Cyan, Optional ByVal elemsBackgroundColor As VBAcolors = VBAcolors.Unchanged)
Attribute SetHightlightFoundElems.VB_Description = "Highlights/unhighlights future found element(s) with color"
    If onOffSwitch = True Then
        'store border and color settings:
        If setHightlightFoundElems_Settings Is Nothing Then Set setHightlightFoundElems_Settings = New Collection
        With setHightlightFoundElems_Settings
            .Add elemBorderSizePx:  .Add elemBorderColor:  .Add elemBackgroundColor:
            .Add elemsBorderSizePx: .Add elemsBorderColor: .Add elemsBackgroundColor:
        End With
    Else 'when set to False, clear any active Highlight
        If Not lastElems Is Nothing Then
            If lastElems.Count > 0 Then
                Dim elem As WebElement
                For Each elem In lastElems
                    SetHighlight False, elem, bForgiving_UnHighlight:=True
                Next
            End If
        End If
        Set setHightlightFoundElems_Settings = Nothing
    End If
End Sub

'@Description("Returns a boolean indicating if the current element has the specified property")
Public Function HasProperty(ByVal propName As String, element As WebElement) As Boolean
Attribute HasProperty.VB_Description = "Returns a boolean indicating if the current element has the specified property"
    If GetProperty(propName, element) = vbNullString Then HasProperty = False Else HasProperty = True
End Function

'@Description("Returns a boolean indicating if the element has the specified attribute")
Public Function HasAttribute(ByVal attrName As String, element As WebElement) As Boolean
Attribute HasAttribute.VB_Description = "Returns a boolean indicating if the element has the specified attribute"
    If GetAttribute(attrName, element) = vbNullString Then HasAttribute = False Else HasAttribute = True
End Function

'@Description("Gets the specified attribute of the element")
Public Function GetAttribute(ByVal attrName As String, element As WebElement) As String
Attribute GetAttribute.VB_Description = "Gets the specified attribute of the element"
    'This returns pre-defined HTML element attributes. Once the browser parses the html code,
    'then a DOM node object with properties is created.
    'To access the state of a DOM node property - such as an input box value, use GetProperty.
    'If the attribute does not exist, then this returns vbNullString
    Dim data As New Dictionary
    Dim resp As Dictionary
    
    data.Add "id", element.elementId
    data.Add "name", attrName
    
    Set resp = execute(tCMD.CMD_GET_ELEMENT_ATTRIBUTE, data)
    If IsNull(resp("value")) Then
        GetAttribute = vbNullString
    Else
        GetAttribute = resp("value")
    End If
End Function

'@Description("Sets the attribute of the element")
Public Sub SetAttribute(element As WebElement, ByVal attrName As String, ByVal val As String)
Attribute SetAttribute.VB_Description = "Sets the attribute of the element"
    'this sets the value of an html element attribute. Consider SendKeys or Action Chain for more control of user input elements.
    'If the attribute does not exist, then this does not throw an error.
    ExecuteScript "arguments[0].setAttribute('" & attrName & "', arguments[1])", element, val
End Sub

'@Description("Gets the Arial role of the element")
Public Function GetAriaRole(element As WebElement) As String
Attribute GetAriaRole.VB_Description = "Gets the Arial role of the element"
    Dim data As New Dictionary
    Dim resp As Dictionary
    
    data.Add "id", element.elementId
    Set resp = execute(tCMD.CMD_GET_ELEMENT_ARIA_ROLE, data)
    If IsNull(resp("value")) Then
        GetAriaRole = vbNullString
    Else
        GetAriaRole = resp("value")
    End If
End Function

'@Description("Gets the Arial label of the element")
Public Function GetAriaLabel(element As WebElement) As String
Attribute GetAriaLabel.VB_Description = "Gets the Arial label of the element"
    Dim data As New Dictionary
    Dim resp As Dictionary
    
    data.Add "id", element.elementId
    Set resp = execute(tCMD.CMD_GET_ELEMENT_ARIA_LABEL, data)
    If IsNull(resp("value")) Then
        GetAriaLabel = vbNullString
    Else
        GetAriaLabel = resp("value")
    End If
End Function

'@Description("Gets the tag name of the element")
Public Function GetTagName(element As WebElement) As String
Attribute GetTagName.VB_Description = "Gets the tag name of the element"
    Dim data As New Dictionary
    data.Add "id", element.elementId
    GetTagName = execute(tCMD.CMD_GET_ELEMENT_TAG_NAME, data)("value")
End Function

'@Description("Selects a dropdown option element with the specified value attribute")
Public Sub SelectByValue(ByVal val As String, element As WebElement)
Attribute SelectByValue.VB_Description = "Selects a dropdown option element with the specified value attribute"
    Dim elem As WebElement
    
    Set elem = FindElement(By.tagName, "option[value =" & """" & val & """" & "]", element)
    If Not IsSelected(elem) Then
        elem.Click
    End If
End Sub

'@Description("Selects a dropdown option element with the specified visible text")
Public Sub SelectByVisibleText(ByVal text As String, element As WebElement)
Attribute SelectByVisibleText.VB_Description = "Selects a dropdown option element with the specified visible text"
    Dim elem As WebElement

    Set elem = FindElement(By.XPath, ".//option[normalize-space(.) = " & """" & text & """" & "]", element)
    If Not IsSelected(elem) Then
        elem.Click
    End If
End Sub

'@Description("Returns all selected option elements for the specified dropdown element")
Public Function GetAllSelectedOptions(element As WebElement) As WebElements
Attribute GetAllSelectedOptions.VB_Description = "Returns all selected option elements for the specified dropdown element"
    Dim elems As WebElements
    Dim selectedElems As New WebElements
    Dim i As Long

    Set elems = FindElements(By.cssSelector, "option", element)
    For i = 1 To elems.Count
        If IsSelected(elems(i)) Then selectedElems.Add elems(i)
    Next i
    
    Set GetAllSelectedOptions = selectedElems
End Function

'@Description("Returns the first selected option element for the specified dropdown element")
Public Function GetSelectedOption(element As WebElement) As WebElement
Attribute GetSelectedOption.VB_Description = "Returns the first selected option element for the specified dropdown element"
    Dim elems As WebElements
    Dim selectedElem As WebElement
    Dim i As Long

    Set elems = FindElements(By.cssSelector, "option", element)
    For i = 1 To elems.Count
        If IsSelected(elems(i)) Then
            Set selectedElem = elems(i)
            Exit For
        End If
    Next i
    
    Set GetSelectedOption = selectedElem
End Function

'@Description("Deselects the dropdown option element having the specified value attribute")
Public Sub DeSelectByValue(ByVal val As String, element As WebElement)
Attribute DeSelectByValue.VB_Description = "Deselects the dropdown option element having the specified value attribute"
    Dim elem As WebElement

    Set elem = FindElement(By.cssSelector, "option[value =" & """" & val & """" & "]", element)
    If IsSelected(elem) Then
        elem.Click
    End If
End Sub

'@Description("Deselects the dropdown option element having the specified visible text")
Public Sub DeSelectByVisibleText(ByVal text As String, element As WebElement)
Attribute DeSelectByVisibleText.VB_Description = "Deselects the dropdown option element having the specified visible text"
    Dim elem As WebElement
    
    Set elem = FindElement(By.XPath, ".//option[normalize-space(.) = " & """" & text & """" & "]", element)
    If IsSelected(elem) Then
        elem.Click
    End If
End Sub

'@Description("Returns a boolean indicating if the dropdown option element is selected")
Public Function IsSelected(element As WebElement) As Boolean
Attribute IsSelected.VB_Description = "Returns a boolean indicating if the dropdown option element is selected"
    Dim data As New Dictionary
    data.Add "id", element.elementId
    IsSelected = execute(tCMD.CMD_IS_ELEMENT_SELECTED, data)("value")
End Function

'@Description("Returns a boolean indicating if the element is a multi-select dropdown")
Public Function IsMultiSelect(selectElement As WebElement) As Boolean
Attribute IsMultiSelect.VB_Description = "Returns a boolean indicating if the element is a multi-select dropdown"
    IsMultiSelect = HasAttribute("multiple", selectElement)
End Function

'@Description("Selects a dropdown option element with the specified index")
Public Sub SelectByIndex(ByVal index As Integer, element As WebElement)
Attribute SelectByIndex.VB_Description = "Selects a dropdown option element with the specified index"
    SetFocus element
    ExecuteScript "arguments[0].options[" & CStr(index - baseDomIdx) & "].selected = true;", element
End Sub

'@Description("Selects all multi-select dropdown option elements")
Public Sub SelectAll(element As WebElement)
Attribute SelectAll.VB_Description = "Selects all multi-select dropdown option elements"
    Dim elems As WebElements
    Dim i As Integer
    Set elems = FindElements(By.cssSelector, "option", element)
    SetFocus element
    For i = 1 To elems.Count
        ExecuteScript "arguments[0].options[" & CStr(i - 1) & "].selected = true;", element
    Next i
End Sub

'@Description("Deselects all dropdown option elements")
Public Sub DeSelectAll(element As WebElement)
Attribute DeSelectAll.VB_Description = "Deselects all dropdown option elements"
    Dim elems As WebElements
    Dim i As Integer
    Set elems = FindElements(By.cssSelector, "option", element)
    SetFocus element
    For i = 1 To elems.Count
        ExecuteScript "arguments[0].options[" & CStr(i - 1) & "].selected = false;", element
    Next i
End Sub

'@Description("Deselects a dropdown option element with the specified index")
Public Sub DeSelectByIndex(ByVal index As Integer, element As WebElement)
Attribute DeSelectByIndex.VB_Description = "Deselects a dropdown option element with the specified index"
    Dim elems As WebElements
    Set elems = FindElements(By.cssSelector, "option", element)
    If IsSelected(elems(index - baseDomIdx + 1)) Then
        ExecuteScript "return arguments[0].selectedIndex = " & CStr(index - baseDomIdx), elems(index - baseDomIdx + 1)
        elems(index - baseDomIdx + 1).Click
    End If
End Sub

'@Description("Executes a java script synchronously")
Public Function ExecuteScript(ByVal script As String, ParamArray argumentList() As Variant) As Variant
Attribute ExecuteScript.VB_Description = "Executes a java script synchronously"
    Dim data As New Dictionary
    Dim resp As Dictionary
    Dim args() As Variant

    data.Add "script", script
    
    'replace WebElement, WebElements, WebShadowRoot object arguments with format needed by WebDriver server
    If UBound(argumentList) >= 0 Then
        args = argumentList
        args = deObjectifyInputArgs(args)
    End If
    
    data.Add "args", args
    
    Set resp = execute(tCMD.CMD_EXECUTE_SCRIPT, data)
    
    'replace the response representation of an element, a collection of elements, and a shadow root
    'with a WebElement object, a WebElements object, and a WebShadowRoot object, respectively
    If IsObject(resp("value")) Then
        Set ExecuteScript = objectifyScriptResponse(resp("value"))
    Else
        ExecuteScript = resp("value")
    End If
End Function

'@Description("Executes a java script asynchronously")
Public Function ExecuteScriptAsync(ByVal script As String, ParamArray argumentList() As Variant) As Variant
Attribute ExecuteScriptAsync.VB_Description = "Executes a java script asynchronously"
    Dim data As New Dictionary
    Dim resp As Dictionary
    Dim args() As Variant

    data.Add "script", script
    
    'replace WebElement, WebElements, WebShadowRoot object arguments with format needed by WebDriver server
    If UBound(argumentList) >= 0 Then
        args = argumentList
        args = deObjectifyInputArgs(args)
    End If
    
    data.Add "args", args
    
    Set resp = execute(tCMD.CMD_EXECUTE_SCRIPT_ASYNC, data)
    
    'replace the response representation of an element, a collection of elements, and a shadow root
    'with a WebElement object, a WebElements object, and a WebShadowRoot object, respectively
    If IsObject(resp("value")) Then
        Set ExecuteScriptAsync = objectifyScriptResponse(resp("value"))
    Else
        ExecuteScriptAsync = resp("value")
    End If
End Function

'@Description("Saves a screenshot of the active browser window to a file")
Public Sub SaveScreenshot(Optional ByVal filePath As String = ".\screenshot.png", Optional element As WebElement)
Attribute SaveScreenshot.VB_Description = "Saves a screenshot of the active browser window to a file"
    Dim data As New Dictionary
    Dim pngstr As String
        
    filePath = WebShared.getFullLocalPath(filePath, defaultIOFolder_)
    
    If element Is Nothing Then
        pngstr = execute(tCMD.CMD_SCREENSHOT, data)("value")
    Else
        data.Add "id", element.elementId
        pngstr = execute(tCMD.CMD_ELEMENT_SCREENSHOT, data)("value")
    End If
    
    saveByteArrayToFile decodeBase64(pngstr), filePath
End Sub

'@Description("Prints the current webpage window to a specified pdf file path")
Public Sub PrintToPDF(Optional ByVal filePath As String = ".\printpage.pdf", Optional settings As WebPrintSettings)
Attribute PrintToPDF.VB_Description = "Prints the current webpage window to a specified pdf file path"
    Dim data As New Dictionary
    Dim pdfstr As String
    
    filePath = WebShared.getFullLocalPath(filePath, defaultIOFolder_)
    
    If Not settings Is Nothing Then
        Set data = settings.ToDictionary
    End If
    
    pdfstr = execute(tCMD.CMD_PRINT_PAGE, data)("value")

    saveByteArrayToFile decodeBase64(pdfstr), filePath
End Sub

'@Description("Returns a dictionary containing x, y, width, height of the element")
Public Function GetRect(element As WebElement) As Dictionary
Attribute GetRect.VB_Description = "Returns a dictionary containing x, y, width, height of the element"
    'returns dictionary with following keys x, y, width, height
    Dim data As New Dictionary
    data.Add "id", element.elementId
    Set GetRect = execute(tCMD.CMD_GET_ELEMENT_RECT, data)("value")
End Function

'@Description("Returns the current webpage source as a string")
Public Function GetPageSource() As String
Attribute GetPageSource.VB_Description = "Returns the current webpage source as a string"
    Dim data As New Dictionary
    GetPageSource = execute(tCMD.CMD_GET_PAGE_SOURCE, data)("value")
End Function

'@Description("Returns the active window's handle as a string")
Public Function GetCurrentWindowHandle() As String
Attribute GetCurrentWindowHandle.VB_Description = "Returns the active window's handle as a string"
    GetCurrentWindowHandle = execute(tCMD.CMD_GET_CURRENT_WINDOW_HANDLE, New Dictionary)("value")
End Function

'@Description("Returns all windows' handles as a string array")
Public Function GetWindowHandles() As String()
Attribute GetWindowHandles.VB_Description = "Returns all windows' handles as a string array"
    Dim handles As Collection
    Dim ret() As String
    Dim i As Integer
    
    Set handles = execute(tCMD.CMD_GET_WINDOW_HANDLES, New Dictionary)("value")
    
    'To array of handles
    ReDim ret(baseArrayIdx To handles.Count - 1 + baseArrayIdx)
    For i = 1 To handles.Count   ' handles is Collection, not array
        ret(i - 1 + baseArrayIdx) = handles.Item(i)
    Next
    
    GetWindowHandles = ret
End Function

'@Description("Opens and activates a new window of the specified type")
Public Function SwitchToNewWindow(Optional ByVal windowType As svbaWindowType = svbaTab) As String
Attribute SwitchToNewWindow.VB_Description = "Opens and activates a new window of the specified type"
    Dim data As New Dictionary
    Dim hndl As String
    Dim wtype As String
    
    If windowType = svbaTab Then wtype = "tab" Else wtype = "window"
    
    data.Add "type", wtype 'tab or window
    hndl = execute(tCMD.CMD_NEW_WINDOW, data)("value")("handle")
    
    SwitchToWindow hndl 'above just creates a new window - we still need to switch to it
    
    SwitchToNewWindow = hndl
End Function

'@Description("Activates the window of the specified index or handle string")
Public Function SwitchToWindow(ByVal handleOrIndex As Variant) As String
Attribute SwitchToWindow.VB_Description = "Activates the window of the specified index or handle string"
    Dim hndl As String
    Dim data As New Dictionary
    If IsNumeric(handleOrIndex) Then
        hndl = GetWindowHandles()(handleOrIndex)
    Else
        hndl = handleOrIndex
    End If
    
    data.Add "handle", hndl
    execute tCMD.CMD_SWITCH_TO_WINDOW, data
    
    SwitchToWindow = hndl
End Function

'@Description("Closes the window of the specified index or handle string - if no window is specified, then the active window is closed")
Public Sub CloseWindow(Optional ByVal handleOrIndex As Variant)
Attribute CloseWindow.VB_Description = "Closes the window of the specified index or handle string - if no window is specified, then the active window is closed"
    If Not IsMissing(handleOrIndex) Then SwitchToWindow handleOrIndex
    execute tCMD.CMD_CLOSE, New Dictionary
End Sub

'@Description("Returns a dictionary containing x, y, width, and height in screen pixels of the active window")
Public Function GetWindowRect() As Dictionary
Attribute GetWindowRect.VB_Description = "Returns a dictionary containing x, y, width, and height in screen pixels of the active window"
    Set GetWindowRect = execute(tCMD.CMD_GET_WINDOW_RECT, New Dictionary)("value")
End Function

'@Description("Sets the position and size in screen pixels for the active window")
Public Function SetWindowRect(Optional ByVal x As Variant, Optional ByVal y As Variant, Optional ByVal width As Variant, Optional ByVal height As Variant) As Dictionary
Attribute SetWindowRect.VB_Description = "Sets the position and size in screen pixels for the active window"
    Dim data As New Dictionary
    Dim rect As Dictionary
    Set rect = GetWindowRect
    If IsMissing(width) Then width = rect("width")
    If IsMissing(height) Then height = rect("height")
    If IsMissing(x) Then x = rect("x")
    If IsMissing(y) Then y = rect("y")
    data.Add "x", x
    data.Add "y", y
    data.Add "width", width
    data.Add "height", height
    Set SetWindowRect = execute(tCMD.CMD_SET_WINDOW_RECT, data)("value")
End Function

'@Description("Sets the size in screen pixels for the active window")
Public Sub SetWindowSize(Optional ByVal width As Variant, Optional ByVal height As Variant)
Attribute SetWindowSize.VB_Description = "Sets the size in screen pixels for the active window"
    Dim rect As Dictionary
    Set rect = GetWindowRect
    If IsMissing(width) Then width = rect("width")
    If IsMissing(height) Then height = rect("height")
    SetWindowRect rect("x"), rect("y"), width, height
End Sub

'@Description("Sets the position in screen pixels for the active window")
Public Sub SetWindowPosition(Optional ByVal x As Variant, Optional ByVal y As Variant)
Attribute SetWindowPosition.VB_Description = "Sets the position in screen pixels for the active window"
    'returns dictionary with following keys x, y, width, height
    Dim rect As Dictionary
    Set rect = GetWindowRect
    If IsMissing(x) Then x = rect("x")
    If IsMissing(y) Then y = rect("y")
    SetWindowRect x, y, rect("width"), rect("height")
End Sub

'@Description("Maximize the active window")
Public Sub MaximizeWindow()
Attribute MaximizeWindow.VB_Description = "Maximize the active window"
    execute tCMD.CMD_MAXIMIZE_WINDOW, New Dictionary
End Sub

'@Description("Sets the current window to full screen mode")
Public Sub FullScreenWindow()
Attribute FullScreenWindow.VB_Description = "Sets the current window to full screen mode"
    execute tCMD.CMD_FULLSCREEN_WINDOW, New Dictionary
End Sub

'@Description("Minimize the active window")
Public Sub MinimizeWindow()
Attribute MinimizeWindow.VB_Description = "Minimize the active window"
    execute tCMD.CMD_MINIMIZE_WINDOW, New Dictionary
End Sub

'@Description("Switches focus to the frame element")
Public Sub SwitchToFrame(frameElement As WebElement)
Attribute SwitchToFrame.VB_Description = "Switches focus to the frame element"
    Dim data As New Dictionary
    Dim elem As New Dictionary
    elem.Add ELEMENT_KEY, frameElement.elementId
    data.Add "id", elem
    execute tCMD.CMD_SWITCH_TO_FRAME, data
End Sub

'@Description("Switches focus to the frame indicated by the specified index")
Public Sub SwitchToFrameByIndex(ByVal index As Integer)
Attribute SwitchToFrameByIndex.VB_Description = "Switches focus to the frame indicated by the specified index"
    Dim data As New Dictionary
    data.Add "id", index - baseDomIdx
    execute tCMD.CMD_SWITCH_TO_FRAME, data
End Sub

'@Description("Switches focus to the frame of the main document")
Public Sub SwitchToDefaultContent()
Attribute SwitchToDefaultContent.VB_Description = "Switches focus to the frame of the main document"
    Dim data As New Dictionary
    data.Add "id", Null
    execute tCMD.CMD_SWITCH_TO_FRAME, data
End Sub

'@Description("Switches focus to the active frame's parent")
Public Sub SwitchToParentFrame()
Attribute SwitchToParentFrame.VB_Description = "Switches focus to the active frame's parent"
    'if top-level frame then this switches to main document
    execute tCMD.CMD_SWITCH_TO_PARENT_FRAME, New Dictionary
End Sub

'@Description("Returns the name of the active frame")
Public Function GetCurrentFrameName() As String
Attribute GetCurrentFrameName.VB_Description = "Returns the name of the active frame"
    GetCurrentFrameName = ExecuteScript("return self.name")
End Function

'@Description("Sets/Gets how long to wait for the element location strategy to complete when locating an element(s)")
Public Property Let ImplicitWait(ByVal milliseconds As Long)
Attribute ImplicitWait.VB_Description = "Sets/Gets how long to wait for the element location strategy to complete when locating an element(s)"
    Dim data As New Dictionary
    data.Add "implicit", milliseconds
    execute tCMD.CMD_SET_TIMEOUTS, data
End Property

Public Property Get ImplicitWait() As Long
    ImplicitWait = execute(tCMD.CMD_GET_TIMEOUTS, New Dictionary)("value")("implicit")
End Property

'@Description("Sets/Gets the timeout limit used to interrupt a navigation attempt")
Public Property Let PageLoadTimeout(ByVal milliseconds As Long)
Attribute PageLoadTimeout.VB_Description = "Sets/Gets the timeout limit used to interrupt a navigation attempt"
    Dim data As New Dictionary
    data.Add "pageLoad", milliseconds
    execute tCMD.CMD_SET_TIMEOUTS, data
End Property

Public Property Get PageLoadTimeout() As Long
    PageLoadTimeout = execute(tCMD.CMD_GET_TIMEOUTS, New Dictionary)("value")("pageLoad")
End Property

'@Description("Sets/Gets how long to wait to interrupt a script that is being evaluated")
Public Property Let ScriptTimeout(ByVal milliseconds As Long)
Attribute ScriptTimeout.VB_Description = "Sets/Gets how long to wait to interrupt a script that is being evaluated"
    Dim data As New Dictionary
    data.Add "script", milliseconds
    execute tCMD.CMD_SET_TIMEOUTS, data
End Property

Public Property Get ScriptTimeout() As Long
    ScriptTimeout = execute(tCMD.CMD_GET_TIMEOUTS, New Dictionary)("value")("script")
End Property

'@Description("Pauses in-line execution for the specified number of milliseconds")
Public Sub Wait(Optional ByVal milliseconds As Long = 300)
Attribute Wait.VB_Description = "Pauses in-line execution for the specified number of milliseconds"
    WebShared.sleep milliseconds
End Sub

'@Description("Returns a boolean indicating if the element is present. User can optionally set an implicit wait, and get a reference to the found element.")
Public Function IsPresent(by_ As By, ByVal val As String, Optional ByVal ImplicitWaitMS As Long = 0, _
                            Optional ByRef elemFound As WebElement, Optional ByRef parentElement As WebElement) As Boolean
Attribute IsPresent.VB_Description = "Returns a boolean indicating if the element is present. User can optionally set an implicit wait, and get a reference to the found element."
    Dim timeout As Long
    Dim resp As Dictionary
    Dim data As Dictionary
    Dim vCMD As Variant
    
    timeout = ImplicitWait()
    If timeout <> ImplicitWaitMS Then ImplicitWait = ImplicitWaitMS
    Set data = toSelector(by_, val)
    
    If Not parentElement Is Nothing Then
        data.Add "id", parentElement.elementId
        vCMD = tCMD.CMD_FIND_ELEMENT_FROM_ELEMENT
    Else
        vCMD = tCMD.CMD_FIND_ELEMENT
    End If
    
    Set resp = execute(vCMD, data, False) 'don't raise an error if not present
    If isResponseError(resp) Then
        IsPresent = False
        Set elemFound = Nothing
    Else
        IsPresent = True
        Set elemFound = toWebElement(resp("value")(ELEMENT_KEY))
    End If
    If timeout <> ImplicitWaitMS Then ImplicitWait = timeout
End Function

'@Description("Returns a boolean indicating if the element is enabled")
Public Function IsEnabled(element As WebElement) As Boolean
Attribute IsEnabled.VB_Description = "Returns a boolean indicating if the element is enabled"
    Dim data As New Dictionary
    data.Add "id", element.elementId
    IsEnabled = execute(tCMD.CMD_IS_ELEMENT_ENABLED, data)("value")
End Function

'@Description("Returns a boolean indicating if the element is displayed")
Public Function IsDisplayed(element As WebElement) As Boolean
Attribute IsDisplayed.VB_Description = "Returns a boolean indicating if the element is displayed"
    Dim data As New Dictionary
    data.Add "id", element.elementId
    IsDisplayed = execute(tCMD.CMD_IS_ELEMENT_DISPLAYED, data)("value")
End Function

'@Description("Gets the browser name")
Public Function GetBrowserName() As String
Attribute GetBrowserName.VB_Description = "Gets the browser name"
    GetBrowserName = WebShared.getBrowserNameString(browser_)
End Function

'@Description("Gets the text of the active alert")
Public Function GetAlertText() As String
Attribute GetAlertText.VB_Description = "Gets the text of the active alert"
    GetAlertText = execute(tCMD.CMD_GET_ALERT_TEXT, New Dictionary)("value")
End Function

'@Description("Sends the specified text to the active alert")
Public Sub SetAlertText(ByVal alertText As String)
Attribute SetAlertText.VB_Description = "Sends the specified text to the active alert"
    Dim data As New Dictionary
    data.Add "text", alertText
    execute tCMD.CMD_SET_ALERT_TEXT, data
End Sub

'@Description("Accepts the active alert")
Public Sub AcceptAlert()
Attribute AcceptAlert.VB_Description = "Accepts the active alert"
    execute tCMD.CMD_ACCEPT_ALERT, New Dictionary
End Sub

'@Description("Dismisses the active alert")
Public Sub DismissAlert()
Attribute DismissAlert.VB_Description = "Dismisses the active alert"
    execute tCMD.CMD_DISMISS_ALERT, New Dictionary
End Sub

'@Description("Returns a boolean indicating whether an active alert is present")
Public Function IsAlertPresent() As Boolean
Attribute IsAlertPresent.VB_Description = "Returns a boolean indicating whether an active alert is present"
    Dim resp As Dictionary
    Set resp = execute(tCMD.CMD_GET_ALERT_TEXT, New Dictionary, False)
    If isResponseError(resp) Then
        IsAlertPresent = False
    Else
        IsAlertPresent = True
    End If
End Function

'@Description("Clears the text field of the element")
Public Function Clear(element As WebElement) As WebElement
Attribute Clear.VB_Description = "Clears the text field of the element"
    Dim data As New Dictionary
    data.Add "id", element.elementId
    data.Add "text", vbNullString
    data.Add "value", vbNullString
    execute tCMD.CMD_CLEAR_ELEMENT, data
    Set Clear = element
End Function

'@Description("Gets the active browser window's title")
Public Function GetTitle() As String
Attribute GetTitle.VB_Description = "Gets the active browser window's title"
    GetTitle = execute(tCMD.CMD_GET_TITLE, New Dictionary)("value")
End Function

'@Description("Gets the shadow root hosted by the element")
Public Function GetShadowRoot(element As WebElement) As WebShadowRoot
Attribute GetShadowRoot.VB_Description = "Gets the shadow root hosted by the element"
    Dim data As New Dictionary
    Dim resp As Dictionary
    data.Add "id", element.elementId
    Set resp = execute(tCMD.CMD_GET_ELEMENT_SHADOW_ROOT, data)
    Set GetShadowRoot = toShadowRoot(resp("value")(SHADOWROOT_KEY))
End Function

'@Description("Gets the specified CSS property of the element")
Public Function GetCSSProperty(ByVal attrName As String, element As WebElement) As String
Attribute GetCSSProperty.VB_Description = "Gets the specified CSS property of the element"
    Dim data As New Dictionary
    data.Add "id", element.elementId
    data.Add "name", attrName
    GetCSSProperty = execute(tCMD.CMD_GET_ELEMENT_VALUE_OF_CSS_PROPERTY, data)("value")
End Function

'@Description("Gets the inner HTML of the element")
Public Function GetInnerHTML(Optional element As WebElement) As String
Attribute GetInnerHTML.VB_Description = "Gets the inner HTML of the element"
    Dim script As String
    If element Is Nothing Then
        script = "return document.documentElement.innerHTML"
        GetInnerHTML = ExecuteScript(script)
    Else
        script = "return arguments[0].innerHTML"
        GetInnerHTML = ExecuteScript(script, element)
    End If
End Function

'@Description("Gets the outer HTML of the element")
Public Function GetOuterHTML(Optional element As WebElement) As String
Attribute GetOuterHTML.VB_Description = "Gets the outer HTML of the element"
    Dim script As String
    If element Is Nothing Then
        script = "return document.documentElement.outerHTML"
        GetOuterHTML = ExecuteScript(script)
    Else
        script = "return arguments[0].outerHTML"
        GetOuterHTML = ExecuteScript(script, element)
    End If
End Function

'@Description("Creates a capabilities object - must be performed after initializing the driver using one of the Start* methods")
Public Function CreateCapabilities(Optional ByVal initializeFromSettingsFile As Boolean = True) As WebCapabilities
Attribute CreateCapabilities.VB_Description = "Creates a capabilities object - must be performed after initializing the driver using one of the Start* methods"
    Dim caps As New WebCapabilities
    
    'insure that this function is being called in the correct order
    If browser_ = 0 Or sessionId_ <> vbNullString Then
        Err.Raise 1, "WebDriver", "The ""CreateCapabilities"" Method must be invoked after the ""StartEdge"" or ""StartChrome"" Methods, and before ""OpenBrowser"" Method."
    End If
    
    'now initialize Capabilities
    caps.initializeFor browser_
    
    Set caps.WebDriver = Me
    
    If initializeFromSettingsFile Then
        Select Case browser_
        Case svbaBrowser.Chrome
            If preloadChromeCapabilitiesFilePath <> vbNullString Then caps.LoadFromFile preloadChromeCapabilitiesFilePath
        Case svbaBrowser.Edge
            If preloadEdgeCapabilitiesFilePath <> vbNullString Then caps.LoadFromFile preloadEdgeCapabilitiesFilePath
        Case svbaBrowser.Firefox
            If preloadFirefoxCapabilitiesFilePath <> vbNullString Then caps.LoadFromFile preloadFirefoxCapabilitiesFilePath
        Case svbaBrowser.IE
            If preloadIECapabilitiesFilePath <> vbNullString Then caps.LoadFromFile preloadIECapabilitiesFilePath
        End Select
    End If

    Set CreateCapabilities = caps
End Function

'@Description("Creates a cookies object - must be performed after initializing the driver using one of the Start* methods")
Public Function CreateCookies() As WebCookies
Attribute CreateCookies.VB_Description = "Creates a cookies object - must be performed after initializing the driver using one of the Start* methods"
    Dim cookies As New WebCookies
    Set cookies.WebDriver = Me
    Set CreateCookies = cookies
End Function

'@Description("Scrolls and set the focus to the element")
Public Function SetFocus(element As WebElement) As WebElement
Attribute SetFocus.VB_Description = "Scrolls and set the focus to the element"
    Dim script As String
    script = "arguments[0].focus({'preventScroll': arguments[1]})"
    ExecuteScript script, element, 0  'Scroll
    Set SetFocus = element
End Function

'@Description("Sets the focus to the element without scrolling")
Public Function SetFocusNoScroll(element As WebElement) As WebElement
Attribute SetFocusNoScroll.VB_Description = "Sets the focus to the element without scrolling"
    Dim script As String
    script = "arguments[0].focus({'preventScroll': arguments[1]})"
    ExecuteScript script, element, 1  'No Scroll
    Set SetFocusNoScroll = element
End Function

'@Description("Scrolls by an XY offset referenced to location of the element")
Public Function ScrollToElement(element As WebElement, Optional ByVal xOffset As Integer = 0, Optional ByVal yOffset As Integer = 0) As WebElement
Attribute ScrollToElement.VB_Description = "Scrolls by an XY offset referenced to location of the element"
    ScrollIntoView element
    ScrollBy xOffset, yOffset
    Set ScrollToElement = element
End Function

'@Description("Scrolls to the element")
Public Function ScrollIntoView(element As WebElement, _
                    Optional enSpeed As ScrollIntoViewOptions_speed, _
                    Optional enAlign_vert As ScrollIntoViewOptions_alignment, _
                    Optional enAlign_horiz As ScrollIntoViewOptions_alignment) As WebElement
Attribute ScrollIntoView.VB_Description = "Scrolls to the element"
    Dim script As String
    
    If enSpeed = 0 And enAlign_vert = 0 And enAlign_horiz = 0 Then
        script = "arguments[0].scrollIntoView();" 'Default, same of "arguments[0].scrollIntoView(true)"
    Else
        Dim sSpeed$, sAlign_vert$, sAlign_horiz$
        Select Case enSpeed
            Case 0: sSpeed = "auto" 'Default
            Case 1: sSpeed = "smooth"
            Case 2: sSpeed = "instant"
            Case 3: sSpeed = "auto"
        End Select
        Select Case enAlign_vert
            Case 0: sAlign_vert = "start" 'Default
            Case 1: sAlign_vert = "start"
            Case 2: sAlign_vert = "center"
            Case 3: sAlign_vert = "end"
            Case 4: sAlign_vert = "nearest"
        End Select
        Select Case enAlign_horiz
            Case 0: sAlign_horiz = "nearest" 'Default
            Case 1: sAlign_horiz = "start"
            Case 2: sAlign_horiz = "center"
            Case 3: sAlign_horiz = "end"
            Case 4: sAlign_horiz = "nearest"
        End Select
        script = "arguments[0].scrollIntoView({ behavior: """ & sSpeed & """, block: """ & sAlign_vert & """, inline: """ & sAlign_horiz & """ });"
    End If
    ExecuteScript script, element
    Set ScrollIntoView = element
End Function

'@Description("Scrolls to a x-y offset in pixels relative to the current window position")
Public Sub ScrollBy(Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0)
Attribute ScrollBy.VB_Description = "Scrolls to a x-y offset in pixels relative to the current window position"
    Dim script As String
    script = "window.scrollBy(" & xOffset & "," & yOffset & ");"
    ExecuteScript script
End Sub

'@Description("Scrolls to a x-y coordinate position in pixels relative to the current window viewport origin")
Public Sub ScrollTo(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0)
Attribute ScrollTo.VB_Description = "Scrolls to a x-y coordinate position in pixels relative to the current window viewport origin"
    Dim script As String
    script = "window.scrollTo(" & x & "," & y & ");"
    ExecuteScript script
End Sub

'@Description("Scrolls to the current window viewport origin")
Public Sub ScrollToTop()
Attribute ScrollToTop.VB_Description = "Scrolls to the current window viewport origin"
    ScrollTo 0, 0
End Sub

'@Description("Scrolls to the bottom of current window viewport")
Public Sub ScrollToBottom()
Attribute ScrollToBottom.VB_Description = "Scrolls to the bottom of current window viewport"
    Dim script As String
    script = "window.scrollTo(0, document.body.scrollHeight);"
    ExecuteScript script
End Sub

'@Description("Gets the heigth of the current window viewport")
Public Function GetScrollHeight() As Integer
Attribute GetScrollHeight.VB_Description = "Gets the heigth of the current window viewport"
    Dim script As String
    script = "return document.body.scrollHeight;"
    GetScrollHeight = ExecuteScript(script)
End Function

'@Description("Returns the collection of cookies for the current webpage")
Public Function GetAllCookies() As WebCookies
Attribute GetAllCookies.VB_Description = "Returns the collection of cookies for the current webpage"
    Dim resp As Collection
    Dim cookie As WebCookie
    Dim cookies As WebCookies
    Set resp = execute(tCMD.CMD_GET_ALL_COOKIES, New Dictionary)("value")
    'Dim jc As New WebJSonConverter
    'Debug.Print jc.ConvertToJson(resp, 4)
    Set cookies = Me.CreateCookies
    Dim i As Long
    For i = 1 To resp.Count
        Set cookie = New WebCookie
        If resp(i).Exists("name") Then cookie.Name = resp(i)("name")
        If resp(i).Exists("domain") Then cookie.Domain = resp(i)("domain")
        If resp(i).Exists("expiry") Then cookie.ExpiryUnix = resp(i)("expiry")
        If resp(i).Exists("httpOnly") Then cookie.HttpOnly = resp(i)("httpOnly")
        If resp(i).Exists("sameSite") Then cookie.SameSite = resp(i)("sameSite")
        If resp(i).Exists("secure") Then cookie.Secure = resp(i)("secure")
        If resp(i).Exists("value") Then cookie.Value = resp(i)("value")
        If resp(i).Exists("path") Then cookie.Path = resp(i)("path")
        cookies.Add cookie
    Next i
    Set GetAllCookies = cookies
End Function

'@Description("Returns a cookie having the specified name")
Public Function GetCookie(ByVal cookieName As String) As WebCookie
Attribute GetCookie.VB_Description = "Returns a cookie having the specified name"
    Dim resp As Dictionary
    Dim data As New Dictionary
    Dim cookie As New WebCookie
    data.Add "name", cookieName
    Set resp = execute(tCMD.CMD_GET_COOKIE, data)("value")
    If resp.Exists("name") Then cookie.Name = resp("name")
    If resp.Exists("domain") Then cookie.Domain = resp("domain")
    If resp.Exists("expiry") Then cookie.ExpiryUnix = resp("expiry")
    If resp.Exists("httpOnly") Then cookie.HttpOnly = resp("httpOnly")
    If resp.Exists("sameSite") Then cookie.SameSite = resp("sameSite")
    If resp.Exists("secure") Then cookie.Secure = resp("secure")
    If resp.Exists("value") Then cookie.Value = resp("value")
    If resp.Exists("path") Then cookie.Path = resp("path")
    Set GetCookie = cookie
End Function

'@Description("Deletes a cookie having the specified name")
Public Sub DeleteCookie(ByVal cookieName As String)
Attribute DeleteCookie.VB_Description = "Deletes a cookie having the specified name"
    Dim data As New Dictionary
    data.Add "name", cookieName
    execute tCMD.CMD_DELETE_COOKIE, data
End Sub

'@Description("Sets a cookie object")
Public Sub SetCookie(cookie As WebCookie)
Attribute SetCookie.VB_Description = "Sets a cookie object"
    Dim data As New Dictionary
    'value can take a number string greater than 15 digits
    'need to set the following in JSonConverter else  will get an error on execute (see sendRequest):
    'jc.OptionsUseDoubleForLargeNumbers = True
    If cookie.Name <> vbNullString And cookie.Value <> vbNullString Then
        data.Add "cookie", cookie.ToDictionary
        execute tCMD.CMD_ADD_COOKIE, data
    End If
End Sub

'@Description("Sets a cookies collection object")
Public Sub SetCookies(cookies As WebCookies)
Attribute SetCookies.VB_Description = "Sets a cookies collection object"
    Dim i As Long
    For i = 1 To cookies.Count
        SetCookie cookies.Item(i)
    Next i
End Sub

'@Description("Deletes all cookies")
Public Sub DeleteAllCookies()
Attribute DeleteAllCookies.VB_Description = "Deletes all cookies"
    execute tCMD.CMD_DELETE_ALL_COOKIES, New Dictionary
End Sub

'@Description("Checks if the current url is equal to the specified url")
Public Function IsPageFound(ByVal url As String) As Boolean
Attribute IsPageFound.VB_Description = "Checks if the current url is equal to the specified url"
    If GetCurrentURL() = url Then
        IsPageFound = True
    Else
        IsPageFound = False
    End If
End Function

'@Description("Returns a dictionary holding various info about the session")
Public Function GetSessionsInfo() As Dictionary
Attribute GetSessionsInfo.VB_Description = "Returns a dictionary holding various info about the session"
    'since only one browser per driver is allowed, this will return a collection of one
    'so lets return a more useful dictionary instead
    If browser_ = svbaBrowser.Firefox Then
        Err.Raise 1, "WebDriver", "GetSessionsInfo not supported for Firefox WebDriver"
    End If
    Set GetSessionsInfo = execute(tCMD.CMD_GET_ALL_SESSIONS)("value")(1)
End Function

'@Description("Returns a dictionary holding various info about the driver")
Public Function GetDriverStatus() As Dictionary
Attribute GetDriverStatus.VB_Description = "Returns a dictionary holding various info about the driver"
    Set GetDriverStatus = execute(tCMD.CMD_STATUS)("value")
    'DriverVersion = resp("build")("version")
    'DriverVersion = Trim$(Left$(DriverVersion, InStr(DriverVersion, "(") - 1))
    'resp("message") 'gets info "MSEdgeDriver ready for new sessions."
    'resp("os")("name") 'gets info "Windows NT"
    'resp("os")("version") 'gets info "10.0.19042"
    'resp("os")("arch") 'gets info "x86_64"
    'resp("ready")'true or false
End Function

'@Description("Creates an initialized action chain object")
Public Function ActionChain() As WebActionChain
Attribute ActionChain.VB_Description = "Creates an initialized action chain object"
    Set ActionChain = New WebActionChain
    'pass a reference of the driver to each link
    Set ActionChain.WebDriver = Me
    Set ActionChain.DataDictionary = New Dictionary
End Function

'@Description("Finds the first element in the shadow root for a given selector and value")
Public Function FindElementFromShadowRoot(by_ As By, ByVal val As String, shadowRoot As WebShadowRoot) As WebElement
Attribute FindElementFromShadowRoot.VB_Description = "Finds the first element in the shadow root for a given selector and value"
    Dim data As Dictionary
    Dim elem As WebElement
    
    Set data = toSelector(by_, val)
    data.Add "sid", shadowRoot.shadowRootId
    Set elem = toWebElement(execute(tCMD.CMD_FIND_ELEMENT_FROM_SHADOW_ROOT, data)("value")(ELEMENT_KEY))
    
    If Not setHightlightFoundElems_Settings Is Nothing Then setHightlightFoundElems_exec elem
    
    Set FindElementFromShadowRoot = elem
End Function

'@Description("Finds all elements in the shadow root for a given selector and value")
Public Function FindElementsFromShadowRoot(by_ As By, ByVal val As String, shadowRoot As WebShadowRoot) As WebElements
Attribute FindElementsFromShadowRoot.VB_Description = "Finds all elements in the shadow root for a given selector and value"
    Dim data As Dictionary
    Dim elements As Collection
    Dim elems As New WebElements
    Dim i As Integer
    
    Set data = toSelector(by_, val)
    data.Add "sid", shadowRoot.shadowRootId
    
    Set elements = execute(tCMD.CMD_FIND_ELEMENTS_FROM_SHADOW_ROOT, data)("value")
    
    For i = 1 To elements.Count
        elems.Add toWebElement(elements(i)(ELEMENT_KEY))
    Next i
    
    If Not setHightlightFoundElems_Settings Is Nothing Then setHightlightFoundElems_exec elems
    
    Set elems.WebDriver = Me
    
    Set FindElementsFromShadowRoot = elems
End Function

'@Description("Uploads the contents of a file to the element")
Public Sub UploadFile(element As WebElement, ByVal filePath As String)
Attribute UploadFile.VB_Description = "Uploads the contents of a file to the element"
    'this is just a special wrapper for sendkeys (upload file mode)
    filePath = WebShared.getFullLocalPath(filePath, defaultIOFolder_)
    SendKeys element, filePath
End Sub

'@Description("Drags and drops the element to the specified target element")
Public Sub DragAndDrop(sourceElement As WebElement, targetElement As WebElement)
Attribute DragAndDrop.VB_Description = "Drags and drops the element to the specified target element"
    ActionChain.DragAndDrop(sourceElement, targetElement).Perform
End Sub

'@Description("Waits until element is interactable (displayed)")
Public Function WaitUntilReady(element As WebElement, Optional ByVal maxWaitTimeMS As Long = -1) As WebElement
Attribute WaitUntilReady.VB_Description = "Waits until element is interactable (displayed)"
    'waits until element is interactable, returns the input element for further action
    'if maxWaitTimeMS is positive, then throw error if limit is exceeded before IsDisplayed condition is met
    'see https://www.w3.org/TR/webdriver/#element-displayedness
    Dim cTimeStart As Currency
    Dim cTimeNow As Currency
    Dim cMaxWaitTimeMS As Currency
    
    Const loopDelay As Long = 50 'in ms
    
    If maxWaitTimeMS >= 0 Then
        GetTime cTimeStart
        Static cPerSecond As Currency
        If cPerSecond = 0 Then GetFrequency cPerSecond
        cMaxWaitTimeMS = CCur(maxWaitTimeMS) * (cPerSecond / 1000)
    Else
        'set up to run until IsDisplayed condition is met (no max time limit)
        cMaxWaitTimeMS = 0
        cTimeNow = 0
    End If
    
    Do
        If element.IsDisplayed Then
            Set WaitUntilReady = element
            Exit Function
        End If
        If maxWaitTimeMS >= 0 Then GetTime cTimeNow
        
        sleep loopDelay
        DoEvents 'yield to other processes
    Loop Until (cTimeNow - cTimeStart) > cMaxWaitTimeMS
    
    Err.Raise 1, "WebDriver", "Error: maximum wait time exceeded while waiting for element to be ready."
End Function

'@Description("Waits until file is finished downloading")
Public Sub WaitForDownload(ByVal filePath As String, Optional ByVal maxWaitTimeMS As Long = -1)
Attribute WaitForDownload.VB_Description = "Waits until file is finished downloading"
    'adapted from WaitNewFile in https://github.com/florentbr/SeleniumBasic/tree/master/Examples/Excel
    'if maxWaitTimeMS is positive, then throw error if limit is exceeded before file ready condition is met
    Dim fso As New FileSystemObject
    Dim cTimeStart As Currency
    Dim cTimeNow As Currency
    Dim cMaxWaitTimeMS As Currency
    
    Const loopDelay As Long = 50 'in ms
    
    filePath = WebShared.getFullLocalPath(filePath, defaultIOFolder_)
    
    If maxWaitTimeMS >= 0 Then
        GetTime cTimeStart 'gets current time
        Static cPerSecond As Currency
        If cPerSecond = 0 Then GetFrequency cPerSecond
        cMaxWaitTimeMS = CCur(maxWaitTimeMS) * (cPerSecond / 1000)
    Else
        'set up to run until file created and closed condition is met (no max time limit)
        cMaxWaitTimeMS = 0
        cTimeNow = 0
    End If
    
    'loop until file exists
    Do
        If fso.FileExists(filePath) Then Exit Do
        sleep loopDelay
        DoEvents 'yield to other processes
        'check if max time as elapsed
        If maxWaitTimeMS >= 0 Then GetTime cTimeNow
        If (cTimeNow - cTimeStart) > cMaxWaitTimeMS Then GoTo MaxTimeExceeded
    Loop
    
    'the file exists but now make sure it is finished writing (closed)
    Do
        If FileClosed(filePath) Then Exit Sub
        sleep loopDelay
        DoEvents 'yield to other processes
        'check if max time as elapsed
        If maxWaitTimeMS >= 0 Then GetTime cTimeNow
    Loop Until (cTimeNow - cTimeStart) > cMaxWaitTimeMS
    
MaxTimeExceeded:
    Err.Raise 1, "WebDriver", "Error: maximum wait time exceeded while waiting for file download."
End Sub

'@Description("Waits until an element is not present anymore")
Public Sub WaitUntilNotPresent(by_ As By, ByVal val As String, Optional ByVal maxWaitTimeMS As Long = -1)
Attribute WaitUntilNotPresent.VB_Description = "Waits until an element is not present anymore"
    'if maxWaitTimeMS is positive, then throw error if limit is exceeded before not IsPresent condition is met
    Dim cTimeStart As Currency
    Dim cTimeNow As Currency
    Dim cMaxWaitTimeMS As Currency
    
    Const loopDelay As Long = 50 'in ms
    
    If maxWaitTimeMS >= 0 Then
        GetTime cTimeStart
        Static cPerSecond As Currency
        If cPerSecond = 0 Then GetFrequency cPerSecond
        cMaxWaitTimeMS = CCur(maxWaitTimeMS) * (cPerSecond / 1000)
    Else
        'set up to run until not IsPresent condition is met (no max time limit)
        cMaxWaitTimeMS = 0
        cTimeNow = 0
    End If
    
    Do
        If Not IsPresent(by_, val) Then Exit Sub
        If maxWaitTimeMS >= 0 Then GetTime cTimeNow
        
        sleep loopDelay
        DoEvents 'yield to other processes
    Loop Until (cTimeNow - cTimeStart) > cMaxWaitTimeMS
    
    Err.Raise 1, "WebDriver", "Error: maximum wait time exceeded while waiting for element to disappear."
End Sub

'@Description("Returns the user agent string - must be called after OpenBrowser")
Public Function GetUserAgent() As String
Attribute GetUserAgent.VB_Description = "Returns the user agent string - must be called after OpenBrowser"
    If sessionId_ <> vbNullString Then
        GetUserAgent = ExecuteScript("return navigator.userAgent;")
    Else
        Err.Raise 1, "WebDriver", "GetUserAgent must be executed after OpenBrowser!"
    End If
End Function

'@Description("Sets/Gets the default input/output folder")
Public Property Let DefaultIOFolder(ByVal folderPath As String)
Attribute DefaultIOFolder.VB_Description = "Sets/Gets the default input/output folder"
    defaultIOFolder_ = WebShared.getFullLocalPath(folderPath)
End Property

Public Property Get DefaultIOFolder() As String
    DefaultIOFolder = defaultIOFolder_
End Property

'@Description("Sets/Gets the default Selenium driver folder")
Public Property Let DefaultDriverFolder(ByVal folderPath As String)
Attribute DefaultDriverFolder.VB_Description = "Sets/Gets the default Selenium driver folder"
'Note: default folder is Downloads dir
    defaultDriverFolder_ = WebShared.getFullLocalPath(folderPath)
End Property

Public Property Get DefaultDriverFolder() As String
    DefaultDriverFolder = defaultDriverFolder_
End Property

'@Description("Deletes a list of files")
Public Sub DeleteFiles(ParamArray fileList() As Variant)
Attribute DeleteFiles.VB_Description = "Deletes a list of files"
    Dim fso As New FileSystemObject
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Dim strlist() As String
    For i = 0 To UBound(fileList)
        If IsArray(fileList(i)) Then
            For j = LBound(fileList(i)) To UBound(fileList(i))
                strlist = Split(fileList(i)(j), ",")
                For k = 0 To UBound(strlist)
                    On Error Resume Next 'file might not exists
                    fso.DeleteFile WebShared.getFullLocalPath(strlist(k), defaultIOFolder_), True
                    On Error GoTo 0
                Next k
            Next j
        Else
            strlist = Split(fileList(i), ",")
            For k = 0 To UBound(strlist)
                On Error Resume Next 'file might not exists 'file might not exists
                fso.DeleteFile WebShared.getFullLocalPath(strlist(k), defaultIOFolder_), True
                On Error GoTo 0
            Next k
        End If
    Next i
End Sub

'@Description("Saves a specified string to a file")
Public Sub SaveStringToFile(ByVal val As String, Optional ByVal filePath As String = ".\snippet.html", Optional ByVal formatUnicode16 As Boolean = True)
Attribute SaveStringToFile.VB_Description = "Saves a specified string to a file"
    'this is used to save test html snippets to file for subsequent browser navigation
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    filePath = WebShared.getFullLocalPath(filePath, defaultIOFolder_)
    Set ts = fso.CreateTextFile(filePath, True, formatUnicode16)
    ts.Write val
    ts.Close
End Sub

'@Description("Returns the parent of element")
Public Function GetParentElement(childElement As WebElement) As WebElement
Attribute GetParentElement.VB_Description = "Returns the parent of element"
    Set GetParentElement = childElement.FindElement(By.XPath, "./parent::*")
End Function

'@Description("Returns the direct descendant child elements of the element")
Public Function GetChildElements(parentElement As WebElement) As WebElements
Attribute GetChildElements.VB_Description = "Returns the direct descendant child elements of the element"
    'this gets the immediate descendant elements - note that this is not same as element.FindElements(by.tagName, "*")
    Set GetChildElements = parentElement.FindElements(By.XPath, "./child::*")
End Function

'@Description("Returns an array holding the text values from a table or tbody element")
Public Function TableToArray(table As WebElement, Optional ByVal skipHeader As Boolean = False, Optional ByVal skipFooter As Boolean = False, Optional ByVal createSpanData As Boolean = True) As Variant()
Attribute TableToArray.VB_Description = "Returns an array holding the text values from a table or tbody element"
    Dim htmlDoc As New HTMLDocument
    Dim htmlTbl As htmlTable

    'convert the WebElement table into an html dom table object
    'input WebElement can either be a table or a tbody element
    'using a html dom table for processing is greater than 25x faster than using Selenium selectors to process table
    
    If htmlDoc.documentMode >= 11 Then
        'this below is more direct and works in my Office 365, but appears to fail on some 64-bit systems, related to
        'FEATURE_BROWSER_EMULATION registry key setting.
        Set htmlTbl = htmlDoc.createElement("table")
        
        Select Case table.GetTagName
        Case "table"
            htmlTbl.innerHTML = table.GetInnerHTML
        Case "tbody"
            htmlTbl.innerHTML = table.GetOuterHTML
        Case Else 'report error
            Err.Raise 1, "WebDriver", "Error: input object to TableToArray method must be a table or tbody element."
        End Select
    Else
        'this is the work-around...
        Select Case table.GetTagName
        Case "table"
            htmlDoc.body.innerHTML = table.GetOuterHTML
        Case "tbody"
            htmlDoc.body.innerHTML = "<table>" & table.GetOuterHTML & "</table>"
        Case Else 'report error
            Err.Raise 1, "WebDriver", "Error: input object to TableToArray method must be a table or tbody element."
        End Select
        
        Set htmlTbl = htmlDoc.body.Children(0)
    End If

    TableToArray = htmlTableToArray(htmlTbl, skipHeader, skipFooter, createSpanData)
End Function

'@Description("Converts webpage to an html document object - use sanitize to remove online dynamic content while leaving the dom tree intact")
Public Function PageToHTMLDoc(Optional ByVal sanitize As Boolean = False) As MSHTML.HTMLDocument
Attribute PageToHTMLDoc.VB_Description = "Converts webpage to an html document object - use sanitize to remove online dynamic content while leaving the dom tree intact"
    'sanitize option attempts to remove all "online" dynamic content while leaving the dom tree intact
    Dim htmlDoc As New MSHTML.HTMLDocument
    Dim pageSource As String
    
    pageSource = Me.GetPageSource
    
    If sanitize Then pageSource = sanitizeHTML(pageSource)

    If htmlDoc.documentMode >= 11 Then
        htmlDoc.DocumentElement.innerHTML = pageSource
        Set PageToHTMLDoc = htmlDoc
        Exit Function
    Else
        'this is a hack that works for older documentModes
        Dim htmlDoc2 As Object
        Set htmlDoc2 = New MSHTML.HTMLDocument
        htmlDoc2.Open
        htmlDoc2.Write pageSource
        htmlDoc2.Close
        Set PageToHTMLDoc = htmlDoc2
    End If
End Function

'@Description("Converts webpage to an xml document object")
Public Function PageToXMLDoc() As MSXML2.DOMDocument60
Attribute PageToXMLDoc.VB_Description = "Converts webpage to an xml document object"
    Dim xmlDoc As New MSXML2.DOMDocument60
    xmlDoc.LoadXML getXMLStringFromPage()
    Set PageToXMLDoc = xmlDoc
End Function

'@Description("Converts webpage to a json dictionary or collection object")
Public Function PageToJSONObject() As Object
Attribute PageToJSONObject.VB_Description = "Converts webpage to a json dictionary or collection object"
    'will return a dictionary or collection depending on json source
    Dim jc As New WebJsonConverter
    Set PageToJSONObject = jc.ParseJson(getJSONStringFromPage())
End Function

'@Description("Converts webpage to an html file - use sanitize to remove online dynamic content while leaving the dom tree intact")
Public Sub PageToHTMLFile(ByVal filePath As String, Optional ByVal sanitize As Boolean = False, Optional ByVal formatUnicode16 As Boolean = True)
Attribute PageToHTMLFile.VB_Description = "Converts webpage to an html file - use sanitize to remove online dynamic content while leaving the dom tree intact"
    'sanitize option attempts to remove all "online" dynamic content while leaving the dom tree intact
    If sanitize Then
        Me.SaveStringToFile sanitizeHTML(Me.GetPageSource), filePath, formatUnicode16
    Else
        Me.SaveStringToFile Me.GetPageSource, filePath, formatUnicode16
    End If
End Sub

'@Description("Converts webpage to an xml file")
Public Sub PageToXMLFile(ByVal filePath As String, Optional ByVal prettyPrint As Boolean = False, Optional ByVal formatUnicode16 As Boolean = True)
Attribute PageToXMLFile.VB_Description = "Converts webpage to an xml file"
    Me.SaveStringToFile getXMLStringFromPage(prettyPrint), filePath, formatUnicode16
End Sub

'@Description("Converts webpage to a json file")
Public Sub PageToJSONFile(ByVal filePath As String, Optional ByVal prettyPrint As Boolean = False, Optional ByVal formatUnicode16 As Boolean = True)
Attribute PageToJSONFile.VB_Description = "Converts webpage to a json file"
    Me.SaveStringToFile getJSONStringFromPage(prettyPrint), filePath, formatUnicode16
End Sub

'@Description("Returns the current webpage source as a string")
Public Function PageToString() As String
Attribute PageToString.VB_Description = "Returns the current webpage source as a string"
    'for consistency...
    PageToString = Me.GetPageSource()
End Function

'@Description("Sets geolocation parameter values")
Public Sub SetGeolocation(ByVal latitude As Double, ByVal longitude As Double, Optional ByVal altitude As Double = 0#, Optional ByVal accuracy As Double = 100#)
Attribute SetGeolocation.VB_Description = "Sets geolocation parameter values"
    Dim data As New Dictionary
    Dim location As New Dictionary
    location.Add "latitude", latitude
    location.Add "longitude", longitude
    location.Add "altitude", altitude
    location.Add "accuracy", accuracy
    data.Add "location", location
    execute tCMD.CMD_SET_LOCATION, data
End Sub

'@Description("Gets geolocation parameter values as a dictionay object")
Public Function GetGeolocation() As Dictionary
Attribute GetGeolocation.VB_Description = "Gets geolocation parameter values as a dictionay object"
    Set GetGeolocation = execute(tCMD.CMD_GET_LOCATION, New Dictionary)("value")
End Function

'@Description("Creates/updates the optional INI settings file - set keepExistingValues to False to reset existing settings to system defaults")
Public Sub CreateSettingsFile(Optional ByVal keepExistingValues As Boolean = True)
Attribute CreateSettingsFile.VB_Description = "Creates/updates the optional INI settings file - set keepExistingValues to False to reset existing settings to system defaults"
    Dim iniFilePath As String
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim overWriteFile As Long
    Dim i As Long
    Dim useDefaultValue As Boolean
    Dim entryCol As New Collection
    
    useDefaultValue = Not keepExistingValues
    
    iniFilePath = getFullLocalPath(thisLibFolderPath) & "\SeleniumVBA.ini"
    
    entryCol.Add "# This settings file is completely optional. For it to have effect,"
    entryCol.Add "# it must be located in the same folder as the SeleniumVBA code"
    entryCol.Add "# library, and be named ""SeleniumVBA.ini""."
    entryCol.Add " "
    entryCol.Add "# If a value for an entry is not specified, then the system"
    entryCol.Add "# default value will be used."
    entryCol.Add " "
    entryCol.Add "# Note that all path-type entry values recognize the %[Environ]% syntax."
    entryCol.Add " "
    entryCol.Add "# A few useful Environ values for reference:"
    entryCol.Add " "
    entryCol.Add "# %USERPROFILE%=C:\Users\[user name]"
    entryCol.Add "# %APPDATA%=C:\Users\[user name]\AppData\Roaming"
    entryCol.Add "# %LOCALAPPDATA%=C:\Users\[user name]\AppData\Local"
    entryCol.Add "# %TEMP%=C:\Users\[user name]\AppData\Local\Temp"
    entryCol.Add " "
    
    entryCol.Add "[GENERAL]"
    entryCol.Add " "
    entryCol.Add "# The driver_location_folder system defaults to Downloads folder."
    entryCol.Add "# The default_io_folder system defaults to the active vba project's"
    entryCol.Add "# document location - leave this blank to use default."
    entryCol.Add "# Valid values for command_window_style are vbHide (default), "
    entryCol.Add "# vbNormalFocus, vbMinimizedFocus, vbMaximizedFocus, vbNormalNoFocus,"
    entryCol.Add "# and vbMinimizedNoFocus."
    entryCol.Add " "
    entryCol.Add "driver_location_folder=" & readIniFileEntry(iniFilePath, "GENERAL", "driver_location_folder", "%USERPROFILE%\Downloads", useDefaultValue)
    entryCol.Add "default_io_folder=" & readIniFileEntry(iniFilePath, "GENERAL", "default_io_folder", vbNullString, useDefaultValue)
    entryCol.Add "command_window_style=" & readIniFileEntry(iniFilePath, "GENERAL", "command_window_style", "vbHide", useDefaultValue)
    entryCol.Add " "
    
    entryCol.Add "[AUTO-DRIVER-UPDATE]"
    entryCol.Add " "
    entryCol.Add "# If auto_detect_and_update=True (system default) then everytime"
    entryCol.Add "# the WebDriver's Start* method is called, the Selenium driver's"
    entryCol.Add "# version is checked against the corresponding browser version."
    entryCol.Add "# If the driver is not compatible with browser, it will be updated."
    entryCol.Add "# min_compatibility_level determines trigger for updating an"
    entryCol.Add "# an out-of-date driver. System default is svbaBuildMajor."
    entryCol.Add "# Use svbaMinor for less frequent updating, and svbaExactMatch"
    entryCol.Add "# for more frequent updating."
    entryCol.Add " "
    entryCol.Add "auto_detect_and_update=" & readIniFileEntry(iniFilePath, "AUTO-DRIVER-UPDATE", "auto_detect_and_update", True, useDefaultValue)
    entryCol.Add "min_compatibility_level=" & readIniFileEntry(iniFilePath, "AUTO-DRIVER-UPDATE", "min_compat_level", "svbaBuildMajor", useDefaultValue)
    entryCol.Add " "
    
    entryCol.Add "# Below are browser-specific initializations."
    entryCol.Add "# To automatically initialize a set of capabilities each time the"
    entryCol.Add "# OpenBrowser method of WebDriver class is invoked, set the"
    entryCol.Add "# preload_capabilities_file_path entry to the path of a valid json"
    entryCol.Add "# capabilities file. Note that if preload_capabilities_file_path is"
    entryCol.Add "# set to a blank value, or the entry is missing or commented out,"
    entryCol.Add "# then this option is ignored. Use the SaveToFile method of the"
    entryCol.Add "# WebCapabilities class to save a default set of capabilities"
    entryCol.Add "# for pre-loading."
    entryCol.Add "# The system defaults for local_host_port:"
    entryCol.Add "# Chrome - 9515, Edge - 9516, Firefox - 4444, IE - 5555"
    entryCol.Add " "
    
    entryCol.Add "[CHROME]"
    entryCol.Add " "
    entryCol.Add "preload_capabilities_file_path=" & readIniFileEntry(iniFilePath, "CHROME", "preload_capabilities_file_path", vbNullString, useDefaultValue)
    entryCol.Add "local_host_port=" & readIniFileEntry(iniFilePath, "CHROME", "local_host_port", 9515, useDefaultValue)
    entryCol.Add " "

    entryCol.Add "[EDGE]"
    entryCol.Add " "
    entryCol.Add "preload_capabilities_file_path=" & readIniFileEntry(iniFilePath, "EDGE", "preload_capabilities_file_path", vbNullString, useDefaultValue)
    entryCol.Add "local_host_port=" & readIniFileEntry(iniFilePath, "EDGE", "local_host_port", 9516, useDefaultValue)
    entryCol.Add " "
    
    entryCol.Add "[FIREFOX]"
    entryCol.Add " "
    entryCol.Add "preload_capabilities_file_path=" & readIniFileEntry(iniFilePath, "FIREFOX", "preload_capabilities_file_path", vbNullString, useDefaultValue)
    entryCol.Add "local_host_port=" & readIniFileEntry(iniFilePath, "FIREFOX", "local_host_port", 4444, useDefaultValue)
    entryCol.Add " "
    
    entryCol.Add "[INTERNET EXPLORER]"
    entryCol.Add " "
    entryCol.Add "preload_capabilities_file_path=" & readIniFileEntry(iniFilePath, "INTERNET EXPLORER", "preload_capabilities_file_path", vbNullString, useDefaultValue)
    entryCol.Add "local_host_port=" & readIniFileEntry(iniFilePath, "INTERNET EXPLORER", "local_host_port", 5555, useDefaultValue)
    entryCol.Add " "
    
    entryCol.Add "[PDF_DEFAULT_PRINT_SETTINGS]"
    entryCol.Add " "
    entryCol.Add "# Valid units values are svbaInches (default) or svbaCentimeters."
    entryCol.Add "# Valid orientation values are svbaPortrait (default) or svbaLandscape."
    entryCol.Add " "
    entryCol.Add "units=" & readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "units", "svbaInches", useDefaultValue)
    entryCol.Add "page_height=" & readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "page_height", 11, useDefaultValue)
    entryCol.Add "page_width=" & readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "page_width", 8.5, useDefaultValue)
    entryCol.Add "margin_bottom=" & readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "margin_bottom", Format(1 / 2.54, ".000000"), useDefaultValue)
    entryCol.Add "margin_top=" & readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "margin_top", Format(1 / 2.54, ".000000"), useDefaultValue)
    entryCol.Add "margin_right=" & readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "margin_right", Format(1 / 2.54, ".000000"), useDefaultValue)
    entryCol.Add "margin_left=" & readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "margin_left", Format(1 / 2.54, ".000000"), useDefaultValue)
    entryCol.Add "background=" & readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "background", False, useDefaultValue)
    entryCol.Add "orientation=" & readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "orientation", "svbaPortrait", useDefaultValue)
    entryCol.Add "print_scale=" & readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "print_scale", 1, useDefaultValue)
    entryCol.Add "shrink_to_fit=" & readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "shrink_to_fit", True, useDefaultValue)
    
    Set ts = fso.CreateTextFile(iniFilePath, True, True)
    
    For i = 1 To entryCol.Count
        ts.WriteLine entryCol.Item(i)
    Next i
    
    ts.Close
End Sub

'@Description("Downloads a resource (such as an image) contained in an element attribute to a file or folder path - if a folder is specified, then the file will inherit the name of the source")
Public Sub DownloadResource(element As WebElement, ByVal srcAttribute As String, Optional ByVal fileOrFolderPath As String = ".\")
Attribute DownloadResource.VB_Description = "Downloads a resource (such as an image) contained in an element attribute to a file or folder path - if a folder is specified, then the file will inherit the name of the source"
    Dim url As String
    Dim baseUrl As String
    Dim srcMax As String
    Dim sizeMax As String
    Dim srcArray As Variant
    Dim srcSizePair As Variant
    Dim src As String
    Dim srcSize As String
    Dim ret As Long
    Dim fso As New FileSystemObject
    Dim filePath As String
    Dim fileName As String
    Dim ipos As Long
    Dim i As Long
    Dim urlLevel As Long
    Dim curUrl As String
    Dim numUpLevels As Long
    Dim numUrlSections As Long
    
    url = element.GetAttribute(srcAttribute)
    If url = vbNullString Then Err.Raise 1, "WebDriver", "Could not find Attribute " & Chr$(34) & srcAttribute & Chr$(34)
    'remove new line characters
    url = Replace$(Replace$(url, vbLf, ""), vbCr, "")
    
    If LCase$(srcAttribute) = "srcset" Then
        'if image then get the highest pixel density src possible
        'split the sources in srcset into an array
        srcArray = Split(url, ", ")
        
        'initialize the maximum size and source
        sizeMax = vbNullString
        srcMax = vbNullString
        
        For i = 0 To UBound(srcArray)
            'loop through and extract src/size pairs
            srcSizePair = Split(Trim$(srcArray(i)), " ")
            src = Trim$(srcSizePair(0))
            
            'process size if available
            If UBound(srcSizePair) > 0 Then
                srcSize = Trim$(srcSizePair(1))
                'remove pixel density and/or width indicator for size compare
                srcSize = Replace$(Replace$(srcSize, "x", vbNullString, , , vbTextCompare), "w", vbNullString, , , vbTextCompare)
            Else
                srcSize = vbNullString
            End If
            
            'if this size is greater than max found size, then store as max
            If srcSize >= sizeMax Then
                sizeMax = srcSize
                srcMax = src
            End If
        Next i
        
        'use the max found source
        url = srcMax
    End If
    
    'handle relative path if applicable
    If Left$(url, 4) <> "http" Then 'it is a relative path
        'construct a full url from base url and the relative path
        Dim urlSections As Variant
        
        'get the current url path
        curUrl = GetCurrentURL
        
        'strip off query params if they exist
        curUrl = Split(curUrl, "?")(0)
        
        'replace double slash in "http[s]://" - will put back later
        urlSections = Split(Replace$(curUrl, "//", "^^"), "/")
        numUrlSections = UBound(urlSections)
        
        'determine relativity level (root, current, up(s))
        Select Case True
        Case url Like "/*"
            'root level
            urlLevel = 0
        Case url Like "[!./]*", url Like "./*"
            'current level
            urlLevel = numUrlSections - 1
            'strip out the rel indicator
            url = Replace$(url, "./", "")
        Case url Like "../*"
            'up-levels
            'count how many up-levels
            numUpLevels = UBound(Split(url, "../"))
            urlLevel = numUrlSections - 1 - numUpLevels
            'strip out the rel indicator(s)
            url = Replace$(url, "../", "")
        End Select
        
        'build the base url
        baseUrl = urlSections(0) 'root
        For i = 1 To urlLevel
            baseUrl = baseUrl & "/" & urlSections(i)
        Next i
        
        'put back the double slash in "http[s]://"
        baseUrl = Replace$(baseUrl, "^^", "//")
        
        'make sure there is a slash between base url and relative url
        If Right$(baseUrl, 1) <> "/" And Left$(url, 1) <> "/" Then baseUrl = baseUrl & "/"
        
        'concat to calculate absolute url
        url = baseUrl & url
    End If
    
    'now calculate the destination path
    'first convert user input path to absolute path
    fileOrFolderPath = WebShared.getFullLocalPath(fileOrFolderPath, defaultIOFolder_)

    'check if user gave a folder path or a file path
    If fso.FolderExists(fileOrFolderPath) Then 'its a folder path
        'file path inherits the name of the target url
        fileName = fso.GetFileName(url)
        'strip off query params if they exist
        'and cat with full path of user input
        filePath = fileOrFolderPath & "\" & Split(fileName, "?")(0)
    Else 'its a file path
        filePath = fileOrFolderPath
        If Not fso.FolderExists(fso.GetParentFolderName(filePath)) Then
            Err.Raise 1, "WebDriver", "Specified folder path for saving file does not exist:" & vbNewLine & vbNewLine & fso.GetParentFolderName(filePath)
        End If
    End If
    
    'download the url to file path
    ret = UrlDownloadToFile(0&, StrPtr(url), StrPtr(filePath), 0&, 0&)
    
    If ret <> 0 Then
         Err.Raise 1, "WebDriver", "Can not download from " & url & " to " & filePath
    End If
End Sub

Friend Function execute(driverCommand As Variant, Optional parameters As Dictionary = Nothing, Optional ByVal raiseError As Boolean = True) As Dictionary
    Dim cmdMethod As String
    Dim cmdPath As String
    Dim cmdArgs As New Dictionary
    Dim parmKey As Variant
    Dim response As Dictionary
    Dim client As New MSXML2.ServerXMLHTTP60
    Dim jc As New WebJsonConverter
    
    cmdMethod = driverCommand(0)
    cmdPath = driverCommand(1)
    
    If parameters Is Nothing Then
        Set parameters = New Dictionary
    End If
    
    'Set session id
    If Not parameters.Exists("sessionId") Then
        parameters.Add "sessionId", sessionId_
    End If
    
    For Each parmKey In parameters.keys
        If InStr(cmdPath, "$" & parmKey) > 0 Then 'path parameter
            cmdPath = Replace$(cmdPath, "$" & parmKey, parameters(parmKey))
        Else 'non-path parameter
            cmdArgs.Add parmKey, parameters(parmKey)
        End If
    Next parmKey
    
    'there are commands that don't require sessionId in path: tCMD.CMD_STATUS, tCMD.CMD_NEW_SESSION, tCMD.CMD_GET_ALL_SESSIONS, tCMD.CMD_SHUTDOWN
    If cmdArgs.Exists("sessionId") Then cmdArgs.Remove "sessionId"
    
    'Send request to selenium server

    client.Open cmdMethod, driverUrl_ & cmdPath
    
    'Set a short 1s receiveTimeout when setting browser capabilities (assuming communication is only with local browser)
    'otherwire leave it to 0 (infinite) to accommodate longer than 30 sec (default) pageloads
    Dim isCapabilities As Boolean
    If cmdArgs.Count > 0 Then If cmdArgs.keys(0) = "capabilities" Then isCapabilities = True
    client.setTimeouts 0, 60000, 30000, IIf(isCapabilities, 3000, 0)
    
    If cmdMethod = "POST" Or cmdMethod = "PUT" Then
        client.setRequestHeader "Content-Type", "application/json; charset=utf-8"
        client.setRequestHeader "Cache-Control", "no-cache"
        On Error GoTo Err
        client.send jc.ConvertToJson(cmdArgs)
        GoTo errEnd
Err:    If Err.Number = -2147012894 Then
            Err.Raise Err.Number, Err.Source, "Timeout in setting browser capabilities." & vbNewLine & vbNewLine _
                        & "If you're trying to attach WebDriver to an existing instance of your Chrome/Edge browser," _
                        & "restart your browser process(es) by adding command line argument " _
                        & """--remote-debugging-port=9222"" (or different port number)", Err.HelpFile, Err.HelpContext
        ElseIf Err.Number <> 0 Then
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        End If
errEnd: On Error GoTo 0
    Else
        client.send
    End If

    Do While client.readyState < 4
        DoEvents
    Loop

    Set response = jc.ParseJson(client.responseText)
    
    If raiseError And isResponseError(response) Then 'raise error if occurs
        Debug.Print getResponseErrorMessage(response)
        Err.Raise client.status, "WebDriver", getResponseErrorMessage(response)
    End If

    Set execute = response 'Pass the response dictionary object - let caller parse info based on context, including error
End Function

Private Function toWebElement(ByVal elementId As String) As WebElement
    Dim element As New WebElement
    Set element.driver = Me
    element.elementId = elementId
    Set toWebElement = element
End Function

Private Function toShadowRoot(ByVal shadowRootId As String) As WebShadowRoot
    Dim sr As New WebShadowRoot
    Set sr.driver = Me
    sr.shadowRootId = shadowRootId
    Set toShadowRoot = sr
End Function

Private Function getResponseErrorMessage(resp As Dictionary) As String
    getResponseErrorMessage = vbNullString
    If TypeName(resp("value")) = "Dictionary" Then
        If resp("value").Exists("error") Then
            getResponseErrorMessage = resp("value")("message")
        End If
    End If
End Function

Private Function isResponseError(resp As Dictionary) As Boolean
    isResponseError = False
    If TypeName(resp("value")) = "Dictionary" Then
        If resp("value").Exists("error") Then
            isResponseError = True
        End If
    End If
End Function

Private Function params(ParamArray keysAndValues() As Variant) As Dictionary
    Dim dict As New Dictionary
    Dim i As Integer
    For i = 0 To UBound(keysAndValues) - 1 Step 2
        dict.Add keysAndValues(i), keysAndValues(i + 1)
    Next i
    Set params = dict
End Function

Private Function rgbLongToString(ByVal rgbColor As Long) As String
    If rgbColor < RGB(0, 0, 0) Or rgbColor > RGB(255, 255, 255) Then Err.Raise 1, "WebDriver", "Error: argument rgbColor out of range (0 to 16777215)"
    Dim r As Long
    Dim g As Long
    Dim b As Long
    r = rgbColor And &HFF
    g = (rgbColor And &H100FF00) / &H100
    b = (rgbColor And &HFF0000) / &H10000
    rgbLongToString = "rgb(" & r & "," & g & "," & b & ")"
End Function

Private Function objToWebElements(elementOrElements As Object) As WebElements
    Select Case TypeName(elementOrElements)
        Case "WebElement": Set objToWebElements = New WebElements: objToWebElements.Add elementOrElements
        Case "WebElements": Set objToWebElements = elementOrElements
        Case Else: Err.Raise 1, "WebDriver", "Error: the elementOrElements argument is not of valid type (WebElement or WebElements)"
    End Select
End Function

Private Function decodeBase64(ByVal strData As String) As Byte()
    Dim domDoc As New MSXML2.DOMDocument60
    Dim domNode As MSXML2.IXMLDOMElement
    'create node with type of base 64 and decode
    Set domNode = domDoc.createElement("b64")
    domNode.DataType = "bin.base64"
    domNode.text = strData
    decodeBase64 = domNode.nodeTypedValue
End Function

Private Sub saveByteArrayToFile(ByteArray() As Byte, ByVal filePath As String)
  'Create Stream object
  Dim binaryStream As New ADODB.Stream

  'Specify stream type - we want To save binary data.
  binaryStream.Type = adTypeBinary
  
  'Open the stream And write binary data To the object
  binaryStream.Open
  binaryStream.Write ByteArray
  
  'Save binary data To disk
  binaryStream.SaveToFile filePath, adSaveCreateOverWrite
  binaryStream.Close
End Sub

Private Sub killDriver(ByVal browser As svbaBrowser, Optional ByVal port As Integer = 0)
    'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-process
    Dim objService As WbemScripting.SWbemServices
    Dim objDriverList As WbemScripting.SWbemObjectSet
    Dim objParentList As WbemScripting.SWbemObjectSet
    Dim objDriver As WbemScripting.SWbemObject
    Dim objParent As WbemScripting.SWbemObject
    Dim sImageName As String
    Dim pid As String
    
    Select Case browser
    Case svbaBrowser.Chrome
        sImageName = "chromedriver%"
    Case svbaBrowser.Edge
        sImageName = "msedgedriver%"
    Case svbaBrowser.Firefox
        sImageName = "geckodriver%"
    Case svbaBrowser.IE
        sImageName = "IEDriverServer%"
    End Select
    
    Set objService = GetObject("winmgmts:\\.\root\CIMV2")
    
    Set objDriverList = objService.ExecQuery("SELECT * FROM Win32_Process" & " WHERE Name like '" & sImageName & "'")
    
    For Each objDriver In objDriverList
        If port > 0 Then
            If InStr(objDriver.commandLine, CStr(port)) = 0 Then
                GoTo 10
            End If
        End If
        Set objParentList = objService.ExecQuery("SELECT * FROM Win32_Process" & " WHERE ProcessID = '" & objDriver.ParentProcessId & "'")
        If objParentList.Count > 0 Then
            Set objParent = objParentList.ItemIndex(0)
            Select Case LCase$(objParent.Name)
            Case "cmd.exe"
                'driver was called using cmd.exe so we kill the parent process and all subprocesses (/t parameter)
                pid = objParent.processID
            Case Else '"excel.exe"
                'driver was called directly using shell and so will kill it and all subprocesses (/t parameter)
                pid = objDriver.processID
            End Select
        Else 'its a stranded process (parent may have already been quit)
            pid = objDriver.processID
        End If
        taskKillByPid pid
10:
    Next objDriver
End Sub

Private Function taskKillByPid(ByVal pid As String) As Variant
    Dim wsh As New IWshRuntimeLibrary.WshShell
    taskKillByPid = wsh.Run("taskkill /f /t /pid " & pid, 0, True)
End Function

Private Function htmlTableToArray(table As htmlTable, Optional ByVal skipHeader As Boolean = False, Optional ByVal skipFooter As Boolean = False, Optional ByVal createSpanData As Boolean = True) As Variant()
    Dim row As HTMLTableRow
    Dim cell As HTMLTableCell
    Dim cellChild As HTMLHtmlElement
    Dim maxCols As Integer
    Dim i As Integer
    Dim foundTable As Boolean
    Dim rowSpan As Integer
    Dim colSpan As Integer
    Dim v() As Variant
    
    If skipHeader Then table.deleteTHead
    If skipFooter Then table.deleteTFoot
    
    If createSpanData Then
        'first handle row spans
        For Each row In table.Rows
            For Each cell In row.Cells
                rowSpan = cell.rowSpan
                If rowSpan > 1 Then
                    For i = row.RowIndex + 1 To row.RowIndex + rowSpan - 1
                        table.Rows(i).insertCell(cell.cellIndex).innerText = cell.innerText
                    Next i
                End If
            Next cell
        Next row
        
        'now process column spans
        For Each row In table.Rows
            For Each cell In row.Cells
                colSpan = cell.colSpan
                If colSpan > 1 Then
                    For i = 1 To colSpan - 1
                        row.insertCell(row.Cells.Length).innerText = cell.innerText
                    Next i
                End If
            Next cell
        Next row
    End If
    
    'calculate the number of array columns needed for output
    For Each row In table.Rows
        If row.Cells.Length > maxCols Then maxCols = row.Cells.Length
    Next row
    
    'dimension output array
    ReDim v(baseArrayIdx To table.Rows.Length - 1 + baseArrayIdx, baseArrayIdx To maxCols - 1 + baseArrayIdx)
    
    'extract cell data (td and th elements) from each row
    For Each row In table.Rows
        For Each cell In row.Cells
            'there can be all sorts of crap in a table cell including <font>, <p>, <img>, and MULTIPLE <table>'s!
            'search to see if there is at least one table
            'if there are more than one in the cell then only process first!
            If cell.Children.Length > 0 Then
                foundTable = False
                For Each cellChild In cell.Children
                    If cellChild.tagName = "TABLE" Then
                        'recursive call to TableToArray
                        v(row.RowIndex + baseArrayIdx, cell.cellIndex + baseArrayIdx) = htmlTableToArray(cellChild, skipHeader, skipFooter, createSpanData)
                        foundTable = True
                        Exit For
                    End If
                Next cellChild
                If Not foundTable Then v(row.RowIndex + baseArrayIdx, cell.cellIndex + baseArrayIdx) = cell.innerText
            Else
                v(row.RowIndex + baseArrayIdx, cell.cellIndex + baseArrayIdx) = cell.innerText
            End If
        Next cell
    Next row
    
    htmlTableToArray = v
End Function

Private Function getXMLStringFromPage(Optional ByVal prettyPrint As Boolean = False) As String
    Dim str As String
    Select Case browser_
    Case svbaBrowser.Chrome, svbaBrowser.Edge, svbaBrowser.IE
        If Me.IsPresent(By.XPath, "//*[@id='webkit-xml-viewer-source-xml']") Then
            str = Me.FindElementByXPath("//*[@id='webkit-xml-viewer-source-xml']").GetInnerHTML
        End If
    Case svbaBrowser.Firefox
        str = Me.GetPageSource
    End Select
    
    If Left$(str, 1) <> "<" Or Left$(str, 5) = "<html" Then Err.Raise 1, "WebDriver", "Error: The page source is not XML format!"
    
    If prettyPrint Then str = formatXML(str)
    
    getXMLStringFromPage = str
End Function

Private Function getJSONStringFromPage(Optional ByVal prettyPrint As Boolean = False) As String
    Dim str As String
    Dim jc As New WebJsonConverter
    Dim j As Long
    Dim i As Long
    
    If browser_ = svbaBrowser.Firefox Then
        'check to see if preference "devtools.jsonview.enabled" = True (default)
        If Not Me.IsPresent(By.tagName, "pre") Then
            'Firefox using jsonview, so now click rawdata tab
            If Me.IsPresent(By.ID, "rawdata-tab") Then
                Me.FindElementByID("rawdata-tab").Click
            Else
                Err.Raise 1, "WebDriver", "Error: The page source is not JSON format!"
            End If
        End If
        
        str = Me.FindElementByTagName("pre").GetText
        
        'firefox returns a wrongly interpreted json if webpage source came from NavigateToFile *.json
        'GCUser99 has reported this bug - see https://bugzilla.mozilla.org/show_bug.cgi?id=1797871
        str = fixFirefoxBug1797871(str)
    Else
        If Me.IsPresent(By.tagName, "pre") Then
            str = Me.FindElementByTagName("pre").GetText
        Else
            Err.Raise 1, "WebDriver", "Error: The page source is not JSON format!"
        End If
    End If
    
    If prettyPrint Then str = jc.ConvertToJson(jc.ParseJson(str), 4)
    
    getJSONStringFromPage = str
End Function

Private Function sanitizeHTML(ByVal str As String) As String
    Dim htmlDoc As New HTMLDocument
    Dim elems As IHTMLDOMChildrenCollection
    Dim elem As HTMLGenericElement
    Dim parentElem As HTMLDivElement
    Dim i As Long, j As Long
    Dim ipos1 As Long, ipos2 As Long
    Dim foundOne As Boolean
    Dim subStr As String
    Dim eventsArray() As Variant
    Dim eventStr As String
    Dim emptyScript As HTMLScriptElement
    Dim scriptNode As IHTMLDOMNode
    Dim styleNode As IHTMLDOMNode
    Dim newStyleNode As IHTMLDOMNode
    Dim tempDoc As HTMLDocument

    'This function renders html source "offline" by disabling dynamic content, while leaving the DOM tree intact.
    
    'This following are disabled:
    
    '- script elements
    '- url redirects
    '- src attributes
    '- srcset attributes
    '- href attributes
    '- rel attributes
    '- url() function in style elements and attributes
    '- action attributes in form elements
    '- as attribute in link element when as='script'
    '- on* event attributes
    '- data attribute in object elements
    
    'To perform above in a reliable and efficient way, the source is loaded into a simulated document fragment using a div element as
    'the root. This way the source can be parsed using DOM while disconnected from a html document root.
    'html, head, and body tags in the input string are renamed prior to loading into the div element, and then renamed back
    'before returning the results.
    
    'rename html, head, and body tags
    'use vbTextCompare for case insensitivity
    str = Replace$(str, "<html", "<html_", Compare:=vbTextCompare, Count:=1)
    str = Replace$(str, "</html", "</html_", Compare:=vbTextCompare, Count:=1)

    str = Replace$(str, "<head", "<head_", Compare:=vbTextCompare, Count:=1)
    str = Replace$(str, "</head", "</head_", Compare:=vbTextCompare, Count:=1)
    
    str = Replace$(str, "<body", "<body_", Compare:=vbTextCompare, Count:=1)
    str = Replace$(str, "</body", "</body_", Compare:=vbTextCompare, Count:=1)
    
    'create the simulated document fragment using a div element
    Set parentElem = htmlDoc.createElement("div")
    'assign the input string to inner html of the parent element
    parentElem.innerHTML = str

    'disable scripts
    Set elems = parentElem.querySelectorAll("script")
    For i = 0 To elems.Length - 1
        If htmlDoc.documentMode >= 11 Then
            'these work in Office 365, but appear not to work on some 64-bit systems
            Set elem = elems.Item(i)
            elem.innerText = vbNullString
        Else
            'so we use the node replace method to empty script
            Set scriptNode = elems.Item(i)
            Set emptyScript = htmlDoc.createElement("script")
            scriptNode.ParentNode.replaceChild emptyScript, scriptNode
        End If
    Next i

    'disable redirects
    Set elems = parentElem.querySelectorAll("meta[http-equiv='refresh']")
    For i = 0 To elems.Length - 1
        Set elem = elems.Item(i)
        elem.SetAttribute "content", vbNullString
    Next i
    
    'disable src attributes
    Set elems = parentElem.querySelectorAll("*[src]")
    For i = 0 To elems.Length - 1
        Set elem = elems.Item(i)
        elem.SetAttribute "src", vbNullString
    Next i
    
    'disable srcset attributes
    Set elems = parentElem.querySelectorAll("*[srcset]")
    For i = 0 To elems.Length - 1
        Set elem = elems.Item(i)
        elem.SetAttribute "srcset", vbNullString
    Next i
    
    'disable href attributes
    Set elems = parentElem.querySelectorAll("*[href]")
    For i = 0 To elems.Length - 1
        Set elem = elems.Item(i)
        If Not (elem.GetAttribute("href") Like "about:blank*") Then
            elem.SetAttribute "href", vbNullString
        End If
    Next i
    
    'disable rel attributes
    Set elems = parentElem.querySelectorAll("*[rel]")
    For i = 0 To elems.Length - 1
        Set elem = elems.Item(i)
        elem.SetAttribute "rel", vbNullString
    Next i
    
    'disable url() function in style elements
    If htmlDoc.documentMode >= 11 Then
        Set elems = parentElem.querySelectorAll("style")
        
        For i = 0 To elems.Length - 1
            Set elem = elems.Item(i)
            subStr = LCase$(elem.innerHTML) 'this is needed to use much faster (>10x!) vbBinaryCompare
            
            foundOne = False
            ipos1 = InStr(1, subStr, "url(")
            Do While ipos1
                ipos2 = InStr(ipos1 + 4, subStr, ")")
                'subStr = Left$(subStr, ipos1 + 3) & Mid$(subStr, ipos2)
                'ipos1 = InStr(ipos1 + 5, subStr, "url(")
                'this takes a fraction of time as the above but replaces with spaces instead of nulls
                Mid$(subStr, ipos1 + 4, ipos2 - ipos1 - 4) = Space$(ipos2 - ipos1 - 4)
                ipos1 = InStr(ipos2 + 1, subStr, "url(")
                foundOne = True
                DoEvents
            Loop
            
            If foundOne Then
                elem.innerHTML = subStr
            End If
        Next i
    End If
    
    'disable url() function in style attributes
    Set elems = parentElem.querySelectorAll("[style*='url(']")
    For i = 0 To elems.Length - 1
        Set elem = elems.Item(i)
        subStr = LCase$(elem.GetAttribute("style"))
        foundOne = False
        ipos1 = InStr(1, subStr, "url(")
        Do While ipos1
            ipos2 = InStr(ipos1 + 4, subStr, ")")
            'subStr = Left$(subStr, ipos1 + 3) & Mid$(subStr, ipos2)
            'ipos1 = InStr(ipos1 + 5, subStr, "url(")
            'this takes a fraction of time as the above but replaces with spaces instead of nulls
            Mid$(subStr, ipos1 + 4, ipos2 - ipos1 - 4) = Space$(ipos2 - ipos1 - 4)
            ipos1 = InStr(ipos2 + 1, subStr, "url(")
            foundOne = True
            DoEvents
        Loop
        If foundOne Then elem.SetAttribute "style", subStr
    Next i

    'disable action attributes in form elements
    Set elems = parentElem.querySelectorAll("form[action]")
    For i = 0 To elems.Length - 1
        Set elem = elems.Item(i)
        elem.SetAttribute "action", vbNullString
    Next i
    
    'disable as attributes equal to "script" in link elements
    Set elems = parentElem.querySelectorAll("link[as='script']")
    For i = 0 To elems.Length - 1
        Set elem = elems.Item(i)
        elem.SetAttribute "as", vbNullString
    Next i
    
    'disable data attributes in object elements
    Set elems = parentElem.querySelectorAll("object[data]")
    For i = 0 To elems.Length - 1
        Set elem = elems.Item(i)
        elem.SetAttribute "data", vbNullString
    Next i
   
    'disable on* event attributes
    eventsArray = Array("onclick", "onload", "onoffline", "ononline", "onunload", "onpageshow", "onpagehide", "onerror", "onresize", "onhashchange", "onbeforeunload", "onpopstate", "onmessage")
    For j = 0 To UBound(eventsArray)
        eventStr = eventsArray(j)
        Set elems = parentElem.querySelectorAll("*[" & eventStr & "]")
        For i = 0 To elems.Length - 1
            Set elem = elems.Item(i)
            'can't use SetAttribute for on* event attributes so let's remove it instead
            elem.removeAttribute eventStr
            'add a replacement attribute with dummy suffix - will remove suffix with replace after dom processing
            elem.SetAttribute eventStr & "___xxx123", vbNullString
        Next i
    Next j

    'load processed source back into str
    str = parentElem.innerHTML
    
    'rename dummy tags back to original tag names
    str = Replace$(str, "<html_", "<html", , 1)
    str = Replace$(str, "</html_", "</html", , 1)

    str = Replace$(str, "<head_", "<head", , 1)
    str = Replace$(str, "</head_", "</head", , 1)
    
    str = Replace$(str, "<body_", "<body", , 1)
    str = Replace$(str, "</body_", "</body", , 1)
    
    str = Replace$(str, "___xxx123", vbNullString)
    
    sanitizeHTML = str
End Function

Private Function formatXML(ByVal val As String) As String
    Dim xslDoc As New MSXML2.DOMDocument60
    Dim xmlDocFormatted As New MSXML2.DOMDocument60
    Dim xmlDoc As New MSXML2.DOMDocument60
    
    xmlDoc.LoadXML val
    
    xslDoc.LoadXML "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & "?>" _
            & "<xsl:stylesheet version=" & Chr$(34) & "1.0" & Chr$(34) _
            & "                xmlns:xsl=" & Chr$(34) & "http://www.w3.org/1999/XSL/Transform" & Chr$(34) & ">" _
            & "<xsl:strip-space elements=" & Chr$(34) & "*" & Chr$(34) & " />" _
            & "<xsl:output method=" & Chr$(34) & "xml" & Chr$(34) & " indent=" & Chr$(34) & "yes" & Chr$(34) & vbNullString _
            & "            encoding=" & Chr$(34) & "UTF-8" & Chr$(34) & "/>" _
            & " <xsl:template match=" & Chr$(34) & "node() | @*" & Chr$(34) & ">" _
            & "  <xsl:copy>" _
            & "   <xsl:apply-templates select=" & Chr$(34) & "node() | @*" & Chr$(34) & " />" _
            & "  </xsl:copy>" _
            & " </xsl:template>" _
            & "</xsl:stylesheet>"
    
    xslDoc.async = False
    xmlDoc.transformNodeToObject xslDoc, xmlDocFormatted
    formatXML = xmlDocFormatted.XML
End Function

Private Function fixFirefoxBug1797871(ByVal str As String) As String
    Dim i As Long
    Dim j As Long
    Dim bytes() As Byte
    'see https://bugzilla.mozilla.org/show_bug.cgi?id=1797871
    'firefox json viewer only accepts utf-8 encoded text files
    'the following works for preference "devtools.jsonview.enabled", [True or False]
    bytes = str
    'check the first 2 bytes for the "mis-interpretion" pattern and if found then unscramble
    If bytes(0) = 253 And bytes(1) = 255 Then
        j = -2
        For i = 0 To UBound(bytes) - 1 Step 2
            'strip out nulls (double zeros) and black-diamond question marks
            If Not (bytes(i) = 0 And bytes(i + 1) = 0) And Not (bytes(i) = 253 And bytes(i + 1) = 255) Then
                j = j + 2
                bytes(j) = bytes(i)
                bytes(j + 1) = bytes(i + 1)
            End If
        Next i
        ReDim Preserve bytes(0 To j + 1)
        str = bytes
        
        fixFirefoxBug1797871 = Replace$(str, vbNewLine & vbNewLine, vbNewLine)
    Else
        fixFirefoxBug1797871 = str
    End If
End Function

Private Function deObjectifyInputArgs(argumentList() As Variant) As Variant()
    'replace WebElement, WebElements, WebShadowRoot object arguments with format needed by WebDriver server
    Dim i As Long, j As Long
    Dim elemCol As Collection
    Dim elemDict As Dictionary
    Dim shadowRootDict As Dictionary
    Dim shadowRoot As WebShadowRoot
    Dim elements As WebElements
    Dim element As WebElement
    Dim args() As Variant
        
    ReDim args(0 To UBound(argumentList))
    For i = 0 To UBound(argumentList)
        Select Case TypeName(argumentList(i))
        Case "WebElement"
            'convert WebElement object to Dictionary that can be handled by WebDriver
            Set element = argumentList(i)
            Set elemDict = New Dictionary
            elemDict.Add ELEMENT_KEY, element.elementId
            Set args(i) = elemDict
        Case "WebShadowRoot"
            'convert WebShadowRoot object to Dictionary that can be handled by WebDriver
            Set shadowRoot = argumentList(i)
            Set shadowRootDict = New Dictionary
            shadowRootDict.Add SHADOWROOT_KEY, shadowRoot.shadowRootId
            Set args(i) = shadowRootDict
        Case "WebElements"
            'convert WebElements object to Dictionary that can be handled by WebDriver
            Set elements = argumentList(i)
            Set elemCol = New Collection
            For j = 1 To elements.Count
                Set elemDict = New Dictionary
                elemDict.Add ELEMENT_KEY, elements(j).elementId
                elemCol.Add elemDict
            Next j
            Set args(i) = elemCol
        Case Else
            args(i) = argumentList(i)
        End Select
    Next i
    
    deObjectifyInputArgs = args
End Function

Private Function objectifyScriptResponse(responseValue As Object) As Object
    'this function replaces the ExecuteScript response representation of an element, a collection of elements,
    'and a shadow root with a WebElement object, a WebElements object, and a WebShadowRoot object, respectively
    'its recursive designed allows to search and replace a nested response structure
    Dim i As Long
    Dim vkey As Variant
    Dim obj As Object
    Dim numItems As Long
    Dim isElementCollection As Boolean
    Dim elems As WebElements
    
    Select Case TypeName(responseValue)
    Case "Collection"
        isElementCollection = True
        numItems = responseValue.Count
        
        'loop through all items of the collection, looking for objects
        For i = 1 To numItems
            If IsObject(responseValue.Item(i)) Then
                Set obj = objectifyScriptResponse(responseValue.Item(i))
                
                If TypeName(obj) <> "WebElement" Then isElementCollection = False
                
                responseValue.Remove i
                If i = numItems Then
                    responseValue.Add obj
                Else
                    responseValue.Add obj, , i
                End If
            Else
                isElementCollection = False
            End If
        Next i
        
        If isElementCollection Then
            'replace the collection with a WebElements object
            Set elems = New WebElements
            For i = 1 To responseValue.Count
                elems.Add responseValue.Item(i)
            Next i
            Set responseValue = elems
        End If
    Case "Dictionary"
        'check if its an element or shadow root representation and if so, replace with WebElement or WebShadowRoot object
        If responseValue.Exists(ELEMENT_KEY) Then 'replace with a WebElement object
            Set responseValue = toWebElement(responseValue(ELEMENT_KEY))
        ElseIf responseValue.Exists(SHADOWROOT_KEY) Then 'replace with a WebShadowRoot object
            Set responseValue = toShadowRoot(responseValue(SHADOWROOT_KEY))
        Else 'check each key value for collection or dictionary
            For Each vkey In responseValue.keys
                If IsObject(responseValue.Item(vkey)) Then
                    Set responseValue.Item(vkey) = objectifyScriptResponse(responseValue.Item(vkey))
                End If
            Next vkey
        End If
    End Select
        
    Set objectifyScriptResponse = responseValue
End Function

Private Function FileClosed(ByVal filePath As String) As Boolean
    Dim fso As New FileSystemObject
    On Error Resume Next
    fso.MoveFile filePath, filePath
    If Err.Number = 0 Then
        FileClosed = True
    Else
        FileClosed = False
    End If
End Function
