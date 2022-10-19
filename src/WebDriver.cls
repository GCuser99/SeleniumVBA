VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================
' SeleniumVBA v2.2
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
' (c) 2021 uezo
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

'References:
'Microsoft Script Control 1.0
'Microsoft XML, v6.0
'Microsoft Scripting Runtime

Option Explicit

'set to True if auto-check-and-update desired
Private Const checkDriverBrowserVersionAlignment = True

'set version compatibility threshold for triggering a driver update
Private Const minDriverBrowserCompatibilityLevel = svbaBuildMajor

'desired lower bound of returned arrays
Private Const baseArrayIdx = 1

'change baseDomIdx to 0 if zero-based Dom (de)selection is desired
'with baseDomIdx=1: SelectByIndex 1 gets first element
'with baseDomIdx=0: SelectByIndex 0 gets first element
Private Const baseDomIdx = 1

Private sessionId_ As String
Private driverUrl_ As String
Private browser_ As svbaBrowser
Private appWinStyle As VbAppWinStyle
Private defaultIOFolder_ As String
Private defaultBinaryFolder_ As String

Private SetHighlight_ElemsOrigStyle As Scripting.Dictionary 'store original style of Highlighted elements to be able to UnHighlight them at a future time

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
Private SetHightlightFoundElems_Settings As Collection
Private lastElems As WebElements

Private driverFilename As String
Private processID As String

Private Const ELEMENT_KEY = "element-6066-11e4-a52e-4f735466cecf"
Private Const SHADOW_KEY = "shadow-6066-11e4-a52e-4f735466cecf"

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
    'CMD_UPLOAD_FILE As Variant                    'not wrapped yet
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
    CMD_SET_ALERT_VALUE As Variant
    CMD_GET_ALERT_TEXT As Variant
    CMD_CLICK As Variant
    'CMD_ACTIONS As Variant                        'used in ActionChain.cls
    'CMD_CLEAR_ACTIONS As Variant                  'used in ActionChain.cls
    CMD_SET_WINDOW_RECT As Variant
    CMD_GET_WINDOW_RECT As Variant
    CMD_MAXIMIZE_WINDOW As Variant
    'CMD_GET_LOCATION As Variant                   'not wrapped yet
    'CMD_SET_LOCATION As Variant                   'not wrapped yet
    'CMD_GET_APP_CACHE_STATUS As Variant           'not wrapped yet
    'CMD_GET_NETWORK_CONNECTION As Variant         'not wrapped yet
    'CMD_SET_NETWORK_CONNECTION As Variant         'not wrapped yet
    'CMD_GET_LOG As Variant                        'see logging option in StartEdge and StartChrome methods
    'CMD_GET_AVAILABLE_LOG_TYPES As Variant        'see logging option in StartEdge and StartChrome methods
    'CMD_CURRENT_CONTEXT_HANDLE As Variant         'not wrapped yet
    'CMD_CONTEXT_HANDLES As Variant                'not wrapped yet
    'CMD_SWITCH_TO_CONTEXT As Variant              'not wrapped yet
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

Public Enum by
    ID = 0
    tagName = 1
    className = 2
    Name = 3
    cssSelector = 4
    XPath = 5
    linkText = 6
    PartialLinkText = 7
End Enum

Public Enum svbaWindowType
    svbaWindow = 0
    svbaTab = 1
End Enum

'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Property Let CommandWindowStyle(ByVal style As VbAppWinStyle)
    appWinStyle = style
End Property

Public Sub StartEdge(Optional ByVal driverPath As String = "msedgedriver.exe", Optional ByVal localHostPort As Long = 9515, Optional ByVal enableLogging As Boolean = False, Optional ByVal logFilePath As String = ".\edge.log")
    Start svbaBrowser.Edge, driverPath, localHostPort, enableLogging, logFilePath
End Sub

Public Sub StartChrome(Optional ByVal driverPath As String = "chromedriver.exe", Optional ByVal localHostPort As Long = 9515, Optional ByVal enableLogging As Boolean = False, Optional ByVal logFilePath As String = ".\chrome.log")
    Start svbaBrowser.Chrome, driverPath, localHostPort, enableLogging, logFilePath
End Sub

Public Sub StartFirefox(Optional ByVal driverPath As String = "geckodriver.exe", Optional ByVal localHostPort As Long = 4444, Optional ByVal enableLogging As Boolean = False, Optional ByVal logFilePath As String = ".\firefox.log")
    Start svbaBrowser.Firefox, driverPath, localHostPort, enableLogging, logFilePath
End Sub

Private Sub Start(ByVal browser As svbaBrowser, ByVal driverPath As String, ByVal localHostPort As Long, ByVal enableLogging As Boolean, ByVal logFilePath As String)
    Dim fso As New Scripting.FileSystemObject, serviceArgs As String, commandStr As String
    
    browser_ = browser
    
    If browser_ = svbaBrowser.Firefox Then
        'for firefox, if previous session is still open on same port, then will get "Session is already started"
        'on OpenBrowser so kill all driver instances on the assigned port
        'could use KillDriver on Chrome/Edge too, but since they support multi-sessions on same port, this will eliminate
        'that functionality if important
        KillDriver browser_, localHostPort
    End If
    
    driverPath = WebShared.GetFullLocalPath(driverPath, defaultBinaryFolder_)

    If checkDriverBrowserVersionAlignment Then
        Dim mngr As New WebDriverManager
        'it's important to run these in silent mode so that if there is a problem, we don't get in the way of automation activity
        mngr.AlignDriverWithBrowser browser_, driverPath, minDriverBrowserCompatibilityLevel, True
    End If
    
    If Not fso.FileExists(driverPath) Then
        Err.raise 1, "WebDriver", "Could not find Selenium WebDriver at :" & vbNewLine & vbNewLine & driverPath & vbNewLine & vbNewLine & "Either manually install the driver or let WebDriverManager class install"
    End If
    
    commandStr = "": serviceArgs = ""

    If enableLogging Then
        logFilePath = WebShared.GetFullLocalPath(logFilePath, defaultIOFolder_)
        Select Case browser_
        Case svbaBrowser.Edge, svbaBrowser.Chrome
            serviceArgs = " --verbose --log-path=" & Chr(34) & logFilePath & Chr(34)
        Case svbaBrowser.Firefox
            'https://firefox-source-docs.mozilla.org/testing/geckodriver/TraceLogs.html
            'apparently firefox does not support command line argument for specifying log file path
            'so using cmd.exe to redirect log from console window
            serviceArgs = " -vv"" > " & Replace(logFilePath, " ", "^ ") 'need to escape spaces in log path
            commandStr = Chr(34) & Environ("comspec") & Chr(34) & " /c " & """"
        End Select
    End If
    
    commandStr = commandStr & Chr(34) & driverPath & Chr(34) & " --port=" & localHostPort & serviceArgs
    
    processID = Shell(commandStr, appWinStyle)

    If processID = 0 Then
        Err.raise 1, "WebDriver", "Failed in starting Selenium WebDriver." & vbNewLine & "WebDriver path: " & commandStr
    End If
    
    'if we've got this far, then store some global variable values for later use
    driverUrl_ = "http://localhost:" & localHostPort
    driverFilename = fso.GetFileName(driverPath)
End Sub

Public Sub Shutdown()
    'this will shut down all WebDriver instances using this driver's assigned port
    'which can be a good thing if there are multiple "rogue" instances lurking
    
    Select Case browser_
    Case svbaBrowser.Edge, svbaBrowser.Chrome
        'this will shutdown driver and session - equivalent to CloseBrowser and kill driver process
        'see "detach" option for chrome/edge to affect this behavior
        Execute tCMD.CMD_SHUTDOWN
    Case svbaBrowser.Firefox
        'firefox does not recognize shutdown command, so need to roll our own...
        'exit as gracefully as possible - first be sure user closed the browser, and if not, then let's do it here
        If sessionId_ <> "" Then CloseBrowser
        TaskKillByPid processID
    End Select
End Sub

Public Sub OpenBrowser(Optional caps As WebCapabilities, Optional ByVal invisible As Boolean = False, Optional ByVal incognito As Boolean = False, Optional capabilitiesFilePath As String = "")
    Dim resp As Dictionary
    
    If sessionId_ <> "" Then
        Err.raise 1, "WebDriver", "Only one OpenBrowser command per WebDriver instance is allowed! Use CloseBrowser before another OpenBrowser to fix this."
    End If
    
    If caps Is Nothing Then
        Set caps = CreateCapabilities
        If capabilitiesFilePath <> "" Then caps.LoadFromFile capabilitiesFilePath
    End If
    
    If invisible Then
        caps.RunInvisible
    End If
    
    If incognito Then
        caps.RunIncognito
    End If
    
    Set resp = Execute(tCMD.CMD_NEW_SESSION, params("capabilities", caps.Dictionary))("value")
    
    sessionId_ = resp("sessionId")
    
    'here's some more info that can be extracted if needed...
    'browserName = resp("capabilities")("browserName")
    'browserVersion = resp("capabilities")("browserVersion")
    'driverVersion = resp("capabilities")("chrome")("chromedriverVersion")
    'driverVersion = resp("capabilities")("msedge")("msedgedriverVersion")
    'Debug.Print resp("capabilities")("chrome")("userDataDir")
    'Debug.Print resp("capabilities")("msedge")("userDataDir")
    'Debug.Print resp("capabilities")("goog:chromeOptions")("debuggerAddress")
    'DebuggerAddress = resp("capabilities")("ms:edgeOptions")("debuggerAddress")
    'Debug.Print resp("capabilities")(GetBrowserName())("userDataDir")
End Sub

Public Sub CloseBrowser()
    Dim Data As New Dictionary
    Execute tCMD.CMD_QUIT, Data
    sessionId_ = ""
End Sub

Public Function NavigateTo(ByVal url As String, Optional ByVal timeOutms, Optional ByVal raise As Boolean = True) As Boolean
    Dim Data As New Dictionary, resp As Dictionary, savTimeOutms As Long
    
    Data.Add "url", url
    If Not IsMissing(timeOutms) Then
        savTimeOutms = GetPageLoadTimeout()
        If savTimeOutms <> timeOutms Then SetPageLoadTimeout timeOutms
    End If
    
    Set resp = Execute(tCMD.CMD_GET, Data, raise)

    If Not IsMissing(timeOutms) Then
        If savTimeOutms <> timeOutms Then SetPageLoadTimeout savTimeOutms
    End If
    
    If IsResponseError(resp) Then
        NavigateTo = False
        Debug.Print GetResponseErrorMessage(resp)
    Else
        NavigateTo = True
    End If
End Function

Public Sub NavigateToFile(ByVal htmlFilePath As String)
    'input file must have html extension to be interpreted as a webpage
    'NavigateTo (tCMD.CMD_GET) will automatically discern whether the input is a file or url, even without the 'file:///'
    'So the only functional value of this method is to handle relative paths!
    NavigateTo "file:///" & WebShared.GetFullLocalPath(htmlFilePath, defaultIOFolder_)
End Sub

Public Function FindElement(by_ As by, ByVal val As String, Optional parentElement As WebElement) As WebElement
    Dim Data As Dictionary
    Dim vCMD As Variant
    Dim elem As New WebElement
    Set Data = ToSelector(by_, val)
    If Not parentElement Is Nothing Then
        Data.Add "id", parentElement.elementId
        vCMD = tCMD.CMD_FIND_ELEMENT_FROM_ELEMENT
    Else
        vCMD = tCMD.CMD_FIND_ELEMENT
    End If

    Set elem = ToWebElement(Execute(vCMD, Data)("value")(ELEMENT_KEY))
    
    If Not SetHightlightFoundElems_Settings Is Nothing Then SetHightlightFoundElems_exec elem
    
    Set FindElement = elem
End Function

Public Function FindElements(by_ As by, ByVal val As String, Optional parentElement As WebElement) As WebElements
    Dim Data As Dictionary
    Dim i As Integer
    Dim vCMD As Variant
    Dim elements As Collection
    Dim elems As New WebElements
    
    Set Data = ToSelector(by_, val)
    
    If Not parentElement Is Nothing Then
        Data.Add "id", parentElement.elementId
        vCMD = tCMD.CMD_FIND_ELEMENTS_FROM_ELEMENT
    Else
        vCMD = tCMD.CMD_FIND_ELEMENTS
    End If
    
    Set elements = Execute(vCMD, Data)("value")
    
    For i = 1 To elements.Count
        elems.Add ToWebElement(elements(i)(ELEMENT_KEY))
    Next

    If Not SetHightlightFoundElems_Settings Is Nothing Then SetHightlightFoundElems_exec elems
    
    Set elems.WebDriver = Me
    
    Set FindElements = elems
End Function

Private Sub SetHightlightFoundElems_exec(elementOrElements As Object)
    With SetHightlightFoundElems_Settings
        Select Case TypeName(elementOrElements)
            Case "WebElement"
                SetHighlight True, elementOrElements, .Item(1), .Item(2), .Item(3)
            Case "WebElements"
                SetHighlight True, elementOrElements, .Item(4), .Item(5), .Item(6)
        End Select
    End With
End Sub

Private Function ToSelector(by_ As by, ByVal val As String) As Dictionary
    Dim Data As New Dictionary
    
    Select Case by_
    Case by.ID
        Data.Add "using", "css selector"
        Data.Add "value", "[id=""" & val & """]"
    Case by.tagName
        Data.Add "using", "css selector"
        Data.Add "value", val
    Case by.className
        Data.Add "using", "css selector"
        Data.Add "value", "." & val
    Case by.Name
        Data.Add "using", "css selector"
        Data.Add "value", "[name=""" & val & """]"
    Case by.cssSelector
        Data.Add "using", "css selector"
        Data.Add "value", val
    Case by.XPath
        Data.Add "using", "xpath"
        Data.Add "value", val
    Case by.linkText
        Data.Add "using", "link text"
        Data.Add "value", val
    Case by.PartialLinkText
        Data.Add "using", "partial link text"
        Data.Add "value", val
    Case Else
        Data.Add "using", "css selector"
        Data.Add "value", val
    End Select
    
    Set ToSelector = Data
End Function

Public Function FindElementByID(ByVal val As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByID = FindElement(by.ID, val, parentElement)
End Function

Public Function FindElementByTagName(ByVal val As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByTagName = FindElement(by.tagName, val, parentElement)
End Function

Public Function FindElementByClassName(ByVal val As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByClassName = FindElement(by.className, val, parentElement)
End Function

Public Function FindElementByCssSelector(ByVal val As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByCssSelector = FindElement(by.cssSelector, val, parentElement)
End Function

Public Function FindElementByLinkText(ByVal val As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByLinkText = FindElement(by.linkText, val, parentElement)
End Function

Public Function FindElementByName(ByVal val As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByName = FindElement(by.Name, val, parentElement)
End Function

Public Function FindElementByPartialLinkText(ByVal val As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByPartialLinkText = FindElement(by.PartialLinkText, val, parentElement)
End Function

Public Function FindElementByXPath(ByVal val As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByXPath = FindElement(by.XPath, val, parentElement)
End Function

Public Function FindElementsByXPath(ByVal val As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByXPath = FindElements(by.XPath, val, parentElement)
End Function

Public Function FindElementsByID(ByVal val As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByID = FindElements(by.ID, val, parentElement)
End Function

Public Function FindElementsByTagName(ByVal val As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByTagName = FindElements(by.tagName, val, parentElement)
End Function

Public Function FindElementsByClassName(ByVal val As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByClassName = FindElements(by.className, val, parentElement)
End Function

Public Function FindElementsByCssSelector(ByVal val As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByCssSelector = FindElements(by.cssSelector, val, parentElement)
End Function

Public Function FindElementsByLinkText(ByVal val As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByLinkText = FindElements(by.linkText, val, parentElement)
End Function

Public Function FindElementsByName(ByVal val As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByName = FindElements(by.Name, val, parentElement)
End Function

Public Function FindElementsByPartialLinkText(ByVal val As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByPartialLinkText = FindElements(by.PartialLinkText, val, parentElement)
End Function

Public Sub SendKeys(element As WebElement, ByVal keys As String, Optional clearBeforeTyping As Boolean = True)
    Dim Data As New Dictionary
    Dim chunk() As String
    Dim chars() As String
    
    Data.Add "id", element.elementId
    If clearBeforeTyping = True Then Execute tCMD.CMD_CLEAR_ELEMENT, Data
    
    'split key sequence into chunks with special keys (pattern "\ue0*") by themselves
    ReDim chunk(0 To Len(keys) - 1)
    Dim i As Long
    i = -1
    Do Until keys = ""
        Dim ipos As String
        ipos = InStr(keys, "\ue0")
        If ipos > 0 Then
            If ipos > 1 Then 'text to left is chunk
                i = i + 1
                chunk(i) = Left(keys, ipos - 1)
            End If
            'now process special key
            i = i + 1
            chunk(i) = Mid(keys, ipos, 6)
            keys = Right(keys, Len(keys) - (ipos + 5))
        Else
            i = i + 1
            chunk(i) = keys
            keys = ""
        End If
    Loop
    
    ReDim Preserve chunk(0 To i)
    
    'initialize param keys
    Data.Add "text", ""
    Data.Add "value", ""
    
    'loop through and send text chunks to target
    For i = 0 To UBound(chunk)
        If Left(chunk(i), 4) = "\ue0" Then
            ReDim chars(0 To 0)
            chars(0) = chunk(i)
        Else
            ReDim chars(0 To Len(chunk(i)) - 1)
            Dim j As Long
            For j = 0 To Len(chunk(i)) - 1
                chars(j) = Mid(chunk(i), j + 1, 1)
            Next j
        End If
        Data("text") = chunk(i)
        Data("value") = chars
        Execute tCMD.CMD_SEND_KEYS_TO_ELEMENT, Data
    Next i
End Sub

Public Sub Click(Optional element As WebElement)
    Dim Data As New Dictionary
    If Not element Is Nothing Then
        Data.Add "id", element.elementId
        Execute tCMD.CMD_CLICK_ELEMENT, Data
    Else
        Execute tCMD.CMD_CLICK, Data
    End If
End Sub

Public Sub Submit(element As WebElement)
    Click element
End Sub

Public Function GetText(element As WebElement) As String
    Dim Data As New Dictionary
    Data.Add "id", element.elementId
    GetText = Execute(tCMD.CMD_GET_ELEMENT_TEXT, Data)("value")
End Function

Public Sub GoBack()
    Dim Data As New Dictionary
    Execute tCMD.CMD_GO_BACK, Data
End Sub

Public Sub GoForward()
    Dim Data As New Dictionary
    Execute tCMD.CMD_GO_FORWARD, Data
End Sub

Public Sub Refresh()
    Dim Data As New Dictionary
    Execute tCMD.CMD_REFRESH, Data
End Sub

Public Function GetCurrentURL() As String
    Dim Data As New Dictionary
    Dim encodeURL As String
    encodeURL = Execute(tCMD.CMD_GET_CURRENT_URL, Data)("value")
    GetCurrentURL = ExecuteScript("return decodeURI('" & encodeURL & "')")
End Function

Public Function GetActiveElement() As WebElement
    Dim Data As New Dictionary
    Dim elem As WebElement
    
    Set elem = ToWebElement(Execute(tCMD.CMD_GET_ACTIVE_ELEMENT, Data)("value")(ELEMENT_KEY))
    
    If Not SetHightlightFoundElems_Settings Is Nothing Then SetHightlightFoundElems_exec elem
    
    Set GetActiveElement = elem
End Function

Public Function GetProperty(ByVal Name As String, element As WebElement) As String
    'This returns the current state of a DOM node object, such as the value of an input box.
    'If the property does not exist, then this returns vbNullString
    Dim Data As New Dictionary, resp As Dictionary
    Data.Add "name", Name
    Data.Add "id", element.elementId
    
    Set resp = Execute(tCMD.CMD_GET_ELEMENT_PROPERTY, Data)
    
    If IsNull(resp("value")) Then
        GetProperty = ""
    Else
        GetProperty = resp("value")
    End If
End Function

Public Sub SetHighlight(OnOffSwitch As Boolean, elementOrElements As Object, _
                Optional borderSizePx As Byte = 4, Optional borderColor As VBAcolors = Blue, _
                Optional backgroundColor As VBAcolors = VBAcolors.Unchanged, _
                Optional ScrollIntoView As Boolean = True, Optional UnHighlightLast As Boolean = True, _
                Optional bForgiving_UnHighlight As Boolean = False)

    Dim elements As WebElements
    Set elements = ObjToWebElements(elementOrElements)
    Dim element As WebElement

    If OnOffSwitch = True Then
        If lastElems Is Nothing Then Set lastElems = New WebElements
        If SetHighlight_ElemsOrigStyle Is Nothing Then Set SetHighlight_ElemsOrigStyle = New Scripting.Dictionary

        'handling of last highlighted elements
        If UnHighlightLast = True And lastElems.Count > 0 Then
            SetHighlight False, lastElems, bForgiving_UnHighlight:=True
            lastElems.RemoveAll
        End If
        
        For Each element In elements
            If Not lastElems.IsMember(element) Then lastElems.Add element
        Next element

        'Create the JS string for the highlighting SetAttribute
        Dim borderColorString As String, backgroundColorString As String
        If (borderSizePx = 0 Or borderColor = VBAcolors.Unchanged) And backgroundColor = VBAcolors.Unchanged Then
            Err.raise 1, "WebDriver", "Error: Highlight function used with no border and no background specified. Choose at least one or leave defaults."
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
            If ScrollIntoView And (element Is elements.Item(1)) Then element.ScrollIntoView
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
    With Err: Err.raise .Number, .source, .Description, .HelpFile, .HelpContext: End With
End If
End Sub

Public Sub SetHightlightFoundElems(OnOffSwitch As Boolean, Optional ElemBorderSizePx As Byte = 4, _
                Optional ElemBorderColor As VBAcolors = Blue, Optional ElemBackgroundColor As VBAcolors = VBAcolors.Unchanged, _
                Optional ElemsBorderSizePx As Byte = 4, _
                Optional ElemsBorderColor As VBAcolors = Cyan, Optional ElemsBackgroundColor As VBAcolors = VBAcolors.Unchanged)
    If OnOffSwitch = True Then
        'store border and color settings:
        If SetHightlightFoundElems_Settings Is Nothing Then Set SetHightlightFoundElems_Settings = New Collection
        With SetHightlightFoundElems_Settings
            .Add ElemBorderSizePx:  .Add ElemBorderColor:  .Add ElemBackgroundColor:
            .Add ElemsBorderSizePx: .Add ElemsBorderColor: .Add ElemsBackgroundColor:
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
        Set SetHightlightFoundElems_Settings = Nothing
    End If
End Sub

Public Function HasProperty(ByVal Name As String, element As WebElement) As Boolean
    If GetProperty(Name, element) = "" Then HasProperty = False Else HasProperty = True
End Function

Public Function HasAttribute(ByVal Name As String, element As WebElement) As Boolean
    If GetAttribute(Name, element) = "" Then HasAttribute = False Else HasAttribute = True
End Function

Public Function GetAttribute(ByVal Name As String, element As WebElement) As String
    'This returns pre-defined HTML element attributes. Once the browser parses the html code,
    'then a DOM node object with properties is created.
    'To access the state of a DOM node property - such as an input box value, use GetProperty.
    'If the attribute does not exist, then this returns vbNullString
    Dim Data As New Dictionary
    Data.Add "id", element.elementId
    Data.Add "name", Name
    Dim resp As Dictionary
    Set resp = Execute(tCMD.CMD_GET_ELEMENT_ATTRIBUTE, Data)
    If IsNull(resp("value")) Then
        GetAttribute = ""
    Else
        GetAttribute = resp("value")
    End If
End Function

Public Sub SetAttribute(element As WebElement, ByVal Name As String, ByVal val As String)
    'this sets the value of an html element attribute. Consider SendKeys or Action Chain for more control of user input elements.
    'If the attribute does not exist, then this does not throw an error.
    ExecuteScript "arguments[0].setAttribute('" & Name & "', arguments[1])", element, val
End Sub

Public Function GetAriaRole(element As WebElement) As String
    Dim Data As New Dictionary, resp As Dictionary
    
    If browser_ = svbaBrowser.Firefox Then
        Err.raise 1, "WebDriver", "Aria methods not supported for Firefox WebDriver"
    End If
    
    Data.Add "id", element.elementId
    Set resp = Execute(tCMD.CMD_GET_ELEMENT_ARIA_ROLE, Data)
    If IsNull(resp("value")) Then
        GetAriaRole = ""
    Else
        GetAriaRole = resp("value")
    End If
End Function

Public Function GetAriaLabel(element As WebElement) As String
    Dim Data As New Dictionary, resp As Dictionary
    
    If browser_ = svbaBrowser.Firefox Then
        Err.raise 1, "WebDriver", "Aria methods not supported for Firefox WebDriver"
    End If
    
    Data.Add "id", element.elementId
    Set resp = Execute(tCMD.CMD_GET_ELEMENT_ARIA_LABEL, Data)
    If IsNull(resp("value")) Then
        GetAriaLabel = ""
    Else
        GetAriaLabel = resp("value")
    End If
End Function

Public Function GetTagName(element As WebElement) As String
    Dim Data As New Dictionary
    Data.Add "id", element.elementId
    GetTagName = Execute(tCMD.CMD_GET_ELEMENT_TAG_NAME, Data)("value")
End Function

Public Sub SelectByValue(ByVal val As String, element As WebElement)
    Dim elem As WebElement
    
    Set elem = FindElement(by.tagName, "option[value =" & """" & val & """" & "]", element)
    If Not IsSelected(elem) Then
        elem.Click
    End If
End Sub

Public Sub SelectByVisibleText(ByVal text As String, element As WebElement)
    Dim elem As WebElement

    Set elem = FindElement(by.XPath, ".//option[normalize-space(.) = " & """" & text & """" & "]", element)
    If Not IsSelected(elem) Then
        elem.Click
    End If
End Sub

Public Function GetSelectTagsOptions(element As WebElement) As String()
    Dim elems As WebElements
    Dim ret() As String
    Dim i As Integer
    
    Set elems = FindElements(by.tagName, "option", element)
    ReDim ret(baseArrayIdx To elems.Count - 1 + baseArrayIdx)
    For i = 1 To elems.Count
        ret(i - 1 + baseArrayIdx) = elems(i).GetText
    Next
    
    GetSelectTagsOptions = ret
End Function

Public Function GetAllSelectedOptionsText(element As WebElement) As String()
    Dim elems As WebElements
    Dim ret() As String
    Dim i As Integer
    Dim j As Integer

    Set elems = FindElements(by.cssSelector, "option", element)
    ReDim ret(baseArrayIdx To elems.Count - 1 + baseArrayIdx)
    j = baseArrayIdx - 1
    For i = 1 To elems.Count
        If IsSelected(elems(i)) Then
            j = j + 1
            ret(j) = elems(i).GetText
        End If
    Next i
    
    If j >= baseArrayIdx Then ReDim Preserve ret(baseArrayIdx To j - 1 + baseArrayIdx) Else ret = VBA.Split(VBA.vbNullString)
    GetAllSelectedOptionsText = ret
End Function

Public Function GetSelectedOptionText(element As WebElement) As String
    Dim elems As WebElements
    Dim i As Integer

    Set elems = FindElements(by.cssSelector, "option", element)
    For i = 1 To elems.Count
        If IsSelected(elems(i)) Then
            GetSelectedOptionText = elems(i).GetText
            Exit Function
        End If
    Next i
    
    GetSelectedOptionText = ""
End Function

Public Sub DeSelectByValue(ByVal val As String, element As WebElement)
    Dim elem As WebElement

    Set elem = FindElement(by.cssSelector, "option[value =" & """" & val & """" & "]", element)
    If IsSelected(elem) Then
        elem.Click
    End If
End Sub

Public Sub DeSelectByVisibleText(ByVal text As String, element As WebElement)
    Dim elem As WebElement
    
    Set elem = FindElement(by.XPath, ".//option[normalize-space(.) = " & """" & text & """" & "]", element)
    If IsSelected(elem) Then
        elem.Click
    End If
End Sub

Public Function IsSelected(element As WebElement) As Boolean
    Dim Data As New Dictionary
    Data.Add "id", element.elementId
    IsSelected = Execute(tCMD.CMD_IS_ELEMENT_SELECTED, Data)("value")
End Function

Public Function IsMultiSelect(selectElement As WebElement) As Boolean
    IsMultiSelect = HasAttribute("multiple", selectElement)
End Function

Public Sub SelectByIndex(ByVal index As Integer, element As WebElement)
    SetFocus element
    ExecuteScript "arguments[0].options[" & CStr(index - baseDomIdx) & "].selected = true;", element
End Sub

Public Sub SelectAll(element As WebElement)
    Dim elems As WebElements
    Dim i As Integer
    Set elems = FindElements(by.cssSelector, "option", element)
    SetFocus element
    For i = 1 To elems.Count
        ExecuteScript "arguments[0].options[" & CStr(i - 1) & "].selected = true;", element
    Next i
End Sub

Public Sub DeSelectAll(element As WebElement)
    Dim elems As WebElements
    Dim i As Integer
    Set elems = FindElements(by.cssSelector, "option", element)
    SetFocus element
    For i = 1 To elems.Count
        ExecuteScript "arguments[0].options[" & CStr(i - 1) & "].selected = false;", element
    Next i
End Sub

Public Sub DeSelectByIndex(ByVal index As Integer, element As WebElement)
    Dim elems As WebElements
    Set elems = FindElements(by.cssSelector, "option", element)
    If IsSelected(elems(index - baseDomIdx + 1)) Then
        ExecuteScript "return arguments[0].selectedIndex = " & CStr(index - baseDomIdx), elems(index - baseDomIdx + 1)
        elems(index - baseDomIdx + 1).Click
    End If
End Sub

Public Function ExecuteScriptAsync(ByVal script As String, ParamArray scriptArgs() As Variant)
    Dim Data As New Dictionary, ElmData As New Dictionary
    Dim resp As Dictionary
    Dim Args()
    
    Data.Add "script", script
    
    If UBound(scriptArgs) >= 0 Then
        ReDim Args(0 To UBound(scriptArgs))
        Dim i As Long
        For i = 0 To UBound(scriptArgs)
            If TypeName(scriptArgs(i)) = "WebElement" Then
                'Convert WebElement to Dictionary that can be handled by WebDriver
                Dim elem As New Dictionary, welem As WebElement
                Set welem = scriptArgs(i)
                elem.Add ELEMENT_KEY, welem.elementId
                Set Args(i) = elem
            Else
                Args(i) = scriptArgs(i)
            End If
        Next i
    End If
    Data.Add "args", Args
    
    Set resp = Execute(tCMD.CMD_EXECUTE_SCRIPT_ASYNC, Data)
    
    Select Case TypeName(resp("value"))
    Case "Collection"
        Dim arry()
        arry = ScriptCollectionToConvertedArray(resp("value"))
        Dim allWebElements As Boolean
        allWebElements = True
        For i = LBound(arry) To UBound(arry)
            If TypeName(arry(i)) <> "WebElement" Then
                allWebElements = False
                Exit For
            End If
        Next i
        If allWebElements Then 'load into WebElements object
            Dim webElems As New WebElements, webElem As WebElement
            For i = LBound(arry) To UBound(arry)
                Set webElem = arry(i)
                webElems.Add webElem
            Next i
            Set ExecuteScriptAsync = webElems
        Else
            ExecuteScriptAsync = arry
        End If
    Case "Dictionary"
        Set ExecuteScriptAsync = ScriptDictToConvertedDict(resp("value"))
    Case Else 'other returns here
        If IsObject(resp("value")) Then
            Set ExecuteScriptAsync = resp("value")
        Else
            ExecuteScriptAsync = resp("value")
        End If
    End Select
End Function

Public Function ExecuteScript(ByVal script As String, ParamArray scriptArgs() As Variant)
    Dim Data As New Dictionary, ElmData As New Dictionary
    Dim resp As Dictionary
    Dim Args()
    
    Data.Add "script", script
    
    If UBound(scriptArgs) >= 0 Then
        ReDim Args(0 To UBound(scriptArgs))
        Dim i As Long
        For i = 0 To UBound(scriptArgs)
            If TypeName(scriptArgs(i)) = "WebElement" Then
                'Convert WebElement to Dictionary that can be handled by WebDriver
                Dim elem As New Dictionary, welem As WebElement
                Set welem = scriptArgs(i)
                elem.Add ELEMENT_KEY, welem.elementId
                Set Args(i) = elem
            Else
                Args(i) = scriptArgs(i)
            End If
        Next i
    End If
    Data.Add "args", Args
    
    Set resp = Execute(tCMD.CMD_EXECUTE_SCRIPT, Data)
    
    Select Case TypeName(resp("value"))
    Case "Collection"
        Dim arry()
        arry = ScriptCollectionToConvertedArray(resp("value"))
        If (Not arry) = -1 Then
            'Dim jc As New WebJSonConverter
            'Debug.Print jc.ConvertToJson(resp, 4)
            'is it ok to assume that this empty collection is of type WebElements?
            Set ExecuteScript = New WebElements
            Exit Function
        End If
        Dim allWebElements As Boolean
        allWebElements = True
        For i = LBound(arry) To UBound(arry)
            If TypeName(arry(i)) <> "WebElement" Then
                allWebElements = False
                Exit For
            End If
        Next i
        If allWebElements Then 'load into WebElements object
            Dim webElems As New WebElements, webElem As WebElement
            For i = LBound(arry) To UBound(arry)
                Set webElem = arry(i)
                webElems.Add webElem
            Next i
            Set ExecuteScript = webElems
        Else
            ExecuteScript = arry
        End If
    Case "Dictionary"
        Set ExecuteScript = ScriptDictToConvertedDict(resp("value"))
    Case Else 'other returns here
        If IsObject(resp("value")) Then
            Set ExecuteScript = resp("value")
        Else
            ExecuteScript = resp("value")
        End If
    End Select
End Function

'Convert dictionary from script to dictionary (converted to web element)
Private Function ScriptDictToConvertedDict(dict As Variant) As Object
    If dict.Exists(ELEMENT_KEY) Then
        ' Convert to WebElement
        Set ScriptDictToConvertedDict = ToWebElement(dict(ELEMENT_KEY))
        ' Return immediately (ignore other items if they exist)
        Exit Function
    End If
    Dim ret As New Dictionary
    Dim key
    For Each key In dict.keys
        If TypeName(dict(key)) = "Collection" Then
            ret(key) = ScriptCollectionToConvertedArray(dict(key))
        ElseIf TypeName(dict(key)) = "Dictionary" Then
            Set ret(key) = ScriptDictToConvertedDict(dict(key))
        ElseIf IsObject(dict(key)) Then
            Set ret(key) = dict(key)
        Else
            ret(key) = dict(key)
        End If
    Next
    Set ScriptDictToConvertedDict = ret
End Function

'Convert collection from script to array
Private Function ScriptCollectionToConvertedArray(col As Variant) As Variant()
    Dim arry()
    If col.Count > 0 Then
    ReDim arry(baseArrayIdx To col.Count - 1 + baseArrayIdx)
    Dim i As Integer
    For i = 1 To col.Count
        If TypeName(col(i)) = "Collection" Then
            arry(i - 1 + baseArrayIdx) = ScriptCollectionToConvertedArray(col(i))
        ElseIf TypeName(col(i)) = "Dictionary" Then
            Set arry(i - 1 + baseArrayIdx) = ScriptDictToConvertedDict(col(i))
        ElseIf IsObject(col(i)) Then
            Set arry(i - 1 + baseArrayIdx) = col(i)
        Else
            arry(i - 1 + baseArrayIdx) = col(i)
        End If
    Next
    End If
    ScriptCollectionToConvertedArray = arry
End Function

Public Sub SaveScreenshot(Optional ByVal filePath As String = ".\screenshot.png", Optional element As WebElement)
    Dim Data As New Dictionary
    Dim pngstr As String
        
    filePath = WebShared.GetFullLocalPath(filePath, defaultIOFolder_)
    
    If element Is Nothing Then
        pngstr = Execute(tCMD.CMD_SCREENSHOT, Data)("value")
    Else
        Data.Add "id", element.elementId
        pngstr = Execute(tCMD.CMD_ELEMENT_SCREENSHOT, Data)("value")
    End If
    
    SaveByteArrayToFile DecodeBase64(pngstr), filePath
End Sub

Public Sub PrintToPDF(Optional ByVal filePath As String = ".\printpage.pdf", Optional settings As WebPrintSettings)
    Dim Data As New Dictionary, pdfstr As String
    
    filePath = WebShared.GetFullLocalPath(filePath, defaultIOFolder_)
    
    If Not settings Is Nothing Then
        Set Data = settings.ToDictionary
    End If
    
    pdfstr = Execute(tCMD.CMD_PRINT_PAGE, Data)("value")

    SaveByteArrayToFile DecodeBase64(pdfstr), filePath
End Sub

Public Function GetRect(element As WebElement) As Dictionary
    'returns dictionary with following keys x, y, width, height
    Dim Data As New Dictionary
    Data.Add "id", element.elementId
    Set GetRect = Execute(tCMD.CMD_GET_ELEMENT_RECT, Data)("value")
End Function

Public Function GetPageSource() As String
    Dim Data As New Dictionary
    GetPageSource = Execute(tCMD.CMD_GET_PAGE_SOURCE, Data)("value")
End Function

Public Function GetCurrentWindowHandle() As String
    Dim Data As New Dictionary
    GetCurrentWindowHandle = Execute(tCMD.CMD_GET_CURRENT_WINDOW_HANDLE, Data)("value")
End Function

Public Function GetWindowHandles() As String()
    Dim Data As New Dictionary
    Dim handles As Collection
    Dim ret() As String
    Dim i As Integer
    
    Set handles = Execute(tCMD.CMD_GET_WINDOW_HANDLES, Data)("value")
    
    'To array of handles
    ReDim ret(baseArrayIdx To handles.Count - 1 + baseArrayIdx)
    For i = 1 To handles.Count   ' handles is Collection, not array
        ret(i - 1 + baseArrayIdx) = handles.Item(i)
    Next
    
    GetWindowHandles = ret
End Function

Public Function SwitchToNewWindow(Optional ByVal windowType As svbaWindowType = svbaTab) As String
    Dim Data As New Dictionary, hndl As String
    
    Dim wtype As String
    If windowType = svbaTab Then wtype = "tab" Else wtype = "window"
    
    Data.Add "type", wtype 'tab or window
    hndl = Execute(tCMD.CMD_NEW_WINDOW, Data)("value")("handle")
    
    SwitchToWindow hndl 'above just creates a new window - we still need to switch to it
    
    SwitchToNewWindow = hndl
End Function

Public Function SwitchToWindow(ByVal handleOrIndex As Variant) As String
    Dim Data As New Dictionary, hndl As String
    
    If IsNumeric(handleOrIndex) Then
        hndl = GetWindowHandles()(handleOrIndex)
    Else
        hndl = handleOrIndex
    End If
    
    Data.Add "handle", hndl
    Execute tCMD.CMD_SWITCH_TO_WINDOW, Data
    
    SwitchToWindow = hndl
End Function

Public Sub CloseWindow(Optional ByVal handleOrIndex As Variant)
    If Not IsMissing(handleOrIndex) Then SwitchToWindow handleOrIndex
    'closes the active window
    Dim Data As New Dictionary
    Execute tCMD.CMD_CLOSE, Data
End Sub

Public Function GetWindowRect() As Dictionary
    'returns dictionary with following keys x, y, width, height
    Dim Data As New Dictionary
    Set GetWindowRect = Execute(tCMD.CMD_GET_WINDOW_RECT, Data)("value")
End Function

Public Function SetWindowRect(Optional ByVal x As Variant, Optional ByVal y As Variant, Optional ByVal width As Variant, Optional ByVal height As Variant) As Dictionary
    'returns dictionary with following keys x, y, width, height
    Dim Data As New Dictionary
    Dim rect As Dictionary
    Set rect = GetWindowRect
    If IsMissing(width) Then width = rect("width")
    If IsMissing(height) Then height = rect("height")
    If IsMissing(x) Then x = rect("x")
    If IsMissing(y) Then y = rect("y")
    Data.Add "x", x
    Data.Add "y", y
    Data.Add "width", width
    Data.Add "height", height
    Set SetWindowRect = Execute(tCMD.CMD_SET_WINDOW_RECT, Data)("value")
End Function

Public Sub SetWindowSize(Optional ByVal width As Variant, Optional ByVal height As Variant)
    Dim rect As Dictionary
    Set rect = GetWindowRect
    If IsMissing(width) Then width = rect("width")
    If IsMissing(height) Then height = rect("height")
    SetWindowRect rect("x"), rect("y"), width, height
End Sub

Public Sub SetWindowPosition(Optional ByVal x As Variant, Optional ByVal y As Variant)
    'returns dictionary with following keys x, y, width, height
    Dim rect As Dictionary
    Set rect = GetWindowRect
    If IsMissing(x) Then x = rect("x")
    If IsMissing(y) Then y = rect("y")
    SetWindowRect x, y, rect("width"), rect("height")
End Sub

Public Sub MaximizeWindow()
    Dim Data As New Dictionary
    Execute tCMD.CMD_MAXIMIZE_WINDOW, Data
End Sub

Public Sub FullScreenWindow()
    Dim Data As New Dictionary
    Execute tCMD.CMD_FULLSCREEN_WINDOW, Data
End Sub

Public Sub MinimizeWindow()
    Dim Data As New Dictionary
    Execute tCMD.CMD_MINIMIZE_WINDOW, Data
End Sub

Public Sub SwitchToFrame(frameElement As WebElement)
    Dim Data As New Dictionary
    Dim elem As New Dictionary
    elem.Add ELEMENT_KEY, frameElement.elementId
    Data.Add "id", elem
    Execute tCMD.CMD_SWITCH_TO_FRAME, Data
End Sub

Public Sub SwitchToFrameByIndex(ByVal index As Integer)
    Dim Data As New Dictionary
    Data.Add "id", index - baseDomIdx
    Execute tCMD.CMD_SWITCH_TO_FRAME, Data
End Sub

Public Sub SwitchToDefaultContent()
    'switches to main document
    Dim Data As New Dictionary
    Data.Add "id", Null
    Execute tCMD.CMD_SWITCH_TO_FRAME, Data
End Sub

Public Sub SwitchToParentFrame()
    'if top-level frame then this switches to main document
    Dim Data As New Dictionary
    Execute tCMD.CMD_SWITCH_TO_PARENT_FRAME, Data
End Sub

Public Function GetCurrentFrameName() As String
    GetCurrentFrameName = ExecuteScript("return self.name")
End Function

Public Sub SetImplicitlyWait(Optional ByVal millisecond As Long = 0)
    Dim Data As New Dictionary
    Data.Add "implicit", millisecond
    Execute tCMD.CMD_SET_TIMEOUTS, Data
End Sub

Public Sub SetPageLoadTimeout(Optional ByVal millisecond As Long = 300000)
    Dim Data As New Dictionary
    Data.Add "pageLoad", millisecond
    Execute tCMD.CMD_SET_TIMEOUTS, Data
End Sub

Public Sub SetScriptTimeout(Optional ByVal millisecond As Long = 30000)
    Dim Data As New Dictionary
    Data.Add "script", millisecond
    Execute tCMD.CMD_SET_TIMEOUTS, Data
End Sub

Public Function GetImplicitlyWait() As Long
    Dim Data As New Dictionary
    GetImplicitlyWait = Execute(tCMD.CMD_GET_TIMEOUTS, Data)("value")("implicit")
End Function

Public Function GetPageLoadTimeout() As Long
    Dim Data As New Dictionary
    GetPageLoadTimeout = Execute(tCMD.CMD_GET_TIMEOUTS, Data)("value")("pageLoad")
End Function

Public Function GetScriptTimeout() As Long
    Dim Data As New Dictionary
    GetScriptTimeout = Execute(tCMD.CMD_GET_TIMEOUTS, Data)("value")("script")
End Function

Public Sub Wait(Optional ByVal milliseconds As Long = 300)
    'both of the system-level timers below work"
    '
    TimerWait milliseconds
    'Sleep milliseconds 'this works great, but lets try to do without windows api if possible
    '
    'these native methods below mysteriously interfere with/accept browser alerts -
    'maybe because they are browser-level, not system-level?
    '
    'ActionChain.Wait(milliseconds).Perform
    '
    'If milliseconds > 30000 Then SetScriptTimeout 2 * milliseconds '30000 is the default, so this isn't needed unless waitTime > 30 secs is needed
    'ExecuteScriptAsync "window.setTimeout(arguments[arguments.length - 1], arguments[0]);", milliseconds
End Sub

Public Function IsPresent(by_ As by, ByVal val As String) As Boolean
    Dim timeout As Long, resp As Dictionary, Data As Dictionary
    
    timeout = GetImplicitlyWait()
    If timeout <> 0 Then SetImplicitlyWait 0
    Set Data = ToSelector(by_, val)
    Set resp = Execute(tCMD.CMD_FIND_ELEMENT, Data, False) 'don't raise an error if not present
    If IsResponseError(resp) Then
        IsPresent = False
    Else
        IsPresent = True
    End If
    If timeout <> 0 Then SetImplicitlyWait timeout
End Function

Public Function IsEnabled(element As WebElement) As Boolean
    Dim Data As New Dictionary
    Data.Add "id", element.elementId
    IsEnabled = Execute(tCMD.CMD_IS_ELEMENT_ENABLED, Data)("value")
End Function

Public Function IsDisplayed(element As WebElement) As Boolean
    Dim Data As New Dictionary
    Data.Add "id", element.elementId
    IsDisplayed = Execute(tCMD.CMD_IS_ELEMENT_DISPLAYED, Data)("value")
End Function

Public Function GetBrowserName() As String
    GetBrowserName = WebShared.GetBrowserName(browser_)
End Function

Public Function GetAlertText() As String
    Dim Data As New Dictionary
    GetAlertText = Execute(tCMD.CMD_GET_ALERT_TEXT, Data)("value")
End Function

Public Sub SetAlertText(ByVal alertText As String)
    Dim Data As New Dictionary
    Data.Add "text", alertText
    Execute tCMD.CMD_SET_ALERT_VALUE, Data
End Sub

Public Sub AcceptAlert()
    Dim Data As New Dictionary
    Execute tCMD.CMD_ACCEPT_ALERT, Data
End Sub

Public Sub DismissAlert()
    Dim Data As New Dictionary
    Execute tCMD.CMD_DISMISS_ALERT, Data
End Sub

Public Function IsAlertPresent() As Boolean
    Dim Data As New Dictionary, resp As Dictionary
    Set resp = Execute(tCMD.CMD_GET_ALERT_TEXT, Data, False)
    If IsResponseError(resp) Then
        IsAlertPresent = False
    Else
        IsAlertPresent = True
    End If
End Function

Public Sub Clear(element As WebElement)
    Dim Data As New Dictionary
    Data.Add "id", element.elementId
    Data.Add "text", vbNullString
    Data.Add "value", vbNullString
    Execute tCMD.CMD_CLEAR_ELEMENT, Data
End Sub

Public Function GetTitle() As String
    Dim Data As New Dictionary
    GetTitle = Execute(tCMD.CMD_GET_TITLE, Data)("value")
End Function

Public Function GetShadowRoot(element As WebElement) As WebShadowRoot
    Dim Data As New Dictionary, resp As Dictionary
    Data.Add "id", element.elementId
    Set resp = Execute(tCMD.CMD_GET_ELEMENT_SHADOW_ROOT, Data)
    Set GetShadowRoot = ToShadowRoot(resp("value")(SHADOW_KEY))
End Function

Public Function GetCSSProperty(ByVal Name As String, element As WebElement) As String
    Dim Data As New Dictionary
    Data.Add "id", element.elementId
    Data.Add "name", Name
    GetCSSProperty = Execute(tCMD.CMD_GET_ELEMENT_VALUE_OF_CSS_PROPERTY, Data)("value")
End Function

Public Function GetInnerHTML(Optional element As WebElement) As String
    Dim script As String
    If element Is Nothing Then
        script = "return document.body.innerHTML"
        GetInnerHTML = ExecuteScript(script)
    Else
        script = "return arguments[0].innerHTML"
        GetInnerHTML = ExecuteScript(script, element)
    End If
End Function

Public Function GetOuterHTML(Optional element As WebElement) As String
    Dim script As String
    If element Is Nothing Then
        script = "return document.body.outerHTML"
        GetOuterHTML = ExecuteScript(script)
    Else
        script = "return arguments[0].outerHTML"
        GetOuterHTML = ExecuteScript(script, element)
    End If
End Function

' Get default capabilities for current browser
Public Function CreateCapabilities() As WebCapabilities
    Dim caps As New WebCapabilities
    
    'insure that this function is being called in the correct order
    If browser_ = 0 Or sessionId_ <> "" Then
        Err.raise 1, "WebDriver", "The ""CreateCapabilities"" Method must be invoked after the ""StartEdge"" or ""StartChrome"" Methods, and before ""OpenBrowser"" Method."
    End If
    
    'now initialize Capabilities
    caps.InitializeFor browser_
    
    Set caps.WebDriver = Me

    Set CreateCapabilities = caps
End Function

Public Function CreateCookies() As WebCookies
    Dim cks As New WebCookies
    Set cks.WebDriver = Me
    Set CreateCookies = cks
End Function

Public Sub SetFocus(element As WebElement)
    Dim script As String
    script = "arguments[0].focus({'preventScroll': arguments[1]})"
    ExecuteScript script, element, 0  'Scroll
End Sub

Public Sub SetFocusNoScroll(element As WebElement)
    Dim script As String
    script = "arguments[0].focus({'preventScroll': arguments[1]})"
    ExecuteScript script, element, 1  'No Scroll
End Sub

Public Sub ScrollToElement(element As WebElement, Optional ByVal xOffset As Integer = 0, Optional ByVal yOffset As Integer = 0)
    ScrollIntoView element, True
    ScrollBy xOffset, yOffset
End Sub

Public Sub ScrollIntoView(element As WebElement, Optional ByVal alignTop As Boolean = True)
    Dim script As String
    script = "arguments[0].scrollIntoView(" & LCase(alignTop) & ");"
    ExecuteScript script, element
End Sub

Public Sub ScrollBy(Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0)
    Dim script As String
    script = "window.scrollBy(" & xOffset & "," & yOffset & ");"
    ExecuteScript script
End Sub

Public Sub ScrollTo(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0)
    Dim script As String
    script = "window.scrollTo(" & x & "," & y & ");"
    ExecuteScript script
End Sub

Public Sub ScrollToTop()
    ScrollTo 0, 0
End Sub

Public Sub ScrollToBottom()
    Dim script As String
    script = "window.scrollTo(0, document.body.scrollHeight);"
    ExecuteScript script
End Sub

Public Function GetScrollHeight() As Integer
    Dim script As String
    script = "return document.body.scrollHeight;"
    GetScrollHeight = ExecuteScript(script)
End Function

Public Function GetAllCookies() As WebCookies
    Dim Data As New Dictionary, resp As Collection, ck As WebCookie, cks As WebCookies
    Set resp = Execute(tCMD.CMD_GET_ALL_COOKIES, Data)("value")
    'Dim jc As New WebJSonConverter
    'Debug.Print jc.ConvertToJson(resp, 4)
    Set cks = Me.CreateCookies
    Dim i As Long
    For i = 1 To resp.Count
        Set ck = New WebCookie
        If resp(i).Exists("name") Then ck.Name = resp(i)("name")
        If resp(i).Exists("domain") Then ck.Domain = resp(i)("domain")
        If resp(i).Exists("expiry") Then ck.ExpiryUnix = resp(i)("expiry")
        If resp(i).Exists("httpOnly") Then ck.HttpOnly = resp(i)("httpOnly")
        If resp(i).Exists("sameSite") Then ck.SameSite = resp(i)("sameSite")
        If resp(i).Exists("secure") Then ck.Secure = resp(i)("secure")
        If resp(i).Exists("value") Then ck.Value = resp(i)("value")
        If resp(i).Exists("path") Then ck.Path = resp(i)("path")
        cks.Add ck
    Next i
    Set GetAllCookies = cks
End Function

Public Function GetCookie(ByVal cookieName) As WebCookie
    Dim Data As New Dictionary, resp As Dictionary, ck As New WebCookie
    Data.Add "name", cookieName
    Set resp = Execute(tCMD.CMD_GET_COOKIE, Data)("value")
    If resp.Exists("name") Then ck.Name = resp("name")
    If resp.Exists("domain") Then ck.Domain = resp("domain")
    If resp.Exists("expiry") Then ck.ExpiryUnix = resp("expiry")
    If resp.Exists("httpOnly") Then ck.HttpOnly = resp("httpOnly")
    If resp.Exists("sameSite") Then ck.SameSite = resp("sameSite")
    If resp.Exists("secure") Then ck.Secure = resp("secure")
    If resp.Exists("value") Then ck.Value = resp("value")
    If resp.Exists("path") Then ck.Path = resp("path")
    Set GetCookie = ck
End Function

Public Sub DeleteCookie(ByVal cookieName As String)
    Dim Data As New Dictionary
    Data.Add "name", cookieName
    Execute tCMD.CMD_DELETE_COOKIE, Data
End Sub

Public Sub SetCookie(ck As WebCookie)
    Dim Data As New Dictionary
    'value can take a number string greater than 15 digits
    'need to set the following in JSonConverter else  will get an error on execute (see sendRequest):
    'jc.OptionsUseDoubleForLargeNumbers = True
    If ck.Name <> "" And ck.Value <> "" Then
        Data.Add "cookie", ck.ToDictionary
        Execute tCMD.CMD_ADD_COOKIE, Data
    End If
End Sub

Public Sub SetCookies(cks As WebCookies)
    Dim i As Long
    For i = 1 To cks.Count
        SetCookie cks.Item(i)
    Next i
End Sub

Public Sub DeleteAllCookies()
    Dim Data As New Dictionary
    Execute tCMD.CMD_DELETE_ALL_COOKIES, Data
End Sub

Public Function IsPageFound(ByVal url As String) As Boolean
    If GetCurrentURL() = url Then
        IsPageFound = True
    Else
        IsPageFound = False
    End If
End Function

Public Function GetSessionsInfo() As Collection
    'change to dictionary output?
    If browser_ = svbaBrowser.Firefox Then
        Err.raise 1, "WebDriver", "GetSessionsInfo not supported for Firefox WebDriver"
    End If
    Set GetSessionsInfo = Execute(tCMD.CMD_GET_ALL_SESSIONS)("value")
End Function

Public Function GetDriverStatus() As Dictionary
    'Dim jc As New WebJSonConverter
    Set GetDriverStatus = Execute(tCMD.CMD_STATUS)("value")
    'DriverVersion = resp("build")("version")
    'DriverVersion = Trim(Left(DriverVersion, InStr(DriverVersion, "(") - 1))
    'resp("message") 'gets info "MSEdgeDriver ready for new sessions."
    'resp("os")("name") 'gets info "Windows NT"
    'resp("os")("version") 'gets info "10.0.19042"
    'resp("os")("arch") 'gets info "x86_64"
    'resp("ready")'true or false
End Function

Public Function ActionChain() As WebActionChain
    Dim data_ As New Dictionary
    Set ActionChain = New WebActionChain
    'pass a reference of the driver to each link
    Set ActionChain.WebDriver = Me
    Set ActionChain.Data = data_
End Function

Public Function FindElementFromShadowRoot(by_ As by, ByVal val As String, sRoot As WebShadowRoot) As WebElement
    Dim Data As Dictionary
    Dim elem As WebElement
    
    Set Data = ToSelector(by_, val)
    Data.Add "sid", sRoot.shadowRootId
    Set elem = ToWebElement(Execute(tCMD.CMD_FIND_ELEMENT_FROM_SHADOW_ROOT, Data)("value")(ELEMENT_KEY))
    
    If Not SetHightlightFoundElems_Settings Is Nothing Then SetHightlightFoundElems_exec elem
    
    Set FindElementFromShadowRoot = elem
End Function

Public Function FindElementsFromShadowRoot(by_ As by, ByVal val As String, sRoot As WebShadowRoot) As WebElements
    Dim Data As Dictionary
    Dim elements As Collection
    Dim elems As WebElements
    Dim i As Integer
    
    Set Data = ToSelector(by_, val)
    Data.Add "sid", sRoot.shadowRootId
    
    Set elements = Execute(tCMD.CMD_FIND_ELEMENT_FROM_SHADOW_ROOT, Data)("value")
    
    For i = 1 To elements.Count
        elems.Add ToWebElement(elements(i)(ELEMENT_KEY))
    Next i
    
    If Not SetHightlightFoundElems_Settings Is Nothing Then SetHightlightFoundElems_exec elems
    
    Set elems.WebDriver = Me
    
    Set FindElementsFromShadowRoot = elems
End Function

Public Sub UploadFile(element As WebElement, ByVal filePath As String)
    'this is just a special wrapper for sendkeys (upload file mode)
    filePath = WebShared.GetFullLocalPath(filePath, defaultIOFolder_)
    SendKeys element, filePath
End Sub

Public Sub DragAndDrop(source As WebElement, target As WebElement)
    ActionChain.DragAndDrop(source, target).Perform
End Sub

Public Function WaitUntilReady(element As WebElement, Optional ByVal maxWaitTimeMS = 30000) As WebElement
    'waits until element is interactable, returns the input element for further action
    'such as "Click" on same line
    'see https://www.w3.org/TR/webdriver/#element-displayedness
    Dim startTime As Single, nowTime As Single, endTime As Single
    startTime = VBA.timer
    nowTime = startTime
    endTime = startTime + maxWaitTimeMS / 1000#
    
    Do While nowTime < endTime
        If element.IsDisplayed Then Exit Do
        nowTime = VBA.timer()
        Dim elapsedTime As Single
        If nowTime < startTime Then 'oops - someone is burning the midnight oil!
            endTime = endTime - elapsedTime
            startTime = 0
        End If
        elapsedTime = nowTime - startTime
        DoEvents    'yield to other processes.
    Loop

    Set WaitUntilReady = element
End Function

Public Function GetUserAgent() As String
    'this needs to be called after OpenBrowser method called
    If sessionId_ <> "" Then
        GetUserAgent = ExecuteScript("return navigator.userAgent;")
    Else
        Err.raise 1, "WebDriver", "GetUserAgent must be executed after OpenBrowser!"
    End If
End Function

Public Property Let DefaultIOFolder(ByVal folderPath As String)
    defaultIOFolder_ = WebShared.GetFullLocalPath(folderPath)
End Property

Public Property Get DefaultIOFolder() As String
    DefaultIOFolder = defaultIOFolder_
End Property

Public Property Let DefaultBinaryFolder(ByVal folderPath As String)
    defaultBinaryFolder_ = WebShared.GetFullLocalPath(folderPath)
End Property

Public Property Get DefaultBinaryFolder() As String
    DefaultBinaryFolder = defaultBinaryFolder_
End Property

Public Sub DeleteFiles(ParamArray filelist() As Variant)
    Dim fso As New FileSystemObject
    Dim i As Integer, k As Integer, j As Integer
    Dim strlist() As String
    For i = 0 To UBound(filelist)
        If IsArray(filelist(i)) Then
            For j = LBound(filelist(i)) To UBound(filelist(i))
                strlist = Split(filelist(i)(j), ",")
                For k = 0 To UBound(strlist)
                    On Error Resume Next 'file might not exists
                    fso.DeleteFile WebShared.GetFullLocalPath(strlist(k), defaultIOFolder_), True
                    On Error GoTo 0
                Next k
            Next j
        Else
            strlist = Split(filelist(i), ",")
            For k = 0 To UBound(strlist)
                On Error Resume Next 'file might not exists 'file might not exists
                fso.DeleteFile WebShared.GetFullLocalPath(strlist(k), defaultIOFolder_), True
                On Error GoTo 0
            Next k
        End If
    Next i
End Sub

Public Sub SaveHTMLToFile(ByVal snippet As String, Optional ByVal filePath As String = ".\snippet.html")
    'this is used to save test html snippets to file for subsequent browser navigation
    Dim fs As New FileSystemObject, ts As textstream
    filePath = WebShared.GetFullLocalPath(filePath, defaultIOFolder_)
    Set ts = fs.CreateTextFile(filePath, True)
    ts.Write snippet
    ts.Close
End Sub

Public Function GetParentElement(childElement As WebElement) As WebElement
    Set GetParentElement = childElement.FindElement(by.XPath, "./parent::*")
End Function

Public Function GetChildElements(parentElement As WebElement) As WebElements
    'this gets the immediate descendant elements - note that this is not same as element.FindElements(by.tagName, "*")
    Set GetChildElements = parentElement.FindElements(by.XPath, "./child::*")
End Function

Public Function TableToArray(table As WebElement, Optional ByVal skipHeader As Boolean = False, Optional ByVal skipFooter As Boolean = False, Optional ByVal createSpanData As Boolean = True) As Variant()
    Dim htmlDoc As New HTMLDocument, htmltable As htmltable

    'check if we really have a table
    If Not table.GetTagName = "table" Then 'report error
        Err.raise 1, "WebDriver", "Error: input element to TableToArray method must be a table"
    End If
    
    'convert the WebElement table into an html dom table object
    'using a html dom table for processing is greater than 25x faster than using Selenium selectors to process table
    htmlDoc.body.innerHTML = table.GetOuterHTML
    Set htmltable = htmlDoc.getElementsByTagName("table").Item(0)

    TableToArray = htmlTableToArray(htmltable, skipHeader, skipFooter, createSpanData)
End Function

Friend Function Execute(driverCommand, Optional parameters As Dictionary = Nothing, Optional ByVal raise As Boolean = True) As Dictionary
    Dim Method As String, Path As String
    Dim cmdArgs As New Dictionary, parmKey As Variant
    Dim response As Dictionary
    
    Method = driverCommand(0)
    Path = driverCommand(1)
    
    If parameters Is Nothing Then
        Set parameters = New Dictionary
    End If
    
    'Set session id
    If Not parameters.Exists("sessionId") Then
        parameters.Add "sessionId", sessionId_
    End If
    
    For Each parmKey In parameters.keys
        If InStr(Path, "$" & parmKey) > 0 Then 'path parameter
            Path = Replace(Path, "$" & parmKey, parameters(parmKey))
        Else 'non-path parameter
            cmdArgs.Add parmKey, parameters(parmKey)
        End If
    Next parmKey
    
    'there are commands that don't require sessionId in path: tCMD.CMD_STATUS, tCMD.CMD_NEW_SESSION, tCMD.CMD_GET_ALL_SESSIONS, tCMD.CMD_SHUTDOWN
    If cmdArgs.Exists("sessionId") Then cmdArgs.Remove "sessionId"
    
    'Send request to selenium server
    Set response = SendRequest(Method, driverUrl_ & Path, cmdArgs)
    
    If raise And IsResponseError(response) Then 'raise error if occurs
        Debug.Print GetResponseErrorMessage(response)
        Dim lHttpStatus As Long
        Select Case response.Item("value").Item("error")
            'Complete list of error codes & numbers taken from here https://www.w3.org/TR/webdriver/#dfn-error-code as of 18-oct-2022
            Case "element click intercepted", "element not interactable", "insecure certificate", "invalid argument", "invalid cookie domain", "invalid element state", "invalid selector"
                lHttpStatus = 400 'Bad Request
            Case "invalid session id", "no such alert", "no such cookie", "no such element", "no such frame", "no such window", "no such shadow root", "stale element reference", "detached shadow root", "unknown command"
                lHttpStatus = 404 'Not Found
            Case "javascript Error", "move target out of bounds", "script timeout error", "session not created", "timeout", "unable to set cookie", "unable to capture screen", "unexpected alert open", "unknown error", "unsupported operation"
                lHttpStatus = 500 'Internal Server Error
            Case "unknown method"
                lHttpStatus = 405 'Method Not Allowed
        End Select
        Err.raise lHttpStatus, "WebDriver", GetResponseErrorMessage(response)
    End If

    Set Execute = response 'always pass a dictionary object - let caller parse info based on context, including error
End Function

Private Function ToWebElement(ByVal elementId As String) As WebElement
    Dim element As New WebElement
    Set element.driver = Me
    element.elementId = elementId
    Set ToWebElement = element
End Function

Private Function ToShadowRoot(ByVal shadowRootId As String) As WebShadowRoot
    Dim sr As New WebShadowRoot
    Set sr.driver = Me
    sr.shadowRootId = shadowRootId
    Set ToShadowRoot = sr
End Function

Private Function GetResponseErrorMessage(resp As Dictionary) As String
    Dim jc As New WebJSonConverter
    GetResponseErrorMessage = ""
    If TypeName(resp("value")) = "Dictionary" Then
        If resp("value").Exists("error") Then
            GetResponseErrorMessage = jc.ConvertToJson(resp("value")("message"), 4)
        End If
    End If
End Function

Private Function IsResponseError(resp As Dictionary) As Boolean
    IsResponseError = False
    If TypeName(resp("value")) = "Dictionary" Then
        If resp("value").Exists("error") Then
            IsResponseError = True
        End If
    End If
End Function

Private Function SendRequest(ByVal Method As String, ByVal url As String, Optional Data As Dictionary = Nothing) As Dictionary
    Dim client As MSXML2.ServerXMLHTTP60, jc As New WebJSonConverter
    Set client = New MSXML2.ServerXMLHTTP60

    client.Open Method, url
    
    'set receiveTimeout to infinite to accommodate longer than 30 sec (default) pageloads
    client.setTimeouts 0, 60000, 30000, 0
    
    If Method = "POST" Or Method = "PUT" Then
        client.setRequestHeader "Content-Type", "application/json; charset=utf-8"
        client.setRequestHeader "Cache-Control", "no-cache"
        client.send jc.ConvertToJson(Data)  '
    Else
        client.send
    End If

    Do While client.readyState < 4
        DoEvents
    Loop

    Set SendRequest = jc.ParseJSON(client.responseText)
End Function

Private Function params(ParamArray keysAndValues()) As Dictionary
    Dim dict As New Dictionary
    Dim i As Integer
    For i = 0 To UBound(keysAndValues) - 1 Step 2
        dict.Add keysAndValues(i), keysAndValues(i + 1)
    Next i
    Set params = dict
End Function

Private Function rgbLongToString(rgbColor As Long) As String
    If rgbColor < RGB(0, 0, 0) Or rgbColor > RGB(255, 255, 255) Then Err.raise 1, "WebDriver", "Error: argument rgbColor out of range (0 to 16777215)"
    Dim r As Long, g As Long, b As Long
    r = rgbColor And &HFF
    g = (rgbColor And &H100FF00) / &H100
    b = (rgbColor And &HFF0000) / &H10000
    rgbLongToString = "rgb(" & r & "," & g & "," & b & ")"
End Function

Private Function ObjToWebElements(elementOrElements As Object) As WebElements
    Select Case TypeName(elementOrElements)
        Case "WebElement": Set ObjToWebElements = New WebElements: ObjToWebElements.Add elementOrElements
        Case "WebElements": Set ObjToWebElements = elementOrElements
        Case Else: Err.raise 1, "WebDriver", "Error: the elementOrElements argument is not of valid type (WebElement or WebElements)"
    End Select
End Function

Private Sub InitCommands()
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
        '.CMD_UPLOAD_FILE = Array("POST", "/session/$sessionId/file")
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
        .CMD_SET_ALERT_VALUE = Array("POST", "/session/$sessionId/alert/text")
        .CMD_GET_ALERT_TEXT = Array("GET", "/session/$sessionId/alert/text")
        .CMD_CLICK = Array("POST", "/session/$sessionId/click")
        '.CMD_ACTIONS = Array("POST", "/session/$sessionId/actions") 'used in ActionChain.cls
        '.CMD_CLEAR_ACTIONS = Array("DELETE", "/session/$sessionId/actions") 'used in ActionChain.cls
        .CMD_SET_WINDOW_RECT = Array("POST", "/session/$sessionId/window/rect")
        .CMD_GET_WINDOW_RECT = Array("GET", "/session/$sessionId/window/rect")
        .CMD_MAXIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/maximize")
        '.CMD_GET_LOCATION = Array("GET", "/session/$sessionId/location")
        '.CMD_SET_LOCATION = Array("POST", "/session/$sessionId/location")
        '.CMD_GET_APP_CACHE_STATUS = Array("GET", "/session/$sessionId/application_cache/status")
        '.CMD_GET_NETWORK_CONNECTION = Array("GET", "/session/$sessionId/network_connection")
        '.CMD_SET_NETWORK_CONNECTION = Array("POST", "/session/$sessionId/network_connection")
        '.CMD_GET_LOG = Array("POST", "/session/$sessionId/se/log")  'see logging option in StartEdge and StartChrome methods
        '.CMD_GET_AVAILABLE_LOG_TYPES = Array("GET", "/session/$sessionId/se/log/types") 'see logging option in StartEdge and StartChrome methods
        '.CMD_CURRENT_CONTEXT_HANDLE = Array("GET", "/session/$sessionId/se/context")
        '.CMD_CONTEXT_HANDLES = Array("GET", "/session/$sessionId/se/contexts")
        '.CMD_SWITCH_TO_CONTEXT = Array("POST", "/session/$sessionId/context")
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

Private Sub Class_Initialize()
    InitCommands
    Me.CommandWindowStyle = vbHide
    'note that user has control through the DefaultIOFolder and DefaultBinaryFolder
    'properties so these below are just reasonable starting points...
    'ThisWorkbook.Path and ActiveVBAProjectFolderPath will be different for a project
    'referencing the SeleniumVBA Add-in, but same otherwise
    Me.DefaultBinaryFolder = ThisWorkbook.Path
    Me.DefaultIOFolder = ActiveVBAProjectFolderPath
End Sub

Private Sub TimerWait(ByVal durationMS As Long)
    'pause in milliseconds
    Dim startTime As Single, endTime As Single, nowTime As Single, elapsedTime As Single
    startTime = VBA.timer()
    nowTime = startTime
    endTime = startTime + durationMS / 1000#
    Do While nowTime < endTime
        nowTime = VBA.timer()
        If nowTime < startTime Then
            endTime = endTime - elapsedTime
            startTime = 0
        End If
        elapsedTime = nowTime - startTime
        DoEvents    'yield to other processes.
    Loop
End Sub

Private Function DecodeBase64(ByVal strData As String) As Byte()
    Dim domDoc As New MSXML2.DOMDocument60
    Dim domNode As MSXML2.IXMLDOMElement
    'create node with type of base 64 and decode
    Set domNode = domDoc.createElement("b64")
    domNode.DataType = "bin.base64"
    domNode.text = strData
    DecodeBase64 = domNode.nodeTypedValue
End Function

Private Sub SaveByteArrayToFile(byteArray() As Byte, ByVal filePath As String)
    Dim fp As Long, fso As New FileSystemObject
    If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
    fp = FreeFile
    Open filePath For Binary Access Write As #fp
    Put #fp, 1, byteArray
    Close fp
End Sub

Private Sub KillDriver(ByVal browser As svbaBrowser, Optional ByVal port As Integer = 0)
    'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-process
    Dim objService As WbemScripting.SWbemServices
    Dim objDriverList As WbemScripting.SWbemObjectSet
    Dim objParentList As WbemScripting.SWbemObjectSet
    Dim objDriver As WbemScripting.SWbemObject
    Dim objParent As WbemScripting.SWbemObject
    Dim sImageName As String, ret As Variant, pid As String
    
    Select Case browser
    Case svbaBrowser.Chrome
        sImageName = "chromedriver%"
    Case svbaBrowser.Edge
        sImageName = "msedgedriver%"
    Case svbaBrowser.Firefox
        sImageName = "geckodriver%"
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
            Select Case LCase(objParent.Name)
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
        ret = TaskKillByPid(pid)
10:
    Next objDriver
End Sub

Private Function TaskKillByPid(ByVal pid As String)
    Dim wsh As New IWshRuntimeLibrary.WshShell
    TaskKillByPid = wsh.Run("taskkill /f /t /pid " & pid, 0, True)
End Function

Private Function ActiveVBAProjectFolderPath() As String
    Dim fso As New FileSystemObject, strpath As String
    strpath = ""
    'if the Excel file holding the active vba project has not yet been saved, then this will throw an error
    '...so return vbNullString in that case
    On Error Resume Next
    strpath = Application.VBE.ActiveVBProject.Filename
    On Error GoTo 0
    If strpath <> "" Then strpath = fso.GetParentFolderName(strpath)
    ActiveVBAProjectFolderPath = strpath
End Function

Private Function htmlTableToArray(table As htmltable, Optional ByVal skipHeader As Boolean = False, Optional ByVal skipFooter As Boolean = False, Optional ByVal createSpanData As Boolean = True) As Variant()
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
        For Each row In table.rows
            For Each cell In row.cells
                rowSpan = cell.rowSpan
                If rowSpan > 1 Then
                    For i = row.RowIndex + 1 To row.RowIndex + rowSpan - 1
                        table.rows(i).insertCell(cell.cellIndex).innerText = cell.innerText
                    Next i
                End If
            Next cell
        Next row
        
        'now process column spans
        For Each row In table.rows
            For Each cell In row.cells
                colSpan = cell.colSpan
                If colSpan > 1 Then
                    For i = 1 To colSpan - 1
                        row.insertCell(row.cells.Length).innerText = cell.innerText
                    Next i
                End If
            Next cell
        Next row
    End If
    
    'calculate the number of array columns needed for output
    For Each row In table.rows
        If row.cells.Length > maxCols Then maxCols = row.cells.Length
    Next row
    
    'dimension output array
    ReDim v(baseArrayIdx To table.rows.Length - 1 + baseArrayIdx, baseArrayIdx To maxCols - 1 + baseArrayIdx)
    
    'extract cell data (td and th elements) from each row
    For Each row In table.rows
        For Each cell In row.cells
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
