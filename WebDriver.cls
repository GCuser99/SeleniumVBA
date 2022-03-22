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
' SeleniumVBA v0.0.2
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
' TinySeleniumVBA v0.1.2
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

'References:
'Microsoft Script Control 1.0 for 'CreateObject("ScriptControl") MSScriptControl.ScriptControl
'Microsoft XML, v6.0  for MSXML2.ServerXMLHTTP60, MSXML2.DOMDocument60 'CreateObject("MSXML2.DOMDocument")
'Microsoft Scripting Runtime for Scripting.FileSystemObject,Scripting.Dictionary  CreateObject("Scripting.Dictionary")
'Microsoft HTML Object Library

Private Const CheckDriverBrowserVersionAlignment = False 'change to True if auto-check-and-update desired
Private sessionId As String
Private urlBase As String

Private Const ELEMENT_KEY = "element-6066-11e4-a52e-4f735466cecf"
Private Const SHADOW_KEY = "shadow-6066-11e4-a52e-4f735466cecf"

Private CMD_STATUS
Private CMD_NEW_SESSION
Private CMD_GET_ALL_SESSIONS
Private CMD_QUIT
Private CMD_W3C_GET_CURRENT_WINDOW_HANDLE
Private CMD_W3C_GET_WINDOW_HANDLES
Private CMD_GET
Private CMD_GO_FORWARD
Private CMD_GO_BACK
Private CMD_REFRESH
Private CMD_W3C_EXECUTE_SCRIPT
Private CMD_W3C_EXECUTE_SCRIPT_ASYNC                 'not wrapped yet
Private CMD_GET_CURRENT_URL
Private CMD_GET_TITLE
Private CMD_GET_PAGE_SOURCE
Private CMD_SCREENSHOT
Private CMD_ELEMENT_SCREENSHOT
Private CMD_FIND_ELEMENT
Private CMD_FIND_ELEMENTS
Private CMD_W3C_GET_ACTIVE_ELEMENT
Private CMD_FIND_CHILD_ELEMENT
Private CMD_FIND_CHILD_ELEMENTS
Private CMD_CLICK_ELEMENT
Private CMD_CLEAR_ELEMENT
Private CMD_GET_ELEMENT_TEXT
Private CMD_SEND_KEYS_TO_ELEMENT
Private CMD_UPLOAD_FILE                              'not wrapped yet
Private CMD_GET_ELEMENT_TAG_NAME
Private CMD_IS_ELEMENT_SELECTED
Private CMD_IS_ELEMENT_ENABLED
Private CMD_IS_ELEMENT_DISPLAYED
Private CMD_GET_ELEMENT_RECT
Private CMD_GET_ELEMENT_ATTRIBUTE
Private CMD_GET_ELEMENT_PROPERTY
Private CMD_GET_ALL_COOKIES
Private CMD_ADD_COOKIE
Private CMD_GET_COOKIE
Private CMD_DELETE_ALL_COOKIES
Private CMD_DELETE_COOKIE
Private CMD_SWITCH_TO_FRAME
Private CMD_SWITCH_TO_PARENT_FRAME
Private CMD_SWITCH_TO_WINDOW
Private CMD_NEW_WINDOW
Private CMD_CLOSE
Private CMD_GET_ELEMENT_VALUE_OF_CSS_PROPERTY
Private CMD_SET_TIMEOUTS
Private CMD_GET_TIMEOUTS
Private CMD_W3C_DISMISS_ALERT
Private CMD_W3C_ACCEPT_ALERT
Private CMD_W3C_SET_ALERT_VALUE
Private CMD_W3C_GET_ALERT_TEXT
Private CMD_CLICK
'Private CMD_W3C_ACTIONS                              'used in Action.cls
'Private CMD_W3C_CLEAR_ACTIONS                        'used in Action.cls
Private CMD_SET_WINDOW_RECT
Private CMD_GET_WINDOW_RECT
Private CMD_W3C_MAXIMIZE_WINDOW
Private CMD_GET_LOCATION                             'not wrapped yet
Private CMD_SET_LOCATION                             'not wrapped yet
Private CMD_GET_APP_CACHE_STATUS                     'not wrapped yet
Private CMD_GET_NETWORK_CONNECTION                   'not wrapped yet
Private CMD_SET_NETWORK_CONNECTION                   'not wrapped yet
'Private CMD_GET_LOG                                  'see logging option in Edge and Chrome methods
'Private CMD_GET_AVAILABLE_LOG_TYPES                  'see logging option in Edge and Chrome methods
Private CMD_CURRENT_CONTEXT_HANDLE                   'not wrapped yet
Private CMD_CONTEXT_HANDLES                          'not wrapped yet
Private CMD_SWITCH_TO_CONTEXT                        'not wrapped yet
Private CMD_FULLSCREEN_WINDOW
Private CMD_MINIMIZE_WINDOW
Private CMD_SHUTDOWN
Private CMD_PRINT_PAGE                               'only works in -headless mode but where does the pdf file go?
Private CMD_GET_ELEMENT_SHADOW_ROOT
Private CMD_FIND_ELEMENT_FROM_SHADOW_ROOT
Private CMD_FIND_ELEMENTS_FROM_SHADOW_ROOT
Private CMD_GET_ELEMENT_ARIA_ROLE
Private CMD_GET_ELEMENT_ARIA_LABEL

Public Enum by
    ID = 0
    tagname = 1
    className = 2
    name = 3
    cssSelector = 4
    XPath = 5
    linkText = 6
    PartialLinkText = 7
End Enum

#If VBA7 And Win64 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal millisecond As Long)
#End If

' GetSystemMetrics              '2021/7/13 add ishi
Private Const SM_CXSCREEN = 0           'Screen Withd (DPI)
Private Const SM_CYSCREEN = 1           'Screen Hright (DPI)
Private Const SM_CMOUSEBUTTONS = 43     'Mouse Bottons
Private Const SM_CMONITORS = 80         'Monitors

#If VBA7 And Win64 Then
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private browserName As String
Private platformName_data As String
Private platform_data As String
Private browserOptions_key As String
Private AppWinStyle As VbAppWinStyle

Public Sub Edge(Optional ByVal driverPath As String = "msedgedriver.exe", Optional ByVal driverUrl As String = "http://localhost:9515", Optional ByVal enableLogging As Boolean = False, Optional ByVal logFilePath As String = ".\edge_verbose.log")
    Dim fso As New Scripting.FileSystemObject
    
    browserName = "msedge"
    driverPath = GetAbsolutePath(driverPath) 'support for relative path such as driverPath=".\msedgedriver.exe"
    
    If enableLogging Then
        logFilePath = GetAbsolutePath(logFilePath)
        serviceArgs = " --verbose --log-path=" & Chr(34) & logFilePath & Chr(34)
    End If
    
    If CheckDriverBrowserVersionAlignment Then
        If Not DriverAndBrowserAligned(browserName, driverPath) Then Exit Sub
    End If
    
    Start driverPath, driverUrl, serviceArgs
    
End Sub

Public Sub Chrome(Optional ByVal driverPath As String = "chromedriver.exe", Optional ByVal driverUrl As String = "http://localhost:9515", Optional ByVal enableLogging As Boolean = False, Optional ByVal logFilePath As String = ".\chrome_verbose.log")
    Dim fso As New Scripting.FileSystemObject
    
    browserName = "chrome"
    driverPath = GetAbsolutePath(driverPath) 'support for relative path such as driverPath=".\chromedriver.exe"
    
    If enableLogging Then
        logFilePath = GetAbsolutePath(logFilePath)
        serviceArgs = " --verbose --log-path=" & Chr(34) & logFilePath & Chr(34)
    End If
    
    If CheckDriverBrowserVersionAlignment Then
        If Not DriverAndBrowserAligned(browserName, driverPath) Then Exit Sub
    End If
    
    Start driverPath, driverUrl, serviceArgs

End Sub

Private Function DriverAndBrowserAligned(ByVal browserName As String, ByVal driverPath As String) As Boolean
    Dim wdmgr As New WebDriverManager
    
    driverPath = GetAbsolutePath(driverPath) 'support for relative path such as driverPath=".\msedgedriver.exe"
    
    DriverAndBrowserAligned = True
    
    If Not wdmgr.IsInstalledDriverCompatible(browserName, , driverPath) Then
        resp = MsgBox("WebDriver is not compatible with installed browser - would you like to install the compatible WebDriver?", vbYesNo)
        If resp = vbYes Then
            bverinstalled = wdmgr.GetInstalledBrowserVersion(browserName)
            dvercompat = wdmgr.GetCompatibleDriverVersion(browserName, bverinstalled)
            wdmgr.DownloadAndInstall browserName, dvercompat, driverPath
        Else
            DriverAndBrowserAligned = False
        End If
    End If

End Function

Private Sub Start(ByVal driverPath As String, Optional ByVal driverUrl As String = "http://localhost:9515", Optional ByVal serviceArgs As String = vbNullString)
    Dim Port As String
    
    'make sure driver path is "quoted" in case we have spaces in path
    If Left(driverPath, 1) <> Chr(34) Then driverPath = Chr(34) & driverPath
    If Right(driverPath, 1) <> Chr(34) Then driverPath = driverPath & Chr(34)
    
    Port = vbNullString
    If InStr(driverPath, "edgedriver") > 0 Then
        Port = " --port=" & mid(driverUrl, InStrRev(driverUrl, ":") + 1)
    ElseIf InStr(driverPath, "chromedriver") > 0 Then
        Port = " --port=" & mid(driverUrl, InStrRev(driverUrl, ":") + 1)
    End If
    
    If serviceArgs <> "" Then
        If Left(serviceArgs, 1) <> " " Then serviceArgs = " " & serviceArgs
    End If
    
    ' Start WebDriver executable
    If shell(driverPath & serviceArgs & Port, AppWinStyle) = 0 Then
        MsgBox "Failed in starting WebDriver." & vbCrLf & _
               "WebDriverPath : " & driverPath & Port, vbCritical + vbOKOnly
        End
    End If
    
    'Set WebDriver url
    urlBase = driverUrl
    
End Sub

Public Sub Shutdown()
    'this will shut down all WebDriver instances, not just the one it is being called from
    Execute CMD_SHUTDOWN
End Sub

Public Function OpenBrowser(Optional caps As Capabilities, Optional ByVal invisible As Boolean = False)
    Dim resp As Dictionary
    
    If sessionId <> "" Then Exit Function 'only one browser per webdriver instance
    
    If caps Is Nothing Then
        Set caps = Me.CreateCapabilities
    End If
        
    If invisible Then
        'user may have already added headless, but selenium doesn't seem to care about multiple arg entries
        caps.AddArgument "--headless"
    End If
    
    Set resp = Execute(CMD_NEW_SESSION, Params("capabilities", caps.Dictionary))("value")
    
    sessionId = resp("sessionId")
    
End Function

Public Sub CloseBrowser()
    Dim data As New Dictionary
    Execute CMD_QUIT, data
    sessionId = ""
End Sub

Public Function Navigate(ByVal url As String, Optional ByVal timeOutms, Optional ByVal raise As Boolean = True) As Boolean
    Dim data As New Dictionary, resp As Object, savtimeOutms, jc As New JSonConverter
    
    data.Add "url", url
    If Not IsMissing(timeOutms) Then
        savtimeOutms = Me.GetPageLoadTimeout()
        If savtimeOutms <> timeOutms Then Me.SetPageLoadTimeout timeOutms
    End If
    
    Set resp = Execute(CMD_GET, data, raise)

    If Not IsMissing(timeOutms) Then
        If savtimeOutms <> timeOutms Then Me.SetPageLoadTimeout savtimeOutms
    End If
    
    If IsResponseError(resp) Then
        Navigate = False
        Debug.Print GetResponseErrorMessage(resp)
    Else
        Navigate = True
    End If
    
End Function

Public Function FindElement(by_ As by, ByVal value As String, Optional parentElement As WebElement) As WebElement
    Dim data As Dictionary
    Set data = ToSelector(by_, value)
    Dim cmd
    If Not parentElement Is Nothing Then
        data.Add "id", parentElement.elementId
        cmd = CMD_FIND_CHILD_ELEMENT
    Else
        cmd = CMD_FIND_ELEMENT
    End If
    
    Set FindElement = ToWebElement(Execute(cmd, data)("value")(ELEMENT_KEY))
End Function

Public Function FindElementByID(ByVal value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByID = FindElement(by.ID, value, parentElement)
End Function

Public Function FindElementByTagName(ByVal value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByTagName = FindElement(by.tagname, value, parentElement)
End Function

Public Function FindElementByClassName(ByVal value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByClassName = FindElement(by.className, value, parentElement)
End Function

Public Function FindElementByCssSelector(ByVal value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByCssSelector = FindElement(by.cssSelector, value, parentElement)
End Function

Public Function FindElementByLinkText(ByVal value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByLinkText = FindElement(by.linkText, value, parentElement)
End Function

Public Function FindElementByName(ByVal value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByName = FindElement(by.name, value, parentElement)
End Function

Public Function FindElementByPartialLinkText(ByVal value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByPartialLinkText = FindElement(by.PartialLinkText, value, parentElement)
End Function

Public Function FindElementByXPath(ByVal value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByXPath = FindElement(by.XPath, value, parentElement)
End Function

Public Function FindElementsByXPath(ByVal value As String, Optional parentElement As WebElement) As WebElement()
    FindElementsByXPath = FindElements(by.XPath, value, parentElement)
End Function

Public Function FindElementsByID(ByVal value As String, Optional parentElement As WebElement) As WebElement()
    FindElementsByID = FindElements(by.ID, value, parentElement)
End Function

Public Function FindElementsByTagName(ByVal value As String, Optional parentElement As WebElement) As WebElement()
    FindElementsByTagName = FindElements(by.tagname, value, parentElement)
End Function

Public Function FindElementsByClassName(ByVal value As String, Optional parentElement As WebElement) As WebElement()
    FindElementsByClassName = FindElements(by.className, value, parentElement)
End Function

Public Function FindElementsByCssSelector(ByVal value As String, Optional parentElement As WebElement) As WebElement()
    FindElementsByCssSelector = FindElements(by.cssSelector, value, parentElement)
End Function

Public Function FindElementsByLinkText(ByVal value As String, Optional parentElement As WebElement) As WebElement()
    FindElementsByLinkText = FindElements(by.linkText, value, parentElement)
End Function

Public Function FindElementsByName(ByVal value As String, Optional parentElement As WebElement) As WebElement()
    FindElementsByName = FindElements(by.name, value, parentElement)
End Function

Public Function FindElementsByPartialLinkText(ByVal value As String, Optional parentElement As WebElement) As WebElement()
    FindElementsByPartialLinkText = FindElements(by.PartialLinkText, value, parentElement)
End Function

Public Function FindElements(by_ As by, ByVal value As String, Optional parentElement As WebElement) As WebElement()
    Dim data As Dictionary
    
    Set data = ToSelector(by_, value)
    
    Dim cmd
    If Not parentElement Is Nothing Then
        data.Add "id", parentElement.elementId
        cmd = CMD_FIND_CHILD_ELEMENTS
    Else
        cmd = CMD_FIND_ELEMENTS
    End If
    
    Dim elements
    Set elements = Execute(cmd, data)("value")
    
    Dim ret() As WebElement
    Dim i As Integer
    For i = 0 To elements.Count - 1
        ReDim Preserve ret(i)
        Set ret(i) = ToWebElement(elements(i + 1)(ELEMENT_KEY))
    Next
    
    FindElements = ret
End Function

Private Function ToSelector(by_ As by, ByVal value As String) As Dictionary
    Dim data As New Dictionary
    Select Case by_
    Case by.ID
        data.Add "using", "css selector"
        data.Add "value", "[id=""" + value + """]"
    Case by.tagname
        data.Add "using", "css selector"
        data.Add "value", value
    Case by.className
        data.Add "using", "css selector"
        data.Add "value", "." + value
    Case by.name
        data.Add "using", "css selector"
        data.Add "value", "[name=""" + value + """]"
    Case by.cssSelector
        data.Add "using", "css selector"
        data.Add "value", value
    Case by.XPath
        data.Add "using", "xpath"
        data.Add "value", value
    Case by.linkText
        data.Add "using", "link text"
        data.Add "value", value
    Case by.PartialLinkText
        data.Add "using", "partial link text"
        data.Add "value", value
    Case Else
        data.Add "using", "css selector"
        data.Add "value", value
    End Select
    
    Set ToSelector = data
End Function

Private Function ToWebElement(ByVal elementId As String) As WebElement
    Dim element As New WebElement
    Set element.Driver = Me
    element.elementId = elementId
    Set ToWebElement = element
End Function

Private Function ToShadowRoot(ByVal shadowRootId As String) As ShadowRoot
    Dim sr As New ShadowRoot
    Set sr.Driver = Me
    sr.shadowRootId = shadowRootId
    Set ToShadowRoot = sr
End Function

Public Sub SendKeys(element As WebElement, ByVal Keys As String, Optional clearBeforeTyping As Boolean = True)
    Dim data As New Dictionary
    
    data.Add "id", element.elementId
    If clearBeforeTyping = True Then Execute CMD_CLEAR_ELEMENT, data
    
    'split key sequence into chunks with special keys (pattern "\ue0*") by themselves
    ReDim chunk(0 To Len(Keys) - 1)
    i = -1
    Do Until Keys = ""
        ipos = InStr(Keys, "\ue0")
        If ipos > 0 Then
            If ipos > 1 Then 'text to left is chunk
                i = i + 1
                chunk(i) = Left(Keys, ipos - 1)
            End If
            'now process special key
            i = i + 1
            chunk(i) = mid(Keys, ipos, 6)
            Keys = Right(Keys, Len(Keys) - (ipos + 5))
        Else
            i = i + 1
            chunk(i) = Keys
            Keys = ""
        End If
    Loop
    
    ReDim Preserve chunk(0 To i)
    
    'initialize param keys
    data.Add "text", ""
    data.Add "value", ""
    
    'loop through and send text chunks to target
    For i = 0 To UBound(chunk)
        If Left(chunk(i), 4) = "\ue0" Then
            ReDim chars(0 To 0)
            chars(0) = chunk(i)
        Else
            ReDim chars(0 To Len(chunk(i)) - 1)
            For j = 0 To Len(chunk(i)) - 1
                chars(j) = mid(chunk(i), j + 1, 1)
            Next j
        End If
        data("text") = chunk(i)
        data("value") = chars
        Execute CMD_SEND_KEYS_TO_ELEMENT, data
    Next i
    
End Sub

Public Sub Click(Optional element As WebElement)
    Dim data As New Dictionary
    If Not element Is Nothing Then
        data.Add "id", element.elementId
        Execute CMD_CLICK_ELEMENT, data
    Else
        Execute CMD_CLICK, data
    End If
End Sub

Public Sub Submit(element As WebElement)
    Me.Click element
End Sub

Public Function GetText(element As WebElement) As String
    Dim data As New Dictionary
    data.Add "id", element.elementId
    GetText = Execute(CMD_GET_ELEMENT_TEXT, data)("value")
End Function

Friend Function Execute(driverCommand, Optional parameters As Dictionary = Nothing, Optional ByVal raise As Boolean = True) As Dictionary
    Dim method As String, path As String
    Dim cmdArgs As New Dictionary, parmKey As Variant
    Dim response As Dictionary
    
    method = driverCommand(0)
    path = driverCommand(1)
    
    If parameters Is Nothing Then
        Set parameters = New Dictionary
    End If
    
    'Set session id
    If Not parameters.Exists("sessionId") Then
        parameters.Add "sessionId", sessionId
    End If
    
    For Each parmKey In parameters
        If InStr(path, "$" & parmKey) > 0 Then 'path parameter
            path = Replace(path, "$" & parmKey, parameters(parmKey))
        Else 'non-path parameter
            cmdArgs.Add parmKey, parameters(parmKey)
        End If
    Next parmKey
    
    'there are commands that don't require sessionId in path: CMD_STATUS, CMD_NEW_SESSION, CMD_GET_ALL_SESSIONS, CMD_SHUTDOWN
    If cmdArgs.Exists("sessionId") Then cmdArgs.Remove "sessionId"
    
    'Send request to selenium server
    Set response = SendRequest(method, urlBase + path, cmdArgs)
    
    If raise And IsResponseError(response) Then 'raise error if occurs
        Err.raise 513, "WebDriver.Execute", GetResponseErrorMessage(response)
    End If

    Set Execute = response 'always pass a dictionary object - let caller parse info based on context, including error
    
End Function

Private Function GetResponseErrorMessage(resp As Dictionary) As String
    Dim jc As New JSonConverter
    GetResponseErrorMessage = ""
    If TypeName(resp("value")) = "Dictionary" Then
        If resp("value").Exists("error") Then
            GetResponseErrorMessage = jc.ConvertToJson(resp("value"))
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

Private Function SendRequest(ByVal method As String, ByVal url As String, Optional data As Dictionary = Nothing) As Dictionary
    Dim client As MSXML2.ServerXMLHTTP60, jc As New JSonConverter
    Set client = New MSXML2.ServerXMLHTTP60

    client.Open method, url
    
    If method = "POST" Or method = "PUT" Then
        'client.setTimeouts
        'client.setRequestHeader "Content-Type", "application/json" 'original
        client.setRequestHeader "Content-Type", "application/json; charset=utf-8"
        client.setRequestHeader "Cache-Control", "no-cache"
        client.send jc.ConvertToJson(data)  '
    Else
        client.send
    End If

    Do While client.readyState < 4
        DoEvents
    Loop

    Dim json As Object
    Set json = jc.ParseJSON(client.responseText)
    Set SendRequest = json
    
    'Debug.Print jc.ConvertToJson(Json, 4) 'this will get browser info
End Function

'see https://github.com/SeleniumHQ/selenium/commit/7a6bf7e30d8bd834b6982b1d28f002bb4b26f380
Private Sub InitCommands()
    CMD_STATUS = Array("GET", "/status")
    CMD_NEW_SESSION = Array("POST", "/session")
    CMD_GET_ALL_SESSIONS = Array("GET", "/sessions")
    CMD_QUIT = Array("DELETE", "/session/$sessionId")
    CMD_W3C_GET_CURRENT_WINDOW_HANDLE = Array("GET", "/session/$sessionId/window")
    CMD_W3C_GET_WINDOW_HANDLES = Array("GET", "/session/$sessionId/window/handles")
    CMD_GET = Array("POST", "/session/$sessionId/url")
    CMD_GO_FORWARD = Array("POST", "/session/$sessionId/forward")
    CMD_GO_BACK = Array("POST", "/session/$sessionId/back")
    CMD_REFRESH = Array("POST", "/session/$sessionId/refresh")
    CMD_W3C_EXECUTE_SCRIPT = Array("POST", "/session/$sessionId/execute/sync")
    CMD_W3C_EXECUTE_SCRIPT_ASYNC = Array("POST", "/session/$sessionId/execute/async")
    CMD_GET_CURRENT_URL = Array("GET", "/session/$sessionId/url")
    CMD_GET_TITLE = Array("GET", "/session/$sessionId/title")
    CMD_GET_PAGE_SOURCE = Array("GET", "/session/$sessionId/source")
    CMD_SCREENSHOT = Array("GET", "/session/$sessionId/screenshot")
    CMD_ELEMENT_SCREENSHOT = Array("GET", "/session/$sessionId/element/$id/screenshot")
    CMD_FIND_ELEMENT = Array("POST", "/session/$sessionId/element")
    CMD_FIND_ELEMENTS = Array("POST", "/session/$sessionId/elements")
    CMD_W3C_GET_ACTIVE_ELEMENT = Array("GET", "/session/$sessionId/element/active")
    CMD_FIND_CHILD_ELEMENT = Array("POST", "/session/$sessionId/element/$id/element")
    CMD_FIND_CHILD_ELEMENTS = Array("POST", "/session/$sessionId/element/$id/elements")
    CMD_CLICK_ELEMENT = Array("POST", "/session/$sessionId/element/$id/click")
    CMD_CLEAR_ELEMENT = Array("POST", "/session/$sessionId/element/$id/clear")
    CMD_GET_ELEMENT_TEXT = Array("GET", "/session/$sessionId/element/$id/text")
    CMD_SEND_KEYS_TO_ELEMENT = Array("POST", "/session/$sessionId/element/$id/value")
    CMD_UPLOAD_FILE = Array("POST", "/session/$sessionId/file")
    CMD_GET_ELEMENT_TAG_NAME = Array("GET", "/session/$sessionId/element/$id/name")
    CMD_IS_ELEMENT_SELECTED = Array("GET", "/session/$sessionId/element/$id/selected")
    CMD_IS_ELEMENT_ENABLED = Array("GET", "/session/$sessionId/element/$id/enabled")
    CMD_IS_ELEMENT_DISPLAYED = Array("GET", "/session/$sessionId/element/$id/displayed")
    CMD_GET_ELEMENT_RECT = Array("GET", "/session/$sessionId/element/$id/rect")
    CMD_GET_ELEMENT_ATTRIBUTE = Array("GET", "/session/$sessionId/element/$id/attribute/$name")
    CMD_GET_ELEMENT_PROPERTY = Array("GET", "/session/$sessionId/element/$id/property/$name")
    CMD_GET_ALL_COOKIES = Array("GET", "/session/$sessionId/cookie")
    CMD_ADD_COOKIE = Array("POST", "/session/$sessionId/cookie")
    CMD_GET_COOKIE = Array("GET", "/session/$sessionId/cookie/$name")
    CMD_DELETE_ALL_COOKIES = Array("DELETE", "/session/$sessionId/cookie")
    CMD_DELETE_COOKIE = Array("DELETE", "/session/$sessionId/cookie/$name")
    CMD_SWITCH_TO_FRAME = Array("POST", "/session/$sessionId/frame")
    CMD_SWITCH_TO_PARENT_FRAME = Array("POST", "/session/$sessionId/frame/parent")
    CMD_SWITCH_TO_WINDOW = Array("POST", "/session/$sessionId/window")
    CMD_NEW_WINDOW = Array("POST", "/session/$sessionId/window/new")
    CMD_CLOSE = Array("DELETE", "/session/$sessionId/window")
    CMD_GET_ELEMENT_VALUE_OF_CSS_PROPERTY = Array("GET", "/session/$sessionId/element/$id/css/$name")
    CMD_SET_TIMEOUTS = Array("POST", "/session/$sessionId/timeouts")
    CMD_GET_TIMEOUTS = Array("GET", "/session/$sessionId/timeouts")
    CMD_W3C_DISMISS_ALERT = Array("POST", "/session/$sessionId/alert/dismiss")
    CMD_W3C_ACCEPT_ALERT = Array("POST", "/session/$sessionId/alert/accept")
    CMD_W3C_SET_ALERT_VALUE = Array("POST", "/session/$sessionId/alert/text")
    CMD_W3C_GET_ALERT_TEXT = Array("GET", "/session/$sessionId/alert/text")
    CMD_CLICK = Array("POST", "/session/$sessionId/click")
    'CMD_W3C_ACTIONS = Array("POST", "/session/$sessionId/actions") 'used in Action.cls
    'CMD_W3C_CLEAR_ACTIONS = Array("DELETE", "/session/$sessionId/actions") 'used in Action.cls
    CMD_SET_WINDOW_RECT = Array("POST", "/session/$sessionId/window/rect")
    CMD_GET_WINDOW_RECT = Array("GET", "/session/$sessionId/window/rect")
    CMD_W3C_MAXIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/maximize")
    CMD_GET_LOCATION = Array("GET", "/session/$sessionId/location")
    CMD_SET_LOCATION = Array("POST", "/session/$sessionId/location")
    CMD_GET_APP_CACHE_STATUS = Array("GET", "/session/$sessionId/application_cache/status")
    CMD_GET_NETWORK_CONNECTION = Array("GET", "/session/$sessionId/network_connection")
    CMD_SET_NETWORK_CONNECTION = Array("POST", "/session/$sessionId/network_connection")
    'CMD_GET_LOG = Array("POST", "/session/$sessionId/se/log")  'see logging option in Edge and Chrome methods
    'CMD_GET_AVAILABLE_LOG_TYPES = Array("GET", "/session/$sessionId/se/log/types") 'see logging option in Edge and Chrome methods
    CMD_CURRENT_CONTEXT_HANDLE = Array("GET", "/session/$sessionId/se/context")
    CMD_CONTEXT_HANDLES = Array("GET", "/session/$sessionId/se/contexts")
    CMD_SWITCH_TO_CONTEXT = Array("POST", "/session/$sessionId/context")
    CMD_FULLSCREEN_WINDOW = Array("POST", "/session/$sessionId/window/fullscreen")
    CMD_MINIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/minimize")
    CMD_SHUTDOWN = Array("GET", "/shutdown")
    CMD_PRINT_PAGE = Array("POST", "/session/$sessionId/print") 'Maybe only in headless mode
    CMD_GET_ELEMENT_SHADOW_ROOT = Array("GET", "/session/$sessionId/element/$id/shadow")
    CMD_FIND_ELEMENT_FROM_SHADOW_ROOT = Array("POST", "/session/$sessionId/shadow/$sid/element")
    CMD_FIND_ELEMENTS_FROM_SHADOW_ROOT = Array("POST", "/session/$sessionId/shadow/$sid/elements")
    CMD_GET_ELEMENT_ARIA_ROLE = Array("GET", "/session/$sessionId/element/$id/computedrole")
    CMD_GET_ELEMENT_ARIA_LABEL = Array("GET", "/session/$sessionId/element/$id/computedlabel")
End Sub

Private Function Params(ParamArray keysAndValues()) As Dictionary
    Dim dict As New Dictionary
    Dim i As Integer
    For i = 0 To UBound(keysAndValues) - 1 Step 2
        dict.Add keysAndValues(i), keysAndValues(i + 1)
    Next i
    Set Params = dict
End Function

Public Function Back()
    Dim data As New Dictionary
    Execute CMD_GO_BACK, data
End Function

Public Function Forward()
    Dim data As New Dictionary
    Execute CMD_GO_FORWARD, data
End Function

Public Function Refresh()
    Dim data As New Dictionary
    Execute CMD_REFRESH, data
End Function

Public Function GetCurrentURL() As String
    Dim data As New Dictionary
    Dim encodeURL As String
    encodeURL = Execute(CMD_GET_CURRENT_URL, data)("value")
    GetCurrentURL = ExecuteScript("return decodeURI('" & encodeURL & "')")
End Function

Public Function GetActiveElement() As WebElement
    Dim data As New Dictionary
   ' Return element
    Set GetActiveElement = ToWebElement(Execute(CMD_W3C_GET_ACTIVE_ELEMENT, data)("value")(ELEMENT_KEY))
End Function

Public Function GetProperty(ByVal name As String, element As WebElement) As String
    'This returns the current state of a DOM node object, such as the value of an input box.
    'If the property does not exist, then this returns vbNullString
    Dim data As New Dictionary, resp As Dictionary
    data.Add "name", name
    data.Add "id", element.elementId
    
    Set resp = Execute(CMD_GET_ELEMENT_PROPERTY, data)
    
    If IsNull(resp("value")) Then
        GetProperty = ""
    Else
        GetProperty = resp("value")
    End If
End Function

Public Function HasProperty(ByVal name As String, element As WebElement) As Boolean
    If GetProperty(name, element) = "" Then HasProperty = False Else HasProperty = True
End Function

Public Function HasAttribute(ByVal name As String, element As WebElement) As Boolean
    If GetAttribute(name, element) = "" Then HasAttribute = False Else HasAttribute = True
End Function

Public Function GetAttribute(ByVal name As String, element As WebElement) As String
    'This returns pre-defined HTML element attributes. Once the browser parses the html code,
    'then a DOM node object with properties is created.
    'To access the state of a DOM node property - such as an input box value, use GetProperty.
    'If the attribute does not exist, then this returns vbNullString
    Dim data As New Dictionary

    data.Add "id", element.elementId
    data.Add "name", name
    
    Set resp = Execute(CMD_GET_ELEMENT_ATTRIBUTE, data)
    
    If IsNull(resp("value")) Then
        GetAttribute = ""
    Else
        GetAttribute = resp("value")
    End If
End Function

Public Sub SetAttribute(element As WebElement, ByVal name As String, ByVal value As String)
    'this sets the value of an html element attribute. Consider SendKeys or Action Chain for more control of user input elements.
    'If the attribute does not exist, then this does not throw an error.
    ExecuteScript "arguments[0].setAttribute('" & name & "', arguments[1])", element, value
End Sub

Public Function GetAriaRole(element As WebElement) As String
    'This returns pre-defined HTML element attributes. Once the browser parses the html code,
    'then a DOM node object with properties is created.
    'To access the state of a DOM node property - such as an input box value, use GetProperty.
    'If the attribute does not exist, then this returns vbNullString
    Dim data As New Dictionary
    data.Add "id", element.elementId
    
    Set resp = Execute(CMD_GET_ELEMENT_ARIA_ROLE, data)
    
    If IsNull(resp("value")) Then
        GetAriaRole = ""
    Else
        GetAriaRole = resp("value")
    End If
End Function

Public Function GetAriaLabel(element As WebElement) As String
    'This returns pre-defined HTML element attributes. Once the browser parses the html code,
    'then a DOM node object with properties is created.
    'To access the state of a DOM node property - such as an input box value, use GetProperty.
    'If the attribute does not exist, then this returns vbNullString
    Dim data As New Dictionary
    data.Add "id", element.elementId
    'data.Add "name", name
    
    Set resp = Execute(CMD_GET_ELEMENT_ARIA_LABEL, data)
    
    If IsNull(resp("value")) Then
        GetAriaLabel = ""
    Else
        GetAriaLabel = resp("value")
    End If
End Function

Public Function GetTagName(element As WebElement) As String
    Dim data As New Dictionary
    data.Add "id", element.elementId
    GetTagName = Execute(CMD_GET_ELEMENT_TAG_NAME, data)("value")
End Function

Public Function SelectByValue(ByVal value As String, element As WebElement)
    Dim data As New Dictionary
    data.Add "using", "css selector"
    data.Add "value", "option[value =" & """" & value & """" & "]"
    data.Add "id", element.elementId
    Dim elmId As WebElement
    Set elmId = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, data)("value")(ELEMENT_KEY))
    If Not IsSelected(elmId) Then
        elmId.Click
    End If
End Function

Public Function SelectByVisibleText(ByVal text As String, element As WebElement)
    Dim data As New Dictionary
    data.Add "using", "xpath"
    data.Add "value", ".//option[normalize-space(.) = " & """" & text & """" & "]"
    data.Add "id", element.elementId
    Dim elmId As WebElement
    Set elmId = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, data)("value")(ELEMENT_KEY))
    If Not IsSelected(elmId) Then
        elmId.Click
    End If
End Function

Public Function SelectTagsOptions(element As WebElement) As String()
    Dim elmTags() As WebElement
    elmTags = FindElements(by.tagname, "option", element)
    Dim ret() As String
    Dim i As Integer
    ret = VBA.Split(VBA.vbNullString)
    For i = 0 To UBound(elmTags)
        ReDim Preserve ret(i)
        ret(i) = elmTags(i).GetText
    Next
    SelectTagsOptions = ret
End Function

Public Function AllSelectedOptions(element As WebElement) As String()
    Dim elmIds() As WebElement
    elmIds = FindChildElements(element)
    Dim ret() As String
    Dim i As Integer
    Dim j As Integer
    j = 0
    ret = VBA.Split(VBA.vbNullString)
    For i = 0 To UBound(elmIds)
        If IsSelected(elmIds(i)) Then
            ReDim Preserve ret(j)
            ret(j) = elmIds(i).GetText
            j = j + 1
        End If
    Next i
    AllSelectedOptions = ret
End Function

Public Function SelectedOptionText(element As WebElement) As String
    Dim elmIds() As WebElement
    elmIds = FindChildElements(element)
    Dim i As Integer
    For i = 0 To UBound(elmIds)
        If IsSelected(elmIds(i)) Then
            SelectedOptionText = elmIds(i).GetText
            Exit Function
        End If
    Next i
    SelectedOptionText = ""
End Function

Public Function DeSelectByValue(ByVal value As String, element As WebElement)
    Dim data As New Dictionary, elmId As WebElement
    data.Add "using", "css selector"
    data.Add "value", "option[value =" & """" & value & """" & "]"
    data.Add "id", element.elementId
    Set elmId = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, data)("value")(ELEMENT_KEY))
    If IsSelected(elmId) = "True" Then
        elmId.Click
    End If
End Function

Public Function DeSelectByVisibleText(ByVal text As String, element As WebElement)
    Dim data As New Dictionary, elmId As WebElement
    data.Add "using", "xpath"
    data.Add "value", ".//option[normalize-space(.) = " & """" & text & """" & "]"
    data.Add "id", element.elementId
    Set elmId = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, data)("value")(ELEMENT_KEY))
    If IsSelected(elmId) = "True" Then
        elmId.Click
    End If
End Function

Public Function FindChildElements(element As WebElement) As WebElement()
    Dim data As New Dictionary, elements, ret() As WebElement, i As Integer
    data.Add "using", "css selector"
    data.Add "value", "option"
    data.Add "id", element.elementId
    Set elements = Execute(CMD_FIND_CHILD_ELEMENTS, data)("value")

    For i = 0 To elements.Count - 1
        ReDim Preserve ret(i)
        Set ret(i) = ToWebElement(elements(i + 1)(ELEMENT_KEY))
    Next
    FindChildElements = ret
End Function

Public Function IsSelected(element As WebElement) As Boolean
    Dim data As New Dictionary
    data.Add "id", element.elementId
    IsSelected = Execute(CMD_IS_ELEMENT_SELECTED, data)("value")
End Function

Public Function IsMultiSelect(selectElement As WebElement) As Boolean
    IsMultiSelect = HasAttribute("multiple", selectElement)
End Function

Public Function SelectByIndex(ByVal index As Integer, element As WebElement)
    Me.Focus element
    ExecuteScript "arguments[0].options[" & CStr(index) & "].selected = true;", element
End Function

Public Function SelectAll(element As WebElement)
    Dim elmIds() As WebElement
    Dim i As Integer
    elmIds = FindChildElements(element)
    Me.Focus element
    For i = 0 To UBound(elmIds)
        ExecuteScript "arguments[0].options[" & CStr(i) & "].selected = true;", element
    Next i
End Function

Public Function DeSelectAll(element As WebElement)
    Dim elmIds() As WebElement
    Dim i As Integer
    elmIds = FindChildElements(element)
    Me.Focus element
    For i = 0 To UBound(elmIds)
        ExecuteScript "arguments[0].options[" & CStr(i) & "].selected = false;", element
    Next i
End Function

Public Function DeSelectByIndex(ByVal index As Integer, element As WebElement)
    Dim elmIds() As WebElement, args As Variant, script As String, optdata As New Dictionary
    elmIds = FindChildElements(element)
    script = "return arguments[0].selectedIndex = " & CStr(index)
    For i = 0 To UBound(elmIds)
        If i = index Then
            If IsSelected(elmIds(index)) = "True" Then
                ExecuteScript script, elmIds(index)
                elmIds(index).Click
                Exit For
            End If
        End If
    Next
End Function

Public Function ExecuteScript(ByVal script As String, ParamArray scriptArgs() As Variant)
    Dim data As New Dictionary, ElmData As New Dictionary
    Dim resp As Dictionary
    Dim args()
    
    data.Add "script", script
    
    If UBound(scriptArgs) >= 0 Then
        For i = 0 To UBound(scriptArgs)
            ReDim Preserve args(i)
            If TypeName(scriptArgs(i)) = "WebElement" Then
                'Convert WebElement to Dictionary that can be handled by WebDriver
                Dim elem As New Dictionary
                elem.Add ELEMENT_KEY, scriptArgs(i).elementId
                Set args(i) = elem
            Else
                args(i) = scriptArgs(i)
            End If
        Next i
    End If
    data.Add "args", args
    
    Set resp = Execute(CMD_W3C_EXECUTE_SCRIPT, data)
    
    Select Case TypeName(resp("value"))
    Case "Collection"
        ExecuteScript = ScriptCollectionToConvertedArray(resp("value"))
    Case "Dictionary"
        Set ExecuteScript = ScriptDictToConvertedDict(resp("value"))
    Case Else
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
    For Each key In dict
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
    Dim ary()
    ReDim ary(col.Count - 1)
    Dim i As Integer
    
    For i = 0 To UBound(ary)
        If TypeName(col(i + 1)) = "Collection" Then
            ary(i) = ScriptCollectionToConvertedArray(col(i + 1))
        ElseIf TypeName(col(i + 1)) = "Dictionary" Then
            Set ary(i) = ScriptDictToConvertedDict(col(i + 1))
        ElseIf IsObject(col(i + 1)) Then
            Set ary(i) = col(i + 1)
        Else
            ary(i) = col(i + 1)
        End If
    Next
    ScriptCollectionToConvertedArray = ary
End Function

Public Sub SaveScreenshot(ByVal filename As String, Optional element As WebElement)
    Dim data As New Dictionary, DomDoc As New MSXML2.DOMDocument60
    Dim png As String
        
    If element Is Nothing Then
        png = Execute(CMD_SCREENSHOT, data)("value")
    Else
        data.Add "id", element.elementId
        png = Execute(CMD_ELEMENT_SCREENSHOT, data)("value")
    End If
    
    ' Conversion from base64 string to Byte array
    Dim chars() As Byte
    With DomDoc.createElement("b64")
        .DataType = "bin.base64"
        .text = png
        chars = .nodeTypedValue
    End With
    
    ' png file put
    Dim fp As Long
    fp = FreeFile
    Open filename For Output As fp
    Close fp
    Open filename For Binary Access Write As fp
        Put #fp, 1, chars
    Close fp
End Sub

Public Function PageSource() As String
    Dim data As New Dictionary
    PageSource = Execute(CMD_GET_PAGE_SOURCE, data)("value")
End Function

Public Function SwitchToNextWindow()
    Dim handles() As String
    handles = GetWindowHandles
    SwitchToWindow handles(UBound(handles))
End Function

Public Function GetWindowHandles() As String()
    Dim data As New Dictionary
    Dim handles
    Set handles = Execute(CMD_W3C_GET_WINDOW_HANDLES, data)("value")
    ' To array of handles
    Dim ret() As String
    Dim i As Integer
    For i = 0 To handles.Count - 1  ' handles is Collection, not array
        ReDim Preserve ret(i)
        ret(i) = handles.Item(i + 1)
    Next
    ' Return handles
    GetWindowHandles = ret
End Function

Public Function SwitchToWindow(ByVal handle As String)
    Dim data As New Dictionary
    data.Add "handle", handle
    Execute CMD_SWITCH_TO_WINDOW, data
End Function

Public Function GetCurrentWindowHandle() As String
    Dim data As New Dictionary
    GetCurrentWindowHandle = Execute(CMD_W3C_GET_CURRENT_WINDOW_HANDLE, data)("value")
End Function

Public Sub CloseWindow()
    Dim data As New Dictionary
    Execute CMD_CLOSE, data
End Sub

Public Function Title() As String
    Dim data As New Dictionary
    Title = Execute(CMD_GET_TITLE, data)("value")
End Function

Public Function SwitchToFrame(element As WebElement)
    Dim data As New Dictionary
    Dim subdata As New Dictionary
    'according to W3C, can also take a null or "number object" - GCUser99
    subdata.Add "ELEMENT", element.elementId
    subdata.Add ELEMENT_KEY, element.elementId
    data.Add "id", subdata
    Execute CMD_SWITCH_TO_FRAME, data
End Function

Public Function SwitchToParentFrame()
    Dim data As New Dictionary
    Execute CMD_SWITCH_TO_PARENT_FRAME, data
End Function

Public Function GetCurrentFrame() As String
    Dim data As New Dictionary
    data.Add "script", "return self.name"
    Dim args As Variant
    args = Array()
    data.Add "args", args
    GetCurrentFrame = Execute(CMD_W3C_EXECUTE_SCRIPT, data)("value")
End Function

Public Function SetImplicitlyWait(Optional ByVal millisecond As Long = 0)
    Dim data As New Dictionary
    data.Add "implicit", millisecond
    Execute CMD_SET_TIMEOUTS, data
End Function

Public Function SetPageLoadTimeout(Optional ByVal millisecond As Long = 300000)
    Dim data As New Dictionary
    data.Add "pageLoad", millisecond
    Execute CMD_SET_TIMEOUTS, data
End Function

Public Function SetScriptTimeout(Optional ByVal millisecond As Long = 30000)
    Dim data As New Dictionary
    data.Add "script", millisecond
    Execute CMD_SET_TIMEOUTS, data
End Function

Public Function GetImplicitlyWait() As Double
    Dim data As New Dictionary
    Dim results
    Set results = Execute(CMD_GET_TIMEOUTS, data)("value")
    GetImplicitlyWait = results.Item("implicit")
End Function

Public Function GetPageLoadTimeout() As Double
    Dim data As New Dictionary
    Dim results
    Set results = Execute(CMD_GET_TIMEOUTS, data)("value")
    GetPageLoadTimeout = results.Item("pageLoad")
    'Dim jc As New JsonConverter
End Function

Public Function GetScriptTimeout() As Double
    Dim data As New Dictionary
    Dim results
    Set results = Execute(CMD_GET_TIMEOUTS, data)("value")
    GetScriptTimeout = results.Item("script")
End Function

Public Function Wait(Optional ByVal millisecond As Long = 1000)
    Sleep millisecond
End Function

Public Function IsPresent(by_ As by, ByVal value As String) As Boolean
    Dim timeout As Double, resp As Dictionary, data As Dictionary
    
    timeout = GetImplicitlyWait()
    If timeout <> 0 Then Me.SetImplicitlyWait 0
    Set data = ToSelector(by_, value)
    Set resp = Execute(CMD_FIND_ELEMENT, data, False) 'don't raise an error if not present
    If IsResponseError(resp) Then
        IsPresent = False
    Else
        IsPresent = True
    End If
    If timeout <> 0 Then Me.SetImplicitlyWait timeout

End Function

Public Function IsEnabled(element As WebElement) As Boolean
    Dim data As New Dictionary
    data.Add "id", element.elementId
    IsEnabled = Execute(CMD_IS_ELEMENT_ENABLED, data)("value")
End Function

Public Function IsDisplayed(element As WebElement) As Boolean
    Dim data As New Dictionary
    data.Add "id", element.elementId
    IsDisplayed = Execute(CMD_IS_ELEMENT_DISPLAYED, data)("value")
End Function

Public Function GetAlertText() As String
    Dim data As New Dictionary
    GetAlertText = Execute(CMD_W3C_GET_ALERT_TEXT, data)("value")
End Function

Public Function SetAlertText(ByVal alertText As String)
    Dim data As New Dictionary
    data.Add "text", alertText
    Execute CMD_W3C_SET_ALERT_VALUE, data
End Function

Public Function AcceptAlert()
    Dim data As New Dictionary
    Execute CMD_W3C_ACCEPT_ALERT, data
End Function

Public Function DismissAlert()
    Dim data As New Dictionary
    Execute CMD_W3C_DISMISS_ALERT, data
End Function

Public Function IsAlertPresent() As Boolean
    Dim data As New Dictionary, resp As Dictionary
    Set resp = Execute(CMD_W3C_GET_ALERT_TEXT, data, False)
    If IsResponseError(resp) Then
        IsAlertPresent = False
    Else
        IsAlertPresent = True
    End If
End Function

Public Function GetWindowRect() As Integer()
    Dim data As New Dictionary, rects, ret() As Integer
    Set rects = Execute(CMD_GET_WINDOW_RECT, data)("value")
    ReDim Preserve ret(rects.Count - 1)
    ret(0) = CInt(rects.Item("x"))
    ret(1) = CInt(rects.Item("y"))
    ret(2) = CInt(rects.Item("width"))
    ret(3) = CInt(rects.Item("height"))
    GetWindowRect = ret
End Function

Public Function SetWindowRect(Optional ByVal x As Integer = 0, Optional ByVal y As Integer = 0, Optional ByVal width As Integer = 0, Optional ByVal height As Integer = 0) As Integer()
    Dim data As New Dictionary
    data.Add "x", x
    data.Add "y", y
    data.Add "width", width
    data.Add "height", height
    Dim rects
    Set rects = Execute(CMD_SET_WINDOW_RECT, data)("value")
    Dim ret() As Integer
    ReDim Preserve ret(rects.Count - 1)
    ret(0) = CInt(rects.Item("x"))
    ret(1) = CInt(rects.Item("y"))
    ret(2) = CInt(rects.Item("width"))
    ret(3) = CInt(rects.Item("height"))
    SetWindowRect = ret
End Function

Public Function GetWindowSize() As Integer()
    Dim rect() As Integer
    rect = GetWindowRect()
    Dim ret(1) As Integer
    ret(0) = rect(2)    'width
    ret(1) = rect(3)    'height
    ' Return
    GetWindowSize = ret
End Function

Public Function SetWindowSize(ByVal width As Variant, ByVal height As Variant)
    Dim rect() As Integer
    rect = GetWindowRect
    rect = SetWindowRect(rect(0), rect(1), width, height)
End Function

Public Function GetWindowPosition() As Integer()
    Dim rect() As Integer
    rect = GetWindowRect()
    Dim ret(1) As Integer
    ret(0) = rect(0)    'x
    ret(1) = rect(1)    'y
    ' Return
    GetWindowPosition = ret
End Function

Public Function SetWindowPosition(ByVal x As Variant, ByVal y As Variant)
    Dim rect() As Integer
    rect = GetWindowRect
    rect = SetWindowRect(x, y, rect(2), rect(3))
End Function

Public Function GetScreenSize() As Integer()
    Dim ret(1) As Integer
    ret(0) = GetSystemMetrics(SM_CXSCREEN)  'x
    ret(1) = GetSystemMetrics(SM_CYSCREEN)  'y
    GetScreenSize = ret
End Function

Public Sub MaximizeWindow()
    Dim data As New Dictionary
    Execute CMD_W3C_MAXIMIZE_WINDOW, data
End Sub

Public Sub FullScreenWindow()
    Dim data As New Dictionary
    Execute CMD_FULLSCREEN_WINDOW, data
End Sub

Public Sub MinimizeWindow()
    Dim data As New Dictionary
    Execute CMD_MINIMIZE_WINDOW, data
End Sub

Public Sub Clear(element As WebElement)
    Dim data As New Dictionary
    data.Add "id", element.elementId
    data.Add "text", vbNullString
    data.Add "value", vbNullString
    Execute CMD_CLEAR_ELEMENT, data
End Sub

Public Function GetShadowRoot(element As WebElement) As ShadowRoot
    Dim data As New Dictionary, resp As Dictionary
    data.Add "id", element.elementId
    Set resp = Execute(CMD_GET_ELEMENT_SHADOW_ROOT, data)
    Set GetShadowRoot = ToShadowRoot(resp("value")(SHADOW_KEY))
End Function

Public Function GetCSSProperty(ByVal name As String, element As WebElement) As String
    Dim data As New Dictionary
    data.Add "id", element.elementId
    data.Add "name", name
    GetCSSProperty = Execute(CMD_GET_ELEMENT_VALUE_OF_CSS_PROPERTY, data)("value")
End Function

Public Function GetRect(element As WebElement) As Dictionary
    Dim data As New Dictionary, rects, ret() As Integer
    data.Add "id", element.elementId
    Set rects = Execute(CMD_GET_ELEMENT_RECT, data)("value")
    Set GetRect = rects
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

Public Function GetHTMLDoc() As HTMLDocument
    Dim htmldoc As New HTMLDocument
    htmldoc.Body.innerHTML = Me.GetInnerHTML()
    Set GetHTMLDoc = htmldoc
End Function

' Get default capabilities for current browser
Public Function CreateCapabilities() As Capabilities
    Dim Capabilities As New Capabilities
    Capabilities.InitializeFor browserName
    Set CreateCapabilities = Capabilities
End Function

Public Function Focus(element As WebElement)
    Dim script As String
    script = "arguments[0].focus({'preventScroll': arguments[1]})"
    ExecuteScript script, element, 0  'Scroll
End Function

Public Function FocusNoScroll(element As WebElement)
    Dim script As String
    script = "arguments[0].focus({'preventScroll': arguments[1]})"
    ExecuteScript script, element, 1  'No Scroll
End Function

Public Function ScrollToElement(element As WebElement, Optional ByVal xOffset As Integer = 0, Optional ByVal yOffset As Integer = 0)
    ScrollIntoView element, True
    ScrollBy xOffset, yOffset
End Function

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

Public Function GetScrollHeight()
    Dim script As String
    script = "return document.body.scrollHeight;"
    GetScrollHeight = ExecuteScript(script)
End Function

Public Function GetAllCookies() As Cookie()
    Dim data As New Dictionary, resp As Collection, ck As Cookie, cks() As Cookie
    Set resp = Execute(CMD_GET_ALL_COOKIES, data)("value")
    'Dim jc As New JSonConverter
    'Debug.Print jc.ConvertToJson(resp, 4)
    ReDim cks(0 To resp.Count - 1)
    For i = 1 To resp.Count
        Set ck = New Cookie
        If resp(i).Exists("name") Then ck.name = resp(i)("name")
        If resp(i).Exists("domain") Then ck.Domain = resp(i)("domain")
        If resp(i).Exists("expiry") Then ck.ExpiryUnix = resp(i)("expiry")
        If resp(i).Exists("httpOnly") Then ck.HttpOnly = resp(i)("httpOnly")
        If resp(i).Exists("sameSite") Then ck.SameSite = resp(i)("sameSite")
        If resp(i).Exists("secure") Then ck.Secure = resp(i)("secure")
        If resp(i).Exists("value") Then ck.value = resp(i)("value")
        If resp(i).Exists("path") Then ck.path = resp(i)("path")
        Set cks(i - 1) = ck
    Next i
    GetAllCookies = cks
End Function

Public Function GetCookie(ByVal cookieName) As Cookie
    Dim data As New Dictionary, resp As Dictionary, ck As New Cookie
    data.Add "name", cookieName
    Set resp = Execute(CMD_GET_COOKIE, data)("value")
    If resp.Exists("name") Then ck.name = resp("name")
    If resp.Exists("domain") Then ck.Domain = resp("domain")
    If resp.Exists("expiry") Then ck.ExpiryUnix = resp("expiry")
    If resp.Exists("httpOnly") Then ck.HttpOnly = resp("httpOnly")
    If resp.Exists("sameSite") Then ck.SameSite = resp("sameSite")
    If resp.Exists("secure") Then ck.Secure = resp("secure")
    If resp.Exists("value") Then ck.value = resp("value")
    If resp.Exists("path") Then ck.path = resp("path")
    Set GetCookie = ck
End Function

Public Sub DeleteCookie(ByVal cookieName As String)
    Dim data As New Dictionary
    data.Add "name", cookieName
    Execute CMD_DELETE_COOKIE, data
End Sub

Public Sub SetCookie(cookieToAdd As Cookie)
    Dim data As New Dictionary, ck As New Dictionary
    'value can take a number string greater than 15 digits
    'need to set the following in JSonConverter else  will get an error on execute (see sendRequest):
    'jc.OptionsUseDoubleForLargeNumbers = True
    If cookieToAdd.name <> "" And cookieToAdd.value <> "" Then
        data.Add "cookie", cookieToAdd.ToDictionary
        Execute CMD_ADD_COOKIE, data
    End If
End Sub

Public Sub SetCookies(cookiesToAdd() As Cookie)
    For i = 0 To UBound(cookiesToAdd)
        Me.SetCookie cookiesToAdd(i)
    Next i
End Sub

Public Sub DeleteAllCookies()
    Dim data As New Dictionary
    Execute CMD_DELETE_ALL_COOKIES, data
End Sub

Public Sub SaveCookiesToFile(ck() As Cookie, Optional ByVal path As String = ".\cookies.txt")
    Dim fs As New FileSystemObject, ts As textstream, darray() As Dictionary
    Dim jc As New JSonConverter
    
    path = GetAbsolutePath(path)
    
    Set ts = fs.CreateTextFile(path, True)
    ReDim darray(0 To UBound(ck))
    For i = 0 To UBound(ck)
        Set darray(i) = ck(i).ToDictionary
    Next i
    ts.Write jc.ConvertToJson(darray(), 4)
    ts.Close
End Sub

Public Function LoadCookiesFromFile(Optional ByVal path As String = ".\cookies.txt") As Cookie()
    Dim fs As New FileSystemObject, ts As textstream, str As String, v As Collection, cks() As Cookie, ck As Cookie
    Dim jc As New JSonConverter
    
    path = GetAbsolutePath(path)
    
    Set ts = fs.OpenTextFile(path)
    str = ts.ReadAll
    ts.Close
    Set v = jc.ParseJSON(str)
    
    ReDim cks(0 To v.Count - 1)
    
    For i = 1 To v.Count
        Set ck = New Cookie
        ck.name = v(i)("name")
        ck.Domain = v(i)("domain")
        ck.ExpiryUnix = v(i)("expiry")
        ck.HttpOnly = v(i)("httpOnly")
        ck.SameSite = v(i)("sameSite")
        ck.Secure = v(i)("secure")
        ck.value = v(i)("value")
        ck.path = v(i)("path")
        Set cks(i - 1) = ck
    Next i
    LoadCookiesFromFile = cks

End Function

Public Function PageFound(ByVal url As String) As Boolean
    If Me.GetCurrentURL() = url Then
        PageFound = True
    Else
        PageFound = False
    End If
End Function

Private Function GetAbsolutePath(ByVal strPath As String) As String
    Dim fso As New FileSystemObject
    saveppath = CurDir()
    ChDrive ThisWorkbook.path
    ChDir ThisWorkbook.path
    GetAbsolutePath = fso.GetAbsolutePathName(strPath)
    ChDrive saveppath
    ChDir saveppath
End Function

Public Function GetSessionsInfo() As Collection 'add optional sessionID to get a particular one
    Set GetSessionsInfo = Me.Execute(CMD_GET_ALL_SESSIONS)("value")
End Function

Public Function GetDriverStatus() As Dictionary
    'Dim jc As New JSonConverter
    Set GetDriverStatus = Me.Execute(CMD_STATUS)("value")
    'DriverVersion = resp("build")("version")
    'DriverVersion = Trim(Left(DriverVersion, InStr(DriverVersion, "(") - 1))
    'resp("message") 'gets info "MSEdgeDriver ready for new sessions."
    'resp("os")("name") 'gets info "Windows NT"
    'resp("os")("version") 'gets info "10.0.19042"
    'resp("os")("arch") 'gets info "x86_64"
    'resp("ready")'true or false
End Function

Public Function ActionChain() As ActionChain
    Dim data_ As New Dictionary
    Set ActionChain = New ActionChain
    'pass a reference of the driver to each link
    Set ActionChain.WebDriver = Me
    Set ActionChain.data = data_
End Function

Public Function FindElementFromShadowRoot(by_ As by, ByVal value As String, sRoot As ShadowRoot) As WebElement
    Dim data As Dictionary
    Set data = ToSelector(by_, value)
    data.Add "sid", sRoot.shadowRootId
    Set FindElementFromShadowRoot = ToWebElement(Execute(CMD_FIND_ELEMENT_FROM_SHADOW_ROOT, data)("value")(ELEMENT_KEY))
End Function

Public Function FindElementsFromShadowRoot(by_ As by, ByVal value As String, sRoot As ShadowRoot) As WebElement()
    Dim data As Dictionary, elements
    Set data = ToSelector(by_, value)
    data.Add "sid", sRoot.shadowRootId
    Set elements = Execute(CMD_FIND_ELEMENT_FROM_SHADOW_ROOT, data)("value")
    ' To array of ids
    Dim ret() As WebElement
    Dim i As Integer
    For i = 0 To elements.Count - 1  ' elements is Collection, not array
        ReDim Preserve ret(i)
        Set ret(i) = ToWebElement(elements(i + 1)(ELEMENT_KEY))
    Next
    ' Return element ids
    FindElementsFromShadowRoot = ret
End Function

Public Sub DragAndDrop(source As WebElement, target As WebElement)
    Me.ActionChain().DragAndDrop(source, target).Perform
End Sub

Private Sub Class_Initialize()
    InitCommands
    AppWinStyle = vbHide 'hides the command window
End Sub

Public Sub SaveHTMLToFile(ByVal snippet As String, Optional ByVal path As String = ".\snippet.html")
    'this is used to save test snippets to file for subsequent browser navigation
    Dim fs As New FileSystemObject, ts As textstream
    path = GetAbsolutePath(path)
    Set ts = fs.CreateTextFile(path, True)
    ts.Write snippet
    ts.Close
End Sub
