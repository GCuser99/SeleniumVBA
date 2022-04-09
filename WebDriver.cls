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
' SeleniumVBA v0.0.9
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
'Microsoft Script Control 1.0 for 'CreateObject("ScriptControl") MSScriptControl.ScriptControl
'Microsoft XML, v6.0  for MSXML2.ServerXMLHTTP60, MSXML2.DOMDocument60 'CreateObject("MSXML2.DOMDocument")
'Microsoft Scripting Runtime for Scripting.FileSystemObject,Scripting.Dictionary  CreateObject("Scripting.Dictionary")
'Microsoft HTML Object Library

'change to True if auto-check-and-update desired
Private Const checkDriverBrowserVersionAlignment = False

'desired lower bound of returned arrays
Private Const baseArrayIdx = 1

'change baseDomIdx to 0 if zero-based Dom (de)selection is desired
'with baseDomIdx=1: SelectByIndex 1 gets first element
'with baseDomIdx=0: SelectByIndex 0 gets first element
Private Const baseDomIdx = 1

Private sessionId_ As String
Private driverUrl_ As String
Private browserName As String
Private AppWinStyle As VbAppWinStyle

Private Const ELEMENT_KEY = "element-6066-11e4-a52e-4f735466cecf"
Private Const SHADOW_KEY = "shadow-6066-11e4-a52e-4f735466cecf"

Private CMD_STATUS
Private CMD_NEW_SESSION
Private CMD_GET_ALL_SESSIONS
Private CMD_QUIT
Private CMD_GET_CURRENT_WINDOW_HANDLE
Private CMD_GET_WINDOW_HANDLES
Private CMD_GET
Private CMD_GO_FORWARD
Private CMD_GO_BACK
Private CMD_REFRESH
Private CMD_EXECUTE_SCRIPT
Private CMD_EXECUTE_SCRIPT_ASYNC
Private CMD_GET_CURRENT_URL
Private CMD_GET_TITLE
Private CMD_GET_PAGE_SOURCE
Private CMD_SCREENSHOT
Private CMD_ELEMENT_SCREENSHOT
Private CMD_FIND_ELEMENT
Private CMD_FIND_ELEMENTS
Private CMD_GET_ACTIVE_ELEMENT
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
Private CMD_DISMISS_ALERT
Private CMD_ACCEPT_ALERT
Private CMD_SET_ALERT_VALUE
Private CMD_GET_ALERT_TEXT
Private CMD_CLICK
'Private CMD_ACTIONS                              'used in ActionChain.cls
'Private CMD_CLEAR_ACTIONS                        'used in ActionChain.cls
Private CMD_SET_WINDOW_RECT
Private CMD_GET_WINDOW_RECT
Private CMD_MAXIMIZE_WINDOW
Private CMD_GET_LOCATION                             'not wrapped yet
Private CMD_SET_LOCATION                             'not wrapped yet
Private CMD_GET_APP_CACHE_STATUS                     'not wrapped yet
Private CMD_GET_NETWORK_CONNECTION                   'not wrapped yet
Private CMD_SET_NETWORK_CONNECTION                   'not wrapped yet
'Private CMD_GET_LOG                                  'see logging option in StartEdge and StartChrome methods
'Private CMD_GET_AVAILABLE_LOG_TYPES                  'see logging option in StartEdge and StartChrome methods
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
    tagName = 1
    className = 2
    name = 3
    cssSelector = 4
    XPath = 5
    linkText = 6
    PartialLinkText = 7
End Enum

Public Enum svbaWindowType
    svbaWindow = 0
    svbaTab = 1
End Enum

'#If VBA7 And Win64 Then
'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'#Else
'Private Declare Sub Sleep Lib "kernel32" (ByVal millisecond As Long)
'#End If

Public Sub StartEdge(Optional ByVal driverPath As String = "msedgedriver.exe", Optional ByVal localHostPort As Long = 9515, Optional ByVal enableLogging As Boolean = False, Optional ByVal logFilePath As String = ".\edge_verbose.log")
    Dim fso As New Scripting.FileSystemObject
    
    browserName = "msedge"
    driverPath = GetAbsolutePath(driverPath) 'support for relative path such as driverPath=".\msedgedriver.exe"
    
    If enableLogging Then
        logFilePath = GetAbsolutePath(logFilePath)
        serviceArgs = " --verbose --log-path=" & Chr(34) & logFilePath & Chr(34)
    End If
    
    If checkDriverBrowserVersionAlignment Then
        If Not AlignDriverAndBrowser(browserName, driverPath) Then End 'execution
    End If
    
    If Not fso.FileExists(driverPath) Then
        MsgBox "Could not find WebDriver at :" & vbCrLf & vbCrLf & driverPath & vbCrLf & vbCrLf & "Either manually install the driver or let WebDriverManager class install" & vbCrLf & "for you. To automatically install change: " & vbCrLf & vbCrLf & "CheckDriverBrowserVersionAlignment=True" & vbCrLf & vbCrLf & "in WebDriver class and then save Excel file.", , "SeleniumVBA"
        End 'execution
    End If
    
    Start driverPath, localHostPort, serviceArgs
    
End Sub

Public Sub StartChrome(Optional ByVal driverPath As String = "chromedriver.exe", Optional ByVal localHostPort As Long = 9515, Optional ByVal enableLogging As Boolean = False, Optional ByVal logFilePath As String = ".\chrome_verbose.log")
    Dim fso As New Scripting.FileSystemObject
    
    browserName = "chrome"
    driverPath = GetAbsolutePath(driverPath) 'support for relative path such as driverPath=".\chromedriver.exe"
    
    If enableLogging Then
        logFilePath = GetAbsolutePath(logFilePath)
        serviceArgs = " --verbose --log-path=" & Chr(34) & logFilePath & Chr(34)
    End If
    
    If checkDriverBrowserVersionAlignment Then
        If Not AlignDriverAndBrowser(browserName, driverPath) Then End 'execution
    End If
    
    If Not fso.FileExists(driverPath) Then
        MsgBox "Could not find WebDriver at :" & vbCrLf & vbCrLf & driverPath & vbCrLf & vbCrLf & "Either manually install the driver or let WebDriverManager class install" & vbCrLf & "for you. To automatically install change: " & vbCrLf & vbCrLf & "CheckDriverBrowserVersionAlignment=True" & vbCrLf & vbCrLf & "in WebDriver class and then save Excel file.", , "SeleniumVBA"
        End 'execution
    End If
    
    Start driverPath, localHostPort, serviceArgs

End Sub

Private Function AlignDriverAndBrowser(ByVal browserName As String, ByVal driverPath As String) As Boolean
    Dim wdmgr As New WebDriverManager
    
    driverPath = GetAbsolutePath(driverPath) 'support for relative path such as driverPath=".\msedgedriver.exe"
    
    AlignDriverAndBrowser = True
    
    If Not wdmgr.IsInstalledDriverCompatible(browserName, , driverPath) Then
        resp = MsgBox("WebDriver is not compatible with installed browser - would you like to install the compatible WebDriver?", vbYesNo)
        If resp = vbYes Then
            bverInstalled = wdmgr.GetInstalledBrowserVersion(browserName)
            dverCompat = wdmgr.GetCompatibleDriverVersion(browserName, bverInstalled)
            wdmgr.DownloadAndInstall browserName, dverCompat, driverPath
        Else
            AlignDriverAndBrowser = False
        End If
    End If

End Function

Private Sub Start(ByVal driverPath As String, ByVal localHostPort As Long, Optional ByVal serviceArgs As String = vbNullString)
    Dim port As String
    
    'make sure driver path is "quoted" in case we have spaces in path
    If Left(driverPath, 1) <> Chr(34) Then driverPath = Chr(34) & driverPath
    If Right(driverPath, 1) <> Chr(34) Then driverPath = driverPath & Chr(34)
    
    port = " --port=" & localHostPort

    If serviceArgs <> "" Then
        If Left(serviceArgs, 1) <> " " Then serviceArgs = " " & serviceArgs
    End If
    
    ' Start WebDriver executable
    If shell(driverPath & serviceArgs & port, AppWinStyle) = 0 Then
        MsgBox "Failed in starting WebDriver." & vbCrLf & _
               "WebDriverPath : " & driverPath & port, vbCritical + vbOKOnly
        End
    End If
    
    'Set WebDriver url
    driverUrl_ = "http://localhost:" & localHostPort
    
End Sub

Public Sub Shutdown()
    'this will shut down all WebDriver instances, not just the one it is being called from
    Execute CMD_SHUTDOWN
End Sub

Public Sub OpenBrowser(Optional caps As Capabilities, Optional ByVal invisible As Boolean = False)
    Dim resp As Dictionary
    
    If sessionId_ <> "" Then
        MsgBox "Only one OpenBrowser command per WebDriver instance is allowed! Use CloseBrowser before another OpenBrowser to fix this.", , "SeleniumVBA"
        End
    End If
    
    If caps Is Nothing Then
        Set caps = CreateCapabilities
    End If
        
    If invisible Then
        'user may have already added headless, but selenium doesn't seem to care about multiple arg entries
        caps.AddArgument "--headless"
    End If
    
    Set resp = Execute(CMD_NEW_SESSION, Params("capabilities", caps.Dictionary))("value")
    
    sessionId_ = resp("sessionId")
    
End Sub

Public Sub CloseBrowser()
    Dim data As New Dictionary
    Execute CMD_QUIT, data
    sessionId_ = ""
End Sub

Public Function NavigateTo(ByVal url As String, Optional ByVal timeOutms, Optional ByVal raise As Boolean = True) As Boolean
    Dim data As New Dictionary, resp As Dictionary, savtimeOutms As Integer
    
    If Left(url, 8) = "file:///" Then
        url = "file:///" & GetAbsolutePath(Mid(url, 9, 99999))
    End If
    
    data.Add "url", url
    If Not IsMissing(timeOutms) Then
        savtimeOutms = GetPageLoadTimeout()
        If savtimeOutms <> timeOutms Then SetPageLoadTimeout timeOutms
    End If
    
    Set resp = Execute(CMD_GET, data, raise)

    If Not IsMissing(timeOutms) Then
        If savtimeOutms <> timeOutms Then SetPageLoadTimeout savtimeOutms
    End If
    
    If IsResponseError(resp) Then
        NavigateTo = False
        Debug.Print GetResponseErrorMessage(resp)
    Else
        NavigateTo = True
    End If
    
End Function

Public Function FindElement(by_ As by, ByVal Value As String, Optional parentElement As WebElement) As WebElement
    Dim data As Dictionary
    Dim cmd As Variant
    Set data = ToSelector(by_, Value)
    If Not parentElement Is Nothing Then
        data.Add "id", parentElement.elementId
        cmd = CMD_FIND_CHILD_ELEMENT
    Else
        cmd = CMD_FIND_ELEMENT
    End If
    
    Set FindElement = ToWebElement(Execute(cmd, data)("value")(ELEMENT_KEY))
End Function

Public Function FindElements(by_ As by, ByVal Value As String, Optional parentElement As WebElement) As WebElements
    Dim data As Dictionary
    Dim i As Integer
    Dim cmd As Variant
    Dim elements As Collection
    Dim elems As New WebElements
    
    Set data = ToSelector(by_, Value)
    
    If Not parentElement Is Nothing Then
        data.Add "id", parentElement.elementId
        cmd = CMD_FIND_CHILD_ELEMENTS
    Else
        cmd = CMD_FIND_ELEMENTS
    End If
    
    Set elements = Execute(cmd, data)("value")
    
    For i = 1 To elements.Count
        elems.Add ToWebElement(elements(i)(ELEMENT_KEY))
    Next
    
    Set FindElements = elems
End Function

Public Function FindChildElements(element As WebElement) As WebElements
    Dim data As New Dictionary, elements As Collection
    Dim elems As New WebElements
    Dim i As Integer
    data.Add "using", "css selector"
    data.Add "value", "option"
    data.Add "id", element.elementId
    Set elements = Execute(CMD_FIND_CHILD_ELEMENTS, data)("value") 'returns 1-based Collection
    For i = 1 To elements.Count
        elems.Add ToWebElement(elements(i)(ELEMENT_KEY))
    Next
    Set FindChildElements = elems
End Function

Private Function ToSelector(by_ As by, ByVal Value As String) As Dictionary
    Dim data As New Dictionary
    Select Case by_
    Case by.ID
        data.Add "using", "css selector"
        data.Add "value", "[id=""" + Value + """]"
    Case by.tagName
        data.Add "using", "css selector"
        data.Add "value", Value
    Case by.className
        data.Add "using", "css selector"
        data.Add "value", "." + Value
    Case by.name
        data.Add "using", "css selector"
        data.Add "value", "[name=""" + Value + """]"
    Case by.cssSelector
        data.Add "using", "css selector"
        data.Add "value", Value
    Case by.XPath
        data.Add "using", "xpath"
        data.Add "value", Value
    Case by.linkText
        data.Add "using", "link text"
        data.Add "value", Value
    Case by.PartialLinkText
        data.Add "using", "partial link text"
        data.Add "value", Value
    Case Else
        data.Add "using", "css selector"
        data.Add "value", Value
    End Select
    
    Set ToSelector = data
End Function

Public Function FindElementByID(ByVal Value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByID = FindElement(by.ID, Value, parentElement)
End Function

Public Function FindElementByTagName(ByVal Value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByTagName = FindElement(by.tagName, Value, parentElement)
End Function

Public Function FindElementByClassName(ByVal Value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByClassName = FindElement(by.className, Value, parentElement)
End Function

Public Function FindElementByCssSelector(ByVal Value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByCssSelector = FindElement(by.cssSelector, Value, parentElement)
End Function

Public Function FindElementByLinkText(ByVal Value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByLinkText = FindElement(by.linkText, Value, parentElement)
End Function

Public Function FindElementByName(ByVal Value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByName = FindElement(by.name, Value, parentElement)
End Function

Public Function FindElementByPartialLinkText(ByVal Value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByPartialLinkText = FindElement(by.PartialLinkText, Value, parentElement)
End Function

Public Function FindElementByXPath(ByVal Value As String, Optional parentElement As WebElement) As WebElement
    Set FindElementByXPath = FindElement(by.XPath, Value, parentElement)
End Function

Public Function FindElementsByXPath(ByVal Value As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByXPath = FindElements(by.XPath, Value, parentElement)
End Function

Public Function FindElementsByID(ByVal Value As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByID = FindElements(by.ID, Value, parentElement)
End Function

Public Function FindElementsByTagName(ByVal Value As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByTagName = FindElements(by.tagName, Value, parentElement)
End Function

Public Function FindElementsByClassName(ByVal Value As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByClassName = FindElements(by.className, Value, parentElement)
End Function

Public Function FindElementsByCssSelector(ByVal Value As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByCssSelector = FindElements(by.cssSelector, Value, parentElement)
End Function

Public Function FindElementsByLinkText(ByVal Value As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByLinkText = FindElements(by.linkText, Value, parentElement)
End Function

Public Function FindElementsByName(ByVal Value As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByName = FindElements(by.name, Value, parentElement)
End Function

Public Function FindElementsByPartialLinkText(ByVal Value As String, Optional parentElement As WebElement) As WebElements
    Set FindElementsByPartialLinkText = FindElements(by.PartialLinkText, Value, parentElement)
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

Public Sub SendKeys(element As WebElement, ByVal keys As String, Optional clearBeforeTyping As Boolean = True)
    Dim data As New Dictionary
    Dim chunk() As String
    Dim chars() As String
    
    data.Add "id", element.elementId
    If clearBeforeTyping = True Then Execute CMD_CLEAR_ELEMENT, data
    
    'split key sequence into chunks with special keys (pattern "\ue0*") by themselves
    ReDim chunk(0 To Len(keys) - 1)
    i = -1
    Do Until keys = ""
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
                chars(j) = Mid(chunk(i), j + 1, 1)
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
    Click element
End Sub

Public Function GetText(element As WebElement) As String
    Dim data As New Dictionary
    data.Add "id", element.elementId
    GetText = Execute(CMD_GET_ELEMENT_TEXT, data)("value")
End Function

Friend Function Execute(driverCommand, Optional parameters As Dictionary = Nothing, Optional ByVal raise As Boolean = True) As Dictionary
    Dim method As String, Path As String
    Dim cmdArgs As New Dictionary, parmKey As Variant
    Dim response As Dictionary
    
    method = driverCommand(0)
    Path = driverCommand(1)
    
    If parameters Is Nothing Then
        Set parameters = New Dictionary
    End If
    
    'Set session id
    If Not parameters.Exists("sessionId") Then
        parameters.Add "sessionId", sessionId_
    End If
    
    For Each parmKey In parameters
        If InStr(Path, "$" & parmKey) > 0 Then 'path parameter
            Path = Replace(Path, "$" & parmKey, parameters(parmKey))
        Else 'non-path parameter
            cmdArgs.Add parmKey, parameters(parmKey)
        End If
    Next parmKey
    
    'there are commands that don't require sessionId in path: CMD_STATUS, CMD_NEW_SESSION, CMD_GET_ALL_SESSIONS, CMD_SHUTDOWN
    If cmdArgs.Exists("sessionId") Then cmdArgs.Remove "sessionId"
    
    'Send request to selenium server
    Set response = SendRequest(method, driverUrl_ + Path, cmdArgs)
    
    If raise And IsResponseError(response) Then 'raise error if occurs
        Debug.Print GetResponseErrorMessage(response)
        MsgBox GetResponseErrorMessage(response), , "SeleniumVBA"
        End 'execution
        'Err.raise 513, "WebDriver.Execute", GetResponseErrorMessage(response)
    End If

    Set Execute = response 'always pass a dictionary object - let caller parse info based on context, including error
    
End Function

Private Function GetResponseErrorMessage(resp As Dictionary) As String
    Dim jc As New JSonConverter
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

Private Function Params(ParamArray keysAndValues()) As Dictionary
    Dim dict As New Dictionary
    Dim i As Integer
    For i = 0 To UBound(keysAndValues) - 1 Step 2
        dict.Add keysAndValues(i), keysAndValues(i + 1)
    Next i
    Set Params = dict
End Function

Public Sub GoBack()
    Dim data As New Dictionary
    Execute CMD_GO_BACK, data
End Sub

Public Sub GoForward()
    Dim data As New Dictionary
    Execute CMD_GO_FORWARD, data
End Sub

Public Sub Refresh()
    Dim data As New Dictionary
    Execute CMD_REFRESH, data
End Sub

Public Function GetCurrentURL() As String
    Dim data As New Dictionary
    Dim encodeURL As String
    encodeURL = Execute(CMD_GET_CURRENT_URL, data)("value")
    GetCurrentURL = ExecuteScript("return decodeURI('" & encodeURL & "')")
End Function

Public Function GetActiveElement() As WebElement
    Dim data As New Dictionary
   ' Return element
    Set GetActiveElement = ToWebElement(Execute(CMD_GET_ACTIVE_ELEMENT, data)("value")(ELEMENT_KEY))
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

Public Sub SetAttribute(element As WebElement, ByVal name As String, ByVal Value As String)
    'this sets the value of an html element attribute. Consider SendKeys or Action Chain for more control of user input elements.
    'If the attribute does not exist, then this does not throw an error.
    ExecuteScript "arguments[0].setAttribute('" & name & "', arguments[1])", element, Value
End Sub

Public Function GetAriaRole(element As WebElement) As String
    Dim data As New Dictionary, resp As Dictionary
    data.Add "id", element.elementId
    Set resp = Execute(CMD_GET_ELEMENT_ARIA_ROLE, data)
    If IsNull(resp("value")) Then
        GetAriaRole = ""
    Else
        GetAriaRole = resp("value")
    End If
End Function

Public Function GetAriaLabel(element As WebElement) As String
    Dim data As New Dictionary, resp As Dictionary
    data.Add "id", element.elementId
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

Public Sub SelectByValue(ByVal Value As String, element As WebElement)
    Dim data As New Dictionary
    Dim elem As WebElement
    data.Add "using", "css selector"
    data.Add "value", "option[value =" & """" & Value & """" & "]"
    data.Add "id", element.elementId
    Set elem = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, data)("value")(ELEMENT_KEY))
    If Not IsSelected(elem) Then
        elem.Click
    End If
End Sub

Public Sub SelectByVisibleText(ByVal Text As String, element As WebElement)
    Dim data As New Dictionary
    Dim elem As WebElement
    data.Add "using", "xpath"
    data.Add "value", ".//option[normalize-space(.) = " & """" & Text & """" & "]"
    data.Add "id", element.elementId
    Set elem = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, data)("value")(ELEMENT_KEY))
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
    Set elems = FindChildElements(element)
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
    Set elems = FindChildElements(element)
    For i = 1 To elems.Count
        If IsSelected(elems(i)) Then
            GetSelectedOptionText = elems(i).GetText
            Exit Function
        End If
    Next i
    GetSelectedOptionText = ""
End Function

Public Sub DeSelectByValue(ByVal Value As String, element As WebElement)
    Dim data As New Dictionary, elem As WebElement
    data.Add "using", "css selector"
    data.Add "value", "option[value =" & """" & Value & """" & "]"
    data.Add "id", element.elementId
    Set elem = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, data)("value")(ELEMENT_KEY))
    If IsSelected(elem) Then
        elem.Click
    End If
End Sub

Public Sub DeSelectByVisibleText(ByVal Text As String, element As WebElement)
    Dim data As New Dictionary, elem As WebElement
    data.Add "using", "xpath"
    data.Add "value", ".//option[normalize-space(.) = " & """" & Text & """" & "]"
    data.Add "id", element.elementId
    Set elem = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, data)("value")(ELEMENT_KEY))
    If IsSelected(elem) Then
        elem.Click
    End If
End Sub

Public Function IsSelected(element As WebElement) As Boolean
    Dim data As New Dictionary
    data.Add "id", element.elementId
    IsSelected = Execute(CMD_IS_ELEMENT_SELECTED, data)("value")
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
    Set elems = FindChildElements(element)
    SetFocus element
    For i = 1 To elems.Count
        ExecuteScript "arguments[0].options[" & CStr(i - 1) & "].selected = true;", element
    Next i
End Sub

Public Sub DeSelectAll(element As WebElement)
    Dim elems As WebElements
    Dim i As Integer
    Set elems = FindChildElements(element)
    SetFocus element
    For i = 1 To elems.Count
        ExecuteScript "arguments[0].options[" & CStr(i - 1) & "].selected = false;", element
    Next i
End Sub

Public Sub DeSelectByIndex(ByVal index As Integer, element As WebElement)
    Dim elems As WebElements
    Set elems = FindChildElements(element)
    If IsSelected(elems(index - baseDomIdx + 1)) Then
        ExecuteScript "return arguments[0].selectedIndex = " & CStr(index - baseDomIdx), elems(index - baseDomIdx + 1)
        elems(index - baseDomIdx + 1).Click
    End If
End Sub

Public Function ExecuteScriptAsync(ByVal script As String, ParamArray scriptArgs() As Variant)
    Dim data As New Dictionary, ElmData As New Dictionary
    Dim resp As Dictionary
    Dim args()
    
    data.Add "script", script
    
    If UBound(scriptArgs) >= 0 Then
        ReDim args(0 To UBound(scriptArgs))
        For i = 0 To UBound(scriptArgs)
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
    
    Set resp = Execute(CMD_EXECUTE_SCRIPT_ASYNC, data)
    
    Select Case TypeName(resp("value"))
    Case "Collection"
        Dim arry()
        arry = ScriptCollectionToConvertedArray(resp("value"))
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
    Dim data As New Dictionary, ElmData As New Dictionary
    Dim resp As Dictionary
    Dim args()
    
    data.Add "script", script
    
    If UBound(scriptArgs) >= 0 Then
        ReDim args(0 To UBound(scriptArgs))
        For i = 0 To UBound(scriptArgs)
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
    
    Set resp = Execute(CMD_EXECUTE_SCRIPT, data)
    
    Select Case TypeName(resp("value"))
    Case "Collection"
        Dim arry()
        arry = ScriptCollectionToConvertedArray(resp("value"))
        If (Not arry) = -1 Then
            'Dim jc As New JSonConverter
            'Debug.Print jc.ConvertToJson(resp, 4)
            'is it ok to assume that this empty collection is of type WebElements?
            Set ExecuteScript = New WebElements
            Exit Function
        End If
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
        .Text = png
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

Public Function GetRect(element As WebElement) As Dictionary
    'returns dictionary with following keys x, y, width, height
    Dim data As New Dictionary
    data.Add "id", element.elementId
    Set GetRect = Execute(CMD_GET_ELEMENT_RECT, data)("value")
End Function

Public Function GetPageSource() As String
    Dim data As New Dictionary
    GetPageSource = Execute(CMD_GET_PAGE_SOURCE, data)("value")
End Function

Public Function GetCurrentWindowHandle() As String
    Dim data As New Dictionary
    GetCurrentWindowHandle = Execute(CMD_GET_CURRENT_WINDOW_HANDLE, data)("value")
End Function

Public Function GetWindowHandles() As String()
    Dim data As New Dictionary
    Dim handles As Collection
    Dim ret() As String
    Dim i As Integer
    Set handles = Execute(CMD_GET_WINDOW_HANDLES, data)("value")
    'To array of handles
    ReDim ret(baseArrayIdx To handles.Count - 1 + baseArrayIdx)
    For i = 1 To handles.Count   ' handles is Collection, not array
        ret(i - 1 + baseArrayIdx) = handles.Item(i)
    Next
    GetWindowHandles = ret
End Function

Public Function SwitchToNewWindow(Optional ByVal windowType As svbaWindowType = svbaTab) As String
    Dim data As New Dictionary, hndl As String
    If windowType = svbaTab Then wtype = "tab" Else wtype = "window"
    data.Add "type", wtype 'tab or window
    hndl = Execute(CMD_NEW_WINDOW, data)("value")("handle")
    SwitchToWindow hndl 'above just creates a new window - we still need to switch to it
    SwitchToNewWindow = hndl
End Function

Public Function SwitchToWindow(ByVal handleOrIndex As Variant) As String
    Dim data As New Dictionary, handle As String
    
    If IsNumeric(handleOrIndex) Then
        handle = GetWindowHandles()(handleOrIndex)
    Else
        handle = handleOrIndex
    End If
    
    data.Add "handle", handle
    Execute CMD_SWITCH_TO_WINDOW, data
    
    SwitchToWindow = handle
End Function

Public Sub CloseWindow(Optional ByVal handleOrIndex As Variant)
    If Not IsMissing(handleOrIndex) Then SwitchToWindow handleOrIndex
    'closes the active window
    Dim data As New Dictionary
    Execute CMD_CLOSE, data
End Sub

Public Function GetWindowRect() As Dictionary
    'returns dictionary with following keys x, y, width, height
    Dim data As New Dictionary
    Set GetWindowRect = Execute(CMD_GET_WINDOW_RECT, data)("value")
End Function

Public Function SetWindowRect(Optional ByVal x As Variant, Optional ByVal y As Variant, Optional ByVal width As Variant, Optional ByVal height As Variant) As Dictionary
    'returns dictionary with following keys x, y, width, height
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
    Set SetWindowRect = Execute(CMD_SET_WINDOW_RECT, data)("value")
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
    Dim data As New Dictionary
    Execute CMD_MAXIMIZE_WINDOW, data
End Sub

Public Sub FullScreenWindow()
    Dim data As New Dictionary
    Execute CMD_FULLSCREEN_WINDOW, data
End Sub

Public Sub MinimizeWindow()
    Dim data As New Dictionary
    Execute CMD_MINIMIZE_WINDOW, data
End Sub

Public Sub SwitchToFrame(frameElement As WebElement)
    Dim data As New Dictionary
    Dim elem As New Dictionary
    elem.Add ELEMENT_KEY, frameElement.elementId
    data.Add "id", elem
    Execute CMD_SWITCH_TO_FRAME, data
End Sub

Public Sub SwitchToFrameByIndex(ByVal index As Integer)
    Dim data As New Dictionary
    data.Add "id", index - baseDomIdx
    Execute CMD_SWITCH_TO_FRAME, data
End Sub

Public Sub SwitchToDefaultContent()
    'switches to main document
    Dim data As New Dictionary
    data.Add "id", Null
    Execute CMD_SWITCH_TO_FRAME, data
End Sub

Public Sub SwitchToParentFrame()
    'if top-level frame then this switches to main document
    Dim data As New Dictionary
    Execute CMD_SWITCH_TO_PARENT_FRAME, data
End Sub

Public Function GetCurrentFrameName() As String
    GetCurrentFrameName = ExecuteScript("return self.name")
End Function

Public Sub SetImplicitlyWait(Optional ByVal millisecond As Long = 0)
    Dim data As New Dictionary
    data.Add "implicit", millisecond
    Execute CMD_SET_TIMEOUTS, data
End Sub

Public Sub SetPageLoadTimeout(Optional ByVal millisecond As Long = 300000)
    Dim data As New Dictionary
    data.Add "pageLoad", millisecond
    Execute CMD_SET_TIMEOUTS, data
End Sub

Public Sub SetScriptTimeout(Optional ByVal millisecond As Long = 30000)
    Dim data As New Dictionary
    data.Add "script", millisecond
    Execute CMD_SET_TIMEOUTS, data
End Sub

Public Function GetImplicitlyWait() As Double
    Dim data As New Dictionary
    GetImplicitlyWait = Execute(CMD_GET_TIMEOUTS, data)("value")("implicit")
End Function

Public Function GetPageLoadTimeout() As Double
    Dim data As New Dictionary
    GetPageLoadTimeout = Execute(CMD_GET_TIMEOUTS, data)("value")("pageLoad")
End Function

Public Function GetScriptTimeout() As Double
    Dim data As New Dictionary
    GetScriptTimeout = Execute(CMD_GET_TIMEOUTS, data)("value")("script")
End Function

Public Sub Wait(Optional ByVal milliseconds As Long = 300)
    'both of the system-level timers below work"
    '
    TimerWait milliseconds
    'Sleep milliseconds 'this works great, but lets try to do without windows api if possible
    '
    'these native methods below mysteriously interfere with/accept browser alerts - maybe because they are browser-level?
    '
    'ActionChain.Wait(milliseconds).Perform
    '
    'If milliseconds > 30000 Then SetScriptTimeout 2 * milliseconds '30000 is the default, so this isn't needed unless waitTime > 30 secs is needed
    'ExecuteScriptAsync "window.setTimeout(arguments[arguments.length - 1], arguments[0]);", milliseconds
    '
End Sub

Public Function IsPresent(by_ As by, ByVal Value As String) As Boolean
    Dim timeout As Double, resp As Dictionary, data As Dictionary
    
    timeout = GetImplicitlyWait()
    If timeout <> 0 Then SetImplicitlyWait 0
    Set data = ToSelector(by_, Value)
    Set resp = Execute(CMD_FIND_ELEMENT, data, False) 'don't raise an error if not present
    If IsResponseError(resp) Then
        IsPresent = False
    Else
        IsPresent = True
    End If
    If timeout <> 0 Then SetImplicitlyWait timeout

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
    GetAlertText = Execute(CMD_GET_ALERT_TEXT, data)("value")
End Function

Public Sub SetAlertText(ByVal alertText As String)
    Dim data As New Dictionary
    data.Add "text", alertText
    Execute CMD_SET_ALERT_VALUE, data
End Sub

Public Sub AcceptAlert()
    Dim data As New Dictionary
    Execute CMD_ACCEPT_ALERT, data
End Sub

Public Sub DismissAlert()
    Dim data As New Dictionary
    Execute CMD_DISMISS_ALERT, data
End Sub

Public Function IsAlertPresent() As Boolean
    Dim data As New Dictionary, resp As Dictionary
    Set resp = Execute(CMD_GET_ALERT_TEXT, data, False)
    If IsResponseError(resp) Then
        IsAlertPresent = False
    Else
        IsAlertPresent = True
    End If
End Function

Public Sub Clear(element As WebElement)
    Dim data As New Dictionary
    data.Add "id", element.elementId
    data.Add "text", vbNullString
    data.Add "value", vbNullString
    Execute CMD_CLEAR_ELEMENT, data
End Sub

Public Function GetTitle() As String
    Dim data As New Dictionary
    GetTitle = Execute(CMD_GET_TITLE, data)("value")
End Function

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
Public Function CreateCapabilities() As Capabilities
    Dim Capabilities As New Capabilities
    'insure that this function is being called in the correct order
    If browserName = "" Or sessionId_ <> "" Then
        MsgBox "The ""CreateCapabilities"" Method must be invoked after the ""StartEdge"" or ""StartChrome"" Methods, and before ""OpenBrowser"" Method."
        End 'execution
    End If
    Capabilities.InitializeFor browserName
    Set CreateCapabilities = Capabilities
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

Public Function GetAllCookies() As Cookies
    Dim data As New Dictionary, resp As Collection, ck As Cookie, cks As New Cookies
    Set resp = Execute(CMD_GET_ALL_COOKIES, data)("value")
    'Dim jc As New JSonConverter
    'Debug.Print jc.ConvertToJson(resp, 4)
    For i = 1 To resp.Count
        Set ck = New Cookie
        If resp(i).Exists("name") Then ck.name = resp(i)("name")
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
    If resp.Exists("value") Then ck.Value = resp("value")
    If resp.Exists("path") Then ck.Path = resp("path")
    Set GetCookie = ck
End Function

Public Sub DeleteCookie(ByVal cookieName As String)
    Dim data As New Dictionary
    data.Add "name", cookieName
    Execute CMD_DELETE_COOKIE, data
End Sub

Public Sub SetCookie(ck As Cookie)
    Dim data As New Dictionary
    'value can take a number string greater than 15 digits
    'need to set the following in JSonConverter else  will get an error on execute (see sendRequest):
    'jc.OptionsUseDoubleForLargeNumbers = True
    If ck.name <> "" And ck.Value <> "" Then
        data.Add "cookie", ck.ToDictionary
        Execute CMD_ADD_COOKIE, data
    End If
End Sub

Public Sub SetCookies(cks As Cookies)
    For i = 1 To cks.Count
        SetCookie cks.Item(i)
    Next i
End Sub

Public Sub DeleteAllCookies()
    Dim data As New Dictionary
    Execute CMD_DELETE_ALL_COOKIES, data
End Sub

Public Function IsPageFound(ByVal url As String) As Boolean
    If GetCurrentURL() = url Then
        IsPageFound = True
    Else
        IsPageFound = False
    End If
End Function

Private Function GetAbsolutePath(ByVal strPath As String) As String
    Dim fso As New FileSystemObject
    saveppath = CurDir()
    ChDrive ThisWorkbook.Path
    ChDir ThisWorkbook.Path
    GetAbsolutePath = fso.GetAbsolutePathName(strPath)
    ChDrive saveppath
    ChDir saveppath
End Function

Public Function GetSessionsInfo() As Collection 'add optional sessionID to get a particular one
    'change to dictionary
    Set GetSessionsInfo = Execute(CMD_GET_ALL_SESSIONS)("value")
End Function

Public Function GetDriverStatus() As Dictionary
    'Dim jc As New JSonConverter
    Set GetDriverStatus = Execute(CMD_STATUS)("value")
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

Public Function FindElementFromShadowRoot(by_ As by, ByVal Value As String, sRoot As ShadowRoot) As WebElement
    Dim data As Dictionary
    Set data = ToSelector(by_, Value)
    data.Add "sid", sRoot.shadowRootId
    Set FindElementFromShadowRoot = ToWebElement(Execute(CMD_FIND_ELEMENT_FROM_SHADOW_ROOT, data)("value")(ELEMENT_KEY))
End Function

Public Function FindElementsFromShadowRoot(by_ As by, ByVal Value As String, sRoot As ShadowRoot) As WebElements
    Dim data As Dictionary
    Dim elements As Collection
    Dim elems As WebElements
    Dim i As Integer
    
    Set data = ToSelector(by_, Value)
    data.Add "sid", sRoot.shadowRootId
    
    Set elements = Execute(CMD_FIND_ELEMENT_FROM_SHADOW_ROOT, data)("value")
    
    For i = 1 To elements.Count
        elems.Add ToWebElement(elements(i)(ELEMENT_KEY))
    Next
    
    Set FindElementsFromShadowRoot = elems
End Function

Public Sub UploadFile(element As WebElement, ByVal filepath As String)
    filepath = GetAbsolutePath(filepath)
    SendKeys element, filepath
End Sub

Public Sub DragAndDrop(source As WebElement, target As WebElement)
    ActionChain().DragAndDrop(source, target).Perform
End Sub

'see https://github.com/SeleniumHQ/selenium/commit/7a6bf7e30d8bd834b6982b1d28f002bb4b26f380
Private Sub InitCommands()
    CMD_STATUS = Array("GET", "/status")
    CMD_NEW_SESSION = Array("POST", "/session")
    CMD_GET_ALL_SESSIONS = Array("GET", "/sessions")
    CMD_QUIT = Array("DELETE", "/session/$sessionId")
    CMD_GET_CURRENT_WINDOW_HANDLE = Array("GET", "/session/$sessionId/window")
    CMD_GET_WINDOW_HANDLES = Array("GET", "/session/$sessionId/window/handles")
    CMD_GET = Array("POST", "/session/$sessionId/url")
    CMD_GO_FORWARD = Array("POST", "/session/$sessionId/forward")
    CMD_GO_BACK = Array("POST", "/session/$sessionId/back")
    CMD_REFRESH = Array("POST", "/session/$sessionId/refresh")
    CMD_EXECUTE_SCRIPT = Array("POST", "/session/$sessionId/execute/sync")
    CMD_EXECUTE_SCRIPT_ASYNC = Array("POST", "/session/$sessionId/execute/async")
    CMD_GET_CURRENT_URL = Array("GET", "/session/$sessionId/url")
    CMD_GET_TITLE = Array("GET", "/session/$sessionId/title")
    CMD_GET_PAGE_SOURCE = Array("GET", "/session/$sessionId/source")
    CMD_SCREENSHOT = Array("GET", "/session/$sessionId/screenshot")
    CMD_ELEMENT_SCREENSHOT = Array("GET", "/session/$sessionId/element/$id/screenshot")
    CMD_FIND_ELEMENT = Array("POST", "/session/$sessionId/element")
    CMD_FIND_ELEMENTS = Array("POST", "/session/$sessionId/elements")
    CMD_GET_ACTIVE_ELEMENT = Array("GET", "/session/$sessionId/element/active")
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
    CMD_DISMISS_ALERT = Array("POST", "/session/$sessionId/alert/dismiss")
    CMD_ACCEPT_ALERT = Array("POST", "/session/$sessionId/alert/accept")
    CMD_SET_ALERT_VALUE = Array("POST", "/session/$sessionId/alert/text")
    CMD_GET_ALERT_TEXT = Array("GET", "/session/$sessionId/alert/text")
    CMD_CLICK = Array("POST", "/session/$sessionId/click")
    'CMD_ACTIONS = Array("POST", "/session/$sessionId/actions") 'used in ActionChain.cls
    'CMD_CLEAR_ACTIONS = Array("DELETE", "/session/$sessionId/actions") 'used in ActionChain.cls
    CMD_SET_WINDOW_RECT = Array("POST", "/session/$sessionId/window/rect")
    CMD_GET_WINDOW_RECT = Array("GET", "/session/$sessionId/window/rect")
    CMD_MAXIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/maximize")
    CMD_GET_LOCATION = Array("GET", "/session/$sessionId/location")
    CMD_SET_LOCATION = Array("POST", "/session/$sessionId/location")
    CMD_GET_APP_CACHE_STATUS = Array("GET", "/session/$sessionId/application_cache/status")
    CMD_GET_NETWORK_CONNECTION = Array("GET", "/session/$sessionId/network_connection")
    CMD_SET_NETWORK_CONNECTION = Array("POST", "/session/$sessionId/network_connection")
    'CMD_GET_LOG = Array("POST", "/session/$sessionId/se/log")  'see logging option in StartEdge and StartChrome methods
    'CMD_GET_AVAILABLE_LOG_TYPES = Array("GET", "/session/$sessionId/se/log/types") 'see logging option in StartEdge and StartChrome methods
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

Private Sub Class_Initialize()
    'baseArrayIdx = LBound(Array())
    InitCommands
    AppWinStyle = vbHide 'hides the command window
End Sub

Public Sub SaveHTMLToFile(ByVal snippet As String, Optional ByVal Path As String = ".\snippet.html")
    'this is used to save test snippets to file for subsequent browser navigation
    Dim fs As New FileSystemObject, ts As textstream
    Path = GetAbsolutePath(Path)
    Set ts = fs.CreateTextFile(Path, True)
    ts.Write snippet
    ts.Close
End Sub

Private Sub TimerWait(ByVal durationMS As Integer)
    'pause in milliseconds
    Dim startTime As Single, endTime As Single, nowTime As Single, elapsedTime As Single
    
    startTime = VBA.timer()
    endTime = startTime + durationMS / 1000#
    
    Do While nowTime < endTime
        nowTime = VBA.timer()
        If nowTime < startTime Then 'oops - someone is burning the midnight oil!
            endTime = endTime - elapsedTime
            startTime = 0
        End If
        elapsedTime = nowTime - startTime
        DoEvents    'yield to other processes.
    Loop

End Sub
