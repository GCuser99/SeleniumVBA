VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' TinySeleniumVBA v0.1.0
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

Option Explicit

Public DefaultSessionId As String
Public UrlBase As String

' Driver commands
Public CMD_STATUS
Public CMD_NEW_SESSION
Public CMD_GET_ALL_SESSIONS
Public CMD_QUIT
Public CMD_GET_CURRENT_WINDOW_HANDLE
Public CMD_W3C_GET_CURRENT_WINDOW_HANDLE
Public CMD_GET_WINDOW_HANDLES
Public CMD_W3C_GET_WINDOW_HANDLES
Public CMD_GET
Public CMD_GO_FORWARD
Public CMD_GO_BACK
Public CMD_REFRESH
Public CMD_EXECUTE_SCRIPT
Public CMD_W3C_EXECUTE_SCRIPT
Public CMD_W3C_EXECUTE_SCRIPT_ASYNC
Public CMD_GET_CURRENT_URL
Public CMD_GET_TITLE
Public CMD_GET_PAGE_SOURCE
Public CMD_SCREENSHOT
Public CMD_ELEMENT_SCREENSHOT
Public CMD_FIND_ELEMENT
Public CMD_FIND_ELEMENTS
Public CMD_W3C_GET_ACTIVE_ELEMENT
Public CMD_GET_ACTIVE_ELEMENT
Public CMD_FIND_CHILD_ELEMENT
Public CMD_FIND_CHILD_ELEMENTS
Public CMD_CLICK_ELEMENT
Public CMD_CLEAR_ELEMENT
Public CMD_SUBMIT_ELEMENT
Public CMD_GET_ELEMENT_TEXT
Public CMD_SEND_KEYS_TO_ELEMENT
Public CMD_SEND_KEYS_TO_ACTIVE_ELEMENT
Public CMD_UPLOAD_FILE
Public CMD_GET_ELEMENT_VALUE
Public CMD_GET_ELEMENT_TAG_NAME
Public CMD_IS_ELEMENT_SELECTED
Public CMD_SET_ELEMENT_SELECTED
Public CMD_IS_ELEMENT_ENABLED
Public CMD_IS_ELEMENT_DISPLAYED
Public CMD_GET_ELEMENT_LOCATION
Public CMD_GET_ELEMENT_LOCATION_ONCE_SCROLLED_INTO_VIEW
Public CMD_GET_ELEMENT_SIZE
Public CMD_GET_ELEMENT_RECT
Public CMD_GET_ELEMENT_ATTRIBUTE
Public CMD_GET_ELEMENT_PROPERTY
Public CMD_GET_ALL_COOKIES
Public CMD_ADD_COOKIE
Public CMD_GET_COOKIE
Public CMD_DELETE_ALL_COOKIES
Public CMD_DELETE_COOKIE
Public CMD_SWITCH_TO_FRAME
Public CMD_SWITCH_TO_PARENT_FRAME
Public CMD_SWITCH_TO_WINDOW
Public CMD_CLOSE
Public CMD_GET_ELEMENT_VALUE_OF_CSS_PROPERTY
Public CMD_IMPLICIT_WAIT
Public CMD_EXECUTE_ASYNC_SCRIPT
Public CMD_SET_SCRIPT_TIMEOUT
Public CMD_SET_TIMEOUTS
Public CMD_DISMISS_ALERT
Public CMD_W3C_DISMISS_ALERT
Public CMD_ACCEPT_ALERT
Public CMD_W3C_ACCEPT_ALERT
Public CMD_SET_ALERT_VALUE
Public CMD_W3C_SET_ALERT_VALUE
Public CMD_GET_ALERT_TEXT
Public CMD_W3C_GET_ALERT_TEXT
Public CMD_SET_ALERT_CREDENTIALS
Public CMD_CLICK
Public CMD_W3C_ACTIONS
Public CMD_W3C_CLEAR_ACTIONS
Public CMD_DOUBLE_CLICK
Public CMD_MOUSE_DOWN
Public CMD_MOUSE_UP
Public CMD_MOVE_TO
Public CMD_GET_WINDOW_SIZE
Public CMD_SET_WINDOW_SIZE
Public CMD_GET_WINDOW_POSITION
Public CMD_SET_WINDOW_POSITION
Public CMD_SET_WINDOW_RECT
Public CMD_GET_WINDOW_RECT
Public CMD_MAXIMIZE_WINDOW
Public CMD_W3C_MAXIMIZE_WINDOW
Public CMD_SET_SCREEN_ORIENTATION
Public CMD_GET_SCREEN_ORIENTATION
Public CMD_SINGLE_TAP
Public CMD_TOUCH_DOWN
Public CMD_TOUCH_UP
Public CMD_TOUCH_MOVE
Public CMD_TOUCH_SCROLL
Public CMD_DOUBLE_TAP
Public CMD_LONG_PRESS
Public CMD_FLICK
Public CMD_EXECUTE_SQL
Public CMD_GET_LOCATION
Public CMD_SET_LOCATION
Public CMD_GET_APP_CACHE
Public CMD_GET_APP_CACHE_STATUS
Public CMD_CLEAR_APP_CACHE
Public CMD_GET_NETWORK_CONNECTION
Public CMD_SET_NETWORK_CONNECTION
Public CMD_GET_LOCAL_STORAGE_ITEM
Public CMD_REMOVE_LOCAL_STORAGE_ITEM
Public CMD_GET_LOCAL_STORAGE_KEYS
Public CMD_SET_LOCAL_STORAGE_ITEM
Public CMD_CLEAR_LOCAL_STORAGE
Public CMD_GET_LOCAL_STORAGE_SIZE
Public CMD_GET_SESSION_STORAGE_ITEM
Public CMD_REMOVE_SESSION_STORAGE_ITEM
Public CMD_GET_SESSION_STORAGE_KEYS
Public CMD_SET_SESSION_STORAGE_ITEM
Public CMD_CLEAR_SESSION_STORAGE
Public CMD_GET_SESSION_STORAGE_SIZE
Public CMD_GET_LOG
Public CMD_GET_AVAILABLE_LOG_TYPES
Public CMD_CURRENT_CONTEXT_HANDLE
Public CMD_CONTEXT_HANDLES
Public CMD_SWITCH_TO_CONTEXT
Public CMD_FULLSCREEN_WINDOW
Public CMD_MINIMIZE_WINDOW
Public CMD_SHUTDOWN

Private Const ELEMENT_KEY = "element-6066-11e4-a52e-4f735466cecf"

Public Enum By
    ID = 0
    TagName = 1
    ClassName = 2
    Name = 3
    CssSelector = 4
    Xpath = 5
End Enum


' ==========================================================================
' Setup and shutdown
' ==========================================================================

' Launch Edge Driver
Public Sub Edge(ByVal driverPath As String, Optional ByVal driverUrl As String = "http://localhost:9515")
    Start driverPath, driverUrl
End Sub

' Launch Chrome Driver
Public Sub Chrome(ByVal driverPath As String, Optional ByVal driverUrl As String = "http://localhost:9515")
    Start driverPath, driverUrl
End Sub

' Start WebDriver
Public Sub Start(ByVal driverPath As String, ByVal driverUrl As String)
    ' Start WebDriver executable
    If Shell(driverPath, vbMinimizedNoFocus) = 0 Then
        MsgBox "Failed in starting WebDriver"
        End
    End If
    
    ' Set WebDriver url
    UrlBase = driverUrl
    
    ' Initialize driver commands
    InitCommands
End Sub

' Shutdown WebDriver
Public Sub Shutdown()
    Execute CMD_SHUTDOWN
End Sub


' ==========================================================================
' Browser operations
' ==========================================================================

' Open browser
Public Function OpenBrowser(Optional capabilities As Dictionary = Nothing, Optional desiredCapabilities As Dictionary = Nothing, Optional ByVal useAsDefault As Boolean = True) As String
    If capabilities Is Nothing Then
        Set capabilities = New Dictionary
    End If
    
    Dim resp As Dictionary
    Set resp = Execute(CMD_NEW_SESSION, Params("capabilities", capabilities, "desiredCapabilities", desiredCapabilities))
    
    If useAsDefault Then
        DefaultSessionId = resp("sessionId")
    End If
    
    OpenBrowser = resp("sessionId")
End Function

' Close browser
Public Sub CloseBrowser(Optional ByVal sessionId As String = vbNullString)
    Dim data As New Dictionary
    If sessionId <> vbNullString Then
        data.Add "sessionId", sessionId
    End If
    
    Execute CMD_QUIT, data
    DefaultSessionId = vbNullString
End Sub

' Open url
Public Sub Navigate(ByVal url As String, Optional ByVal sessionId As String = vbNullString)
    Dim data As New Dictionary
    If sessionId <> vbNullString Then
        data.Add "sessionId", sessionId
    End If
    data.Add "url", url
    
    Execute CMD_GET, data
End Sub


' ==========================================================================
' DOM operations
' ==========================================================================

' Find DOM element
Public Function FindElement(by_ As By, value As String, Optional parentElementId As String = vbNullString, Optional ByVal sessionId As String = vbNullString) As WebElement
    Dim data As Dictionary
    Set data = ToSelector(by_, value)
    If sessionId <> vbNullString Then
        data.Add "sessionId", sessionId
    End If
    
    Dim cmd
    If parentElementId <> vbNullString Then
        data.Add "id", parentElementId
        cmd = CMD_FIND_CHILD_ELEMENT
    Else
        cmd = CMD_FIND_ELEMENT
    End If
    
    ' Return element
    Set FindElement = ToWebElement(Execute(cmd, data)(ELEMENT_KEY), sessionId)
End Function

' Find multiple DOM elements
Public Function FindElements(by_ As By, value As String, Optional parentElementId As String = vbNullString, Optional ByVal sessionId As String = vbNullString) As WebElement()
    Dim data As Dictionary
    Set data = ToSelector(by_, value)
    If sessionId <> vbNullString Then
        data.Add "sessionId", sessionId
    End If
    
    Dim cmd
    If parentElementId <> vbNullString Then
        data.Add "id", parentElementId
        cmd = CMD_FIND_CHILD_ELEMENTS
    Else
        cmd = CMD_FIND_ELEMENTS
    End If
    
    Dim elements
    Set elements = Execute(cmd, data)
    
    ' To array of ids
    Dim ret() As WebElement
    Dim i As Integer
    For i = 0 To elements.Count - 1  ' elements is Collection, not array
        ReDim Preserve ret(i)
        Set ret(i) = ToWebElement(elements(i + 1)(ELEMENT_KEY), sessionId)
    Next
    
    ' Return element ids
    FindElements = ret
End Function

' by* to CSS selector or Xpath
Private Function ToSelector(by_ As By, ByVal value As String) As Dictionary
    Dim data As New Dictionary
    If (by_ = By.Xpath) Then
        'Locator Strategy XPath
        data.Add "using", "xpath"
    Else:
        'Use Css Selector for Locator strategies other than
        data.Add "using", "css selector"
        If by_ = By.ID Then
            value = "[id=""" + value + """]"
        ElseIf by_ = By.ClassName Then
            value = "." + value
        ElseIf by_ = By.Name Then
            value = "[name=""" + value + """]"
        End If
    End If
    data.Add "value", value
    Set ToSelector = data
End Function

' Create element
Private Function ToWebElement(ByVal elementId As String, Optional ByVal sessionId As String = vbNullString) As WebElement
    Dim element As New WebElement
    Set element.Driver_ = Me
    If sessionId = vbNullString Then
        element.SessionId_ = sessionId
    Else
        element.SessionId_ = DefaultSessionId
    End If
    element.ElementId_ = elementId
    Set ToWebElement = element
End Function

' Returns element.value
Public Function GetValue(elementId As String, Optional ByVal sessionId As String = vbNullString) As String
    Dim data As New Dictionary
    If sessionId <> vbNullString Then
        data.Add "sessionId", sessionId
    End If
    data.Add "id", elementId
    data.Add "name", "value"

    GetValue = Execute(CMD_GET_ELEMENT_ATTRIBUTE, data)
End Function

' Set value to element
Public Sub SetValue(elementId As String, text As String, Optional ByVal sessionId As String = vbNullString)
    Dim data As New Dictionary
    If sessionId <> vbNullString Then
        data.Add "sessionId", sessionId
    End If
    data.Add "id", elementId
    data.Add "text", text
    
    Dim chars() As String
    ReDim chars(Len(text) - 1)
    Dim i As Integer
    For i = 0 To UBound(chars)
        chars(i) = Mid(text, i + 1, 1)
    Next
    data.Add "value", chars
    
    Execute CMD_CLEAR_ELEMENT, data
    Execute CMD_SEND_KEYS_TO_ELEMENT, data
End Sub

' Click element
Public Sub Click(elementId As String, Optional ByVal sessionId As String = vbNullString)
    Dim data As New Dictionary
    If sessionId <> vbNullString Then
        data.Add "sessionId", sessionId
    End If
    data.Add "id", elementId

    Execute CMD_CLICK_ELEMENT, data
End Sub

' Get text
Public Function GetText(elementId As String, Optional ByVal sessionId As String = vbNullString) As String
    Dim data As New Dictionary
    If sessionId <> vbNullString Then
        data.Add "sessionId", sessionId
    End If
    data.Add "id", elementId

    GetText = Execute(CMD_GET_ELEMENT_TEXT, data)
End Function


' ==========================================================================
' Common functions
' ==========================================================================

' Execute driver command
Public Function Execute(driverCommand, Optional parameters As Dictionary = Nothing)
    Dim method As String: method = driverCommand(0)
    Dim path As String: path = driverCommand(1)
    If parameters Is Nothing Then
        Set parameters = New Dictionary
    End If
    
    ' Set default session id if session id is missing
    If Not parameters.Exists("sessionId") Then
        parameters.Add "sessionId", DefaultSessionId
    End If
    
    ' Set params to path
    Dim paramKey As Variant
    For Each paramKey In parameters
        If VarType(parameters(paramKey)) = vbString Then
            path = Replace(path, "$" + paramKey, parameters(paramKey))
        End If
    Next

    ' Send request to selenium server
    Dim resp As Dictionary
    Set resp = SendRequest(method, UrlBase + path, parameters)
    
    ' Return value(s)
    If IsNull(resp("value")) Then
        Set Execute = New Dictionary
    ElseIf TypeName(resp("value")) = "Collection" Then
        Set Execute = resp("value")
    ElseIf VarType(resp("value")) = vbObject Then
        If resp("value").Exists("error") Then
            Err.Raise 513, "WebDriver.Execute", JsonConverter.ConvertToJson(resp("value"))
        Else
            Set Execute = resp("value")
        End If
    Else
        Execute = resp("value")
    End If
End Function

' Send HTTP request
Private Function SendRequest(method As String, url As String, Optional data As Dictionary = Nothing) As Dictionary
    Dim client As Object
    Set client = CreateObject("MSXML2.ServerXMLHTTP")
    
    client.Open method, url
    
    If method = "POST" Or method = "PUT" Then
        client.setRequestHeader "Content-Type", "application/json"
        client.send JsonConverter.ConvertToJson(data)
    Else
        client.send
    End If

    Do While client.readyState < 4
        DoEvents
    Loop

    Dim Json As Object
    Set Json = JsonConverter.ParseJson(client.responseText)
    Set SendRequest = Json
End Function


' ==========================================================================
' Driver commands
' ==========================================================================
Private Sub InitCommands()
    CMD_STATUS = Array("GET", "/status")
    CMD_NEW_SESSION = Array("POST", "/session")
    CMD_GET_ALL_SESSIONS = Array("GET", "/sessions")
    CMD_QUIT = Array("DELETE", "/session/$sessionId")
    CMD_GET_CURRENT_WINDOW_HANDLE = Array("GET", "/session/$sessionId/window_handle")
    CMD_W3C_GET_CURRENT_WINDOW_HANDLE = Array("GET", "/session/$sessionId/window")
    CMD_GET_WINDOW_HANDLES = Array("GET", "/session/$sessionId/window_handles")
    CMD_W3C_GET_WINDOW_HANDLES = Array("GET", "/session/$sessionId/window/handles")
    CMD_GET = Array("POST", "/session/$sessionId/url")
    CMD_GO_FORWARD = Array("POST", "/session/$sessionId/forward")
    CMD_GO_BACK = Array("POST", "/session/$sessionId/back")
    CMD_REFRESH = Array("POST", "/session/$sessionId/refresh")
    CMD_EXECUTE_SCRIPT = Array("POST", "/session/$sessionId/execute")
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
    CMD_GET_ACTIVE_ELEMENT = Array("POST", "/session/$sessionId/element/active")
    CMD_FIND_CHILD_ELEMENT = Array("POST", "/session/$sessionId/element/$id/element")
    CMD_FIND_CHILD_ELEMENTS = Array("POST", "/session/$sessionId/element/$id/elements")
    CMD_CLICK_ELEMENT = Array("POST", "/session/$sessionId/element/$id/click")
    CMD_CLEAR_ELEMENT = Array("POST", "/session/$sessionId/element/$id/clear")
    CMD_SUBMIT_ELEMENT = Array("POST", "/session/$sessionId/element/$id/submit")
    CMD_GET_ELEMENT_TEXT = Array("GET", "/session/$sessionId/element/$id/text")
    CMD_SEND_KEYS_TO_ELEMENT = Array("POST", "/session/$sessionId/element/$id/value")
    CMD_SEND_KEYS_TO_ACTIVE_ELEMENT = Array("POST", "/session/$sessionId/keys")
    CMD_UPLOAD_FILE = Array("POST", "/session/$sessionId/file")
    CMD_GET_ELEMENT_VALUE = Array("GET", "/session/$sessionId/element/$id/value")
    CMD_GET_ELEMENT_TAG_NAME = Array("GET", "/session/$sessionId/element/$id/name")
    CMD_IS_ELEMENT_SELECTED = Array("GET", "/session/$sessionId/element/$id/selected")
    CMD_SET_ELEMENT_SELECTED = Array("POST", "/session/$sessionId/element/$id/selected")
    CMD_IS_ELEMENT_ENABLED = Array("GET", "/session/$sessionId/element/$id/enabled")
    CMD_IS_ELEMENT_DISPLAYED = Array("GET", "/session/$sessionId/element/$id/displayed")
    CMD_GET_ELEMENT_LOCATION = Array("GET", "/session/$sessionId/element/$id/location")
    CMD_GET_ELEMENT_LOCATION_ONCE_SCROLLED_INTO_VIEW = Array("GET", "/session/$sessionId/element/$id/location_in_view")
    CMD_GET_ELEMENT_SIZE = Array("GET", "/session/$sessionId/element/$id/size")
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
    CMD_CLOSE = Array("DELETE", "/session/$sessionId/window")
    CMD_GET_ELEMENT_VALUE_OF_CSS_PROPERTY = Array("GET", "/session/$sessionId/element/$id/css/$propertyName")
    CMD_IMPLICIT_WAIT = Array("POST", "/session/$sessionId/timeouts/implicit_wait")
    CMD_EXECUTE_ASYNC_SCRIPT = Array("POST", "/session/$sessionId/execute_async")
    CMD_SET_SCRIPT_TIMEOUT = Array("POST", "/session/$sessionId/timeouts/async_script")
    CMD_SET_TIMEOUTS = Array("POST", "/session/$sessionId/timeouts")
    CMD_DISMISS_ALERT = Array("POST", "/session/$sessionId/dismiss_alert")
    CMD_W3C_DISMISS_ALERT = Array("POST", "/session/$sessionId/alert/dismiss")
    CMD_ACCEPT_ALERT = Array("POST", "/session/$sessionId/accept_alert")
    CMD_W3C_ACCEPT_ALERT = Array("POST", "/session/$sessionId/alert/accept")
    CMD_SET_ALERT_VALUE = Array("POST", "/session/$sessionId/alert_text")
    CMD_W3C_SET_ALERT_VALUE = Array("POST", "/session/$sessionId/alert/text")
    CMD_GET_ALERT_TEXT = Array("GET", "/session/$sessionId/alert_text")
    CMD_W3C_GET_ALERT_TEXT = Array("GET", "/session/$sessionId/alert/text")
    CMD_SET_ALERT_CREDENTIALS = Array("POST", "/session/$sessionId/alert/credentials")
    CMD_CLICK = Array("POST", "/session/$sessionId/click")
    CMD_W3C_ACTIONS = Array("POST", "/session/$sessionId/actions")
    CMD_W3C_CLEAR_ACTIONS = Array("DELETE", "/session/$sessionId/actions")
    CMD_DOUBLE_CLICK = Array("POST", "/session/$sessionId/doubleclick")
    CMD_MOUSE_DOWN = Array("POST", "/session/$sessionId/buttondown")
    CMD_MOUSE_UP = Array("POST", "/session/$sessionId/buttonup")
    CMD_MOVE_TO = Array("POST", "/session/$sessionId/moveto")
    CMD_GET_WINDOW_SIZE = Array("GET", "/session/$sessionId/window/$windowHandle/size")
    CMD_SET_WINDOW_SIZE = Array("POST", "/session/$sessionId/window/$windowHandle/size")
    CMD_GET_WINDOW_POSITION = Array("GET", "/session/$sessionId/window/$windowHandle/position")
    CMD_SET_WINDOW_POSITION = Array("POST", "/session/$sessionId/window/$windowHandle/position")
    CMD_SET_WINDOW_RECT = Array("POST", "/session/$sessionId/window/rect")
    CMD_GET_WINDOW_RECT = Array("GET", "/session/$sessionId/window/rect")
    CMD_MAXIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/$windowHandle/maximize")
    CMD_W3C_MAXIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/maximize")
    CMD_SET_SCREEN_ORIENTATION = Array("POST", "/session/$sessionId/orientation")
    CMD_GET_SCREEN_ORIENTATION = Array("GET", "/session/$sessionId/orientation")
    CMD_SINGLE_TAP = Array("POST", "/session/$sessionId/touch/click")
    CMD_TOUCH_DOWN = Array("POST", "/session/$sessionId/touch/down")
    CMD_TOUCH_UP = Array("POST", "/session/$sessionId/touch/up")
    CMD_TOUCH_MOVE = Array("POST", "/session/$sessionId/touch/move")
    CMD_TOUCH_SCROLL = Array("POST", "/session/$sessionId/touch/scroll")
    CMD_DOUBLE_TAP = Array("POST", "/session/$sessionId/touch/doubleclick")
    CMD_LONG_PRESS = Array("POST", "/session/$sessionId/touch/longclick")
    CMD_FLICK = Array("POST", "/session/$sessionId/touch/flick")
    CMD_EXECUTE_SQL = Array("POST", "/session/$sessionId/execute_sql")
    CMD_GET_LOCATION = Array("GET", "/session/$sessionId/location")
    CMD_SET_LOCATION = Array("POST", "/session/$sessionId/location")
    CMD_GET_APP_CACHE = Array("GET", "/session/$sessionId/application_cache")
    CMD_GET_APP_CACHE_STATUS = Array("GET", "/session/$sessionId/application_cache/status")
    CMD_CLEAR_APP_CACHE = Array("DELETE", "/session/$sessionId/application_cache/clear")
    CMD_GET_NETWORK_CONNECTION = Array("GET", "/session/$sessionId/network_connection")
    CMD_SET_NETWORK_CONNECTION = Array("POST", "/session/$sessionId/network_connection")
    CMD_GET_LOCAL_STORAGE_ITEM = Array("GET", "/session/$sessionId/local_storage/key/$key")
    CMD_REMOVE_LOCAL_STORAGE_ITEM = Array("DELETE", "/session/$sessionId/local_storage/key/$key")
    CMD_GET_LOCAL_STORAGE_KEYS = Array("GET", "/session/$sessionId/local_storage")
    CMD_SET_LOCAL_STORAGE_ITEM = Array("POST", "/session/$sessionId/local_storage")
    CMD_CLEAR_LOCAL_STORAGE = Array("DELETE", "/session/$sessionId/local_storage")
    CMD_GET_LOCAL_STORAGE_SIZE = Array("GET", "/session/$sessionId/local_storage/size")
    CMD_GET_SESSION_STORAGE_ITEM = Array("GET", "/session/$sessionId/session_storage/key/$key")
    CMD_REMOVE_SESSION_STORAGE_ITEM = Array("DELETE", "/session/$sessionId/session_storage/key/$key")
    CMD_GET_SESSION_STORAGE_KEYS = Array("GET", "/session/$sessionId/session_storage")
    CMD_SET_SESSION_STORAGE_ITEM = Array("POST", "/session/$sessionId/session_storage")
    CMD_CLEAR_SESSION_STORAGE = Array("DELETE", "/session/$sessionId/session_storage")
    CMD_GET_SESSION_STORAGE_SIZE = Array("GET", "/session/$sessionId/session_storage/size")
    CMD_GET_LOG = Array("POST", "/session/$sessionId/log")
    CMD_GET_AVAILABLE_LOG_TYPES = Array("GET", "/session/$sessionId/log/types")
    CMD_CURRENT_CONTEXT_HANDLE = Array("GET", "/session/$sessionId/context")
    CMD_CONTEXT_HANDLES = Array("GET", "/session/$sessionId/contexts")
    CMD_SWITCH_TO_CONTEXT = Array("POST", "/session/$sessionId/context")
    CMD_FULLSCREEN_WINDOW = Array("POST", "/session/$sessionId/window/fullscreen")
    CMD_MINIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/minimize")
    CMD_SHUTDOWN = Array("GET", "/shutdown")
End Sub


' ==========================================================================
' Utility functions
' ==========================================================================

' KeyValue arguments to parameters as dictionaly
Private Function Params(ParamArray keysAndValues()) As Dictionary
    Dim dict As New Dictionary
    Dim i As Integer
    For i = 0 To UBound(keysAndValues) - 1 Step 2
        dict.Add keysAndValues(i), keysAndValues(i + 1)
    Next i
    Set Params = dict
End Function

