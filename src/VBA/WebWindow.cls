VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used to manage the browser window - must be instantiated WebDriver's ActiveWindow method, or through the WebWindows object"
'@ModuleDescription "This class is used to manage the browser window - must be instantiated WebDriver's ActiveWindow method, or through the WebWindows object"
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
' ==========================================================================
' MIT License
'
' Copyright (c) 2023-2024, GCUser99 and 6DiegoDiego9 (https://github.com/GCuser99/SeleniumVBA)
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

Private Type typCMD
    CMD_GET_CURRENT_WINDOW_HANDLE As Variant
    CMD_SWITCH_TO_WINDOW As Variant
    CMD_SET_WINDOW_RECT As Variant
    CMD_GET_WINDOW_RECT As Variant
    CMD_MAXIMIZE_WINDOW As Variant
    CMD_FULLSCREEN_WINDOW As Variant
    CMD_MINIMIZE_WINDOW As Variant
    CMD_GET_TITLE As Variant
    CMD_GET_URL As Variant
    CMD_CLOSE  As Variant
    CMD_GET_WINDOW_HANDLES As Variant
End Type

Private tCMD As typCMD

Private driver_ As WebDriver
Private thisHandle As String
Private previousActiveHandle As String

'@Description("Activates this window")
Public Function Activate() As WebWindow
Attribute Activate.VB_Description = "Activates this window"
    Dim activeHandle As String
    
    'first check if this window is closed - raise error if so
    If isClosed(thisHandle) Then Err.Raise 404, , "Error in Activate method: This window has previously been closed"
    
    'get the currently active window handle
    activeHandle = getCurrentActiveHandle
    
    'if thisHandle is not the current active handle then activate this window
    If activeHandle <> thisHandle Then
        'save current active window's handle to try to activate later if/when this window is closed
        If activeHandle <> vbNullString Then previousActiveHandle = activeHandle
        'now activate this window
        switchToWindow thisHandle
    End If
    
    'make chainable
    Set Activate = Me
End Function

'@Description("Activates this window")
Public Function SwitchTo() As WebWindow
Attribute SwitchTo.VB_Description = "Activates this window"
    Set SwitchTo = Activate
End Function

'@Description("Returns a boolean value indicating whether this window is the active window")
Public Function IsActive() As Boolean
Attribute IsActive.VB_Description = "Returns a boolean value indicating whether this window is the active window"
    If getCurrentActiveHandle = thisHandle Then IsActive = True Else IsActive = False
End Function

'@Description("Closes this window")
Public Sub CloseIt()
Attribute CloseIt.VB_Description = "Closes this window"
    Dim handleToSwitchTo As String
    Dim winHandles As Collection
    Dim activeHandle As String
    
    'first check if this window is closed and if not, return handles collection for use below
    If isClosed(thisHandle, winHandles) Then Err.Raise 404, , "Error in CloseIt method: This window has previously been closed"
    
    'next get the active window handle if there is one
    activeHandle = getCurrentActiveHandle()
    
    If (thisHandle = activeHandle) Or (activeHandle = vbNullString) Then
        'this window is either the active window, or there is no active window at all
        'below attempts to determine which window to switch to after closing
        'first determine if previousActiveHandle is open, which could have been
        'set when this window was activated
        
        If previousActiveHandle <> vbNullString Then
            If Not isClosed(previousActiveHandle, winHandles) Then handleToSwitchTo = previousActiveHandle
        End If
        
        'if previousActiveHandle is not defined or is already closed, then check if there are only
        'two windows open (including this window) - if so then it makes sense to activate the remaining window
        'if there are more than two open, then window to switch to is indeterminate
        If handleToSwitchTo = vbNullString Then
            'if there are only two open windows, then find the window that is not this window
            If winHandles.Count = 2 Then
                If thisHandle <> winHandles(1) Then handleToSwitchTo = winHandles(1) Else handleToSwitchTo = winHandles(2)
            End If
        End If
    Else
        'a different window is currently active
        'save the active window handle for switching to later
        handleToSwitchTo = activeHandle
        'in order to close this window, it must be activated first
        switchToWindow thisHandle
    End If
    
    'close the window now
    driver_.execute tCMD.CMD_CLOSE
    
    'activate the window to switch to after closing this window
    If handleToSwitchTo <> vbNullString Then switchToWindow handleToSwitchTo
End Sub

'@Description("Gets this window's browser title")
Public Function Title() As String
Attribute Title.VB_Description = "Gets this window's browser title"
    Title = getWindowAttribute(tCMD.CMD_GET_TITLE)
End Function

'@Description("Gets this window's url")
Public Function Url(Optional ByVal decode As Boolean = False) As String
Attribute Url.VB_Description = "Gets this window's url"
    Url = getWindowAttribute(tCMD.CMD_GET_URL) 'this returns the encoded url
    If decode Then Url = driver_.decodeUrl(Url)
End Function

'@Description("Gets this window's attribute")
Private Function getWindowAttribute(AttribCMD As Variant) As String
Attribute getWindowAttribute.VB_Description = "Gets this window's attribute"
    Dim activeHandle As String
    
    'first check if this window is closed
    If isClosed(thisHandle) Then Err.Raise 404, , "Error in Title method: This window has previously been closed"
    
    'next get the active window handle if there is one
    activeHandle = getCurrentActiveHandle()
    
    'activate this window if not the active one
    If thisHandle <> activeHandle Then switchToWindow thisHandle
    
    'get the window's attribute
    getWindowAttribute = driver_.execute(AttribCMD)("value")
        
    'put window active state back to where it was before this action
    If thisHandle <> activeHandle Then switchToWindow activeHandle
End Function

'@Description("Sets/Returns a dictionary object specifying the x, y, width, and height of this window in screen pixels")
Public Property Get Bounds() As Dictionary
Attribute Bounds.VB_Description = "Sets/Returns a dictionary object specifying the x, y, width, and height of this window in screen pixels"
    Dim activeHandle As String
    
    'first check if this window is closed
    If isClosed(thisHandle) Then Err.Raise 404, , "Error in Bounds Get property: This window has previously been closed"
    
    'next get the active window handle if there is one
    activeHandle = getCurrentActiveHandle()
    
    'activate this window if not the active one
    If thisHandle <> activeHandle Then switchToWindow thisHandle
        
    'get window rectangle
    Set Bounds = driver_.execute(tCMD.CMD_GET_WINDOW_RECT)("value")
        
    'put window active state back to where it was before this action
    If thisHandle <> activeHandle Then switchToWindow activeHandle
End Property

Public Property Set Bounds(windowBounds As Dictionary)
    Dim activeHandle As String
    
    'first check if this window is closed
    If isClosed(thisHandle) Then Err.Raise 404, , "Error in Bounds Set property: This window has previously been closed"
    
    'next get the active window handle if there is one
    activeHandle = getCurrentActiveHandle()
    
    'activate this window if not the active one
    If thisHandle <> activeHandle Then switchToWindow thisHandle
    
    'check if bounds dictionary is complete - if not then fill in the missing parameters
    With windowBounds
        If Not (.Exists("x") And .Exists("y") And .Exists("height") And .Exists("width")) Then
            Dim currentWinBounds As Dictionary
            Set currentWinBounds = driver_.execute(tCMD.CMD_GET_WINDOW_RECT)("value")
            If Not .Exists("x") Then windowBounds.Add "x", currentWinBounds("x")
            If Not .Exists("y") Then windowBounds.Add "y", currentWinBounds("y")
            If Not .Exists("height") Then windowBounds.Add "height", currentWinBounds("height")
            If Not .Exists("width") Then windowBounds.Add "width", currentWinBounds("width")
        End If
    End With
    
    'set the window bounds
    driver_.execute tCMD.CMD_SET_WINDOW_RECT, windowBounds
        
    'put window active state back to where it was before this action
    If thisHandle <> activeHandle Then switchToWindow activeHandle
End Property

'@Description("Sets the size for this window in screen pixels")
Public Sub SetSize(Optional ByVal width As Variant, Optional ByVal height As Variant)
Attribute SetSize.VB_Description = "Sets the size for this window in screen pixels"
    Dim winBounds As New Dictionary
    If Not IsMissing(width) Then winBounds.Add "width", width
    If Not IsMissing(height) Then winBounds.Add "height", height
    Set Me.Bounds = winBounds
End Sub

'@Description("Sets the position for this window in screen pixels")
Public Sub SetPosition(Optional ByVal x As Variant, Optional ByVal y As Variant)
Attribute SetPosition.VB_Description = "Sets the position for this window in screen pixels"
    Dim winBounds As New Dictionary
    If Not IsMissing(x) Then winBounds.Add "x", x
    If Not IsMissing(y) Then winBounds.Add "y", y
    Set Me.Bounds = winBounds
End Sub

'@Description("Maximize this window")
Public Sub Maximize()
Attribute Maximize.VB_Description = "Maximize this window"
    Dim activeHandle As String
    
    'first check if this window is closed
    If isClosed(thisHandle) Then Err.Raise 404, , "Error in Maximize method: This window has previously been closed"
    
    'next get the active window handle if there is one
    activeHandle = getCurrentActiveHandle()
    
    'activate this window if not the active one
    If thisHandle <> activeHandle Then switchToWindow thisHandle
        
    'maximize the window
    driver_.execute tCMD.CMD_MAXIMIZE_WINDOW
        
    'put window active state back to where it was before this action
    If thisHandle <> activeHandle Then switchToWindow activeHandle
End Sub

'@Description("Sets this window to full screen mode")
Public Sub FullScreen()
Attribute FullScreen.VB_Description = "Sets this window to full screen mode"
    Dim activeHandle As String
    
    'first check if this window is closed
    If isClosed(thisHandle) Then Err.Raise 404, , "Error in FullScreen method: This window has previously been closed"
    
    'next get the active window handle if there is one
    activeHandle = getCurrentActiveHandle()
    
    'activate this window if not the active one
    If thisHandle <> activeHandle Then switchToWindow thisHandle
        
    'make window full-screen
    driver_.execute tCMD.CMD_FULLSCREEN_WINDOW
        
    'put window active state back to where it was before this action
    If thisHandle <> activeHandle Then switchToWindow activeHandle
End Sub

'@Description("Minimize this window")
Public Sub Minimize()
Attribute Minimize.VB_Description = "Minimize this window"
    Dim activeHandle As String
    
    'first check if this window is closed
    If isClosed(thisHandle) Then Err.Raise 404, , "Error in Minimize method: This window has previously been closed"
    
    'next get the active window handle if there is one
    activeHandle = getCurrentActiveHandle()
    
    'activate this window if not the active one
    If thisHandle <> activeHandle Then switchToWindow thisHandle
        
    'minimize the window
    driver_.execute tCMD.CMD_MINIMIZE_WINDOW
        
    'put window active state back to where it was before this action
    If thisHandle <> activeHandle Then switchToWindow activeHandle
End Sub

'@Description("Returns boolean indicating whether the specified window is same as this window")
Public Function IsSameAs(win As WebWindow) As Boolean
Attribute IsSameAs.VB_Description = "Returns boolean indicating whether the specified window is same as this window"
    IsSameAs = (win.Handle = thisHandle)
End Function

'@Description("Returns boolean indicating whether the specified window is not the same as this window")
Public Function IsNotSameAs(win As WebWindow) As Boolean
Attribute IsNotSameAs.VB_Description = "Returns boolean indicating whether the specified window is not the same as this window"
    IsNotSameAs = (win.Handle <> thisHandle)
End Function

Friend Property Set WebDriver(ByRef obj As WebDriver)
    Set driver_ = obj
End Property

'@Description("Returns this window's internal handle string")
Public Property Get Handle() As String
Attribute Handle.VB_Description = "Returns this window's internal handle string"
    Handle = thisHandle
End Property

Friend Property Let Handle(ByVal winHandle As String)
    thisHandle = winHandle
End Property

Private Function isClosed(ByVal winHandle As String, Optional ByRef winHandles As Collection) As Boolean
    'determine if a window is closed, given its handle
    'optionally pass/return the handles collection for improved efficiency here and/or in caller
    Dim i As Long
    If winHandles Is Nothing Then Set winHandles = driver_.execute(tCMD.CMD_GET_WINDOW_HANDLES)("value")
    isClosed = True
    For i = 1 To winHandles.Count
        If winHandles.Item(i) = winHandle Then
            isClosed = False
            Exit Function
        End If
    Next i
End Function

Private Function getCurrentActiveHandle() As String
    Dim resp As Dictionary
    'get the active window handle if there is one
    Set resp = driver_.execute(tCMD.CMD_GET_CURRENT_WINDOW_HANDLE, , False)
    If isResponseError(resp) Then
        If resp("value")("errorNumber") = 404 Then
            'no window is currently active
            getCurrentActiveHandle = vbNullString
        Else
            'some other error occurred so report to user
            Err.Raise resp.Item("value").Item("errorNumber"), , resp.Item("value").Item("message")
        End If
    Else
        getCurrentActiveHandle = resp("value")
    End If
End Function

Private Sub switchToWindow(ByVal winHandle As String)
    'activate the window associated with specified handle
    If winHandle <> vbNullString Then
        Dim data As New Dictionary
        data.Add "handle", winHandle
        driver_.execute tCMD.CMD_SWITCH_TO_WINDOW, data
        'nominal delay to let finish loading on switch -
        'reduces chance of 500 unknown error: cannot determine loading status from no such window
        WebShared.sleep 50
    End If
End Sub

Private Sub Class_Initialize()
    With tCMD
        .CMD_GET_CURRENT_WINDOW_HANDLE = Array("GET", "/session/$sessionId/window")
        .CMD_SET_WINDOW_RECT = Array("POST", "/session/$sessionId/window/rect")
        .CMD_GET_WINDOW_RECT = Array("GET", "/session/$sessionId/window/rect")
        .CMD_MAXIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/maximize")
        .CMD_FULLSCREEN_WINDOW = Array("POST", "/session/$sessionId/window/fullscreen")
        .CMD_MINIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/minimize")
        .CMD_SWITCH_TO_WINDOW = Array("POST", "/session/$sessionId/window")
        .CMD_GET_TITLE = Array("GET", "/session/$sessionId/title")
        .CMD_GET_URL = Array("GET", "/session/$sessionId/url")
        .CMD_CLOSE = Array("DELETE", "/session/$sessionId/window")
        .CMD_GET_WINDOW_HANDLES = Array("GET", "/session/$sessionId/window/handles")
    End With
End Sub
