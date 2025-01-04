VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used to manage a collection of WebWindow objects - must be instantiated using WebDriver's Windows method"
'@ModuleDescription "This class is used to manage a collection of WebWindow objects - must be instantiated using WebDriver's Windows method"
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

Option Explicit

Private colWindows As Collection
Private driver_ As WebDriver

Private Type typCMD
    CMD_GET_WINDOW_HANDLES As Variant
    CMD_GET_CURRENT_WINDOW_HANDLE As Variant
    CMD_NEW_WINDOW As Variant
End Type

Private Enum SearchBy
    byTitle
    byUrl
End Enum

Private tCMD As typCMD

'@Description("Opens, activates, and returns a new window of the specified type")
Public Function SwitchToNew(Optional ByVal windowType As svbaWindowType = svbaTab) As WebWindow
Attribute SwitchToNew.VB_Description = "Opens, activates, and returns a new window of the specified type"
    Dim win As New WebWindow
    Dim data As New Dictionary
    Dim wtype As String
    
    If windowType = svbaTab Then wtype = "tab" Else wtype = "window"
    
    data.Add "type", wtype 'tab or window
    win.Handle = driver_.execute(tCMD.CMD_NEW_WINDOW, data)("value")("handle")
    
    Set win.WebDriver = driver_
    colWindows.Add win, win.Handle
    
    'above just creates a new window - we still need to activate it
    win.Activate
    
    Set SwitchToNew = win
End Function

'@Description("Activates and returns the first window found after the active window with the specified (partial) window title text")
Public Function SwitchToByTitle(ByVal titleText As String, Optional ByVal partialMatch As Boolean = True) As WebWindow
Attribute SwitchToByTitle.VB_Description = "Activates and returns the first window found after the active window with the specified (partial) window title text"
    Set SwitchToByTitle = switchToBy(byTitle, titleText, partialMatch)
End Function

'@Description("Activates and returns the first window found after the active window with the specified (partial) window url text")
Public Function SwitchToByUrl(ByVal urlText As String, Optional ByVal partialMatch As Boolean = True) As WebWindow
Attribute SwitchToByUrl.VB_Description = "Activates and returns the first window found after the active window with the specified (partial) window url text"
    Set SwitchToByUrl = switchToBy(byUrl, urlText, partialMatch)
End Function

'@Description("Activates and returns the first window found after the active window with the specified (partial) window attribute text")
Private Function switchToBy(By As SearchBy, ByVal searchText As String, Optional ByVal partialMatch As Boolean = True) As WebWindow
Attribute switchToBy.VB_Description = "Activates and returns the first window found after the active window with the specified (partial) window attribute text"
    Dim activeHandle As String
    Dim i As Long
    Dim pos As Long
    Dim WndAttribValue As String
    
    'decode url search string to force apples-to-apples compare
    If By = SearchBy.byUrl Then searchText = driver_.decodeUrl(searchText)
    
    searchText = LCase$(searchText) 'for case-insensitive compare
    
    If partialMatch Then searchText = "*" & searchText & "*"
    
    'find the position of the currently activated window in the windows collection
    activeHandle = getCurrentActiveHandle
    If activeHandle <> vbNullString Then pos = getOrdinalPosInWindows(activeHandle)
    
    'starting from the next position in the windows collection, look for the window
    'containing the specified title text - wrap to the beginning if needed
    For i = 1 To colWindows.Count
        If pos = colWindows.Count Then pos = 1 Else pos = pos + 1
        Select Case By
            Case SearchBy.byTitle
                WndAttribValue = LCase$(colWindows.Item(pos).Title)
            Case SearchBy.byUrl
                'decode url string to force apples-to-apples compare
                WndAttribValue = LCase$(colWindows.Item(pos).Url(decode:=True))
        End Select
        If WndAttribValue Like searchText Then
            colWindows.Item(pos).Activate
            Set switchToBy = colWindows.Item(pos)
            Exit Function
        End If
    Next i
    
    If switchToBy Is Nothing Then Err.Raise 404, , "Error in SwitchToBy method: Window with specified string not found"
End Function

'@Description("Activates and returns the first window found after the active one in the open windows collection")
Public Function SwitchToNext() As WebWindow
Attribute SwitchToNext.VB_Description = "Activates and returns the first window found after the active one in the open windows collection"
    Dim activeHandle As String
    Dim pos As Long
    activeHandle = getCurrentActiveHandle
    If activeHandle <> vbNullString Then pos = getOrdinalPosInWindows(activeHandle)
    If pos = colWindows.Count Then pos = 1 Else pos = pos + 1
    Set SwitchToNext = colWindows.Item(pos).Activate
End Function

'@Description("Returns the active window object")
Public Function Active() As WebWindow
Attribute Active.VB_Description = "Returns the active window object"
    Dim win As New WebWindow
    Dim activeHandle As String

    activeHandle = getCurrentActiveHandle
    If activeHandle = vbNullString Then Err.Raise 404, , "Error in GetActive method: There is currently no Active window"
    
    win.Handle = activeHandle
    Set win.WebDriver = driver_
    Set Active = win
End Function

'@Description("Returns the collection of open window handles")
Public Function Handles() As Collection
Attribute Handles.VB_Description = "Returns the collection of open window handles"
    Dim winHandles As New Collection
    Dim i As Long
    For i = 1 To colWindows.Count
        winHandles.Add colWindows.Item(i).Handle
    Next i
    Set Handles = winHandles
End Function

'@Description("Returns the collection of open window titles")
Public Function Titles() As Collection
Attribute Titles.VB_Description = "Returns the collection of open window titles"
    Dim winTitles As New Collection
    Dim i As Long
    For i = 1 To colWindows.Count
        winTitles.Add colWindows.Item(i).Title
    Next i
    Set Titles = winTitles
End Function

'@Description("Returns the collection of open window urls")
Public Function Urls(Optional ByVal decode As Boolean = False) As Collection
Attribute Urls.VB_Description = "Returns the collection of open window urls"
    Dim winUrls As New Collection
    Dim i As Long
    For i = 1 To colWindows.Count
        winUrls.Add colWindows.Item(i).Url(decode)
    Next i
    Set Urls = winUrls
End Function

'@Description("Refreshes the windows collection with all currently open windows")
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refreshes the windows collection with all currently open windows"
    Dim win As WebWindow
    Dim i As Long
    Dim winHandles As Collection
    RemoveAll
    Set winHandles = driver_.execute(tCMD.CMD_GET_WINDOW_HANDLES)("value")
    For i = 1 To winHandles.Count
        Set win = New WebWindow
        Set win.WebDriver = driver_
        win.Handle = winHandles(i)
        colWindows.Add win, winHandles(i)
    Next i
End Sub

'@Description("Returns the number of windows in the collection")
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of windows in the collection"
   Count = colWindows.Count
End Property

'@DefaultMember
'@Description("Returns the window associated with the provided index or handle")
Public Property Get Item(ByVal indexOrHandle As Variant) As WebWindow
Attribute Item.VB_Description = "Returns the window associated with the provided index or handle"
Attribute Item.VB_UserMemId = 0
    Set Item = colWindows.Item(indexOrHandle)
End Property

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colWindows.[_NewEnum]
End Function

Private Function Add(win As WebWindow) As WebWindow
    colWindows.Add win, win.Handle
    Set Add = win
End Function

Private Sub Remove(index As Long)
    colWindows.Remove index
End Sub

Private Sub RemoveAll()
    Do While colWindows.Count > 0
        colWindows.Remove 1
    Loop
End Sub

Friend Property Set WebDriver(obj As WebDriver)
    Set driver_ = obj
End Property

Private Sub Class_Initialize()
    Set colWindows = New Collection
    With tCMD
        .CMD_GET_CURRENT_WINDOW_HANDLE = Array("GET", "/session/$sessionId/window")
        .CMD_GET_WINDOW_HANDLES = Array("GET", "/session/$sessionId/window/handles")
        .CMD_NEW_WINDOW = Array("POST", "/session/$sessionId/window/new")
    End With
End Sub

Private Sub Class_Terminate()
    Set colWindows = Nothing
End Sub

Private Function getOrdinalPosInWindows(ByVal winHandle As String) As Long
    Dim i As Long
    For i = 1 To colWindows.Count
        If colWindows.Item(i).Handle = winHandle Then
            getOrdinalPosInWindows = i
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
            'an unexpected error occurred so report to user
            Err.Raise resp.Item("value").Item("errorNumber"), , resp.Item("value").Item("message")
        End If
    Else
        getCurrentActiveHandle = resp("value")
    End If
End Function
