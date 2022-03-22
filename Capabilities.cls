VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Capabilities"
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

' Spec for capabilities:
' - Chrome: https://chromedriver.chromium.org/capabilities
' - Edge: https://docs.microsoft.com/en-us/microsoft-edge/webdriver-chromium/capabilities-edge-options

Private data_ As New Dictionary
Private browserOptionKey As String

Public Property Let args(value() As String)
    SetOption "args", value
End Property

Public Property Get args() As String()
    args = GetOption("args")
End Property

Public Property Let Binary(ByVal value As String)
    SetOption "binary", value
End Property

Public Property Get Binary() As String
    Binary = GetOption("binary")
End Property

Public Property Let Extensions(value() As String)
    SetOption "extensions", value
End Property

Public Property Get Extensions() As String()
    Extensions = GetOption("extensions")
End Property

Public Property Let LocalState(ByVal value As Dictionary)
    SetOption "localState", value
End Property

Public Property Get LocalState() As Dictionary
    Set LocalState = GetOption("localState")
End Property

Public Property Let Prefs(value As Dictionary)
    SetOption "prefs", value
End Property

Public Property Get Prefs() As Dictionary
    Set Prefs = GetOption("prefs")
End Property

Public Property Let Detach(value As Boolean)
    SetOption "detach", value
End Property

Public Property Get Detach() As Boolean
    Detach = GetOption("detach")
End Property

Public Property Let DebuggerAddress(ByVal value As String)
    SetOption "debuggerAddress", value
End Property

Public Property Get DebuggerAddress() As String
    DebuggerAddress = GetOption("debuggerAddress")
End Property

Public Property Let ExcludeSwitches(value() As String)
    SetOption "excludeSwitches", value
End Property

Public Property Get ExcludeSwitches() As String()
    ExcludeSwitches = GetOption("excludeSwitches")
End Property

Public Property Let MinidumpPath(ByVal value As String)
    SetOption "minidumpPath", value
End Property

Public Property Get MinidumpPath() As String
    MinidumpPath = GetOption("minidumpPath")
End Property

Public Property Let MobileEmulation(value As Dictionary)
    SetOption "mobileEmulation", value
End Property

Public Property Get MobileEmulation() As Dictionary
    Set MobileEmulation = GetOption("mobileEmulation")
End Property

Public Property Let PerfLoggingPrefs(value As Dictionary)
    SetOption "perfLoggingPrefs", value
End Property

Public Property Get PerfLoggingPrefs() As Dictionary
    Set PerfLoggingPrefs = GetOption("perfLoggingPrefs")
End Property

Public Property Let WindowTypes(ByVal value As String)
    SetOption "windowTypes", value
End Property

Public Property Get WindowTypes() As String
    WindowTypes = GetOption("windowTypes")
End Property

Public Sub Add(ByVal key As String, value As Variant)
    If IsObject(value) Then
        Set data_("alwaysMatch")(key) = value
    Else
        data_("alwaysMatch")(key) = value
    End If
End Sub

Public Sub SetOption(ByVal key As String, value As Variant)
    If IsObject(value) Then
        Set data_("alwaysMatch")(browserOptionKey)(key) = value
    Else
        data_("alwaysMatch")(browserOptionKey)(key) = value
    End If
End Sub

Public Function GetOption(ByVal key As String) As Variant
    If IsObject(data_("alwaysMatch")(browserOptionKey)(key)) Then
        Set GetOption = data_("alwaysMatch")(browserOptionKey)(key)
    Else
        GetOption = data_("alwaysMatch")(browserOptionKey)(key)
    End If
End Function

Public Sub AddArgument(ByVal argument As String)
    Dim arguments() As String
    Dim idx As Integer: idx = -1

    On Error Resume Next
    idx = UBound(args)
    On Error GoTo 0
    
    If idx >= 0 Then
        arguments = args
    End If
        
    ReDim Preserve arguments(idx + 1)
    arguments(UBound(arguments)) = argument
    
    args = arguments
End Sub

Public Sub SetArguments(ByVal value As String, Optional ByVal delimiter As String = " ")
    SetOption "args", Split(value, delimiter)
End Sub

' Helper method to add prefs
Public Sub AddPref(ByVal key As String, ByVal value As String)
    If Prefs Is Nothing Then
        SetOption "prefs", New Dictionary
    End If
    Prefs(key) = value
End Sub

' Convert to Dictionary (WebDriver takes Dictionary)
Public Function Dictionary() As Dictionary
Attribute Dictionary.VB_UserMemId = 0
    'edit class in text editor and uncomment line below to make Dictionary the default method for the class
    'this would then allow this syntax: cap() -> cap.Dictionary
    'Attribute Dictionary.VB_UserMemId = 0
    Set Dictionary = data_
End Function

' Convert to JSON string for debugging
Public Function ToJson() As String
    Dim jc As New JSonConverter
    ToJson = jc.ConvertToJson(data_, 4)
End Function

Public Sub InitializeFor(ByVal browserName As String, Optional ByVal optionKey As String)
    Set data_ = New Dictionary
    data_.Add "browserName", browserName
    data_.Add "alwaysMatch", New Dictionary
    
    Dim browserOptions As New Dictionary
    browserOptions.Add "args", Nothing
    browserOptions.Add "binary", Nothing
    browserOptions.Add "extensions", Nothing
    browserOptions.Add "localState", Nothing
    browserOptions.Add "prefs", Nothing
    browserOptions.Add "detach", Nothing
    browserOptions.Add "debuggerAddress", Nothing
    browserOptions.Add "excludeSwitches", Nothing
    browserOptions.Add "minidumpPath", Nothing
    browserOptions.Add "mobileEmulation", Nothing
    browserOptions.Add "perfLoggingPrefs", Nothing
    browserOptions.Add "windowTypes", Nothing
    
    Select Case LCase(browserName)
        Case "chrome"
            browserOptionKey = "goog:chromeOptions"

        Case "msedge"
            browserOptionKey = "ms:edgeOptions"

        Case Else
            browserOptionKey = optionKey
    End Select
    
    data_("alwaysMatch").Add browserOptionKey, browserOptions
End Sub

Public Sub SetDownloadPrefs(Optional ByVal downloadDirectoryPath As String = ".\", Optional userPrompt As Boolean = False)
    
    downloadDirectoryPath = GetAbsolutePath(downloadDirectoryPath)
    'these work for both Chrome and Edge
    
    Me.AddPref "download.default_directory", downloadDirectoryPath
    Me.AddPref "download.directory_upgrade", True
    Me.AddPref "download.prompt_for_download", userPrompt 'false disables prompting user for download permission
    Me.AddPref "safebrowsing.enabled", True
    
End Sub

Public Sub InitializeForChrome()
    InitializeFor "chrome"
End Sub

Public Sub InitializeForEdge()
    InitializeFor "msedge"
End Sub

Public Sub RunInvisible()
    Me.AddArgument "--headless"
End Sub

Private Function GetAbsolutePath(ByVal strPath As String) As String
    Dim fso As New FileSystemObject
    saveppath = CurDir()
    ChDrive ThisWorkbook.path
    ChDir ThisWorkbook.path
    GetAbsolutePath = fso.GetAbsolutePathName(strPath)
    ChDrive saveppath
    ChDir saveppath
End Function

