VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebCookies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used to manage a collection of cookie objects - must be instantiated using WebDriver's CreateCookies method"
'@ModuleDescription "This class is used to manage a collection of cookie objects - must be instantiated using WebDriver's CreateCookies method"
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
'
' This class was written to support collection of Cookie objects
'
' ==========================================================================
' Notes:
'
' - WebCookies object should be instantiated using WebDriver.MakeCookies, as in following example
'
'   Dim cks as WebCookies, driver as new WebDriver
'
'   Set cks = driver.CreateCookies
'
'
Option Explicit

Private colCookies As Collection
Private driver_ As WebDriver

'@Description("Adds a cookie to the collection")
Public Function Add(cookie As WebCookie) As WebCookie
Attribute Add.VB_Description = "Adds a cookie to the collection"
    colCookies.Add cookie
    Set Add = cookie
End Function

'@Description("Returns the number of cookies in the collection")
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of cookies in the collection"
   Count = colCookies.Count
End Property

'@Description("Returns the number of cookies in the collection")
Public Property Get Size() As Long
Attribute Size.VB_Description = "Returns the number of cookies in the collection"
   Size = colCookies.Count
End Property

'@Description("Removes a cookie corresponding to the given index")
Public Sub Remove(index As Long)
Attribute Remove.VB_Description = "Removes a cookie corresponding to the given index"
    colCookies.Remove index
End Sub

'@DefaultMember
'@Description("Returns the cookie of the provided index")
Public Property Get Item(ByVal index As Long) As WebCookie
Attribute Item.VB_Description = "Returns the cookie of the provided index"
Attribute Item.VB_UserMemId = 0
    Set Item = colCookies.Item(index)
End Property

'@Description("Saves the cookie collection to a file path")
Public Sub SaveToFile(Optional ByVal filePath As String = ".\cookies.txt")
Attribute SaveToFile.VB_Description = "Saves the cookie collection to a file path"
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim darray() As Dictionary
    Dim i As Long
    
    If Me.Count = 0 Then Err.Raise 1, , "Error in WebCookies: SaveToFile method failed because there are no cookies to save."
    
    filePath = driver_.ResolvePath(filePath)
    Set ts = fso.CreateTextFile(filePath, True, True)
    ReDim darray(1 To Me.Count)
    For i = 1 To Me.Count
        Set darray(i) = Me.Item(i).ToDictionary
    Next i
    ts.Write WebJsonConverter.ConvertToJson(darray(), 4)
    ts.Close
End Sub

'@Description("Loads the cookie collection from a file path")
Public Function LoadFromFile(Optional ByVal filePath As String = ".\cookies.txt") As WebCookies
Attribute LoadFromFile.VB_Description = "Loads the cookie collection from a file path"
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim str As String
    Dim v As Collection
    Dim ck As WebCookie
    
    Me.RemoveAll
    
    filePath = driver_.ResolvePath(filePath)
    
    If Not fso.FileExists(filePath) Then Err.Raise 1, , "Error in WebCookies: LoadToFile method failed because the specified file does not exist."
    
    'note that this will fail if input file format is utf-8 with non-ascii chars - see https://github.com/GCuser99/SeleniumVBA/discussions/32
    Set ts = fso.OpenTextFile(filePath, , , TristateUseDefault)
    str = ts.ReadAll
    ts.Close
    Set v = WebJsonConverter.ParseJson(str)
    
    Dim i As Long
    For i = 1 To v.Count
        Set ck = New WebCookie
        ck.Name = v(i)("name")
        ck.Domain = v(i)("domain")
        ck.ExpiryUnix = v(i)("expiry")
        ck.HttpOnly = v(i)("httpOnly")
        ck.SameSite = v(i)("sameSite")
        ck.Secure = v(i)("secure")
        ck.Value = v(i)("value")
        ck.Path = v(i)("path")
        Me.Add ck
    Next i
    Set LoadFromFile = Me
End Function

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colCookies.[_NewEnum]
End Function

'@Description("Removes all member cookie objects from the collection")
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Removes all member cookie objects from the collection"
    Do While colCookies.Count > 0
        colCookies.Remove 1
    Loop
End Sub

Friend Property Set WebDriver(obj As WebDriver)
    Set driver_ = obj
End Property

Private Sub Class_Initialize()
    Set colCookies = New Collection
End Sub

Private Sub Class_Terminate()
    Set colCookies = Nothing
End Sub
