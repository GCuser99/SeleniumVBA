VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebCookies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used to manage a collection of cookie objects - must be instantiated using WebDriver's CreateCookies method"
'%ModuleDescription "This class is used to manage a collection of cookie objects - must be instantiated using WebDriver's CreateCookies method"
'@Exposed
'@folder("SeleniumVBA.Source")
' ==========================================================================
' SeleniumVBA v3.3
' A Selenium wrapper for Edge, Chrome, Firefox, and IE written in Windows VBA based on JSon wire protocol.
'
' (c) GCUser99
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' ==========================================================================
' MIT License
'
' Copyright (c) 2022 GCUser99
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

Private colCookie As Collection
Private driver_ As WebDriver

'%Description("Adds a cookie to the collection")
Public Function Add(cookie As WebCookie) As WebCookie
Attribute Add.VB_Description = "Adds a cookie to the collection"
    colCookie.Add cookie, cookie.Name
    Set Add = cookie
End Function

'%Description("Returns the number of cookies in the collection")
Public Property Get Count() As Integer
Attribute Count.VB_Description = "Returns the number of cookies in the collection"
   Count = colCookie.Count
End Property

'%Description("Returns the number of cookies in the collection")
Public Property Get Size() As Integer
Attribute Size.VB_Description = "Returns the number of cookies in the collection"
   Size = colCookie.Count
End Property

'%Description("Removes a cookie corresponding to the given index or cookie object")
Public Sub Remove(indexOrCookie As Variant)
Attribute Remove.VB_Description = "Removes a cookie corresponding to the given index or cookie object"
    If TypeName(indexOrCookie) = "Cookie" Then
        colCookie.Remove indexOrCookie.Name
    Else
        colCookie.Remove indexOrCookie
    End If
End Sub

'@DefaultMember
'%Description("Returns the cookie of the provided name")
Public Property Get Item(ByVal cookieName As Variant) As WebCookie
Attribute Item.VB_Description = "Returns the cookie of the provided name"
Attribute Item.VB_UserMemId = 0
    'Attribute Dictionary.VB_UserMemId = 0
    'uncomment line above in text editor to make this the default property
    Set Item = colCookie.Item(cookieName)
End Property

'%Description("Returns True if cookie object is already in the collection, False otherwise")
Public Function Exists(ByRef cookie As WebCookie) As Boolean
Attribute Exists.VB_Description = "Returns True if cookie object is already in the collection, False otherwise"
    On Error GoTo Err
    Exists = True
    colCookie.Item cookie.Name
    Exit Function
Err:
    Exists = False
End Function

'%Description("Saves the cookie collection to a file path")
Public Sub SaveToFile(Optional ByVal filePath As String = ".\cookies.txt")
Attribute SaveToFile.VB_Description = "Saves the cookie collection to a file path"
    Dim fs As New FileSystemObject
    Dim ts As TextStream
    Dim darray() As Dictionary
    Dim i As Long
    Dim jc As New WebJsonConverter
    filePath = WebShared.GetFullLocalPath(filePath, driver_.DefaultIOFolder)
    Set ts = fs.CreateTextFile(filePath, True, True)
    ReDim darray(1 To Me.Count)
    For i = 1 To Me.Count
        Set darray(i) = Me.Item(i).ToDictionary
    Next i
    ts.Write jc.ConvertToJson(darray(), 4)
    ts.Close
End Sub

'%Description("Loads the cookie collection from a file path")
Public Function LoadFromFile(Optional ByVal filePath As String = ".\cookies.txt") As WebCookies
Attribute LoadFromFile.VB_Description = "Loads the cookie collection from a file path"
    Dim fs As New FileSystemObject
    Dim ts As TextStream
    Dim str As String
    Dim v As Collection
    Dim ck As WebCookie
    Dim jc As New WebJsonConverter
    
    Me.RemoveAll
    
    filePath = WebShared.GetFullLocalPath(filePath, driver_.DefaultIOFolder)
    'note that this will fail if input file format is utf-8 with non-ascii chars - see https://github.com/GCuser99/SeleniumVBA/discussions/32
    Set ts = fs.OpenTextFile(filePath, , , TristateUseDefault)
    str = ts.ReadAll
    ts.Close
    Set v = jc.ParseJson(str)
    
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
Public Function _NewEnum() As IUnknown
Attribute _NewEnum.VB_UserMemId = -4
    Set _NewEnum = colCookie.[_NewEnum]
End Function

Friend Sub RemoveAll()
    Do While colCookie.Count > 0
        colCookie.Remove 1
    Loop
End Sub

Friend Property Set WebDriver(obj As WebDriver)
    Set driver_ = obj
End Property

Private Sub Class_Initialize()
    Set colCookie = New Collection
End Sub

Private Sub Class_Terminate()
    RemoveAll
    Set colCookie = Nothing
End Sub
