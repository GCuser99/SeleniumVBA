VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cookies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================
' SeleniumVBA v0.0.8
' A Selenium wrapper for Edge and Chrome written in Windows VBA based on JSon wire protocol.
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

'Option Explicit
Private colCookie As Collection

Public Function Add(cke As Cookie) As Cookie
    colCookie.Add cke, cke.name
    Set Add = cke
End Function

Public Property Get Count() As Integer
   Count = colCookie.Count
End Property

Public Sub Remove(indexOrCookie As Variant)
    If TypeName(indexOrCookie) = "Cookie" Then
        colCookie.Remove indexOrCookie.name
    Else
        colCookie.Remove indexOrCookie
    End If
End Sub

Public Property Get Item(ByVal name As Variant) As Cookie
Attribute Item.VB_UserMemId = 0
    'Attribute Dictionary.VB_UserMemId = 0
    'uncomment line above in text editor to make this the default property
    Set Item = colCookie.Item(name)
End Property

Friend Sub RemoveAll()
    Do While colCookie.Count > 0
        colCookie.Remove 1
    Loop
End Sub

Public Function IsMember(ByVal cke As Cookie) As Boolean
    Dim i As Long, ckeName As String
    ckeName = cke.name
    IsMember = False
    For i = 1 To Me.Count
        If Me.Item(i).name = ckeName Then
            IsMember = True
            Exit For
        End If
    Next i
End Function

Public Sub SaveToFile(Optional ByVal Path As String = ".\cookies.txt")
    Dim fs As New FileSystemObject, ts As textstream, darray() As Dictionary, i
    Dim jc As New JSonConverter
    Path = GetAbsolutePath(Path)
    Set ts = fs.CreateTextFile(Path, True)
    ReDim darray(1 To Me.Count)
    For i = 1 To Me.Count
        Set darray(i) = Me.Item(i).ToDictionary
    Next i
    ts.Write jc.ConvertToJson(darray(), 4)
    ts.Close
End Sub

Public Function LoadFromFile(Optional ByVal Path As String = ".\cookies.txt") As Cookies
    Dim fs As New FileSystemObject, ts As textstream, str As String, v As Collection, ck As Cookie
    Dim jc As New JSonConverter
    
    Me.RemoveAll
    
    Path = GetAbsolutePath(Path)
    
    Set ts = fs.OpenTextFile(Path)
    str = ts.ReadAll
    ts.Close
    Set v = jc.ParseJSON(str)
    
    For i = 1 To v.Count
        Set ck = New Cookie
        ck.name = v(i)("name")
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

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Attribute NewEnum.VB_UserMemID = -4
    'uncomment line above in text editor to enable For Each iteration in for loop
    Set NewEnum = colCookie.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set colCookie = New Collection
End Sub

Private Sub Class_Terminate()
    RemoveAll
    Set colCookie = Nothing
End Sub

Private Function GetAbsolutePath(ByVal strPath As String) As String
    Dim fso As New FileSystemObject
    saveppath = CurDir()
    ChDrive ThisWorkbook.Path
    ChDir ThisWorkbook.Path
    GetAbsolutePath = fso.GetAbsolutePathName(strPath)
    ChDrive saveppath
    ChDir saveppath
End Function

