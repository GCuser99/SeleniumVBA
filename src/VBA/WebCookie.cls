VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebCookie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used to manage/modify a cookie object"
'@ModuleDescription "This class is used to manage/modify a cookie object"
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

'Note: need to be on the domain that the cookie will be valid for before getting/setting cookies.
'If you are trying to preset cookies before you start interacting with a site and your homepage is large /
'takes a while to load an alternative is to find a smaller page on the site (typically the 404 page is small
'e.g. http://example.com/some404page)
'see https://www.selenium.dev/documentation/webdriver/browser/cookies/

Option Explicit

Private mDomain As String
Private mExpiry  As Long
Private mHttpOnly As Boolean
Private mName  As String
Private mPath  As String
Private mSecure As Boolean
Private mValue  As String
Private mSameSite As String

'@Description("Attribute to let servers specify whether/when cookies are sent with cross-site requests. Values are Strict, Lax (default), and None")
Public Property Get SameSite() As String
Attribute SameSite.VB_Description = "Attribute to let servers specify whether/when cookies are sent with cross-site requests. Values are Strict, Lax (default), and None"
    SameSite = mSameSite
End Property
Public Property Let SameSite(ByVal val As String)
    mSameSite = val
End Property

'@Description("Specifies name of the domain to associate the cookie with. The default value is the current domain")
Public Property Get Domain() As String
Attribute Domain.VB_Description = "Specifies name of the domain to associate the cookie with. The default value is the current domain"
    Domain = mDomain
End Property
Public Property Let Domain(ByVal val As String)
    mDomain = val
End Property

'@Description("Gets or sets the expiration date and time (Unix format) for the cookie")
Public Property Get ExpiryUnix() As Long
Attribute ExpiryUnix.VB_Description = "Gets or sets the expiration date and time (Unix format) for the cookie"
    ExpiryUnix = mExpiry
End Property
Public Property Let ExpiryUnix(ByVal val As Long)
    mExpiry = val
End Property

'@Description("Gets or sets the expiration date and time (Date format) for the cookie")
Public Property Get ExpiryDate() As Date
Attribute ExpiryDate.VB_Description = "Gets or sets the expiration date and time (Date format) for the cookie"
    ExpiryDate = DateAdd("s", mExpiry, "1/1/1970 00:00:00")
End Property
Public Property Let ExpiryDate(ByVal val As Date)
    mExpiry = DateDiff("s", "1/1/1970 00:00:00", val)
End Property

'@Description("Indicates whether the cookie has to be used only over HTTP or not. True if the cookie has the HttpOnly attribute and cannot be accessed through a client-side script; otherwise, false. The default is false")
Public Property Get HttpOnly() As Boolean
Attribute HttpOnly.VB_Description = "Indicates whether the cookie has to be used only over HTTP or not. True if the cookie has the HttpOnly attribute and cannot be accessed through a client-side script; otherwise, false. The default is false"
    HttpOnly = mHttpOnly
End Property
Public Property Let HttpOnly(ByVal val As Boolean)
    mHttpOnly = val
End Property

'@Description("Gets or sets the name of the cookie")
Public Property Get Name() As String
Attribute Name.VB_Description = "Gets or sets the name of the cookie"
    Name = mName
End Property
Public Property Let Name(ByVal val As String)
    mName = val
End Property

'@Description("Indicates whether the cookie can only be sent to the server over a secure HTTPS connection. The default value is false")
Public Property Get Secure() As Boolean
Attribute Secure.VB_Description = "Indicates whether the cookie can only be sent to the server over a secure HTTPS connection. The default value is false"
    Secure = mSecure
End Property
Public Property Let Secure(ByVal val As Boolean)
    mSecure = val
End Property

'@Description("Gets or sets the value of the cookie")
Public Property Get Value() As String
Attribute Value.VB_Description = "Gets or sets the value of the cookie"
    Value = mValue
End Property
Public Property Let Value(ByVal val As String)
    mValue = val
End Property

'@Description("Gets or sets the virtual path to transmit with the cookie. The default is /, which is the server root")
Public Property Get Path() As String
Attribute Path.VB_Description = "Gets or sets the virtual path to transmit with the cookie. The default is /, which is the server root"
    Path = mPath
End Property
Public Property Let Path(ByVal val As String)
    mPath = val
End Property

'@Description("Returns the cookie as a Dictionary object")
Public Function ToDictionary() As Dictionary
Attribute ToDictionary.VB_Description = "Returns the cookie as a Dictionary object"
    Dim ck As New Dictionary
    ck.Add "name", mName
    ck.Add "value", mValue
    If mDomain <> vbNullString Then ck.Add "domain", mDomain
    ck.Add "path", mPath
    ck.Add "secure", mSecure
    ck.Add "httpOnly", mHttpOnly
    If mExpiry > 0 Then ck.Add "expiry", mExpiry
    ck.Add "sameSite", mSameSite
    Set ToDictionary = ck
End Function

Private Sub Class_Initialize()
    'initialize defaults
    Me.Path = "/"
    Me.Secure = False
    Me.HttpOnly = False
    Me.SameSite = "Lax"
End Sub
