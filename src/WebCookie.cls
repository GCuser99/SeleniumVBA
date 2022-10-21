VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebCookie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================
' SeleniumVBA v2.3
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

'Lax, Strict, None
Public Property Get SameSite() As String
    SameSite = mSameSite
End Property
Public Property Let SameSite(ByVal val As String)
    mSameSite = val
End Property

'The name of the domain to associate the cookie with. The default value is the current domain.
Public Property Get Domain() As String
    Domain = mDomain
End Property

Public Property Let Domain(ByVal val As String)
    mDomain = val
End Property

'Gets or sets the expiration date and time for the cookie.
Public Property Get ExpiryUnix() As Long
    ExpiryUnix = mExpiry
End Property
Public Property Let ExpiryUnix(ByVal val As Long)
    mExpiry = val
End Property

Public Property Get ExpiryDate() As Date
    ExpiryDate = DateAdd("s", mExpiry, "1/1/1970 00:00:00")
End Property
Public Property Let ExpiryDate(ByVal val As Date)
    mExpiry = DateDiff("s", "1/1/1970 00:00:00", val)
End Property

'this field indicates whether the cookie has to be used only over HTTP or not.
'true if the cookie has the HttpOnly attribute and cannot be accessed through a client-side script; otherwise, false. The default is false.
Public Property Get HttpOnly() As Boolean
    HttpOnly = mHttpOnly
End Property
Public Property Let HttpOnly(ByVal val As Boolean)
    mHttpOnly = val
End Property

Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(ByVal val As String)
    mName = val
End Property

'This field indicates that the cookie can only be sent to the server over a secure HTTPS connection (set as true).
'Gets or sets a value indicating whether to transmit the cookie using Secure Sockets Layer (SSL)--that is, over HTTPS only.
'true to transmit the cookie over an SSL connection (HTTPS); otherwise, false. The default value is false
Public Property Get Secure() As Boolean
    Secure = mSecure
End Property
Public Property Let Secure(ByVal val As Boolean)
    mSecure = val
End Property

Public Property Get Value() As String
    Value = mValue
End Property
Public Property Let Value(ByVal val As String)
    mValue = val
End Property

'The virtual path to transmit with the cookie. The default is /, which is the server root.
Public Property Get Path() As String
    Path = mPath
End Property
Public Property Let Path(ByVal val As String)
    mPath = val
End Property

Public Function ToDictionary() As Dictionary
    Dim ck As New Dictionary
    ck.Add "name", mName
    ck.Add "value", mValue
    If mDomain <> "" Then ck.Add "domain", mDomain
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
