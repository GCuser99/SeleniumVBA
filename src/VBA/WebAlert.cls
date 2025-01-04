VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used to manage browser alerts - must be instantiated using the SwitchToAlert method of the WebDriver class"
'@ModuleDescription "This class is used to manage browser alerts - must be instantiated using the SwitchToAlert method of the WebDriver class"
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
    CMD_DISMISS_ALERT As Variant
    CMD_ACCEPT_ALERT As Variant
    CMD_SET_ALERT_TEXT As Variant
    CMD_GET_ALERT_TEXT As Variant
End Type

Private tCMD As typCMD

Private driver_ As WebDriver

'@Description("Returns the prompt text of the active alert")
Public Function GetText() As String
Attribute GetText.VB_Description = "Returns the prompt text of the active alert"
    GetText = driver_.execute(tCMD.CMD_GET_ALERT_TEXT)("value")
End Function

'@Description("Sends the specified text to the active prompt-type alert - Selenium special keys not recognized for alerts")
Public Function SendKeys(ByVal alertText As String) As WebAlert
Attribute SendKeys.VB_Description = "Sends the specified text to the active prompt-type alert - Selenium special keys not recognized for alerts"
    Dim data As New Dictionary
    If alertText Like "*\u[eE]0[0-5][0-9a-fA-F]*" Then
        Err.Raise 1, , "Error in alert SendKeys method: special keys not recognized in the alertText argument."
    End If
    data.Add "text", alertText
    driver_.execute tCMD.CMD_SET_ALERT_TEXT, data
    Set SendKeys = Me
End Function

'@Description("Accepts the active alert")
Public Sub Accept()
Attribute Accept.VB_Description = "Accepts the active alert"
    driver_.execute tCMD.CMD_ACCEPT_ALERT
End Sub

'@Description("Dismisses the active alert")
Public Sub Dismiss()
Attribute Dismiss.VB_Description = "Dismisses the active alert"
    driver_.execute tCMD.CMD_DISMISS_ALERT
End Sub

Friend Property Set WebDriver(ByRef obj As WebDriver)
    Set driver_ = obj
End Property

Private Sub Class_Initialize()
    With tCMD
        .CMD_DISMISS_ALERT = Array("POST", "/session/$sessionId/alert/dismiss")
        .CMD_ACCEPT_ALERT = Array("POST", "/session/$sessionId/alert/accept")
        .CMD_SET_ALERT_TEXT = Array("POST", "/session/$sessionId/alert/text")
        .CMD_GET_ALERT_TEXT = Array("GET", "/session/$sessionId/alert/text")
    End With
End Sub
