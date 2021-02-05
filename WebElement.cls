VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebElement"
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

Public Driver_ As WebDriver
Public SessionId_ As String
Public ElementId_ As String

' Find DOM element
Public Function FindElement(by_ As By, value As String) As WebElement
    Set FindElement = Driver_.FindElement(by_, value, ElementId_, SessionId_)
End Function

' Find multiple DOM elements
Public Function FindElements(by_ As By, value As String) As WebElement()
    FindElements = Driver_.FindElements(by_, value, ElementId_, SessionId_)
End Function

' Returns element.value
Public Function GetValue() As String
    GetValue = Driver_.GetValue(ElementId_, SessionId_)
End Function

' Set value to element
Public Sub SetValue(text As String)
    Driver_.SetValue ElementId_, text, SessionId_
End Sub

' Click
Public Sub Click()
    Driver_.Click ElementId_, SessionId_
End Sub

' Returns element.innerText
Public Function GetText() As String
    GetText = Driver_.GetText(ElementId_, SessionId_)
End Function

