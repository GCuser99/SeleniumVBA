VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================
' SeleniumVBA v0.0.6
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
' TinySeleniumVBA v0.1.3
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

Private driver_ As WebDriver
Private elementId_ As String

Friend Property Set driver(objWebDriver As WebDriver)
    Set driver_ = objWebDriver
End Property

Public Property Let elementId(ByVal val)
    elementId_ = val
End Property

Public Property Get elementId()
    elementId = elementId_
End Property

Public Sub SaveScreenshot(ByVal filename As String)
    driver_.SaveScreenshot filename, Me
End Sub

Public Function FindElement(by_ As by, ByVal value As String) As WebElement
    Set FindElement = driver_.FindElement(by_, value, Me)
End Function

Public Function FindElements(by_ As by, ByVal value As String) As WebElement()
    FindElements = driver_.FindElements(by_, value, Me)
End Function

Public Sub SetAttribute(ByVal name As String, ByVal value As String)
    driver_.SetAttribute Me, name, value
End Sub

Public Sub SendKeys(ByVal keys As String, Optional ByVal clearBeforeTyping As Boolean = True)
    driver_.SendKeys Me, keys, clearBeforeTyping
End Sub

Public Sub Click()
    driver_.Click Me
End Sub

Public Sub Submit()
    driver_.Submit Me
End Sub

Public Function GetText() As String
    GetText = driver_.GetText(Me)
End Function

Public Function GetProperty(ByVal name As String) As String
    GetProperty = driver_.GetProperty(name, Me)
End Function

Public Function HasProperty(ByVal name As String) As Boolean
    HasProperty = driver_.HasProperty(name, Me)
End Function

Public Function GetTagName() As String
    GetTagName = driver_.GetTagName(Me)
End Function

Public Function GetRect() As Dictionary
    Set GetRect = driver_.GetRect(Me)
End Function

Public Function GetCSSProperty(ByVal name As String) As String
    GetCSSProperty = driver_.GetCSSProperty(name, Me)
End Function

Public Function HasAttribute(ByVal name As String) As Boolean
    HasAttribute = driver_.HasAttribute(name, Me)
End Function

Public Function GetAttribute(ByVal name As String) As String
    GetAttribute = driver_.GetAttribute(name, Me)
End Function

Public Function GetAriaLabel() As String
    GetAriaLabel = driver_.GetAriaLabel(Me)
End Function

Public Function GetAriaRole() As String
    GetAriaRole = driver_.GetAriaRole(Me)
End Function

Public Function IsMultiSelect() As Boolean
    IsMultiSelect = driver_.IsMultiSelect(Me)
End Function

Public Function IsSelected() As Boolean
    IsSelected = driver_.IsSelected(Me)
End Function

Public Function SelectByValue(ByVal value As String)
    driver_.SelectByValue value, Me
End Function

Public Function SelectByVisibleText(ByVal Text As String)
    driver_.SelectByVisibleText Text, Me
End Function

Public Function SelectByIndex(ByVal index As Integer)
    driver_.SelectByIndex index, Me
End Function

Public Function GetSelectTagsOptions() As String()
    GetSelectTagsOptions = driver_.GetSelectTagsOptions(Me)
End Function

Public Function GetAllSelectedOptionsText() As String()
    GetAllSelectedOptionsText = driver_.GetAllSelectedOptionsText(Me)
End Function

Public Function GetSelectedOptionText() As String
    GetSelectedOptionText = driver_.GetSelectedOptionText(Me)
End Function

Public Function SelectAll() As String()
    driver_.SelectAll Me
End Function

Public Function DeSelectAll() As String()
    driver_.DeSelectAll Me
End Function

Public Function DeSelectByValue(ByVal value As String)
    driver_.DeSelectByValue value, Me
End Function

Public Function DeSelectByVisibleText(ByVal Text As String)
    driver_.DeSelectByVisibleText Text, Me
End Function

Public Function DeSelectByIndex(ByVal index As Integer)
    driver_.DeSelectByIndex index, Me
End Function

Public Function SwitchToFrame()
    driver_.SwitchToFrame Me
End Function

Public Function IsEnabled() As Boolean
    IsEnabled = driver_.IsEnabled(Me)
End Function

Public Function IsDisplayed() As Boolean
    IsDisplayed = driver_.IsDisplayed(Me)
End Function

Public Sub Clear()
    driver_.Clear Me
End Sub

Public Function GetInnerHTML() As String
    GetInnerHTML = driver_.GetInnerHTML(Me)
End Function

Public Function GetOuterHTML() As String
    GetOuterHTML = driver_.GetOuterHTML(Me)
End Function

Public Sub SetFocus()
    driver_.SetFocus Me
End Sub

Public Sub SetFocusNoScroll()
    driver_.SetFocusNoScroll Me
End Sub

Public Function ScrollToElement(Optional ByVal xOffset As Integer = 0, Optional ByVal yOffset As Integer = 0)
    driver_.ScrollToElement Me, xOffset, yOffset
End Function

Public Sub ScrollIntoView(Optional ByVal alignTop As Boolean = True)
    driver_.ScrollIntoView Me, alignTop
End Sub

Public Function GetShadowRoot() As ShadowRoot
    Set GetShadowRoot = driver_.GetShadowRoot(Me)
End Function

Public Sub DragAndDrop(target As WebElement)
    driver_.DragAndDrop Me, target
End Sub

Public Sub UploadFile(ByVal filepath As String)
    driver_.UploadFile Me, filepath
End Sub

Private Sub Class_Terminate()
    Set driver_ = Nothing
End Sub
