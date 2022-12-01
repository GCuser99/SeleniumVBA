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
' SeleniumVBA v2.6
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

Friend Property Let elementId(ByVal val As String)
    elementId_ = val
End Property

Friend Property Get elementId() As String
    elementId = elementId_
End Property

Public Sub SaveScreenshot(Optional ByVal filePath As String = ".\screenshot.png")
    driver_.SaveScreenshot filePath, Me
End Sub

Public Function FindElement(by_ As By, ByVal val As String) As WebElement
    Set FindElement = driver_.FindElement(by_, val, Me)
End Function

Public Function FindElements(by_ As By, ByVal val As String) As WebElements
    Set FindElements = driver_.FindElements(by_, val, Me)
End Function

Public Sub SetAttribute(ByVal Name As String, ByVal val As String)
    driver_.SetAttribute Me, Name, val
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

Public Function GetProperty(ByVal Name As String) As String
    GetProperty = driver_.GetProperty(Name, Me)
End Function

Public Function HasProperty(ByVal Name As String) As Boolean
    HasProperty = driver_.HasProperty(Name, Me)
End Function

Public Function GetTagName() As String
    GetTagName = driver_.GetTagName(Me)
End Function

Public Function GetRect() As Dictionary
    Set GetRect = driver_.GetRect(Me)
End Function

Public Function GetCSSProperty(ByVal Name As String) As String
    GetCSSProperty = driver_.GetCSSProperty(Name, Me)
End Function

Public Function HasAttribute(ByVal Name As String) As Boolean
    HasAttribute = driver_.HasAttribute(Name, Me)
End Function

Public Function GetAttribute(ByVal Name As String) As String
    GetAttribute = driver_.GetAttribute(Name, Me)
End Function

Public Function GetAriaLabel() As String
    GetAriaLabel = driver_.GetAriaLabel(Me)
End Function

Public Function GetAriaRole() As String
    GetAriaRole = driver_.GetAriaRole(Me)
End Function

Public Function Highlight(Optional borderSizePx As Byte = 4, Optional borderColor As VBAcolors = Blue, _
                     Optional backgroundColor As VBAcolors = VBAcolors.Unchanged, _
                     Optional ScrollIntoView As Boolean = True, Optional UnHighlightLast As Boolean = True) _
                     As WebElement
    driver_.SetHighlight True, Me, borderSizePx, borderColor, backgroundColor, ScrollIntoView, UnHighlightLast
    Set Highlight = Me
End Function

Public Sub UnHighlight()
    driver_.SetHighlight False, Me
End Sub

Public Function IsMultiSelect() As Boolean
    IsMultiSelect = driver_.IsMultiSelect(Me)
End Function

Public Function IsSelected() As Boolean
    IsSelected = driver_.IsSelected(Me)
End Function

Public Sub SelectByValue(ByVal val As String)
    driver_.SelectByValue val, Me
End Sub

Public Sub SelectByVisibleText(ByVal text As String)
    driver_.SelectByVisibleText text, Me
End Sub

Public Sub SelectByIndex(ByVal index As Integer)
    driver_.SelectByIndex index, Me
End Sub

Public Function GetSelectTagsOptions() As String()
    GetSelectTagsOptions = driver_.GetSelectTagsOptions(Me)
End Function

Public Function GetAllSelectedOptionsText() As String()
    GetAllSelectedOptionsText = driver_.GetAllSelectedOptionsText(Me)
End Function

Public Function GetSelectedOptionText() As String
    GetSelectedOptionText = driver_.GetSelectedOptionText(Me)
End Function

Public Sub SelectAll()
    driver_.SelectAll Me
End Sub

Public Sub DeSelectAll()
    driver_.DeSelectAll Me
End Sub

Public Sub DeSelectByValue(ByVal val As String)
    driver_.DeSelectByValue val, Me
End Sub

Public Sub DeSelectByVisibleText(ByVal text As String)
    driver_.DeSelectByVisibleText text, Me
End Sub

Public Sub DeSelectByIndex(ByVal index As Integer)
    driver_.DeSelectByIndex index, Me
End Sub

Public Sub SwitchToFrame()
    driver_.SwitchToFrame Me
End Sub

Public Function IsEnabled() As Boolean
    IsEnabled = driver_.IsEnabled(Me)
End Function

Public Function IsDisplayed() As Boolean
    IsDisplayed = driver_.IsDisplayed(Me)
End Function

Public Function Clear() As WebElement
    Set Clear = driver_.Clear(Me)
End Function

Public Function GetInnerHTML() As String
    GetInnerHTML = driver_.GetInnerHTML(Me)
End Function

Public Function GetOuterHTML() As String
    GetOuterHTML = driver_.GetOuterHTML(Me)
End Function

Public Function SetFocus() As WebElement
    Set SetFocus = driver_.SetFocus(Me)
End Function

Public Function SetFocusNoScroll() As WebElement
    Set SetFocusNoScroll = driver_.SetFocusNoScroll(Me)
End Function

Public Function ScrollToElement(Optional ByVal xOffset As Integer = 0, Optional ByVal yOffset As Integer = 0) As WebElement
    Set ScrollToElement = driver_.ScrollToElement(Me, xOffset, yOffset)
End Function

Public Function ScrollIntoView(Optional ByVal alignTop As Boolean = True) As WebElement
    Set ScrollIntoView = driver_.ScrollIntoView(Me, alignTop)
End Function

Public Function GetShadowRoot() As WebShadowRoot
    Set GetShadowRoot = driver_.GetShadowRoot(Me)
End Function

Public Sub DragAndDrop(target As WebElement)
    driver_.DragAndDrop Me, target
End Sub

Public Sub UploadFile(ByVal filePath As String)
    driver_.UploadFile Me, filePath
End Sub

Public Function WaitUntilReady(Optional ByVal maxWaitTimeMS = 30000) As WebElement
    Set WaitUntilReady = driver_.WaitUntilReady(Me, maxWaitTimeMS)
End Function

Public Function GetParentElement() As WebElement
    Set GetParentElement = driver_.GetParentElement(Me)
End Function

Public Function GetChildElements() As WebElements
    Set GetChildElements = driver_.GetChildElements(Me)
End Function

Public Function TableToArray(Optional ByVal skipHeader As Boolean = False, Optional ByVal skipFooter As Boolean = False, Optional ByVal createSpanData As Boolean = True) As Variant()
    TableToArray = driver_.TableToArray(Me, skipHeader, skipFooter, createSpanData)
End Function

Private Sub Class_Terminate()
    Set driver_ = Nothing
End Sub
