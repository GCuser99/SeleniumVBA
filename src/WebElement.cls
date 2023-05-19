VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class provides a mechanism to represent webpage HTML elements as objects and perform various actions on the related elements"
'@ModuleDescription "This class provides a mechanism to represent webpage HTML elements as objects and perform various actions on the related elements"
'@Exposed
'@folder("SeleniumVBA.Source")
' ==========================================================================
' SeleniumVBA v4.1
'
' A Selenium wrapper for browser automation developed for MS Office VBA
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' Contact Info:
'
' https://github.com/6DiegoDiego9
' https://github.com/GCUser99
'
' Based originally on uezo's TinySeleniumVBA with extensive changes/additions
' ==========================================================================
' MIT License
'
' Copyright (c) 2023, GCUser99 and 6DiegoDiego9 (https://github.com/GCuser99/SeleniumVBA)
' Copyright (c) 2021, uezo (https://github.com/uezo/TinySeleniumVBA)
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

'@Description("Saves a screenshot of this element")
Public Sub SaveScreenshot(Optional ByVal filePath As String = ".\screenshot.png")
Attribute SaveScreenshot.VB_Description = "Saves a screenshot of this element"
    driver_.SaveScreenshot filePath, Me
End Sub

'@Description("Finds the first element for a given selector and value")
Public Function FindElement(by_ As By, ByVal val As String) As WebElement
Attribute FindElement.VB_Description = "Finds the first element for a given selector and value"
    Set FindElement = driver_.FindElement(by_, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified selector and value")
Public Function FindElements(by_ As By, ByVal val As String) As WebElements
Attribute FindElements.VB_Description = "Finds all elements that satisfy the specified selector and value"
    Set FindElements = driver_.FindElements(by_, val, Me)
End Function

'@Description("Finds the first element by the specified id value")
Public Function FindElementByID(ByVal val As String) As WebElement
Attribute FindElementByID.VB_Description = "Finds the first element by the specified id value"
    Set FindElementByID = driver_.FindElement(By.ID, val, Me)
End Function

'@Description("Finds the first element by the specified tag name value")
Public Function FindElementByTagName(ByVal val As String) As WebElement
Attribute FindElementByTagName.VB_Description = "Finds the first element by the specified tag name value"
    Set FindElementByTagName = driver_.FindElement(By.tagName, val, Me)
End Function

'@Description("Finds the first element by the specified class name value")
Public Function FindElementByClassName(ByVal val As String) As WebElement
Attribute FindElementByClassName.VB_Description = "Finds the first element by the specified class name value"
    Set FindElementByClassName = driver_.FindElement(By.className, val, Me)
End Function

'@Description("Finds the first element by the specified css selector value")
Public Function FindElementByCssSelector(ByVal val As String) As WebElement
Attribute FindElementByCssSelector.VB_Description = "Finds the first element by the specified css selector value"
    Set FindElementByCssSelector = driver_.FindElement(By.cssSelector, val, Me)
End Function

'@Description("Finds the first element by the specified link text value")
Public Function FindElementByLinkText(ByVal val As String) As WebElement
Attribute FindElementByLinkText.VB_Description = "Finds the first element by the specified link text value"
    Set FindElementByLinkText = driver_.FindElement(By.linkText, val, Me)
End Function

'@Description("Finds the first element by the specified name value")
Public Function FindElementByName(ByVal val As String) As WebElement
Attribute FindElementByName.VB_Description = "Finds the first element by the specified name value"
    Set FindElementByName = driver_.FindElement(By.Name, val, Me)
End Function

'@Description("Finds the first element by the specified partial link text value")
Public Function FindElementByPartialLinkText(ByVal val As String) As WebElement
Attribute FindElementByPartialLinkText.VB_Description = "Finds the first element by the specified partial link text value"
    Set FindElementByPartialLinkText = driver_.FindElement(By.PartialLinkText, val, Me)
End Function

'@Description("Finds the first element by the specified xpath query value")
Public Function FindElementByXPath(ByVal val As String) As WebElement
Attribute FindElementByXPath.VB_Description = "Finds the first element by the specified xpath query value"
    Set FindElementByXPath = driver_.FindElement(By.XPath, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified xpath query value")
Public Function FindElementsByXPath(ByVal val As String) As WebElements
Attribute FindElementsByXPath.VB_Description = "Finds all elements that satisfy the specified xpath query value"
    Set FindElementsByXPath = driver_.FindElements(By.XPath, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified id value")
Public Function FindElementsByID(ByVal val As String) As WebElements
Attribute FindElementsByID.VB_Description = "Finds all elements that satisfy the specified id value"
    Set FindElementsByID = driver_.FindElements(By.ID, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified tag name value")
Public Function FindElementsByTagName(ByVal val As String) As WebElements
Attribute FindElementsByTagName.VB_Description = "Finds all elements that satisfy the specified tag name value"
    Set FindElementsByTagName = driver_.FindElements(By.tagName, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified class name value")
Public Function FindElementsByClassName(ByVal val As String) As WebElements
Attribute FindElementsByClassName.VB_Description = "Finds all elements that satisfy the specified class name value"
    Set FindElementsByClassName = driver_.FindElements(By.className, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified css selector value")
Public Function FindElementsByCssSelector(ByVal val As String) As WebElements
Attribute FindElementsByCssSelector.VB_Description = "Finds all elements that satisfy the specified css selector value"
    Set FindElementsByCssSelector = driver_.FindElements(By.cssSelector, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified link text value")
Public Function FindElementsByLinkText(ByVal val As String) As WebElements
Attribute FindElementsByLinkText.VB_Description = "Finds all elements that satisfy the specified link text value"
    Set FindElementsByLinkText = driver_.FindElements(By.linkText, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified name value")
Public Function FindElementsByName(ByVal val As String) As WebElements
Attribute FindElementsByName.VB_Description = "Finds all elements that satisfy the specified name value"
    Set FindElementsByName = driver_.FindElements(By.Name, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified partial link text value")
Public Function FindElementsByPartialLinkText(ByVal val As String) As WebElements
Attribute FindElementsByPartialLinkText.VB_Description = "Finds all elements that satisfy the specified partial link text value"
    Set FindElementsByPartialLinkText = driver_.FindElements(By.PartialLinkText, val, Me)
End Function

'@Description("Returns a boolean indicating if the element searched for by the selector is present")
Public Function IsPresent(by_ As By, ByVal val As String, Optional ByVal ImplicitWaitMS As Long = 0, _
                            Optional ByRef elemFound As WebElement) As Boolean
Attribute IsPresent.VB_Description = "Returns a boolean indicating if the element searched for by the selector is present"
    IsPresent = driver_.IsPresent(by_, val, ImplicitWaitMS, elemFound, Me)
End Function

'@Description("Sets the attribute of this element")
Public Sub SetAttribute(ByVal attrName As String, ByVal val As String)
Attribute SetAttribute.VB_Description = "Sets the attribute of this element"
    driver_.SetAttribute Me, attrName, val
End Sub

'@Description("Sends a key sequence to this element")
Public Sub SendKeys(ByVal keys As String, Optional ByVal clearBeforeTyping As Boolean = True)
Attribute SendKeys.VB_Description = "Sends a key sequence to this element"
    driver_.SendKeys Me, keys, clearBeforeTyping
End Sub

'@Description("Clicks this element")
Public Sub Click()
Attribute Click.VB_Description = "Clicks this element"
    driver_.Click Me
End Sub

'@Description("Submits a form element")
Public Sub Submit()
Attribute Submit.VB_Description = "Submits a form element"
    driver_.Submit Me
End Sub

'@Description("Gets the text of this element - set optional visibleOnly to False to return text whether hidden or not")
Public Function GetText(Optional ByVal visibleOnly As Boolean = True) As String
Attribute GetText.VB_Description = "Gets the text of this element - set optional visibleOnly to False to return text whether hidden or not"
    GetText = driver_.GetText(Me, visibleOnly)
End Function

'@Description("Gets the property state of this element, such as the value of an input box - if the property does not exist, then it returns vbNullString")
Public Function GetProperty(ByVal propName As String) As String
Attribute GetProperty.VB_Description = "Gets the property state of this element, such as the value of an input box - if the property does not exist, then it returns vbNullString"
    GetProperty = driver_.GetProperty(propName, Me)
End Function

'@Description("Returns a boolean indicating if this element has the specified property")
Public Function HasProperty(ByVal propName As String) As Boolean
Attribute HasProperty.VB_Description = "Returns a boolean indicating if this element has the specified property"
    HasProperty = driver_.HasProperty(propName, Me)
End Function

'@Description("Gets the Tag name of this element")
Public Function GetTagName() As String
Attribute GetTagName.VB_Description = "Gets the Tag name of this element"
    GetTagName = driver_.GetTagName(Me)
End Function

'@Description("Returns a dictionary containing x, y, width, height of this element")
Public Function GetRect() As Dictionary
Attribute GetRect.VB_Description = "Returns a dictionary containing x, y, width, height of this element"
    Set GetRect = driver_.GetRect(Me)
End Function

'@Description("Gets the specified CSS property of this element")
Public Function GetCSSProperty(ByVal propName As String) As String
Attribute GetCSSProperty.VB_Description = "Gets the specified CSS property of this element"
    GetCSSProperty = driver_.GetCSSProperty(propName, Me)
End Function

'@Description("Returns a boolean indicating if this element has the specified attribute")
Public Function HasAttribute(ByVal attrName As String) As Boolean
Attribute HasAttribute.VB_Description = "Returns a boolean indicating if this element has the specified attribute"
    HasAttribute = driver_.HasAttribute(attrName, Me)
End Function

'@Description("Gets the specified attribute of this element")
Public Function GetAttribute(ByVal attrName As String) As String
Attribute GetAttribute.VB_Description = "Gets the specified attribute of this element"
    GetAttribute = driver_.GetAttribute(attrName, Me)
End Function

'@Description("Gets the Arial label of this element")
Public Function GetAriaLabel() As String
Attribute GetAriaLabel.VB_Description = "Gets the Arial label of this element"
    GetAriaLabel = driver_.GetAriaLabel(Me)
End Function

'@Description("Gets the Arial role of this element")
Public Function GetAriaRole() As String
Attribute GetAriaRole.VB_Description = "Gets the Arial role of this element"
    GetAriaRole = driver_.GetAriaRole(Me)
End Function

'@Description("Highlights with color this element")
Public Function Highlight(Optional ByVal borderSizePx As Byte = 4, Optional ByVal borderColor As VBAcolors = Blue, _
                     Optional ByVal backgroundColor As VBAcolors = VBAcolors.Unchanged, _
                     Optional ByVal bScrollIntoView As Boolean = True, Optional ByVal unHighlightLast As Boolean = True) _
                     As WebElement
Attribute Highlight.VB_Description = "Highlights with color this element"
    driver_.SetHighlight True, Me, borderSizePx, borderColor, backgroundColor, bScrollIntoView, unHighlightLast
    Set Highlight = Me
End Function

'@Description("Unhighlights this element")
Public Sub UnHighlight()
Attribute UnHighlight.VB_Description = "Unhighlights this element"
    driver_.SetHighlight False, Me
End Sub

'@Description("Returns a boolean indicating if this element is a multi-select dropdown")
Public Function IsMultiSelect() As Boolean
Attribute IsMultiSelect.VB_Description = "Returns a boolean indicating if this element is a multi-select dropdown"
    IsMultiSelect = driver_.IsMultiSelect(Me)
End Function

'@Description("Returns a boolean indicating if this dropdown option element is selected")
Public Function IsSelected() As Boolean
Attribute IsSelected.VB_Description = "Returns a boolean indicating if this dropdown option element is selected"
    IsSelected = driver_.IsSelected(Me)
End Function

'@Description("Selects a dropdown option element with the specified value attribute")
Public Sub SelectByValue(ByVal val As String)
Attribute SelectByValue.VB_Description = "Selects a dropdown option element with the specified value attribute"
    driver_.SelectByValue val, Me
End Sub

'@Description("Selects a dropdown option element with the specified visible text")
Public Sub SelectByVisibleText(ByVal text As String)
Attribute SelectByVisibleText.VB_Description = "Selects a dropdown option element with the specified visible text"
    driver_.SelectByVisibleText text, Me
End Sub

'@Description("Selects a dropdown option element with the specified index")
Public Sub SelectByIndex(ByVal index As Integer)
Attribute SelectByIndex.VB_Description = "Selects a dropdown option element with the specified index"
    driver_.SelectByIndex index, Me
End Sub

'@Description("Returns all selected dropdown option elements")
Public Function GetAllSelectedOptions() As WebElements
Attribute GetAllSelectedOptions.VB_Description = "Returns all selected dropdown option elements"
    Set GetAllSelectedOptions = driver_.GetAllSelectedOptions(Me)
End Function

'@Description("Returns the first selected dropdown option element")
Public Function GetSelectedOption() As WebElement
Attribute GetSelectedOption.VB_Description = "Returns the first selected dropdown option element"
    Set GetSelectedOption = driver_.GetSelectedOption(Me)
End Function

'@Description("Selects all multi-select dropdown option elements")
Public Sub SelectAll()
Attribute SelectAll.VB_Description = "Selects all multi-select dropdown option elements"
    driver_.SelectAll Me
End Sub

'@Description("Deselects all dropdown option elements")
Public Sub DeSelectAll()
Attribute DeSelectAll.VB_Description = "Deselects all dropdown option elements"
    driver_.DeSelectAll Me
End Sub

'@Description("Deselects the dropdown option element having the specified value attribute")
Public Sub DeSelectByValue(ByVal val As String)
Attribute DeSelectByValue.VB_Description = "Deselects the dropdown option element having the specified value attribute"
    driver_.DeSelectByValue val, Me
End Sub

'@Description("Deselects the dropdown option element having the specified visible text")
Public Sub DeSelectByVisibleText(ByVal text As String)
Attribute DeSelectByVisibleText.VB_Description = "Deselects the dropdown option element having the specified visible text"
    driver_.DeSelectByVisibleText text, Me
End Sub

'@Description("Deselects a dropdown option element with the specified index")
Public Sub DeSelectByIndex(ByVal index As Integer)
Attribute DeSelectByIndex.VB_Description = "Deselects a dropdown option element with the specified index"
    driver_.DeSelectByIndex index, Me
End Sub

'@Description("Switches to this frame element")
Public Sub SwitchToFrame()
Attribute SwitchToFrame.VB_Description = "Switches to this frame element"
    driver_.SwitchToFrame Me
End Sub

'@Description("Returns a boolean indicating if this element is enabled")
Public Function IsEnabled() As Boolean
Attribute IsEnabled.VB_Description = "Returns a boolean indicating if this element is enabled"
    IsEnabled = driver_.IsEnabled(Me)
End Function

'@Description("Returns a boolean indicating if this element is displayed")
Public Function IsDisplayed() As Boolean
Attribute IsDisplayed.VB_Description = "Returns a boolean indicating if this element is displayed"
    IsDisplayed = driver_.IsDisplayed(Me)
End Function

'@Description("Clears the text field of this element")
Public Function Clear() As WebElement
Attribute Clear.VB_Description = "Clears the text field of this element"
    Set Clear = driver_.Clear(Me)
End Function

'@Description("Gets the inner HTML of this element")
Public Function GetInnerHTML() As String
Attribute GetInnerHTML.VB_Description = "Gets the inner HTML of this element"
    GetInnerHTML = driver_.GetInnerHTML(Me)
End Function

'@Description("Gets the outer HTML of this element")
Public Function GetOuterHTML() As String
Attribute GetOuterHTML.VB_Description = "Gets the outer HTML of this element"
    GetOuterHTML = driver_.GetOuterHTML(Me)
End Function

'@Description("Scrolls and set the focus to this element")
Public Function SetFocus() As WebElement
Attribute SetFocus.VB_Description = "Scrolls and set the focus to this element"
    Set SetFocus = driver_.SetFocus(Me)
End Function

'@Description("Sets the focus to this element without scrolling")
Public Function SetFocusNoScroll() As WebElement
Attribute SetFocusNoScroll.VB_Description = "Sets the focus to this element without scrolling"
    Set SetFocusNoScroll = driver_.SetFocusNoScroll(Me)
End Function

'@Description("Scrolls by an XY offset referenced to location of this element")
Public Function ScrollToElement(Optional ByVal xOffset As Integer = 0, Optional ByVal yOffset As Integer = 0) As WebElement
Attribute ScrollToElement.VB_Description = "Scrolls by an XY offset referenced to location of this element"
    Set ScrollToElement = driver_.ScrollToElement(Me, xOffset, yOffset)
End Function

'@Description("Scrolls to this element")
Public Function ScrollIntoView(Optional enSpeed As ScrollIntoViewOptions_speed, _
                               Optional enAlign_vert As ScrollIntoViewOptions_alignment, _
                               Optional enAlign_horiz As ScrollIntoViewOptions_alignment) As WebElement
    Set ScrollIntoView = driver_.ScrollIntoView(Me, enSpeed, enAlign_vert, enAlign_horiz)
End Function

'@Description("Gets the shadow root hosted by this element")
Public Function GetShadowRoot() As WebShadowRoot
Attribute GetShadowRoot.VB_Description = "Gets the shadow root hosted by this element"
    Set GetShadowRoot = driver_.GetShadowRoot(Me)
End Function

'@Description("Drags and drops this element to the specified target element")
Public Sub DragAndDrop(targetElement As WebElement)
Attribute DragAndDrop.VB_Description = "Drags and drops this element to the specified target element"
    driver_.DragAndDrop Me, targetElement
End Sub

'@Description("Uploads the contents of a file to this element")
Public Sub UploadFile(ByVal filePath As String)
Attribute UploadFile.VB_Description = "Uploads the contents of a file to this element"
    driver_.UploadFile Me, filePath
End Sub

'@Description("Waits until this element is interactable (displayed)")
Public Function WaitUntilReady(Optional ByVal maxWaitTimeMS As Long = -1) As WebElement
Attribute WaitUntilReady.VB_Description = "Waits until this element is interactable (displayed)"
    Set WaitUntilReady = driver_.WaitUntilReady(Me, maxWaitTimeMS)
End Function

'@Description("Returns the parent of this element")
Public Function GetParentElement() As WebElement
Attribute GetParentElement.VB_Description = "Returns the parent of this element"
    Set GetParentElement = driver_.GetParentElement(Me)
End Function

'@Description("Returns the direct descendant child elements of this element")
Public Function GetChildElements() As WebElements
Attribute GetChildElements.VB_Description = "Returns the direct descendant child elements of this element"
    Set GetChildElements = driver_.GetChildElements(Me)
End Function

'@Description("Returns an array holding the text values from a table or tbody element")
Public Function TableToArray(Optional ByVal skipHeader As Boolean = False, Optional ByVal skipFooter As Boolean = False, Optional ByVal createSpanData As Boolean = True) As Variant()
Attribute TableToArray.VB_Description = "Returns an array holding the text values from a table or tbody element"
    TableToArray = driver_.TableToArray(Me, skipHeader, skipFooter, createSpanData)
End Function

'@Description("Downloads a resource (such as an image) contained in an element attribute to a file or folder path - if a folder is specified, then the file will inherit the name of the source")
Public Sub DownloadResource(ByVal srcAttribute As String, Optional ByVal fileOrFolderPath As String = ".\")
Attribute DownloadResource.VB_Description = "Downloads a resource (such as an image) contained in an element attribute to a file or folder path - if a folder is specified, then the file will inherit the name of the source"
    driver_.DownloadResource Me, srcAttribute, fileOrFolderPath
End Sub

Private Sub Class_Terminate()
    Set driver_ = Nothing
End Sub
