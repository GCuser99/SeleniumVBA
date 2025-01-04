VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class provides a mechanism to represent webpage HTML elements as objects and perform various actions on the related elements - must be instantiated using one of WebDriver's FindElement methods"
'@ModuleDescription "This class provides a mechanism to represent webpage HTML elements as objects and perform various actions on the related elements - must be instantiated using one of WebDriver's FindElement methods"
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
'
' Based originally on uezo's TinySeleniumVBA with extensive changes/additions
' ==========================================================================
' MIT License
'
' Copyright (c) 2023-2024, GCUser99 and 6DiegoDiego9 (https://github.com/GCuser99/SeleniumVBA)
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
    Set FindElementByTagName = driver_.FindElement(By.TagName, val, Me)
End Function

'@Description("Finds the first element by the specified class name value")
Public Function FindElementByClassName(ByVal val As String) As WebElement
Attribute FindElementByClassName.VB_Description = "Finds the first element by the specified class name value"
    Set FindElementByClassName = driver_.FindElement(By.ClassName, val, Me)
End Function

'@Description("Finds the first element by the specified css selector value")
Public Function FindElementByCssSelector(ByVal val As String) As WebElement
Attribute FindElementByCssSelector.VB_Description = "Finds the first element by the specified css selector value"
    Set FindElementByCssSelector = driver_.FindElement(By.CssSelector, val, Me)
End Function

'@Description("Finds the first element by the specified css selector value")
Public Function QuerySelector(ByVal val As String) As WebElement
Attribute QuerySelector.VB_Description = "Finds the first element by the specified css selector value"
    Set QuerySelector = driver_.FindElement(By.CssSelector, val, Me)
End Function

'@Description("Finds the first element by the specified link text value")
Public Function FindElementByLinkText(ByVal val As String) As WebElement
Attribute FindElementByLinkText.VB_Description = "Finds the first element by the specified link text value"
    Set FindElementByLinkText = driver_.FindElement(By.LinkText, val, Me)
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
    Set FindElementsByTagName = driver_.FindElements(By.TagName, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified class name value")
Public Function FindElementsByClassName(ByVal val As String) As WebElements
Attribute FindElementsByClassName.VB_Description = "Finds all elements that satisfy the specified class name value"
    Set FindElementsByClassName = driver_.FindElements(By.ClassName, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified css selector value")
Public Function FindElementsByCssSelector(ByVal val As String) As WebElements
Attribute FindElementsByCssSelector.VB_Description = "Finds all elements that satisfy the specified css selector value"
    Set FindElementsByCssSelector = driver_.FindElements(By.CssSelector, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified css selector value")
Public Function QuerySelectorAll(ByVal val As String) As WebElements
Attribute QuerySelectorAll.VB_Description = "Finds all elements that satisfy the specified css selector value"
    Set QuerySelectorAll = driver_.FindElements(By.CssSelector, val, Me)
End Function

'@Description("Finds all elements that satisfy the specified link text value")
Public Function FindElementsByLinkText(ByVal val As String) As WebElements
Attribute FindElementsByLinkText.VB_Description = "Finds all elements that satisfy the specified link text value"
    Set FindElementsByLinkText = driver_.FindElements(By.LinkText, val, Me)
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

'@Description("Returns a boolean indicating if the element is present. User can optionally set an implicit wait, and obtain a reference to the found element.")
Public Function IsPresent(by_ As By, ByVal val As String, Optional ByVal ImplicitWaitMS As Long = 0, _
                            Optional ByRef elemFound As WebElement) As Boolean
Attribute IsPresent.VB_Description = "Returns a boolean indicating if the element is present. User can optionally set an implicit wait, and obtain a reference to the found element."
    IsPresent = driver_.IsPresent(by_, val, ImplicitWaitMS, Me, elemFound)
End Function

'@Description("Sets the attribute of this element")
Public Sub SetAttribute(ByVal attrName As String, ByVal val As String)
Attribute SetAttribute.VB_Description = "Sets the attribute of this element"
    driver_.SetAttribute Me, attrName, val
End Sub

'@Description("Sends a key sequence to this element")
Public Sub SendKeys(ByVal keys As String, Optional ByVal clearBeforeTyping As Boolean = False)
Attribute SendKeys.VB_Description = "Sends a key sequence to this element"
    driver_.SendKeys Me, keys, clearBeforeTyping
End Sub

'@Description("Clicks this element")
Public Sub Click(Optional ByVal modifierKeys As String)
Attribute Click.VB_Description = "Clicks this element"
    driver_.Click Me, modifierKeys
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

'@Description("Returns a dictionary containing x, y, width, height, left, top, right, and bottom of the element's rectangle")
Public Function GetRect() As Dictionary
Attribute GetRect.VB_Description = "Returns a dictionary containing x, y, width, height, left, top, right, and bottom of the element's rectangle"
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
Public Function IsSelected(Optional ByVal maxWaitTimeMS As Long = 0) As Boolean
Attribute IsSelected.VB_Description = "Returns a boolean indicating if this dropdown option element is selected"
    IsSelected = driver_.IsSelected(Me, maxWaitTimeMS)
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
Public Sub SelectByIndex(ByVal index As Long)
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
Public Sub DeSelectByIndex(ByVal index As Long)
Attribute DeSelectByIndex.VB_Description = "Deselects a dropdown option element with the specified index"
    driver_.DeSelectByIndex index, Me
End Sub

'@Description("Switches to this frame element")
Public Sub SwitchToFrame()
Attribute SwitchToFrame.VB_Description = "Switches to this frame element"
    driver_.SwitchToFrame Me
End Sub

'@Description("Returns a boolean indicating if this element is enabled")
Public Function IsEnabled(Optional ByVal maxWaitTimeMS As Long = 0) As Boolean
Attribute IsEnabled.VB_Description = "Returns a boolean indicating if this element is enabled"
    IsEnabled = driver_.IsEnabled(Me, maxWaitTimeMS)
End Function

'@Description("Returns a boolean indicating if this element is displayed")
Public Function IsDisplayed(Optional ByVal maxWaitTimeMS As Long = 0) As Boolean
Attribute IsDisplayed.VB_Description = "Returns a boolean indicating if this element is displayed"
    IsDisplayed = driver_.IsDisplayed(Me, maxWaitTimeMS)
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
Public Function SetFocus(Optional ByVal preventScroll As Boolean = False, Optional ByVal focusVisible As Boolean = False) As WebElement
Attribute SetFocus.VB_Description = "Scrolls and set the focus to this element"
    Set SetFocus = driver_.SetFocus(Me, preventScroll, focusVisible)
End Function

'@Description("Returns a boolean indicating if the element has the focus")
Public Function HasFocus() As Boolean
Attribute HasFocus.VB_Description = "Returns a boolean indicating if the element has the focus"
    HasFocus = driver_.HasFocus(Me)
End Function

'@Description("Scrolls this element into the visible viewport")
Public Function ScrollIntoView(Optional ByVal enSpeed As ScrollIntoViewOptions_speed = jump_auto, _
                               Optional ByVal enAlign_vert As ScrollIntoViewOptions_alignment = align_start, _
                               Optional ByVal enAlign_horiz As ScrollIntoViewOptions_alignment = align_nearest, _
                               Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As WebElement
Attribute ScrollIntoView.VB_Description = "Scrolls this element into the visible viewport"
    Set ScrollIntoView = driver_.ScrollIntoView(Me, enSpeed, enAlign_vert, enAlign_horiz, xOffset, yOffset)
End Function

'@Description("Scrolls by an x-y offset in pixels relative to this scrollable element's content position")
Public Sub ScrollBy(Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0, Optional ByVal enSpeed As ScrollIntoViewOptions_speed = jump_auto)
Attribute ScrollBy.VB_Description = "Scrolls by an x-y offset in pixels relative to this scrollable element's content position"
    driver_.ScrollBy xOffset, yOffset, enSpeed, Me
End Sub

'@Description("Scrolls to an x-y position in pixels relative to this scrollable element's content origin")
Public Sub ScrollTo(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal enSpeed As ScrollIntoViewOptions_speed = jump_auto)
Attribute ScrollTo.VB_Description = "Scrolls to an x-y position in pixels relative to this scrollable element's content origin"
    driver_.ScrollTo x, y, enSpeed, Me
End Sub

'@Description("Scrolls to the top of this scrollable element's content")
Public Sub ScrollToTop(Optional ByVal enSpeed As ScrollIntoViewOptions_speed = jump_auto)
Attribute ScrollToTop.VB_Description = "Scrolls to the top of this scrollable element's content"
    driver_.ScrollToTop enSpeed, Me
End Sub

'@Description("Scrolls to the bottom of this scrollable element's content")
Public Sub ScrollToBottom(Optional ByVal enSpeed As ScrollIntoViewOptions_speed = jump_auto)
Attribute ScrollToBottom.VB_Description = "Scrolls to the bottom of this scrollable element's content"
    driver_.ScrollToBottom enSpeed, Me
End Sub

'@Description("Gets the height of this scrollable element's content")
Public Function GetScrollHeight() As Long
Attribute GetScrollHeight.VB_Description = "Gets the height of this scrollable element's content"
    GetScrollHeight = driver_.GetScrollHeight(Me)
End Function

'@Description("Gets the width of of this scrollable element's content")
Public Function GetScrollWidth() As Long
Attribute GetScrollWidth.VB_Description = "Gets the width of of this scrollable element's content"
    GetScrollWidth = driver_.GetScrollWidth(Me)
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

'@Description("Waits until this element is displayed (interactable)")
Public Function WaitUntilDisplayed(Optional ByVal maxWaitTimeMS As Long = 10000) As WebElement
Attribute WaitUntilDisplayed.VB_Description = "Waits until this element is displayed (interactable)"
    Set WaitUntilDisplayed = driver_.WaitUntilDisplayed(Me, maxWaitTimeMS)
End Function

'@Description("Waits until this element is not displayed (not interactable)")
Public Function WaitUntilNotDisplayed(Optional ByVal maxWaitTimeMS As Long = 10000) As WebElement
Attribute WaitUntilNotDisplayed.VB_Description = "Waits until this element is not displayed (not interactable)"
    Set WaitUntilNotDisplayed = driver_.WaitUntilNotDisplayed(Me, maxWaitTimeMS)
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

'@Description("Returns a 1-based array holding the text values from a table or tbody element")
Public Function TableToArray(Optional ByVal skipHeader As Boolean = False, Optional ByVal skipFooter As Boolean = False, Optional ByVal createSpanData As Boolean = True, Optional ByVal ignoreCellFormatting As Boolean = False) As Variant()
Attribute TableToArray.VB_Description = "Returns a 1-based array holding the text values from a table or tbody element"
    TableToArray = driver_.TableToArray(Me, skipHeader, skipFooter, createSpanData, ignoreCellFormatting)
End Function

'@Description("Downloads a resource (such as an image) contained in an element attribute to a file or folder path - if a folder is specified, then the file will inherit the name of the source")
Public Sub DownloadResource(Optional ByVal srcAttribute As String = "src", Optional ByVal fileOrFolderPath As String = ".\")
Attribute DownloadResource.VB_Description = "Downloads a resource (such as an image) contained in an element attribute to a file or folder path - if a folder is specified, then the file will inherit the name of the source"
    driver_.DownloadResource Me, srcAttribute, fileOrFolderPath
End Sub

'@Description("Returns the Element Id string - useful for testing purposes")
Public Function GetElementId() As String
Attribute GetElementId.VB_Description = "Returns the Element Id string - useful for testing purposes"
    GetElementId = driver_.GetElementId(Me)
End Function

Private Sub Class_Terminate()
    Set driver_ = Nothing
End Sub
