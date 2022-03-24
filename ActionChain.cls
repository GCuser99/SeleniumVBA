VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionChain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================
' SeleniumVBA v0.0.2
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
'
'This class was written to emulate Selenium Webdriver ActionChains
'
'USAGE(S):
'
'Method 1:
'   Dim Actions as ActionChain
'   Set elememail = Driver.FindElement(by.name, "email")
'   Set Actions = Driver.ActionChain
'   Actions.MoveToElement(elememail).Click
'   Actions.SendKeys("my@email.com").Perform
'
'Method 2:
'   Set elememail = Driver.FindElement(by.name, "email")
'   Driver.ActionChain.MoveToElement(elememail).Click.SendKeys("my@email.com").Perform
'
'   Note that entire chain + perform must be on one line for Method 2 to work
'
'NOTES:
'
'   I have designed this version to clear actions that are already stored locally and on the remote end
'   after the Perform Method is executed.
'
'   Use Keyboard class for advanced SendKeys support
'
'METHOD                 DESCRIPTION
'------                 -----------
'Clear                  Clears the focused element's text field
'Click                  Clicks an element
'ClickAndHold           Holds down the left mouse button on an element
'ClickContext           Performs a context-click (right click) on an element
'CtrlKey                Sends a Control-[Key] sequence to focused element
'DoubleClick            Double-clicks an element
'DragAndDrop            Holds down the left mouse button on the source element,
'                       then moves to the target element and releases the mouse button
'DragAndDropByOffset    Holds down the left mouse button on the source element,
'                       then moves to the target offset and releases the mouse button
'KeyDown                Sends a key press only, without releasing it
'KeyUp                  Releases a modifier key
'MoveByOffset           Moves the mouse to an offset from current mouse position. If element is specified,
'                       then move the mouse by an offset relative to the top-left corner of the element.
'MoveToElement          Moves the mouse to the middle of an element.
'Perform                Executes the action chain and then resets the action queue on both the local and remote ends
'ReleaseButton          Releases a held mouse button on an element
'ScrollBy               Scrolls window by x and y offsets
'ScrollToElement        Scrolls to element position with optional x and y offsets
'SendKeys               Sends keys to current focused element (see Keyboard class for special keys support)
'ShiftKey               Sends a Shift-[Key] sequence to focused element
'Wait                   Pause all inputs for the specified duration in milliseconds

Private Const DEFAULT_PAUSE_DURATION = 250
Private driver_ As WebDriver
Private data_ As Dictionary

Private CMD_W3C_ACTIONS
Private CMD_W3C_CLEAR_ACTIONS

Private Enum ButtonType
    leftButton = 0
    middleButton = 1
    rightButton = 2
End Enum

'These are helper properties for passing webdriver and data dictionary references to each link of chain
Friend Property Set WebDriver(ByRef obj As WebDriver)
    Set driver_ = obj
End Property

Friend Property Set data(ByRef obj As Dictionary)
    Set data_ = obj
End Property

'Action Chain Methods:

Public Function MoveToElement(element As WebElement, Optional SetFocus As Boolean = False) As ActionChain
    'Moves the mouse pointer to the middle of an element.
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If SetFocus Then 'add click after move
        ReDim actions(0 To 2)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerDown
        Set actions(2) = ActionPointerUp
    Else
        ReDim actions(0 To 0)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set MoveToElement = New ActionChain
    Set MoveToElement.WebDriver = driver_
    Set MoveToElement.data = data_
    
End Function

Public Function Click(Optional element As WebElement) As ActionChain
    'clicks an element
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        ReDim actions(0 To 1)
        Set actions(0) = ActionPointerDown
        Set actions(1) = ActionPointerUp
    Else
        ReDim actions(0 To 2)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerDown
        Set actions(2) = ActionPointerUp
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set Click = New ActionChain
    Set Click.WebDriver = driver_
    Set Click.data = data_
    
End Function

Public Function ClickAndHold(Optional element As WebElement) As ActionChain
    'Holds down the left mouse button on an element
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        ReDim actions(0 To 0)
        Set actions(0) = ActionPointerDown
    Else
        ReDim actions(0 To 1)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerDown
    End If
        
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ClickAndHold = New ActionChain
    Set ClickAndHold.WebDriver = driver_
    Set ClickAndHold.data = data_
    
End Function

Public Function ClickContext(Optional element As WebElement) As ActionChain
    'Performs a context-click (right click) on an element
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        ReDim actions(0 To 1)
        Set actions(0) = ActionPointerDown(rightButton)
        Set actions(1) = ActionPointerUp(rightButton)
    Else
        ReDim actions(0 To 2)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerDown(rightButton)
        Set actions(2) = ActionPointerUp(rightButton)
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ClickContext = New ActionChain
    Set ClickContext.WebDriver = driver_
    Set ClickContext.data = data_
    
End Function

Public Function DoubleClick(Optional element As WebElement) As ActionChain
    'Double-clicks an element
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        ReDim actions(0 To 3)
        Set actions(0) = ActionPointerDown
        Set actions(1) = ActionPointerUp
        Set actions(2) = ActionPointerDown
        Set actions(3) = ActionPointerUp
    Else
        ReDim actions(0 To 4)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerDown
        Set actions(2) = ActionPointerUp
        Set actions(3) = ActionPointerDown
        Set actions(4) = ActionPointerUp
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DoubleClick = New ActionChain
    Set DoubleClick.WebDriver = driver_
    Set DoubleClick.data = data_
    
End Function

Public Function ReleaseButton(Optional element As WebElement) As ActionChain
    'Releasing a held mouse button
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        ReDim actions(0 To 0)
        Set actions(0) = ActionPointerUp
    Else
        ReDim actions(0 To 1)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerUp
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ReleaseButton = New ActionChain
    Set ReleaseButton.WebDriver = driver_
    Set ReleaseButton.data = data_
    
End Function

Public Function DragAndDrop(sourceElement As WebElement, targetElement As WebElement) As ActionChain
    'Holds down the left mouse button on the source element, then moves to the target element and releases the mouse button.
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    ReDim actions(0 To 3)
    Set actions(0) = ActionPointerMove(, , , , , sourceElement.elementId)
    Set actions(1) = ActionPointerDown
    Set actions(2) = ActionPointerMove(, , , , , targetElement.elementId)
    Set actions(3) = ActionPointerUp
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DragAndDrop = New ActionChain
    Set DragAndDrop.WebDriver = driver_
    Set DragAndDrop.data = data_
    
End Function

Public Function DragAndDropByOffset(element As WebElement, ByVal xOffset As Long, ByVal yOffset As Long) As ActionChain
    'Holds down the left mouse button on the source element, then moves to the x and y offsets
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    ReDim actions(0 To 3)
    Set actions(0) = ActionPointerMove(, , , , , element.elementId)
    Set actions(1) = ActionPointerDown
    Set actions(2) = ActionPointerMove(xOffset, yOffset, , , "pointer")
    Set actions(3) = ActionPointerUp
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DragAndDropByOffset = New ActionChain
    Set DragAndDropByOffset.WebDriver = driver_
    Set DragAndDropByOffset.data = data_
    
End Function

Public Function DragAndDropToPosition(element As WebElement, ByVal x As Long, ByVal y As Long) As ActionChain
    'Holds down the left mouse button on the source element, then moves to the x and y offsets
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    ReDim actions(0 To 3)
    Set actions(0) = ActionPointerMove(, , , , , element.elementId)
    Set actions(1) = ActionPointerDown
    Set actions(2) = ActionPointerMove(x, y, , , "viewport")
    Set actions(3) = ActionPointerUp
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DragAndDropToPosition = New ActionChain
    Set DragAndDropToPosition.WebDriver = driver_
    Set DragAndDropToPosition.data = data_
    
End Function


Public Function KeyDown(ByVal modifierKey As String, Optional element As WebElement) As ActionChain
    'Sends a key press only, without releasing it. Should only be used with modifier keys (Control, Alt and Shift).
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If
    
    ReDim actions(0 To 0)
    Set actions(0) = ActionKeyDown(modifierKey)
    
    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set KeyDown = New ActionChain
    Set KeyDown.WebDriver = driver_
    Set KeyDown.data = data_
    
End Function

Public Function KeyUp(ByVal modifierKey As String, Optional element As WebElement) As ActionChain
    'Sends a key press only, without releasing it. Should only be used with modifier keys (Control, Alt and Shift).
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If

    ReDim actions(0 To 0)
    Set actions(0) = ActionKeyUp(modifierKey)
    
    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set KeyUp = New ActionChain
    Set KeyUp.WebDriver = driver_
    Set KeyUp.data = data_
    
End Function

Public Function Wait(Optional ByVal durationms As Long = DEFAULT_PAUSE_DURATION) As ActionChain
    'pauses for time in milliseconds
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    ReDim actions(0 To 0)
    Set actions(0) = ActionPause(durationms)
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set Wait = New ActionChain
    Set Wait.WebDriver = driver_
    Set Wait.data = data_
    
End Function

Public Function MoveByOffset(ByVal xOffset As Long, ByVal yOffset As Long, Optional element As WebElement) As ActionChain
    'Moves the mouse pointer by offsets to current pointer position, or by offset from element top-left
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If Not element Is Nothing Then
        Set rect = element.GetRect()
        x = rect("x") + xOffset
        y = rect("y") + yOffset
        ReDim actions(0 To 0)
        Set actions(0) = ActionPointerMove(x, y)
    Else
        ReDim actions(0 To 0)
        Set actions(0) = ActionPointerMove(xOffset, yOffset, , , "pointer")
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set MoveByOffset = New ActionChain
    Set MoveByOffset.WebDriver = driver_
    Set MoveByOffset.data = data_

End Function

Public Function SendKeys(ByVal Keys As String, Optional element As WebElement) As ActionChain
    'Sends keys to an element
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary

    ReDim chararray(0 To Len(Keys) - 1)
    
    j = -1
    For i = 1 To Len(Keys)
        If Mid(Keys, i, 4) = "\ue0" Then
            j = j + 1
            chararray(j) = Mid(Keys, i, 6)
            i = i + 5
        Else
            j = j + 1
            chararray(j) = Mid(Keys, i, 1)
        End If
    Next i
    
    ReDim Preserve chararray(0 To j)
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If
    
    ReDim actions(0 To 2 * UBound(chararray) + 1)
    For i = 0 To UBound(chararray)
        Set actions(i * 2) = ActionKeyDown(chararray(i))
        Set actions(i * 2 + 1) = ActionKeyUp(chararray(i))
    Next i
    
    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set SendKeys = New ActionChain
    Set SendKeys.WebDriver = driver_
    Set SendKeys.data = data_
    
End Function

Public Function Clear(Optional element As WebElement) As ActionChain
    'Sends a Ctrl-a to select all text and then delete key
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If
    
    ReDim actions(0 To 5)
    Set actions(0) = ActionKeyDown("\ue009") 'control key
    Set actions(1) = ActionKeyDown("a")
    Set actions(2) = ActionKeyUp("a")
    Set actions(3) = ActionKeyUp("\ue009")
    Set actions(4) = ActionKeyDown("\ue017") 'delete key
    Set actions(5) = ActionKeyUp("\ue017")

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set Clear = New ActionChain
    Set Clear.WebDriver = driver_
    Set Clear.data = data_

End Function

Public Function CtrlKey(ByVal key As String, Optional element As WebElement) As ActionChain
    'Sends a Ctrl-[key] down and up
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If

    ReDim actions(0 To 3)
    Set actions(0) = ActionKeyDown("\ue009")
    Set actions(1) = ActionKeyDown(key)
    Set actions(2) = ActionKeyUp(key)
    Set actions(3) = ActionKeyUp("\ue009")

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set CtrlKey = New ActionChain
    Set CtrlKey.WebDriver = driver_
    Set CtrlKey.data = data_

End Function

Public Function ShiftKey(ByVal key As String, Optional element As WebElement) As ActionChain
    'Sends a Shift-[key] down and up
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If
    
    ReDim actions(0 To 3)
    Set actions(0) = ActionKeyDown("\ue008")
    Set actions(1) = ActionKeyDown(key)
    Set actions(2) = ActionKeyUp(key)
    Set actions(3) = ActionKeyUp("\ue008")

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ShiftKey = New ActionChain
    Set ShiftKey.WebDriver = driver_
    Set ShiftKey.data = data_
    
End Function

Public Function ScrollToElement(element As WebElement, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As ActionChain
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    ReDim actions(0 To 0)
    Set actions(0) = ActionScroll(, , xOffset, yOffset, , element.elementId)
    
    inputChan.Add "type", "wheel"
    inputChan.Add "id", "wheelDevice"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ScrollToElement = New ActionChain
    Set ScrollToElement.WebDriver = driver_
    Set ScrollToElement.data = data_

End Function

Public Function ScrollBy(Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As ActionChain
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    ReDim actions(0 To 0)
    Set actions(0) = ActionScroll(, , xOffset, yOffset)
    
    inputChan.Add "type", "wheel"
    inputChan.Add "id", "wheelDevice"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_("actions")
        data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ScrollBy = New ActionChain
    Set ScrollBy.WebDriver = driver_
    Set ScrollBy.data = data_

End Function

Public Sub Perform()
    Dim inputChans() As Variant
    Dim synchronize As Boolean
    
    'we cannot pass a multi-chan Data_("actions") as-is without synchronizing the input channels
    'see https://www.w3.org/TR/webdriver/#actions
    'we can either aggregate and synchronize the inputs into 3 channels (keyboard, pointer, and wheel)
    'and make one call to CMD_W3C_ACTIONS
    'or we execute one input channel at a time in the order in which the inputs were added,
    'making multi-calls to CMD_W3C_ACTIONS
    
    If data_.Exists("actions") Then
        
        'multi-call method:
        'inputChans = data_("actions")
        'For i = 0 To UBound(inputChans)
        '    data_("actions") = Array(inputChans(i))
        '    driver_.Execute CMD_W3C_ACTIONS, data_
        'Next i
        
        '... or synchronize input channels and call Execute once method
        'for typical action chains, this is 5-10% faster...
        
        driver_.Execute CMD_W3C_ACTIONS, SyncChannels(data_)
        
        'it seems that reset actions should always been done on Perform, but Selenium Python doesn't(?)
        'comment below if Selenium Python compatibility is important
        
        ResetActions
        
    End If
    
End Sub

Private Function SyncChannels(data As Dictionary) As Dictionary
    'this private function aggregates an unsynchronized set of action input channels
    'to create a three-channel (keyboard, pointer, and wheel) synchronized system
    'that can be passed as a single call to the Selenium action processor
    'see https://www.w3.org/TR/webdriver/#actions
    
    Dim inputKeyChan As New Dictionary
    Dim inputMouseChan As New Dictionary
    Dim inputScrollChan As New Dictionary
    Dim inputChan As Dictionary
    Dim parameters As New Dictionary
    Dim actionsToAppend() As Dictionary
    Dim mouseActions() As Dictionary
    Dim keyActions() As Dictionary
    Dim scrollActions() As Dictionary
    Dim hasKeyAction As Boolean, hasMouseAction As Boolean, hasScrollAction As Boolean
    Dim lastKeyAction, lastMouseAction, lastScrollAction
    Dim inputChans() As Variant
    Dim syncdInputChans() As Variant
    Dim Pause As Dictionary
    
    inputChans = data("actions")
                
    If UBound(inputChans) = 0 Then 'only one channel - no need to sync
        Set SyncChannels = data
        Exit Function
    End If
    
    'this is a computationally expensive step so do it once up front!!
    Set Pause = ActionPause(0)
    
    'determine in advance the sync'd action array sizes
    numActions = 0
    For i = 0 To UBound(inputChans)
        numActions = numActions + UBound(inputChans(i)("actions")) + 1
    Next i
    
    'dim the sync'd action arrays
    ReDim mouseActions(0 To numActions - 1)
    ReDim keyActions(0 To numActions - 1)
    ReDim scrollActions(0 To numActions - 1)
    
    'loop through each inputChan actions array and load into the appropriate
    'output channels - pad with pauses where needed in order to sync the three output channels
    numActions = 0
    For i = 0 To UBound(inputChans)
        
        Set inputChan = inputChans(i)
        actionsToAppend = inputChan("actions")
        
        Select Case inputChan("id")
        Case "pointerDevice"
            For j = 0 To UBound(actionsToAppend)
                numActions = numActions + 1
                Set mouseActions(numActions - 1) = actionsToAppend(j)
                Set keyActions(numActions - 1) = Pause
                Set scrollActions(numActions - 1) = Pause
                lastMouseAction = numActions 'last non-pause action
            Next j
            hasMouseAction = True
        Case "keyboardDevice"
            For j = 0 To UBound(actionsToAppend)
                numActions = numActions + 1
                Set mouseActions(numActions - 1) = Pause
                Set keyActions(numActions - 1) = actionsToAppend(j)
                Set scrollActions(numActions - 1) = Pause
                lastKeyAction = numActions 'last non-pause action
            Next j
            hasKeyAction = True
        Case "wheelDevice"
            For j = 0 To UBound(actionsToAppend)
                numActions = numActions + 1
                Set mouseActions(numActions - 1) = Pause
                Set keyActions(numActions - 1) = Pause
                Set scrollActions(numActions - 1) = actionsToAppend(j)
                lastScrollAction = numActions 'last non-pause action
            Next j
            hasScrollAction = True
        End Select
    Next i
    
    numsyncdInputChans = -(hasKeyAction + hasMouseAction + hasScrollAction) 'either 1, 2, or 3 output channels
    ReDim syncdInputChans(0 To numsyncdInputChans - 1)
    
    channum = 0
    
    If hasMouseAction Then
        'if channel has a string of pauses until the end, then we should eliminate these for better performance
        ReDim Preserve mouseActions(0 To lastMouseAction - 1)
        parameters.Add "pointerType", "mouse" 'required for pointer type
        inputMouseChan.Add "id", "pointerDevice"
        inputMouseChan.Add "type", "pointer"
        inputMouseChan.Add "parameters", parameters
        inputMouseChan.Add "actions", mouseActions
        channum = channum + 1
        Set syncdInputChans(channum - 1) = inputMouseChan
    End If
    
    If hasKeyAction Then
        'if channel has a string of pauses until the end, then we should eliminate these for better performance
        ReDim Preserve keyActions(0 To lastKeyAction - 1)
        inputKeyChan.Add "id", "keyboardDevice"
        inputKeyChan.Add "type", "key"
        inputKeyChan.Add "actions", keyActions
        channum = channum + 1
        Set syncdInputChans(channum - 1) = inputKeyChan
    End If
    
    If hasScrollAction Then
        'if channel has a string of pauses until the end, then we should eliminate these for better performance
        ReDim Preserve scrollActions(0 To lastScrollAction - 1)
        inputScrollChan.Add "id", "wheelDevice"
        inputScrollChan.Add "type", "wheel"
        inputScrollChan.Add "actions", scrollActions
        channum = channum + 1
        Set syncdInputChans(channum - 1) = inputScrollChan
    End If
    
    data("actions") = syncdInputChans
    
    Set SyncChannels = data

End Function

'make this public for compatibility with Selenium Python (see Perform method)
Private Sub ResetActions()
    'clears actions that are already stored locally and on the remote end
    driver_.Execute CMD_W3C_CLEAR_ACTIONS, data_ 'remote end clear
    data_.Remove "actions" 'local end clear
End Sub

'Following are private single action functions

Private Function ActionKeyDown(Optional ByVal value As String)
    Dim Action As New Dictionary
    Action.Add "type", "keyDown"
    Action.Add "value", value
    Set ActionKeyDown = Action
End Function

Private Function ActionKeyUp(Optional ByVal value As String)
    Dim Action As New Dictionary
    Action.Add "type", "keyUp"
    Action.Add "value", value
    Set ActionKeyUp = Action
End Function

Private Function ActionPointerDown(Optional ByVal button As ButtonType = 0) As Dictionary
    Dim Action As New Dictionary
    Action.Add "type", "pointerDown" 'pause, keyUp, KeyDown, pointerUp, pointerDown, pointerCancel
    Action.Add "button", button '0 for left, 1 for middle,  2 for right
    
    'these are optional for pointerMove and pointerDown
    'action.Add "width", 0
    'action.Add "height", 0
    'action.Add "pressure", 0
    'action.Add "tangentialPressure", 0
    'action.Add "tiltX", 0
    'action.Add "tiltY", 0
    'action.Add "altitudeAngle", 0
    'action.Add "azimuthAngle", 0
    'action.Add "twist", 0
    
    Set ActionPointerDown = Action
End Function

Private Function ActionPointerUp(Optional ByVal button As ButtonType = 0)
    Dim Action As New Dictionary
    Action.Add "type", "pointerUp"
    Action.Add "button", button
    Set ActionPointerUp = Action
End Function

Private Function ActionPause(Optional ByVal duration As Long = DEFAULT_PAUSE_DURATION)
    Dim Action As New Dictionary
    Action.Add "type", "pause"
    If duration > 0 Then Action.Add "duration", duration
    Set ActionPause = Action
End Function

Private Function ActionPointerCancel()
    Dim Action As New Dictionary
    Action.Add "type", "pointerCancel"
    Set ActionPointerCancel = Action
End Function

Private Function ActionPointerMove(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal button As ButtonType = 0, Optional ByVal duration As Long = DEFAULT_PAUSE_DURATION, Optional ByVal origin = "viewport", Optional elementId As String) As Dictionary
    Dim Action As New Dictionary
    
    Action.Add "type", "pointerMove"
    Action.Add "button", button
    
    If elementId = "" Then
        Action.Add "origin", origin  'x,y are relative to. Can be 'viewport', 'pointer' or {"element-6066-11e4-a52e-4f735466cecf": element.id}
    Else
        Dim elemdict As New Dictionary
        elemdict.Add "element-6066-11e4-a52e-4f735466cecf", elementId
        Action.Add "origin", elemdict
    End If
    'action.Add "origin", elemdict
    
    Action.Add "duration", duration 'Only applicable to 'pause' and 'pointerMove'.
    Action.Add "x", x
    Action.Add "y", y
    
    'these are optional for pointerMove and pointerDown
    'action.Add "width", 0
    'action.Add "height", 0
    'action.Add "pressure", 0
    'action.Add "tangentialPressure", 0
    'action.Add "tiltX", 0
    'action.Add "tiltY", 0
    'action.Add "altitudeAngle", 0
    'action.Add "azimuthAngle", 0
    'action.Add "twist", 0
    
    Set ActionPointerMove = Action
    
End Function

Private Function ActionScroll(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal deltaX As Long = 0, Optional ByVal deltaY As Long = 0, Optional ByVal duration As Long = DEFAULT_PAUSE_DURATION, Optional elementId As String) As Dictionary
    Dim Action As New Dictionary
    
    Action.Add "type", "scroll"
    
    If elementId = "" Then
        Action.Add "origin", "viewport"
    Else
        Dim elemdict As New Dictionary
        elemdict.Add "element-6066-11e4-a52e-4f735466cecf", elementId
        Action.Add "origin", elemdict
    End If
    
    Action.Add "duration", duration
    Action.Add "x", x
    Action.Add "y", y
    Action.Add "deltaX", deltaX
    Action.Add "deltaY", deltaY
    
    Set ActionScroll = Action

End Function

'this is a utility function to append an input channel to the input channel array
Private Function AppendInputChannel(inputChans() As Variant, inputChan As Dictionary) As Variant()
    ReDim Preserve inputChans(0 To UBound(inputChans) + 1)
    Set inputChans(UBound(inputChans)) = inputChan
    AppendInputChannel = inputChans
End Function

Private Sub Class_Initialize()
    CMD_W3C_ACTIONS = Array("POST", "/session/$sessionId/actions")
    CMD_W3C_CLEAR_ACTIONS = Array("DELETE", "/session/$sessionId/actions")
End Sub
