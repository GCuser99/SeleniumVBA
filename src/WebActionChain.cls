VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebActionChain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================
' SeleniumVBA v2.2
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
'DESCRIPTION:
'
'This class was written to emulate Selenium Webdriver ActionChains
'
'USAGE(S):
'
'Method 1:
'   Dim Actions As WebActionChain
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
'   This version clears actions that are already stored locally and on the remote end
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
'

Option Explicit

Private Const DEFAULT_PAUSE_DURATION = 300
Private driver_ As WebDriver
Private data_ As Dictionary

Private Type typCMD
    CMD_ACTIONS As Variant
    CMD_CLEAR_ACTIONS As Variant
End Type
Private tCMD As typCMD

Private Enum ButtonType
    leftButton = 0
    middleButton = 1
    rightButton = 2
End Enum

'These are helper properties for passing webdriver and data dictionary references to each link of chain
Friend Property Set WebDriver(ByRef obj As WebDriver)
    Set driver_ = obj
End Property

Friend Property Set Data(ByRef obj As Dictionary)
    Set data_ = obj
End Property

'Action Chain Methods:

Public Function MoveToElement(element As WebElement, Optional SetFocus As Boolean = False) As WebActionChain
    'Moves the mouse pointer to the middle of an element.
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If SetFocus Then 'add click after move
        ReDim actions(0 To 3)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerDown
        Set actions(2) = ActionPointerUp
        Set actions(3) = ActionPause
    Else
        ReDim actions(0 To 0)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        'ActionPointerMove already has pause built in
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
    Set MoveToElement = New WebActionChain
    Set MoveToElement.WebDriver = driver_
    Set MoveToElement.Data = data_
End Function

Public Function Click(Optional element As WebElement) As WebActionChain
    'clicks an element
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        ReDim actions(0 To 3)
        Set actions(0) = ActionPointerDown
        Set actions(1) = ActionPointerUp
        Set actions(2) = ActionPause
        Set actions(3) = ActionPause
    Else
        ReDim actions(0 To 4)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerDown
        Set actions(2) = ActionPointerUp
        Set actions(3) = ActionPause
        Set actions(4) = ActionPause
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
    Set Click = New WebActionChain
    Set Click.WebDriver = driver_
    Set Click.Data = data_
End Function

Public Function ClickAndHold(Optional element As WebElement) As WebActionChain
    'Holds down the left mouse button on an element
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        ReDim actions(0 To 1)
        Set actions(0) = ActionPointerDown
        Set actions(1) = ActionPause
    Else
        ReDim actions(0 To 2)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerDown
        Set actions(2) = ActionPause
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
    Set ClickAndHold = New WebActionChain
    Set ClickAndHold.WebDriver = driver_
    Set ClickAndHold.Data = data_
End Function

Public Function ClickContext(Optional element As WebElement) As WebActionChain
    'Performs a context-click (right click) on an element
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        ReDim actions(0 To 3)
        Set actions(0) = ActionPointerDown(rightButton)
        Set actions(1) = ActionPointerUp(rightButton)
        Set actions(2) = ActionPause
        Set actions(3) = ActionPause
    Else
        ReDim actions(0 To 4)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerDown(rightButton)
        Set actions(2) = ActionPointerUp(rightButton)
        Set actions(3) = ActionPause
        Set actions(4) = ActionPause
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
    Set ClickContext = New WebActionChain
    Set ClickContext.WebDriver = driver_
    Set ClickContext.Data = data_
End Function

Public Function DoubleClick(Optional element As WebElement) As WebActionChain
    'Double-clicks an element
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        ReDim actions(0 To 5)
        Set actions(0) = ActionPointerDown
        Set actions(1) = ActionPointerUp
        Set actions(2) = ActionPointerDown
        Set actions(3) = ActionPointerUp
        Set actions(4) = ActionPause
        Set actions(5) = ActionPause
    Else
        ReDim actions(0 To 6)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerDown
        Set actions(2) = ActionPointerUp
        Set actions(3) = ActionPointerDown
        Set actions(4) = ActionPointerUp
        Set actions(5) = ActionPause
        Set actions(6) = ActionPause
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
    Set DoubleClick = New WebActionChain
    Set DoubleClick.WebDriver = driver_
    Set DoubleClick.Data = data_
End Function

Public Function ReleaseButton(Optional element As WebElement) As WebActionChain
    'Releasing a held mouse button
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        ReDim actions(0 To 1)
        Set actions(0) = ActionPointerUp
        Set actions(1) = ActionPause
    Else
        ReDim actions(0 To 2)
        Set actions(0) = ActionPointerMove(, , , , , element.elementId)
        Set actions(1) = ActionPointerUp
        Set actions(2) = ActionPause
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
    Set ReleaseButton = New WebActionChain
    Set ReleaseButton.WebDriver = driver_
    Set ReleaseButton.Data = data_
End Function

Public Function DragAndDrop(sourceElement As WebElement, targetElement As WebElement) As WebActionChain
    'Holds down the left mouse button on the source element, then moves to the target element and releases the mouse button.
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    ReDim actions(0 To 4)
    Set actions(0) = ActionPointerMove(, , , , , sourceElement.elementId)
    Set actions(1) = ActionPointerDown
    Set actions(2) = ActionPointerMove(, , , , , targetElement.elementId)
    Set actions(3) = ActionPointerUp
    Set actions(4) = ActionPause
    
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
    Set DragAndDrop = New WebActionChain
    Set DragAndDrop.WebDriver = driver_
    Set DragAndDrop.Data = data_
End Function

Public Function DragAndDropByOffset(element As WebElement, ByVal xOffset As Long, ByVal yOffset As Long) As WebActionChain
    'Holds down the left mouse button on the source element, then moves to the x and y offsets
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    ReDim actions(0 To 4)
    Set actions(0) = ActionPointerMove(, , , , , element.elementId)
    Set actions(1) = ActionPointerDown
    Set actions(2) = ActionPointerMove(xOffset, yOffset, , , "pointer")
    Set actions(3) = ActionPointerUp
    Set actions(4) = ActionPause
    
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
    Set DragAndDropByOffset = New WebActionChain
    Set DragAndDropByOffset.WebDriver = driver_
    Set DragAndDropByOffset.Data = data_
End Function

Public Function DragAndDropToPosition(element As WebElement, ByVal x As Long, ByVal y As Long) As WebActionChain
    'Holds down the left mouse button on the source element, then moves to the x and y offsets
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    ReDim actions(0 To 4)
    Set actions(0) = ActionPointerMove(, , , , , element.elementId)
    Set actions(1) = ActionPointerDown
    Set actions(2) = ActionPointerMove(x, y, , , "viewport")
    Set actions(3) = ActionPointerUp
    Set actions(4) = ActionPause
    
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
    Set DragAndDropToPosition = New WebActionChain
    Set DragAndDropToPosition.WebDriver = driver_
    Set DragAndDropToPosition.Data = data_
End Function

Public Function KeyDown(ByVal modifierKey As String, Optional element As WebElement) As WebActionChain
    'Sends a key press only, without releasing it. Should only be used with modifier keys (Control, Alt and Shift).
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If
    
    ReDim actions(0 To 1)
    Set actions(0) = ActionKeyDown(modifierKey)
    Set actions(1) = ActionPause
    
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
    Set KeyDown = New WebActionChain
    Set KeyDown.WebDriver = driver_
    Set KeyDown.Data = data_
End Function

Public Function KeyUp(ByVal modifierKey As String, Optional element As WebElement) As WebActionChain
    'Sends a key press only, without releasing it. Should only be used with modifier keys (Control, Alt and Shift).
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If

    ReDim actions(0 To 1)
    Set actions(0) = ActionKeyUp(modifierKey)
    Set actions(1) = ActionPause
    
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
    Set KeyUp = New WebActionChain
    Set KeyUp.WebDriver = driver_
    Set KeyUp.Data = data_
End Function

Public Function Wait(Optional ByVal durationMS As Long = DEFAULT_PAUSE_DURATION) As WebActionChain
    'pauses for time in milliseconds
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    ReDim actions(0 To 0)
    Set actions(0) = ActionPause(durationMS)
    
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
    Set Wait = New WebActionChain
    Set Wait.WebDriver = driver_
    Set Wait.Data = data_
End Function

Public Function MoveByOffset(ByVal xOffset As Long, ByVal yOffset As Long, Optional element As WebElement) As WebActionChain
    'Moves the mouse pointer by offsets to current pointer position, or by offset from element top-left
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If Not element Is Nothing Then
        Dim rect As Dictionary
        Set rect = element.GetRect()
        Dim x As Long, y As Long
        x = rect("x") + xOffset
        y = rect("y") + yOffset
        ReDim actions(0 To 1)
        Set actions(0) = ActionPointerMove(x, y)
        'ActionPointerMove has built in pause
    Else
        ReDim actions(0 To 0)
        Set actions(0) = ActionPointerMove(xOffset, yOffset, , , "pointer")
        'ActionPointerMove has built in pause
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
    Set MoveByOffset = New WebActionChain
    Set MoveByOffset.WebDriver = driver_
    Set MoveByOffset.Data = data_
End Function

Public Function SendKeys(ByVal keys As String, Optional element As WebElement) As WebActionChain
    'Sends keys to an element
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary

    ReDim charArray(0 To Len(keys) - 1)
    
    Dim j As Long, i As Long
    j = -1
    For i = 1 To Len(keys)
        If Mid(keys, i, 4) = "\ue0" Then 'its a special key
            j = j + 1
            charArray(j) = Mid(keys, i, 6)
            i = i + 5
        Else
            j = j + 1
            charArray(j) = Mid(keys, i, 1)
        End If
    Next i
    
    ReDim Preserve charArray(0 To j)
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If
    
    ReDim actions(0 To 2 * UBound(charArray) + 2)
    For i = 0 To UBound(charArray)
        Set actions(i * 2) = ActionKeyDown(charArray(i))
        Set actions(i * 2 + 1) = ActionKeyUp(charArray(i))
    Next i
    
    Set actions(2 * UBound(charArray) + 2) = ActionPause
    
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
    Set SendKeys = New WebActionChain
    Set SendKeys.WebDriver = driver_
    Set SendKeys.Data = data_
End Function

Public Function Clear(Optional element As WebElement) As WebActionChain
    'Sends a Ctrl-a to select all text and then delete key
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If
    
    ReDim actions(0 To 6)
    Set actions(0) = ActionKeyDown("\ue009") 'control key
    Set actions(1) = ActionKeyDown("a")
    Set actions(2) = ActionKeyUp("a")
    Set actions(3) = ActionKeyUp("\ue009")
    Set actions(4) = ActionKeyDown("\ue017") 'delete key
    Set actions(5) = ActionKeyUp("\ue017")
    Set actions(6) = ActionPause

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
    Set Clear = New WebActionChain
    Set Clear.WebDriver = driver_
    Set Clear.Data = data_
End Function

Public Function CtrlKey(ByVal key As String, Optional element As WebElement) As WebActionChain
    'Sends a Ctrl-[key] down and up
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If

    ReDim actions(0 To 4)
    Set actions(0) = ActionKeyDown("\ue009") 'control key
    Set actions(1) = ActionKeyDown(key)
    Set actions(2) = ActionKeyUp(key)
    Set actions(3) = ActionKeyUp("\ue009")
    Set actions(4) = ActionPause

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
    Set CtrlKey = New WebActionChain
    Set CtrlKey.WebDriver = driver_
    Set CtrlKey.Data = data_
End Function

Public Function ShiftKey(ByVal key As String, Optional element As WebElement) As WebActionChain
    'Sends a Shift-[key] down and up
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If
    
    ReDim actions(0 To 4)
    Set actions(0) = ActionKeyDown("\ue008") 'shift key
    Set actions(1) = ActionKeyDown(key)
    Set actions(2) = ActionKeyUp(key)
    Set actions(3) = ActionKeyUp("\ue008")
    Set actions(4) = ActionPause

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
    Set ShiftKey = New WebActionChain
    Set ShiftKey.WebDriver = driver_
    Set ShiftKey.Data = data_
End Function

Public Function ScrollToElement(element As WebElement, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As WebActionChain
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If driver_.GetBrowserName = "firefox" Then
        Err.raise 1, "WebActionChain", "Scroll actions not supported for Firefox WebDriver"
    End If
    
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
    Set ScrollToElement = New WebActionChain
    Set ScrollToElement.WebDriver = driver_
    Set ScrollToElement.Data = data_
End Function

Public Function ScrollBy(Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As WebActionChain
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If driver_.GetBrowserName = "firefox" Then
        Err.raise 1, "WebActionChain", "Scroll actions not supported for Firefox WebDriver"
    End If
    
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
    Set ScrollBy = New WebActionChain
    Set ScrollBy.WebDriver = driver_
    Set ScrollBy.Data = data_
End Function

Public Sub Perform()
    Dim inputChans() As Variant
    Dim synchronize As Boolean
    
    'we cannot pass a multi-chan Data_("actions") as-is without synchronizing the input channels
    'see https://www.w3.org/TR/webdriver/#actions
    'we can either aggregate and synchronize the inputs into 3 channels (keyboard, pointer, and wheel)
    'and make one call to tCMD.CMD_ACTIONS
    'or we execute one input channel at a time in the order in which the inputs were added,
    'making multi-calls to tCMD.CMD_ACTIONS
    
    'Note for future: Firefox and IEdriverServer doe not support wheel actions
    
    If data_.Exists("actions") Then
        'multi-call method:
        'inputChans = data_("actions")
        'For i = 0 To UBound(inputChans)
        '    data_("actions") = Array(inputChans(i))
        '    driver_.Execute tCMD.CMD_ACTIONS, data_
        'Next i
        
        '... or synchronize input channels and call Execute method once
        'for typical action chains, this is 5-10% faster...
        driver_.Execute tCMD.CMD_ACTIONS, SyncChannels(data_)
        
        'it seems that reset actions should always been done on Perform, but Selenium Python doesn't(?)
        
        ResetActions
    End If
End Sub

Private Function SyncChannels(Data As Dictionary) As Dictionary
    'this private function aggregates an unsynchronized set of action input channels
    'to create a three-channel (keyboard, pointer, and wheel) synchronized set
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
    Dim pause As Dictionary
    
    inputChans = Data("actions")
                
    If UBound(inputChans) = 0 Then 'only one channel - no need to sync
        Set SyncChannels = Data
        Exit Function
    End If
    
    'this is a computationally expensive step so do it once up front!!
    Set pause = ActionPause(0) 'dummy pause
    
    'determine in advance the sync'd action array sizes
    Dim numActions As Long, i As Long
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
        
        Dim j As Long
        Select Case inputChan("id")
        Case "pointerDevice"
            For j = 0 To UBound(actionsToAppend)
                numActions = numActions + 1
                Set mouseActions(numActions - 1) = actionsToAppend(j)
                Set keyActions(numActions - 1) = pause
                Set scrollActions(numActions - 1) = pause
                lastMouseAction = numActions 'last non-dummy-pause action
            Next j
            hasMouseAction = True
        Case "keyboardDevice"
            For j = 0 To UBound(actionsToAppend)
                numActions = numActions + 1
                Set mouseActions(numActions - 1) = pause
                Set keyActions(numActions - 1) = actionsToAppend(j)
                Set scrollActions(numActions - 1) = pause
                lastKeyAction = numActions 'last non-dummy-pause action
            Next j
            hasKeyAction = True
        Case "wheelDevice"
            For j = 0 To UBound(actionsToAppend)
                numActions = numActions + 1
                Set mouseActions(numActions - 1) = pause
                Set keyActions(numActions - 1) = pause
                Set scrollActions(numActions - 1) = actionsToAppend(j)
                lastScrollAction = numActions 'last non-dummy-pause action
            Next j
            hasScrollAction = True
        End Select
    Next i
    
    Dim numSyncdInputChans As Long
    numSyncdInputChans = -(hasKeyAction + hasMouseAction + hasScrollAction) 'either 1, 2, or 3 output channels
    ReDim syncdInputChans(0 To numSyncdInputChans - 1)
    
    Dim chanNum As Long
    chanNum = 0
    
    If hasMouseAction Then
        'if channel has a string of dummy-pauses until the end, then we should eliminate these for better performance
        ReDim Preserve mouseActions(0 To lastMouseAction - 1)
        parameters.Add "pointerType", "mouse" 'required for pointer type
        inputMouseChan.Add "id", "pointerDevice"
        inputMouseChan.Add "type", "pointer"
        inputMouseChan.Add "parameters", parameters
        inputMouseChan.Add "actions", mouseActions
        chanNum = chanNum + 1
        Set syncdInputChans(chanNum - 1) = inputMouseChan
    End If
    
    If hasKeyAction Then
        'if channel has a string of dummy-pauses until the end, then we should eliminate these for better performance
        ReDim Preserve keyActions(0 To lastKeyAction - 1)
        inputKeyChan.Add "id", "keyboardDevice"
        inputKeyChan.Add "type", "key"
        inputKeyChan.Add "actions", keyActions
        chanNum = chanNum + 1
        Set syncdInputChans(chanNum - 1) = inputKeyChan
    End If
    
    If hasScrollAction Then
        'if channel has a string of dummy-pauses until the end, then we should eliminate these for better performance
        ReDim Preserve scrollActions(0 To lastScrollAction - 1)
        inputScrollChan.Add "id", "wheelDevice"
        inputScrollChan.Add "type", "wheel"
        inputScrollChan.Add "actions", scrollActions
        chanNum = chanNum + 1
        Set syncdInputChans(chanNum - 1) = inputScrollChan
    End If
    
    Data("actions") = syncdInputChans
    
    Set SyncChannels = Data
End Function

'make this public for compatibility with Selenium Python? (see Perform method)
Private Sub ResetActions()
    'clears actions that are already stored locally and on the remote end
    driver_.Execute tCMD.CMD_CLEAR_ACTIONS, data_ 'remote end clear
    data_.Remove "actions" 'local end clear
End Sub

'Following are private single action functions

Private Function ActionKeyDown(Optional ByVal val As String)
    Dim action As New Dictionary
    action.Add "type", "keyDown"
    action.Add "value", val
    Set ActionKeyDown = action
End Function

Private Function ActionKeyUp(Optional ByVal val As String)
    Dim action As New Dictionary
    action.Add "type", "keyUp"
    action.Add "value", val
    Set ActionKeyUp = action
End Function

Private Function ActionPointerDown(Optional ByVal button As ButtonType = 0) As Dictionary
    Dim action As New Dictionary
    action.Add "type", "pointerDown" 'pause, keyUp, KeyDown, pointerUp, pointerDown, pointerCancel
    action.Add "button", button '0 for left, 1 for middle,  2 for right
    
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
    
    Set ActionPointerDown = action
End Function

Private Function ActionPointerUp(Optional ByVal button As ButtonType = 0)
    Dim action As New Dictionary
    action.Add "type", "pointerUp"
    action.Add "button", button
    Set ActionPointerUp = action
End Function

Private Function ActionPause(Optional ByVal duration As Long = DEFAULT_PAUSE_DURATION)
    Dim action As New Dictionary
    action.Add "type", "pause"
    If duration > 0 Then action.Add "duration", duration
    Set ActionPause = action
End Function

Private Function ActionPointerCancel()
    Dim action As New Dictionary
    action.Add "type", "pointerCancel"
    Set ActionPointerCancel = action
End Function

Private Function ActionPointerMove(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal button As ButtonType = 0, Optional ByVal duration As Long = DEFAULT_PAUSE_DURATION, Optional ByVal origin = "viewport", Optional elementId As String) As Dictionary
    Dim action As New Dictionary
    
    action.Add "type", "pointerMove"
    action.Add "button", button
    
    If elementId = "" Then
        action.Add "origin", origin  'x,y are relative to. Can be 'viewport', 'pointer' or {"element-6066-11e4-a52e-4f735466cecf": element.id}
    Else
        Dim elem As New Dictionary
        elem.Add "element-6066-11e4-a52e-4f735466cecf", elementId
        action.Add "origin", elem
    End If
    'action.Add "origin", elem
    
    action.Add "duration", duration 'Only applicable to 'pause' and 'pointerMove'.
    action.Add "x", x
    action.Add "y", y
    
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
    
    Set ActionPointerMove = action
End Function

Private Function ActionScroll(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal deltaX As Long = 0, Optional ByVal deltaY As Long = 0, Optional ByVal duration As Long = DEFAULT_PAUSE_DURATION, Optional elementId As String) As Dictionary
    Dim action As New Dictionary
    
    action.Add "type", "scroll"
    
    If elementId = "" Then
        action.Add "origin", "viewport"
    Else
        Dim elem As New Dictionary
        elem.Add "element-6066-11e4-a52e-4f735466cecf", elementId
        action.Add "origin", elem
    End If
    
    action.Add "duration", duration
    action.Add "x", x
    action.Add "y", y
    action.Add "deltaX", deltaX
    action.Add "deltaY", deltaY
    
    Set ActionScroll = action
End Function

'this is a utility function to append an input channel to the input channel array
Private Function AppendInputChannel(inputChans() As Variant, inputChan As Dictionary) As Variant()
    ReDim Preserve inputChans(0 To UBound(inputChans) + 1)
    Set inputChans(UBound(inputChans)) = inputChan
    AppendInputChannel = inputChans
End Function

Private Sub Class_Initialize()
    tCMD.CMD_ACTIONS = Array("POST", "/session/$sessionId/actions")
    tCMD.CMD_CLEAR_ACTIONS = Array("DELETE", "/session/$sessionId/actions")
End Sub
