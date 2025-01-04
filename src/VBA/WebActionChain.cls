VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebActionChain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used to emulate a human interaction sequence with the webpage - must be instantiated using WebDriver's ActionChain method"
'@ModuleDescription "This class is used to emulate a human interaction sequence with the webpage - must be instantiated using WebDriver's ActionChain method"
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
'AltShiftKey            Sends a Alt-Shift-[Key] sequence to focused element
'Clear                  Clears the focused element's text field
'Click                  Clicks an element
'ClickAndHold           Holds down the left mouse button on an element
'ClickContext           Performs a context-click (right click) on an element
'CtrlAltKey             Sends a Ctrl-Alt-[Key] sequence to focused element
'CtrlKey                Sends a Ctrl-[Key] sequence to focused element
'CtrlShiftKey           Sends a Ctrl-Shift-[Key] sequence to focused element
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
'MoveToXY               Moves the mouse pointer to (x,y) relative to viewport origin (top-left).
'Perform                Executes the action chain and then resets the action queue on both the local and remote ends
'ReleaseButton          Releases a held mouse button on an element
'ResetActions           Resets the action queue and key/pointer states on both the local and remote ends
'ScrollBy               Scrolls window by x and y offsets
'ScrollToElement        Scrolls to element position with optional x and y offsets
'SendKeys               Sends key sequence to current focused element (see Keyboard class for special keys support)
'SetFocus               Sets the focus on the given element
'ShiftKey               Sends a Shift-[Key] sequence to focused element
'Wait                   Pause all inputs for the specified duration in milliseconds

Option Explicit

Private Const DEFAULT_PAUSE_DURATION As Long = 350
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

'Action Chain Methods:

'@Description("Moves the mouse to the middle of an element")
Public Function MoveToElement(element As WebElement, Optional ByVal SetFocus As Boolean = False) As WebActionChain
Attribute MoveToElement.VB_Description = "Moves the mouse to the middle of an element"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If SetFocus Then 'add click after move
        actions.Add actionPointerMove(, , , , , element.elementId)
        actions.Add actionPointerDown
        actions.Add actionPointerUp
        actions.Add actionPause
    Else
        actions.Add actionPointerMove(, , , , , element.elementId)
        'ActionPointerMove already has pause built in
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set MoveToElement = Me
End Function

'@Description("Clicks an element")
Public Function Click(Optional element As WebElement) As WebActionChain
Attribute Click.VB_Description = "Clicks an element"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        actions.Add actionPointerDown
        actions.Add actionPointerUp
        actions.Add actionPause
        actions.Add actionPause
    Else
        actions.Add actionPointerMove(, , , , , element.elementId)
        actions.Add actionPointerDown
        actions.Add actionPointerUp
        actions.Add actionPause
        actions.Add actionPause
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set Click = Me
End Function

'@Description("Holds down the left mouse button on an element")
Public Function ClickAndHold(Optional element As WebElement) As WebActionChain
Attribute ClickAndHold.VB_Description = "Holds down the left mouse button on an element"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        actions.Add actionPointerDown
        actions.Add actionPause
    Else
        actions.Add actionPointerMove(, , , , , element.elementId)
        actions.Add actionPointerDown
        actions.Add actionPause
    End If
        
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ClickAndHold = Me
End Function

'@Description("Performs a context-click (right click) on an element")
Public Function ClickContext(Optional element As WebElement) As WebActionChain
Attribute ClickContext.VB_Description = "Performs a context-click (right click) on an element"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        actions.Add actionPointerDown(rightButton)
        actions.Add actionPointerUp(rightButton)
        actions.Add actionPause
        actions.Add actionPause
    Else
        actions.Add actionPointerMove(, , , , , element.elementId)
        actions.Add actionPointerDown(rightButton)
        actions.Add actionPointerUp(rightButton)
        actions.Add actionPause
        actions.Add actionPause
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ClickContext = Me
End Function

'@Description("Double-clicks an element")
Public Function DoubleClick(Optional element As WebElement) As WebActionChain
Attribute DoubleClick.VB_Description = "Double-clicks an element"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        actions.Add actionPointerDown
        actions.Add actionPointerUp
        actions.Add actionPointerDown
        actions.Add actionPointerUp
        actions.Add actionPause
        actions.Add actionPause
    Else
        actions.Add actionPointerMove(, , , , , element.elementId)
        actions.Add actionPointerDown
        actions.Add actionPointerUp
        actions.Add actionPointerDown
        actions.Add actionPointerUp
        actions.Add actionPause
        actions.Add actionPause
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DoubleClick = Me
End Function

'@Description("Releases a held mouse button on an element")
Public Function ReleaseButton(Optional element As WebElement) As WebActionChain
Attribute ReleaseButton.VB_Description = "Releases a held mouse button on an element"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If element Is Nothing Then
        actions.Add actionPointerUp
        actions.Add actionPause
    Else
        actions.Add actionPointerMove(, , , , , element.elementId)
        actions.Add actionPointerUp
        actions.Add actionPause
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ReleaseButton = Me
End Function

'@Description("Holds down the left mouse button on the source element, then moves to the target element and releases the mouse button")
Public Function DragAndDrop(sourceElement As WebElement, targetElement As WebElement) As WebActionChain
Attribute DragAndDrop.VB_Description = "Holds down the left mouse button on the source element, then moves to the target element and releases the mouse button"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    actions.Add actionPointerMove(, , , , , sourceElement.elementId)
    actions.Add actionPointerDown
    actions.Add actionPointerMove(, , , , , targetElement.elementId)
    actions.Add actionPointerUp
    actions.Add actionPause
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DragAndDrop = Me
End Function

'@Description("Holds down the left mouse button on the source element, then moves to the target offset and releases the mouse button")
Public Function DragAndDropByOffset(element As WebElement, ByVal xOffset As Long, ByVal yOffset As Long) As WebActionChain
Attribute DragAndDropByOffset.VB_Description = "Holds down the left mouse button on the source element, then moves to the target offset and releases the mouse button"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    actions.Add actionPointerMove(, , , , , element.elementId)
    actions.Add actionPointerDown
    actions.Add actionPointerMove(xOffset, yOffset, , , "pointer")
    actions.Add actionPointerUp
    actions.Add actionPause
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DragAndDropByOffset = Me
End Function

'@Description("Holds down the left mouse button on the source element, then moves to the x and y offsets")
Public Function DragAndDropToPosition(element As WebElement, ByVal x As Long, ByVal y As Long) As WebActionChain
Attribute DragAndDropToPosition.VB_Description = "Holds down the left mouse button on the source element, then moves to the x and y offsets"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    actions.Add actionPointerMove(, , , , , element.elementId)
    actions.Add actionPointerDown
    actions.Add actionPointerMove(x, y, , , "viewport")
    actions.Add actionPointerUp
    actions.Add actionPause
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DragAndDropToPosition = Me
End Function

'@Description("Sends a key press only, without releasing it. Should only be used with modifier keys (Ctrl, Alt, and Shift)")
Public Function KeyDown(ByVal modifierKey As String, Optional element As WebElement) As WebActionChain
Attribute KeyDown.VB_Description = "Sends a key press only, without releasing it. Should only be used with modifier keys (Ctrl, Alt, and Shift)"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    
    modifierKey = unEscapeUnicode(modifierKey)
    If Not (AscWL(modifierKey) >= &HE008& And AscWL(modifierKey) <= &HE00A&) Then Err.Raise 1, , "Error in KeyDown method - modifierKey must be either a single Ctrl, Alt, or Shift key"
        
    If Not element Is Nothing Then Me.SetFocus element
    
    actions.Add actionKeyDown(modifierKey)
    
    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set KeyDown = Me
End Function

'@Description("Releases a modifier key")
Public Function KeyUp(ByVal modifierKey As String, Optional element As WebElement) As WebActionChain
Attribute KeyUp.VB_Description = "Releases a modifier key"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    
    modifierKey = unEscapeUnicode(modifierKey)
    If Not (AscWL(modifierKey) >= &HE008& And AscWL(modifierKey) <= &HE00A&) Then Err.Raise 1, , "Error in KeyUp method - modifierKey must be either a single Ctrl, Alt, or Shift key"
    
    If Not element Is Nothing Then Me.SetFocus element

    actions.Add actionKeyUp(modifierKey)
    
    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set KeyUp = Me
End Function

'@Description("Pauses execution for time in milliseconds")
Public Function Wait(Optional ByVal durationMS As Long = DEFAULT_PAUSE_DURATION) As WebActionChain
Attribute Wait.VB_Description = "Pauses execution for time in milliseconds"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    actions.Add actionPause(durationMS)
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set Wait = Me
End Function

'@Description("Moves the mouse to an offset from current mouse position - if element is specified, then moves the mouse by an offset relative to the top-left corner of the element")
Public Function MoveByOffset(ByVal xOffset As Long, ByVal yOffset As Long, Optional element As WebElement) As WebActionChain
Attribute MoveByOffset.VB_Description = "Moves the mouse to an offset from current mouse position - if element is specified, then moves the mouse by an offset relative to the top-left corner of the element"
    'Note that the mouse pointer gets (re)set to (0,0) at the start of each action chain. If element is not provided,
    'the offsets are relative to the mouse pointer in the current action chain.
    'So for instance, actions.MoveByOffset(300,200).MoveByOffset(200,300).Perform will move the mouse pointer to (500,500)
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If Not element Is Nothing Then
        Dim rect As Dictionary
        Set rect = element.GetRect()
        Dim x As Long
        Dim y As Long
        x = rect.Item("x") + xOffset
        y = rect.Item("y") + yOffset
        actions.Add actionPointerMove(x, y)
        'actionPointerMove has built in pause
    Else
        actions.Add actionPointerMove(xOffset, yOffset, , , "pointer")
        'actionPointerMove has built in pause
    End If
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set MoveByOffset = Me
End Function

'@Description("Moves the mouse pointer to (x,y) relative to viewport origin (top-left)")
Public Function MoveToXY(ByVal x As Long, ByVal y As Long) As WebActionChain
Attribute MoveToXY.VB_Description = "Moves the mouse pointer to (x,y) relative to viewport origin (top-left)"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    actions.Add actionPointerMove(x, y)
    
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set MoveToXY = Me
End Function

'@Description("Sends key sequence to current focused element (see WebKeyboard class for special keys support)")
Public Function SendKeys(ByVal keys As String, Optional element As WebElement) As WebActionChain
Attribute SendKeys.VB_Description = "Sends key sequence to current focused element (see WebKeyboard class for special keys support)"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    'Dim charList As New Collection
    Dim i As Long
    Dim key As String
    
    If Not element Is Nothing Then Me.SetFocus element
    
    keys = unEscapeUnicode(keys)
    
    For i = 1 To Len(keys)
        key = Mid$(keys, i, 1)
        Select Case AscWL(key)
        Case &HE008&, &HE009&, &HE00A&
            'Shift, Ctrl, or Alt modifier key - assume sticky key
            actions.Add actionKeyDown(key)
        Case &HE000&
            'null key - release all modifier keys - since we don't
            'keep track of key state use shotgun approach
            actions.Add actionKeyUp(ChrW$(&HE008&)) 'shift
            actions.Add actionKeyUp(ChrW$(&HE009&)) 'ctrl
            actions.Add actionKeyUp(ChrW$(&HE00A&)) 'alt
        Case Else
            'normal key press down and up
            actions.Add actionKeyDown(key)
            actions.Add actionKeyUp(key)
        End Select
    Next i
    
    actions.Add actionPause

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set SendKeys = Me
End Function

'@Description("Clears the focused element's text field")
Public Function Clear(Optional element As WebElement) As WebActionChain
Attribute Clear.VB_Description = "Clears the focused element's text field"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    
    If Not element Is Nothing Then Me.SetFocus element
    
    actions.Add actionKeyDown(ChrW$(&HE009&)) 'Ctrl key
    actions.Add actionKeyDown("a")
    actions.Add actionKeyUp("a")
    actions.Add actionKeyUp(ChrW$(&HE009&))
    actions.Add actionKeyDown(ChrW$(&HE017&)) 'delete key
    actions.Add actionKeyUp(ChrW$(&HE017&))
    actions.Add actionPause

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set Clear = Me
End Function

'@Description("Sends a Ctrl-[Key] sequence to focused element")
Public Function CtrlKey(ByVal key As String, Optional element As WebElement) As WebActionChain
Attribute CtrlKey.VB_Description = "Sends a Ctrl-[Key] sequence to focused element"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    
    key = unEscapeUnicode(key)
    If Len(key) <> 1 Then Err.Raise 1, , "Error in CtrlKey method - input string must be a singular key representation"
    
    If Not element Is Nothing Then Me.SetFocus element

    actions.Add actionKeyDown(ChrW$(&HE009&)) 'Ctrl key
    actions.Add actionKeyDown(key)
    actions.Add actionKeyUp(key)
    actions.Add actionKeyUp(ChrW$(&HE009&))
    actions.Add actionPause

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set CtrlKey = Me
End Function

'@Description("Sends a Shift-[Key] sequence to focused element")
Public Function ShiftKey(ByVal key As String, Optional element As WebElement) As WebActionChain
Attribute ShiftKey.VB_Description = "Sends a Shift-[Key] sequence to focused element"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    
    key = unEscapeUnicode(key)
    If Len(key) <> 1 Then Err.Raise 1, , "Error in ShiftKey method - input string must be a singular key representation"

    If Not element Is Nothing Then Me.SetFocus element
    
    actions.Add actionKeyDown(ChrW$(&HE008&)) 'shift key
    actions.Add actionKeyDown(key)
    actions.Add actionKeyUp(key)
    actions.Add actionKeyUp(ChrW$(&HE008&))
    actions.Add actionPause

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ShiftKey = Me
End Function

'@Description("Sends a Ctrl-Shift-[Key] sequence to focused element")
Public Function CtrlShiftKey(ByVal key As String, Optional element As WebElement) As WebActionChain
Attribute CtrlShiftKey.VB_Description = "Sends a Ctrl-Shift-[Key] sequence to focused element"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    
    key = unEscapeUnicode(key)
    If Len(key) <> 1 Then Err.Raise 1, , "Error in CtrlShiftKey method - input string must be a singular key representation"

    If Not element Is Nothing Then Me.SetFocus element

    actions.Add actionKeyDown(ChrW$(&HE009&)) 'Ctrl key
    actions.Add actionKeyDown(ChrW$(&HE008&)) 'shift key
    actions.Add actionKeyDown(key)
    actions.Add actionKeyUp(key)
    actions.Add actionKeyUp(ChrW$(&HE008&))
    actions.Add actionKeyUp(ChrW$(&HE009&))
    actions.Add actionPause

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set CtrlShiftKey = Me
End Function

'@Description("Sends a Ctrl-Alt-[Key] sequence to focused element")
Public Function CtrlAltKey(ByVal key As String, Optional element As WebElement) As WebActionChain
Attribute CtrlAltKey.VB_Description = "Sends a Ctrl-Alt-[Key] sequence to focused element"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    
    key = unEscapeUnicode(key)
    If Len(key) <> 1 Then Err.Raise 1, , "Error in CtrlAltKey method - input string must be a singular key representation"
    
    If Not element Is Nothing Then Me.SetFocus element

    actions.Add actionKeyDown(ChrW$(&HE009&)) 'Ctrl key
    actions.Add actionKeyDown(ChrW$(&HE00A&)) 'alt key
    actions.Add actionKeyDown(key)
    actions.Add actionKeyUp(key)
    actions.Add actionKeyUp(ChrW$(&HE00A&))
    actions.Add actionKeyUp(ChrW$(&HE009&))
    actions.Add actionPause

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set CtrlAltKey = Me
End Function

'@Description("Sends a Alt-Shift-[Key] sequence to focused element")
Public Function AltShiftKey(ByVal key As String, Optional element As WebElement) As WebActionChain
Attribute AltShiftKey.VB_Description = "Sends a Alt-Shift-[Key] sequence to focused element"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    
    key = unEscapeUnicode(key)
    If Len(key) <> 1 Then Err.Raise 1, , "Error in CtrlShiftKey method - input string must be a singular key representation"

    If Not element Is Nothing Then Me.SetFocus element

    actions.Add actionKeyDown(ChrW$(&HE00A&)) 'alt key
    actions.Add actionKeyDown(ChrW$(&HE008&)) 'shift key
    actions.Add actionKeyDown(key)
    actions.Add actionKeyUp(key)
    actions.Add actionKeyUp(ChrW$(&HE008&))
    actions.Add actionKeyUp(ChrW$(&HE00A&))
    actions.Add actionPause

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set AltShiftKey = Me
End Function

'@Description("Scrolls to element position with optional x and y offsets")
Public Function ScrollToElement(element As WebElement, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As WebActionChain
Attribute ScrollToElement.VB_Description = "Scrolls to element position with optional x and y offsets"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    
    actions.Add actionScroll(, , xOffset, yOffset, , element.elementId)
    
    inputChan.Add "type", "wheel"
    inputChan.Add "id", "wheelDevice"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ScrollToElement = Me
End Function

'@Description("Scrolls window by x and y offsets")
Public Function ScrollBy(Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As WebActionChain
Attribute ScrollBy.VB_Description = "Scrolls window by x and y offsets"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    
    actions.Add actionScroll(, , xOffset, yOffset)
    
    inputChan.Add "type", "wheel"
    inputChan.Add "id", "wheelDevice"
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ScrollBy = Me
End Function

'@Description("Sets the focus on the given element for key input - note this clicks the element - for focus without clicking, see SetFocus method of WebDriver/WebElement classes")
Public Function SetFocus(element As WebElement) As WebActionChain
Attribute SetFocus.VB_Description = "Sets the focus on the given element for key input - note this clicks the element - for focus without clicking, see SetFocus method of WebDriver/WebElement classes"
    Dim inputChan As New Dictionary
    Dim inputChans() As Variant
    Dim actions As New Collection
    Dim parameters As New Dictionary
    
    'Note - this clicks the element to acheive focus - for focus without clicking,
    'see SetFocus method of WebDriver/WebElement classes. There does not seem to be
    'a more elegant way with actions.
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    'move to and then click
    actions.Add actionPointerMove(, , , , , element.elementId)
    actions.Add actionPointerDown
    actions.Add actionPointerUp
 
    inputChan.Add "id", "pointerDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "pointer"
    inputChan.Add "parameters", parameters
    inputChan.Add "actions", actions
    
    If data_.Exists("actions") Then
        inputChans = data_.Item("actions")
        data_.Item("actions") = appendInputChannel(inputChans, inputChan)
    Else
        data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set SetFocus = Me
End Function

'@Description("Executes the action chain")
Public Sub Perform()
Attribute Perform.VB_Description = "Executes the action chain"
    'we cannot pass a multi-chan data_.item("actions") as-is without synchronizing the input channels
    'see https://www.w3.org/TR/webdriver/#actions
    'we can either aggregate and synchronize the inputs into 3 channels (keyboard, pointer, and wheel)
    'and make one call to tCMD.CMD_ACTIONS or we execute one input channel at a time in the order
    'in which the inputs were added, making multi-calls to tCMD.CMD_ACTIONS.
    
    If data_.Exists("actions") Then
        'synchronize input channels and call Execute method once
        driver_.execute tCMD.CMD_ACTIONS, syncChannels(data_)
    End If
End Sub

'@Description("Resets the action queue and key/pointer states on both the local and remote ends")
Public Sub ResetActions()
Attribute ResetActions.VB_Description = "Resets the action queue and key/pointer states on both the local and remote ends"
    'The clear actions command is used to release all the keys and pointer buttons that are currently
    'depressed. This causes events to be fired as if the state was released by an explicit series of actions.
    'It also clears all the internal state of the virtual devices.
    driver_.execute tCMD.CMD_CLEAR_ACTIONS  'remote end clear key/pointer states
    If data_.Exists("actions") Then data_.Remove "actions" 'clear the action queue on local end
End Sub

Private Function syncChannels(data As Dictionary) As Dictionary
    'this private function aggregates an unsynchronized set of action input channels
    'to create a three-channel (keyboard, pointer, and wheel) synchronized set
    'that can be passed as a single call to the Selenium action processor
    'see https://www.w3.org/TR/webdriver/#actions
    
    Dim inputKeyChan As New Dictionary
    Dim inputMouseChan As New Dictionary
    Dim inputScrollChan As New Dictionary
    Dim inputChan As Dictionary
    Dim parameters As New Dictionary
    Dim actionsToAppend As Collection
    Dim mouseActions() As Dictionary
    Dim keyActions() As Dictionary
    Dim scrollActions() As Dictionary
    Dim hasKeyAction As Boolean
    Dim hasMouseAction As Boolean
    Dim hasScrollAction As Boolean
    Dim lastKeyAction As Long
    Dim lastMouseAction As Long
    Dim lastScrollAction As Long
    Dim inputChans() As Variant
    Dim syncdInputChans() As Variant
    Dim pause As Dictionary
    
    inputChans = data.Item("actions")
                
    If UBound(inputChans) = 0 Then 'only one channel - no need to sync
        Set syncChannels = data
        Exit Function
    End If
    
    'this is a computationally expensive step so do it once up front!!
    Set pause = actionPause(0) 'dummy pause
    
    'determine in advance the sync'd action array sizes
    Dim numActions As Long
    Dim i As Long
    numActions = 0
    For i = 0 To UBound(inputChans)
        numActions = numActions + inputChans(i)("actions").Count
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
        Set actionsToAppend = inputChan.Item("actions")
        
        Dim j As Long
        Select Case inputChan.Item("id")
        Case "pointerDevice"
            For j = 1 To actionsToAppend.Count
                numActions = numActions + 1
                Set mouseActions(numActions - 1) = actionsToAppend(j)
                Set keyActions(numActions - 1) = pause
                Set scrollActions(numActions - 1) = pause
                lastMouseAction = numActions 'last non-dummy-pause action
            Next j
            hasMouseAction = True
        Case "keyboardDevice"
            For j = 1 To actionsToAppend.Count
                numActions = numActions + 1
                Set mouseActions(numActions - 1) = pause
                Set keyActions(numActions - 1) = actionsToAppend(j)
                Set scrollActions(numActions - 1) = pause
                lastKeyAction = numActions 'last non-dummy-pause action
            Next j
            hasKeyAction = True
        Case "wheelDevice"
            For j = 1 To actionsToAppend.Count
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
    
    'create a new instance of data to preserve original unsync'd action que in case user wants to add more actions
    'on a subsequent Perform
    Dim newData As New Dictionary
    newData.Add "actions", syncdInputChans
    
    Set syncChannels = newData
End Function

'Following are private single action functions

Private Function actionKeyDown(Optional ByVal val As String) As Dictionary
    Dim action As New Dictionary
    action.Add "type", "keyDown"
    action.Add "value", val
    Set actionKeyDown = action
End Function

Private Function actionKeyUp(Optional ByVal val As String) As Dictionary
    Dim action As New Dictionary
    action.Add "type", "keyUp"
    action.Add "value", val
    Set actionKeyUp = action
End Function

Private Function actionPointerDown(Optional ByVal button As ButtonType = 0) As Dictionary
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
    
    Set actionPointerDown = action
End Function

Private Function actionPointerUp(Optional ByVal button As ButtonType = 0) As Dictionary
    Dim action As New Dictionary
    action.Add "type", "pointerUp"
    action.Add "button", button
    Set actionPointerUp = action
End Function

Private Function actionPause(Optional ByVal duration As Long = DEFAULT_PAUSE_DURATION) As Dictionary
    Dim action As New Dictionary
    action.Add "type", "pause"
    If duration > 0 Then action.Add "duration", duration
    Set actionPause = action
End Function

Private Function actionPointerCancel() As Dictionary
    Dim action As New Dictionary
    action.Add "type", "pointerCancel"
    Set actionPointerCancel = action
End Function

Private Function actionPointerMove(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal button As ButtonType = 0, Optional ByVal duration As Long = DEFAULT_PAUSE_DURATION, Optional ByVal origin As String = "viewport", Optional elementId As String) As Dictionary
    Dim action As New Dictionary
    
    action.Add "type", "pointerMove"
    action.Add "button", button
    
    If elementId = vbNullString Then
        action.Add "origin", origin  'x,y are relative to. Can be 'viewport', 'pointer' or {"element-6066-11e4-a52e-4f735466cecf": element.id}
    Else
        Dim elem As New Dictionary
        elem.Add "element-6066-11e4-a52e-4f735466cecf", elementId
        action.Add "origin", elem
    End If
    
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
    
    Set actionPointerMove = action
End Function

Private Function actionScroll(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal deltaX As Long = 0, Optional ByVal deltaY As Long = 0, Optional ByVal duration As Long = DEFAULT_PAUSE_DURATION, Optional elementId As String) As Dictionary
    Dim action As New Dictionary
    
    action.Add "type", "scroll"
    
    If elementId = vbNullString Then
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
    
    Set actionScroll = action
End Function

'this is a utility function to append an input channel to the input channel array
Private Function appendInputChannel(inputChans() As Variant, inputChan As Dictionary) As Variant()
    ReDim Preserve inputChans(0 To UBound(inputChans) + 1)
    Set inputChans(UBound(inputChans)) = inputChan
    appendInputChannel = inputChans
End Function

'helper property for passing webdriver instance
Friend Property Set WebDriver(obj As WebDriver)
    Set driver_ = obj
End Property

Private Sub Class_Initialize()
    With tCMD
        .CMD_ACTIONS = Array("POST", "/session/$sessionId/actions")
        .CMD_CLEAR_ACTIONS = Array("DELETE", "/session/$sessionId/actions")
    End With
    Set data_ = New Dictionary
End Sub
