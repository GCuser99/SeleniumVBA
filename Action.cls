VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Action"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Action.cls
'
'
'This class was written to emulate Selenium Webdriver ActionChains
'
'USAGE(S):
'
'Method 1:
'   Dim Actions as Action
'   Set elememail = Driver.FindElement(by.name, "email")
'   Set Actions = Driver.ActionChain
'   Actions.MoveToElement(elememail).Click
'   Actions.SendKeys("my@email.com").Perform
'
'Method 2:
'   Set elememail = Driver.FindElement(by.name, "email")
'   Driver.ActionChain.MoveToElement(elememail).Click.SendKeys("my@email.com").Perform
'
'   Please note that entire chain + perform must be on one line for Method 2 to work
'
'NOTES:
'
'   To pass a particular sessionId to the chain:
'
'   Set Actions = Driver.ActionChain(sessionId)
'       ...or...
'   Driver.ActionChain(sessionId).MoveToElement(elememail).Click.SendKeys("my@email.com").Perform
'
'   Note: I have designed this version to clear actions that are already stored locally and on the remote end
'         after the Perform Method is executed. But if compatibility with Selenium Python is desired
'         then make ResetActions public, and comment out the same in Perform method.
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
'SendKeys               Sends keys to current focused element
'ShiftKey               Sends a Shift-[Key] sequence to focused element
'Wait                   Pause all inputs for the specified duration in milliseconds

Private Const DEFAULT_PAUSE_DURATION = 250
Private Driver_ As WebDriver
Private Data_ As Dictionary

Private CMD_W3C_ACTIONS
Private CMD_W3C_CLEAR_ACTIONS

Private Enum ButtonType
    leftButton = 0
    middleButton = 1
    rightButton = 2
End Enum

'These are helper properties for passing webdriver and data dictionary references to each link of chain
Friend Property Set WebDriver(ByRef obj As WebDriver)
    Set Driver_ = obj
End Property

Friend Property Set data(ByRef obj As Dictionary)
    Set Data_ = obj
End Property

'Action Chain Methods:

Public Function MoveToElement(element As WebElement, Optional setFocus As Boolean = False) As Action
    'Moves the mouse pointer to the middle of an element.
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    Dim parameters As New Dictionary
    
    parameters.Add "pointerType", "mouse" 'required for 'pointer' inputs
    
    If setFocus Then 'add click after move
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set MoveToElement = New Action
    Set MoveToElement.WebDriver = Driver_
    Set MoveToElement.data = Data_
    
End Function

Public Function Click(Optional element As WebElement) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set Click = New Action
    Set Click.WebDriver = Driver_
    Set Click.data = Data_
    
End Function

Public Function ClickAndHold(Optional element As WebElement) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ClickAndHold = New Action
    Set ClickAndHold.WebDriver = Driver_
    Set ClickAndHold.data = Data_
    
End Function

Public Function ClickContext(Optional element As WebElement) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ClickContext = New Action
    Set ClickContext.WebDriver = Driver_
    Set ClickContext.data = Data_
    
End Function

Public Function DoubleClick(Optional element As WebElement) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DoubleClick = New Action
    Set DoubleClick.WebDriver = Driver_
    Set DoubleClick.data = Data_
    
End Function

Public Function ReleaseButton(Optional element As WebElement) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ReleaseButton = New Action
    Set ReleaseButton.WebDriver = Driver_
    Set ReleaseButton.data = Data_
    
End Function

Public Function DragAndDrop(sourceElement As WebElement, targetElement As WebElement) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DragAndDrop = New Action
    Set DragAndDrop.WebDriver = Driver_
    Set DragAndDrop.data = Data_
    
End Function

Public Function DragAndDropByOffset(element As WebElement, ByVal xOffset As Long, ByVal yOffset As Long) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DragAndDropByOffset = New Action
    Set DragAndDropByOffset.WebDriver = Driver_
    Set DragAndDropByOffset.data = Data_
    
End Function

Public Function DragAndDropToPosition(element As WebElement, ByVal x As Long, ByVal y As Long) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set DragAndDropToPosition = New Action
    Set DragAndDropToPosition.WebDriver = Driver_
    Set DragAndDropToPosition.data = Data_
    
End Function


Public Function KeyDown(ByVal modifierKey As String, Optional element As WebElement) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set KeyDown = New Action
    Set KeyDown.WebDriver = Driver_
    Set KeyDown.data = Data_
    
End Function

Public Function KeyUp(ByVal modifierKey As String, Optional element As WebElement) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set KeyUp = New Action
    Set KeyUp.WebDriver = Driver_
    Set KeyUp.data = Data_
    
End Function

Public Function Wait(Optional ByVal durationms As Long = DEFAULT_PAUSE_DURATION) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set Wait = New Action
    Set Wait.WebDriver = Driver_
    Set Wait.data = Data_
    
End Function

Public Function MoveByOffset(ByVal xOffset As Long, ByVal yOffset As Long, Optional element As WebElement) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set MoveByOffset = New Action
    Set MoveByOffset.WebDriver = Driver_
    Set MoveByOffset.data = Data_

End Function

Public Function SendKeys(ByVal Keys As String, Optional element As WebElement) As Action
    'Sends keys to an element
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary

    'this false if keys=""
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set SendKeys = New Action
    Set SendKeys.WebDriver = Driver_
    Set SendKeys.data = Data_
    
End Function

Public Function Clear(Optional element As WebElement) As Action
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
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set Clear = New Action
    Set Clear.WebDriver = Driver_
    Set Clear.data = Data_

End Function

Public Function CtrlKey(ByVal Key As String, Optional element As WebElement) As Action
    'Sends a Ctrl-[key] down and up
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If

    ReDim actions(0 To 3)
    Set actions(0) = ActionKeyDown("\ue009")
    Set actions(1) = ActionKeyDown(Key)
    Set actions(2) = ActionKeyUp(Key)
    Set actions(3) = ActionKeyUp("\ue009")

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set CtrlKey = New Action
    Set CtrlKey.WebDriver = Driver_
    Set CtrlKey.data = Data_

End Function

Public Function ShiftKey(ByVal Key As String, Optional element As WebElement) As Action
    'Sends a Shift-[key] down and up
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    If Not element Is Nothing Then
        Me.Click element  'get focus
    End If
    
    ReDim actions(0 To 3)
    Set actions(0) = ActionKeyDown("\ue008")
    Set actions(1) = ActionKeyDown(Key)
    Set actions(2) = ActionKeyUp(Key)
    Set actions(3) = ActionKeyUp("\ue008")

    inputChan.Add "id", "keyboardDevice" 'for a given device, this needs to be same for entire chain
    inputChan.Add "type", "key"
    inputChan.Add "actions", actions
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ShiftKey = New Action
    Set ShiftKey.WebDriver = Driver_
    Set ShiftKey.data = Data_
    
End Function

Public Function ScrollToElement(element As WebElement, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As Action
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    ReDim actions(0 To 0)
    Set actions(0) = ActionScroll(, , xOffset, yOffset, , element.elementId)
    
    inputChan.Add "type", "wheel"
    inputChan.Add "id", "wheelDevice"
    inputChan.Add "actions", actions
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ScrollToElement = New Action
    Set ScrollToElement.WebDriver = Driver_
    Set ScrollToElement.data = Data_

End Function

Public Function ScrollBy(Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0) As Action
    Dim inputChan As New Dictionary, inputChans() As Variant
    Dim actions() As Dictionary
    
    ReDim actions(0 To 0)
    Set actions(0) = ActionScroll(, , xOffset, yOffset)
    
    inputChan.Add "type", "wheel"
    inputChan.Add "id", "wheelDevice"
    inputChan.Add "actions", actions
    
    If Data_.Exists("actions") Then
        inputChans = Data_("actions")
        Data_("actions") = AppendInputChannel(inputChans, inputChan)
    Else
        Data_.Add "actions", Array(inputChan)
    End If
    
    'now set things up for next link in chain
    Set ScrollBy = New Action
    Set ScrollBy.WebDriver = Driver_
    Set ScrollBy.data = Data_

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
    
    If Data_.Exists("actions") Then
        
        'multi-call method:
        'inputChans = Data_("actions")
        'For i = 0 To UBound(inputChans)
        '    Data_("actions") = Array(inputChans(i))
        '    Driver_.Execute Driver_.CMD_W3C_ACTIONS, Data_
        'Next i
        
        '... or synchronize input channels and call Execute once method
        'for typical action chains, this is 5-10% faster...
        
        Driver_.Execute CMD_W3C_ACTIONS, SyncChannels(Data_)
        
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
    Dim pause As Dictionary
    
    inputChans = data("actions")
                
    If UBound(inputChans) = 0 Then 'only one channel - no need to sync
        Set SyncChannels = data
        Exit Function
    End If
    
    'this is a computationally expensive step so do it once up front!!
    Set pause = ActionPause(0)
    
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
                Set keyActions(numActions - 1) = pause
                Set scrollActions(numActions - 1) = pause
                lastMouseAction = numActions 'last non-pause action
            Next j
            hasMouseAction = True
        Case "keyboardDevice"
            For j = 0 To UBound(actionsToAppend)
                numActions = numActions + 1
                Set mouseActions(numActions - 1) = pause
                Set keyActions(numActions - 1) = actionsToAppend(j)
                Set scrollActions(numActions - 1) = pause
                lastKeyAction = numActions 'last non-pause action
            Next j
            hasKeyAction = True
        Case "wheelDevice"
            For j = 0 To UBound(actionsToAppend)
                numActions = numActions + 1
                Set mouseActions(numActions - 1) = pause
                Set keyActions(numActions - 1) = pause
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
    Driver_.Execute CMD_W3C_CLEAR_ACTIONS, Data_ 'remote end clear
    Data_.Remove "actions" 'local end clear
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
