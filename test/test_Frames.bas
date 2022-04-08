Attribute VB_Name = "test_Frames"
Sub test_frames_with_frameset()
    Dim Driver As New WebDriver
    Dim elem As WebElement
    
    Driver.StartEdge
    Driver.OpenBrowser
    
    'save content for top frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the top frame source</h2></div></body></html>"
    filepath = ".\snippettop.html"
    Driver.SaveHTMLToFile htmlstr, filepath
    
    'save content for bottom frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the bottom frame source</h2></div></body></html>"
    filepath = ".\snippetbottom.html"
    Driver.SaveHTMLToFile htmlstr, filepath
    
    'save the main snippet
    htmlstr = "<html><div><frameset rows='50%,50%'><frame name='top' id='topid' src='./snippettop.html'/><frame name='bottom' id='bottomid' src='./snippetbottom.html'/><noframes><body>Your browser does not support frames.</body></noframes></frameset></div></html>"
    filepath = ".\snippet.html"
    Driver.SaveHTMLToFile htmlstr, filepath
    
    Driver.NavigateTo "file:///" & filepath
    Driver.Wait
    
    Debug.Print "Number of windows: " & Driver.ExecuteScript("return window.length") 'this includes embed, iframes, frames objects
    Debug.Print "Number of frames: " & Driver.FindElements(by.tagName, "frame").Count
    
    Set elem = Driver.FindElementByName("bottom")
    
    Driver.SwitchToFrame elem
    Driver.Wait
    Debug.Print "Switch by element to frame: " & Driver.GetCurrentFrameName
    
    Driver.SwitchToDefaultContent 'must move up the tree to see sibling frame
    Driver.Wait
    
    Driver.SwitchToFrameByIndex 1
    Driver.Wait
    Debug.Print "Switch by index to frame: " & Driver.GetCurrentFrameName
    
    Driver.CloseBrowser
    Driver.Shutdown
End Sub

Sub test_frames_with_embed_objects()
    Dim Driver As New WebDriver
    Dim elemObject As WebElement, elemEmbed As WebElement
    
    Driver.StartChrome
    Driver.OpenBrowser
    
    'save content for top frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the top frame source</h2></div></body></html>"
    filepath = ".\snippettop.html"
    Driver.SaveHTMLToFile htmlstr, filepath
    
    'save content for bottom frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the bottom frame source</h2></div></body></html>"
    filepath = ".\snippetbottom.html"
    Driver.SaveHTMLToFile htmlstr, filepath
    
    'save the main snippet
    htmlstr = "<html><body><div><embed name='embed frame' type='text/html' src='./snippettop.html' width='500' height='200'></div><div><object name='object frame' data='./snippetbottom.html' width='500' height='200'></object></div></body></html>"
    filepath = ".\snippet.html"
    Driver.SaveHTMLToFile htmlstr, filepath
    
    Driver.NavigateTo "file:///" & filepath
    Driver.Wait 1000
    
    Debug.Print "Number of windows: " & Driver.ExecuteScript("return window.length") 'this includes embed, iframes, frames objects
    Debug.Print "Number of frames: " & Driver.FindElements(by.tagName, "embed").Count + Driver.FindElements(by.tagName, "object").Count
    
    Set elemObject = Driver.FindElementByName("object frame")
    Set elemEmbed = Driver.FindElementByName("embed frame")
    
    Driver.SwitchToFrame elemEmbed
    Driver.Wait
    Debug.Print "Switch by element to frame: " & Driver.GetCurrentFrameName
    
    Driver.SwitchToDefaultContent 'must move up the tree to see sibling frame
    Driver.Wait
    
    'unfortunately, for embedded objects, switching to frame by index does not work
    'Driver.SwitchToFrameByIndex 1
    'Driver.Wait
    'Debug.Print "Switch by index to frame: " & Driver.GetCurrentFrameName
    
    Driver.SwitchToFrame elemObject
    Driver.Wait
    Debug.Print "Switch by element to frame: " & Driver.GetCurrentFrameName
    
    Driver.CloseBrowser
    Driver.Shutdown
    
End Sub

Sub test_frames_with_iframes()
    Dim Driver As New WebDriver
    Dim elem As WebElement
    
    Driver.StartChrome
    Driver.OpenBrowser
    
    'save content for top frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the top frame source</h2></div></body></html>"
    filepath = ".\snippettop.html"
    Driver.SaveHTMLToFile htmlstr, filepath
    
    'save content for bottom frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the bottom frame source</h2></div></body></html>"
    filepath = ".\snippetbottom.html"
    Driver.SaveHTMLToFile htmlstr, filepath
    
    'save the main snippet
    htmlstr = "<html><body><div class='box'><iframe name='iframe1' id='IF1' height='50%' width='50%' src='./snippettop.html'></div></iframe>  <div class='box'><iframe name='iframe2' id='IF2' height='50%' width='50%'  align='left' src='.\snippetbottom.html'></iframe></div></body></html>"
    filepath = ".\snippet.html"
    Driver.SaveHTMLToFile htmlstr, filepath
    
    Driver.NavigateTo "file:///" & filepath
    Driver.Wait 1000
    
    Debug.Print "Number of windows: " & Driver.ExecuteScript("return window.length") 'this includes embed, iframes, frames objects
    Debug.Print "Number of frames: " & Driver.FindElements(by.tagName, "iframe").Count
    
    Set elem = Driver.FindElementByName("iframe2")
    
    Driver.SwitchToFrame elem
    Driver.Wait
    Debug.Print "Switch by element to frame: " & Driver.GetCurrentFrameName
    
    Driver.SwitchToDefaultContent 'must move up the tree to see sibling frame
    Driver.Wait
    
    Driver.SwitchToFrameByIndex 1
    Driver.Wait
    Debug.Print "Switch by index to frame: " & Driver.GetCurrentFrameName
    
    Driver.CloseBrowser
    Driver.Shutdown
    
End Sub

Sub test_frames_with_nested_iframes()
    Dim Driver As New WebDriver
    Dim elems As WebElements, elem As WebElement
    
    Driver.StartChrome
    Driver.OpenBrowser
    
    Driver.NavigateTo "https://demoqa.com/nestedframes"
    Driver.Wait 1000
    
    Debug.Print "Number of windows: " & Driver.ExecuteScript("return window.length") 'this includes embed, iframes, frames objects
    Debug.Print "Number of frames: " & Driver.FindElements(by.tagName, "iframe").Count
    
    Set elem = Driver.FindElementByID("frame1")
    
    'switch to top-level (parent) frame
    Driver.SwitchToFrame elem
    Driver.Wait
    Debug.Print "Parent frame text: " & Driver.FindElementByTagName("body").GetText
    Debug.Print "Number of child frames: " & Driver.FindElements(by.tagName, "iframe").Count
    
    'switch to child frame
    Driver.SwitchToFrameByIndex 1
    Driver.Wait
    Debug.Print "Child frame text: " & Driver.FindElementByTagName("body").GetText
    
    'switch back to top-level (parent) frame
    Driver.SwitchToParentFrame 'must move up the tree to see sibling frame
    Driver.Wait
    Debug.Print "Parent frame text: " & Driver.FindElementByTagName("body").GetText
    
    'switch to main document
    Driver.SwitchToDefaultContent
    Driver.Wait
    
    Driver.CloseBrowser
    Driver.Shutdown
    
End Sub

