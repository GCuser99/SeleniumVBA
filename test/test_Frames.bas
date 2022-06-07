Attribute VB_Name = "test_Frames"
Sub test_frames_with_frameset()
    Dim driver As New WebDriver
    Dim elem As WebElement

    driver.StartChrome
    driver.OpenBrowser
    
    'save content for top frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the top frame source</h2></div></body></html>"
    filePath = ".\snippettop.html"
    driver.SaveHTMLToFile htmlstr, filePath
    
    'save content for bottom frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the bottom frame source</h2></div></body></html>"
    filePath = ".\snippetbottom.html"
    driver.SaveHTMLToFile htmlstr, filePath
    
    'save the main snippet
    htmlstr = "<html><div><frameset rows='50%,50%'><frame name='top' id='topid' src='./snippettop.html'/><frame name='bottom' id='bottomid' src='./snippetbottom.html'/><noframes><body>Your browser does not support frames.</body></noframes></frameset></div></html>"
    filePath = ".\snippet.html"
    driver.SaveHTMLToFile htmlstr, filePath
    
    driver.NavigateTo "file:///" & filePath
    driver.Wait
    
    Debug.Print "Number of windows: " & driver.ExecuteScript("return window.length") 'this includes embed, iframes, frames objects
    Debug.Print "Number of frames: " & driver.FindElements(by.tagName, "frame").Count
    
    Set elem = driver.FindElementByName("bottom")
    
    driver.SwitchToFrame elem
    driver.Wait
    Debug.Print "Switch by element to frame: " & driver.GetCurrentFrameName
    
    driver.SwitchToDefaultContent 'must move up the tree to see sibling frame
    driver.Wait
    
    driver.SwitchToFrameByIndex 1
    driver.Wait
    Debug.Print "Switch by index to frame: " & driver.GetCurrentFrameName
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_frames_with_embed_objects()
    Dim driver As New WebDriver
    Dim elemObject As WebElement, elemEmbed As WebElement

    driver.StartEdge
    driver.OpenBrowser
    
    'save content for top frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the top frame source</h2></div></body></html>"
    filePath = ".\snippettop.html"
    driver.SaveHTMLToFile htmlstr, filePath
    
    'save content for bottom frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the bottom frame source</h2></div></body></html>"
    filePath = ".\snippetbottom.html"
    driver.SaveHTMLToFile htmlstr, filePath
    
    'save the main snippet
    htmlstr = "<html><body><div><embed name='embed frame' type='text/html' src='./snippettop.html' width='500' height='200'></div><div><object name='object frame' data='./snippetbottom.html' width='500' height='200'></object></div></body></html>"
    filePath = ".\snippet.html"
    driver.SaveHTMLToFile htmlstr, filePath
    
    driver.NavigateTo "file:///" & filePath
    driver.Wait 1000
    
    Debug.Print "Number of windows: " & driver.ExecuteScript("return window.length") 'this includes embed, iframes, frames objects
    Debug.Print "Number of frames: " & driver.FindElements(by.tagName, "embed").Count + driver.FindElements(by.tagName, "object").Count
    
    Set elemObject = driver.FindElementByName("object frame")
    Set elemEmbed = driver.FindElementByName("embed frame")
    
    driver.SwitchToFrame elemEmbed
    driver.Wait
    Debug.Print "Switch by element to frame: " & driver.GetCurrentFrameName
    
    driver.SwitchToDefaultContent 'must move up the tree to see sibling frame
    driver.Wait
    
    'unfortunately, for embedded objects, switching to frame by index does not work
    'Driver.SwitchToFrameByIndex 1
    'Driver.Wait
    'Debug.Print "Switch by index to frame: " & Driver.GetCurrentFrameName
    
    driver.SwitchToFrame elemObject
    driver.Wait
    Debug.Print "Switch by element to frame: " & driver.GetCurrentFrameName
    
    driver.CloseBrowser
    driver.Shutdown
    
End Sub

Sub test_frames_with_iframes()
    Dim driver As New WebDriver
    Dim elem As WebElement

    driver.StartChrome
    driver.OpenBrowser
    
    'save content for top frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the top frame source</h2></div></body></html>"
    filePath = ".\snippettop.html"
    driver.SaveHTMLToFile htmlstr, filePath
    
    'save content for bottom frame
    htmlstr = "<html><body><div class='myDiv'><h2>This is the bottom frame source</h2></div></body></html>"
    filePath = ".\snippetbottom.html"
    driver.SaveHTMLToFile htmlstr, filePath
    
    'save the main snippet
    htmlstr = "<html><body><div class='box'><iframe name='iframe1' id='IF1' height='50%' width='50%' src='./snippettop.html'></div></iframe>  <div class='box'><iframe name='iframe2' id='IF2' height='50%' width='50%'  align='left' src='.\snippetbottom.html'></iframe></div></body></html>"
    filePath = ".\snippet.html"
    driver.SaveHTMLToFile htmlstr, filePath
    
    driver.NavigateTo "file:///" & filePath
    driver.Wait 1000
    
    Debug.Print "Number of windows: " & driver.ExecuteScript("return window.length") 'this includes embed, iframes, frames objects
    Debug.Print "Number of frames: " & driver.FindElements(by.tagName, "iframe").Count
    
    Set elem = driver.FindElementByName("iframe2")
    
    driver.SwitchToFrame elem
    driver.Wait
    Debug.Print "Switch by element to frame: " & driver.GetCurrentFrameName
    
    driver.SwitchToDefaultContent 'must move up the tree to see sibling frame
    driver.Wait
    
    driver.SwitchToFrameByIndex 1
    driver.Wait
    Debug.Print "Switch by index to frame: " & driver.GetCurrentFrameName
    
    driver.CloseBrowser
    driver.Shutdown
    
End Sub

Sub test_frames_with_nested_iframes()
    Dim driver As New WebDriver
    Dim elems As WebElements, elem As WebElement

    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://demoqa.com/nestedframes"
    driver.Wait 1000
    
    Debug.Print "Number of windows: " & driver.ExecuteScript("return window.length") 'this includes embed, iframes, frames objects
    Debug.Print "Number of frames: " & driver.FindElements(by.tagName, "iframe").Count
    
    Set elem = driver.FindElementByID("frame1") 'cant find this element
    
    'switch to top-level (parent) frame
    driver.SwitchToFrame elem
    driver.Wait
    Debug.Print "Parent frame text: " & driver.FindElementByTagName("body").GetText
    Debug.Print "Number of child frames: " & driver.FindElements(by.tagName, "iframe").Count
    
    'switch to child frame
    driver.SwitchToFrameByIndex 1
    driver.Wait
    Debug.Print "Child frame text: " & driver.FindElementByTagName("body").GetText
    
    'switch back to top-level (parent) frame
    driver.SwitchToParentFrame 'must move up the tree to see sibling frame
    driver.Wait
    Debug.Print "Parent frame text: " & driver.FindElementByTagName("body").GetText
    
    'switch to main document
    driver.SwitchToDefaultContent
    driver.Wait
    
    driver.CloseBrowser
    driver.Shutdown
    
End Sub

