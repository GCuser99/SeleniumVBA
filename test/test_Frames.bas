Attribute VB_Name = "test_Frames"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_frames_with_frameset()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim htmlStr As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    driver.StartChrome
    driver.OpenBrowser
    
    'save content for top frame
    htmlStr = "<html><body><div class='myDiv'><h2>This is the top frame source</h2></div></body></html>"
    driver.SaveStringToFile htmlStr, ".\snippettop.html"
    
    'save content for bottom frame
    htmlStr = "<html><body><div class='myDiv'><h2>This is the bottom frame source</h2></div></body></html>"
    driver.SaveStringToFile htmlStr, ".\snippetbottom.html"
    
    'save the main snippet
    htmlStr = "<html><div><frameset rows='50%,50%'><frame name='top' id='topid' src='./snippettop.html'/><frame name='bottom' id='bottomid' src='./snippetbottom.html'/><noframes><body>Your browser does not support frames.</body></noframes></frameset></div></html>"
    driver.SaveStringToFile htmlStr, ".\snippet.html"
    
    driver.NavigateToFile ".\snippet.html"
    driver.Wait
    
    Debug.Assert driver.ExecuteScript("return window.length") = 2 'this includes embed, iframes, frames objects
    Debug.Assert driver.FindElements(By.TagName, "frame").Count = 2
    
    Set elem = driver.FindElementByName("bottom")
    
    driver.SwitchToFrame elem
    driver.Wait
    Debug.Assert driver.GetCurrentFrameName = "bottom"
    
    driver.SwitchToDefaultContent 'must move up the tree to see sibling frame
    driver.Wait
    
    driver.SwitchToFrameByIndex 1
    driver.Wait
    Debug.Assert driver.GetCurrentFrameName = "top"
    
    driver.DeleteFiles ".\snippettop.html", ".\snippetbottom.html", ".\snippet.html"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_frames_with_embed_objects()
    Dim driver As SeleniumVBA.WebDriver
    Dim elemObject As SeleniumVBA.WebElement, elemEmbed As SeleniumVBA.WebElement
    Dim htmlStr As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartEdge
    driver.OpenBrowser
    
    'save content for top frame
    htmlStr = "<html><body><div class='myDiv'><h2>This is the top frame source</h2></div></body></html>"
    driver.SaveStringToFile htmlStr, ".\snippettop.html"
    
    'save content for bottom frame
    htmlStr = "<html><body><div class='myDiv'><h2>This is the bottom frame source</h2></div></body></html>"
    driver.SaveStringToFile htmlStr, ".\snippetbottom.html"
    
    'save the main snippet
    htmlStr = "<html><body><div><embed name='embed frame' type='text/html' src='./snippettop.html' width='500' height='200'></div><div><object name='object frame' data='./snippetbottom.html' width='500' height='200'></object></div></body></html>"
    driver.SaveStringToFile htmlStr, ".\snippet.html"
    
    driver.NavigateToFile ".\snippet.html"
    driver.Wait 1000
    
    Debug.Assert driver.ExecuteScript("return window.length") = 2 'this includes embed, iframes, frames objects
    Debug.Assert driver.FindElements(By.TagName, "embed").Count + driver.FindElements(By.TagName, "object").Count = 2
    
    Set elemObject = driver.FindElementByName("object frame")
    Set elemEmbed = driver.FindElementByName("embed frame")
    
    driver.SwitchToFrame elemEmbed
    driver.Wait
    Debug.Assert driver.GetCurrentFrameName = "embed frame"
    
    driver.SwitchToDefaultContent 'must move up the tree to see sibling frame
    driver.Wait
    
    'unfortunately, for embedded objects, switching to frame by index does not work
    'Driver.SwitchToFrameByIndex 1
    'Driver.Wait
    'Debug.Print "Switch by index to frame: " & Driver.GetCurrentFrameName
    
    driver.SwitchToFrame elemObject
    driver.Wait
    Debug.Assert driver.GetCurrentFrameName = "object frame"
    
    driver.DeleteFiles ".\snippettop.html", ".\snippetbottom.html", ".\snippet.html"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_frames_with_iframes()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim htmlStr As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    driver.StartChrome
    driver.OpenBrowser
    
    'save content for top frame
    htmlStr = "<html><body><div class='myDiv'><h2>This is the top frame source</h2></div></body></html>"
    driver.SaveStringToFile htmlStr, ".\snippettop.html"
    
    'save content for bottom frame
    htmlStr = "<html><body><div class='myDiv'><h2>This is the bottom frame source</h2></div></body></html>"
    driver.SaveStringToFile htmlStr, ".\snippetbottom.html"
    
    'save the main snippet
    htmlStr = "<html><body><div class='box'><iframe name='iframe1' id='IF1' height='50%' width='50%' src='./snippettop.html'></iframe></div>  <div class='box'><iframe name='iframe2' id='IF2' height='50%' width='50%'  align='left' src='.\snippetbottom.html'></iframe></div></body></html>"
    driver.SaveStringToFile htmlStr, ".\snippet.html"
    
    driver.NavigateToFile ".\snippet.html"
    driver.Wait 1000
    
    Debug.Assert driver.ExecuteScript("return window.length") = 2 'this includes embed, iframes, frames objects
    Debug.Assert driver.FindElements(By.TagName, "iframe").Count = 2
    
    Set elem = driver.FindElementByName("iframe2")
    
    driver.SwitchToFrame elem
    driver.Wait
    Debug.Assert driver.GetCurrentFrameName = "iframe2"
    
    driver.SwitchToDefaultContent 'must move up the tree to see sibling frame
    driver.Wait
    
    driver.SwitchToFrameByIndex 1
    driver.Wait
    Debug.Assert driver.GetCurrentFrameName = "iframe1"
    
    driver.DeleteFiles ".\snippettop.html", ".\snippetbottom.html", ".\snippet.html"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_frames_with_nested_iframes()
    Dim driver As SeleniumVBA.WebDriver
    Dim elems As SeleniumVBA.WebElements, elem As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://demoqa.com/nestedframes"
    driver.Wait 1000
    
    Debug.Assert driver.ExecuteScript("return window.length") > 0 'this includes embed, iframes, frames objects
    Debug.Assert driver.FindElements(By.TagName, "iframe").Count > 0
    
    Set elem = driver.FindElementByID("frame1") 'cant find this element
    
    'switch to top-level (parent) frame
    driver.SwitchToFrame elem
    driver.Wait
    Debug.Assert driver.FindElementByTagName("body").GetText = "Parent frame"
    Debug.Assert driver.FindElements(By.TagName, "iframe").Count = 1
    
    'switch to child frame
    driver.SwitchToFrameByIndex 1
    driver.Wait
    Debug.Assert driver.FindElementByTagName("body").GetText = "Child Iframe"
    
    'switch back to top-level (parent) frame
    driver.SwitchToParentFrame 'must move up the tree to see sibling frame
    driver.Wait
    Debug.Assert driver.FindElementByTagName("body").GetText = "Parent frame"
    
    'switch to main document
    driver.SwitchToDefaultContent
    driver.Wait
    
    driver.CloseBrowser
    driver.Shutdown
End Sub


