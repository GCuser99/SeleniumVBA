Attribute VB_Name = "test_Frames"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_frames_with_frameset()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim html As String
    
    html = vbNullString
    html = html & "<html lang=""en"">" & vbCrLf
    html = html & "    <head>" & vbCrLf
    html = html & "        <title>Test Frameset</title>" & vbCrLf
    html = html & "        <script type='text/javascript'>" & vbCrLf
    html = html & "            var contents_of_frame1 = '<html><head><title>Top</title></head><body><div><h2>Top</h2></div></body></html>';" & vbCrLf
    html = html & "            var contents_of_frame2 = '<html><head><title>Bottom</title></head><body><div><h2>Bottom</h2></div></body></html>';" & vbCrLf
    html = html & "            var contents_of_frame3 = '<html><head><title>Side</title></head><body><div><h2>Side</h2></div></body></html>';" & vbCrLf
    html = html & "        </script>" & vbCrLf
    html = html & "    </head>" & vbCrLf
    html = html & "    <frameset cols=""20%, 80%"">" & vbCrLf
    html = html & "        <frameset rows=""100, 200"">" & vbCrLf
    html = html & "        <frame name=""top"" title=""top"" src=""javascript:top.contents_of_frame1""/>" & vbCrLf
    html = html & "        <frame name=""bottom"" title=""bottom"" src=""javascript:top.contents_of_frame2""/>" & vbCrLf
    html = html & "        </frameset>" & vbCrLf
    html = html & "    <frame name=""side"" title=""side"" src=""javascript:top.contents_of_frame3""/>" & vbCrLf
    html = html & "    </frameset>" & vbCrLf
    html = html & "</html>"
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateToString html
    driver.Wait
    
    Debug.Assert driver.ExecuteScript("return window.length") = 3 'this includes embed, iframes, frames objects
    Debug.Assert driver.FindElements(By.TagName, "frame").Count = 3
    
    Set elem = driver.FindElementByName("bottom")
    
    driver.SwitchToFrame elem
    driver.Wait
    Debug.Assert driver.GetCurrentFrameName = "bottom"
    
    driver.SwitchToDefaultContent 'must move up the tree to see sibling frame
    driver.Wait
    
    driver.SwitchToFrameByIndex 1
    driver.Wait
    Debug.Assert driver.GetCurrentFrameName = "top"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_frames_with_embed_objects()
    Dim driver As SeleniumVBA.WebDriver
    Dim elemObject As SeleniumVBA.WebElement, elemEmbed As SeleniumVBA.WebElement
    Dim html As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser

    html = vbNullString
    html = html & "<html lang=""en"">" & vbCrLf
    html = html & "    <head>" & vbCrLf
    html = html & "        <title>Test Embedded Objects</title>" & vbCrLf
    html = html & "    </head>" & vbCrLf
    html = html & "    <body>" & vbCrLf
    html = html & "        <div>" & vbCrLf
    html = html & "            <embed name=""embed frame"" title=""embed frame"" type=""text/html"" src=""data:text/html,<html><body><div><h2>Embedded%20Frame</h2></div></body></html>"" width=""500"" height=""200"" >" & vbCrLf
    html = html & "        </div>" & vbCrLf
    html = html & "        <div>" & vbCrLf
    html = html & "            <object name=""object frame"" title=""object frame"" data=""data:text/html,<html><body><div><h2>Object%20Frame</h2></div></body></html>"" width=""500"" height=""200""></object>" & vbCrLf
    html = html & "        </div>" & vbCrLf
    html = html & "    </body>" & vbCrLf
    html = html & "</html>"
    
    driver.NavigateToString html
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
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_frames_with_iframes()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim html As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    html = vbNullString
    html = html & "<html lang=""en"">" & vbCrLf
    html = html & "    <head>" & vbCrLf
    html = html & "        <title>Test IFrames</title>" & vbCrLf
    html = html & "    </head>" & vbCrLf
    html = html & "    <body>" & vbCrLf
    html = html & "        <div class=""box"">" & vbCrLf
    html = html & "            <iframe name=""iframe1"" title=""iframe1"" id=""iframe1"" height=""50%"" width=""50%"" srcdoc=""<html><body><div class='myDiv'><h2>This is the top frame</h2></div></body></html>"">" & vbCrLf
    html = html & "            </iframe>" & vbCrLf
    html = html & "        </div>  " & vbCrLf
    html = html & "        <div class=""box"">" & vbCrLf
    html = html & "            <iframe name=""iframe2"" title=""iframe2"" id=""iframe2"" height=""50%"" width=""50%"" align=""left"" srcdoc=""<html><body><div class='myDiv'><h2>This is the bottom frame</h2></div></body></html>"">" & vbCrLf
    html = html & "            </iframe>" & vbCrLf
    html = html & "        </div>" & vbCrLf
    html = html & "    </body>" & vbCrLf
    html = html & "</html>" & vbCrLf
    
    driver.NavigateToString html
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
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_frames_with_nested_iframes()
    Dim driver As SeleniumVBA.WebDriver
    Dim elems As SeleniumVBA.WebElements, elem As SeleniumVBA.WebElement
    Dim html As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    html = vbNullString
    html = html & "<html lang=""en"">" & vbCrLf
    html = html & "    <head>" & vbCrLf
    html = html & "        <title>Test Nested IFrames</title>" & vbCrLf
    html = html & "    </head>" & vbCrLf
    html = html & "    <body>" & vbCrLf
    html = html & "        <div id=""top-level host"" class=""box"">" & vbCrLf
    html = html & "            <iframe name=""parent"" title=""parent"" id=""parent"" height=""50%"" width=""50%"" srcdoc=""<html><body><div class='myDiv'><h2>This is the parent frame</h2></div><div class='box'><iframe name='child' title='child' id='child' height='50%' width='50%' align='left' srcdoc='<html><body><div class=`myDiv`><h2>This is the nested child frame</h2></div></body></html>'></iframe></div></body></html>""></iframe>" & vbCrLf
    html = html & "        </div>" & vbCrLf
    html = html & "    </body>" & vbCrLf
    html = html & "</html>"
    
    driver.NavigateToString html
    driver.Wait 1000
    
    Debug.Assert driver.ExecuteScript("return window.length") = 1 'this includes embed, iframes, frames objects, but not nested frames
    Debug.Assert driver.FindElements(By.TagName, "iframe").Count = 1
    
    Set elem = driver.FindElementByID("parent")
    
    'switch to top-level (parent) frame
    driver.SwitchToFrame elem
    driver.Wait
    Debug.Assert driver.FindElementByTagName("body").GetText = "This is the parent frame"
    Debug.Assert driver.FindElements(By.TagName, "iframe").Count = 1
    
    'switch to child frame
    driver.SwitchToFrameByIndex 1
    driver.Wait
    Debug.Assert driver.FindElementByTagName("body").GetText = "This is the nested child frame"
    
    'switch back to top-level (parent) frame
    driver.SwitchToParentFrame 'must move up the tree to see sibling frame
    driver.Wait
    Debug.Assert driver.FindElementByTagName("body").GetText = "This is the parent frame"
    
    'switch to main document
    driver.SwitchToDefaultContent
    driver.Wait
    Debug.Assert driver.FindElementByID("top-level host").GetTagName = "div"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
