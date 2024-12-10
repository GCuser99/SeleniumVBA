Attribute VB_Name = "test_Scroll"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_scrollIntoView()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 2000
    driver.ActiveWindow.Maximize
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/scrolling_tests/page_with_scrolling_frame.html"
    
    driver.FindElement(By.Name, "scrolling_frame").SwitchToFrame
    
    Set elem = driver.FindElement(By.Name, "scroll_checkbox")
    
    elem.ScrollIntoView(jump_smooth).Click
    
    Debug.Assert elem.GetProperty("checked")
    
    elem.Click 'click off
    driver.ScrollToTop
    
    elem.ScrollIntoView().Click
    
    Debug.Assert elem.GetProperty("checked")
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/scrolling_tests/page_with_tall_frame.html"
    
    driver.FindElement(By.Name, "tall_frame").SwitchToFrame
    
    Set elem = driver.FindElement(By.Name, "checkbox")
    
    elem.ScrollIntoView jump_smooth
    
    elem.Click
    
    Debug.Assert elem.GetProperty("checked")
    
    elem.Click 'click off
    driver.ScrollToTop
    
    elem.ScrollIntoView().Click
    
    Debug.Assert elem.GetProperty("checked")
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/scrolling_tests/page_with_scrolling_frame_out_of_view.html"
    
    driver.FindElement(By.Name, "scrolling_frame").ScrollIntoView(jump_smooth).SwitchToFrame
    
    Set elem = driver.FindElement(By.Name, "scroll_checkbox")
    
    elem.ScrollIntoView(jump_smooth).Click
    
    Debug.Assert elem.GetProperty("checked")
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/scrolling_tests/page_with_nested_scrolling_frames.html"
    
    driver.FindElement(By.Name, "scrolling_frame").SwitchToFrame
    driver.FindElement(By.Name, "nested_scrolling_frame").ScrollIntoView(jump_smooth).SwitchToFrame
    
    Set elem = driver.FindElement(By.Name, "scroll_checkbox")
    
    elem.ScrollIntoView(jump_smooth).Click
    
    Debug.Assert elem.GetProperty("checked")
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/scrolling_tests/page_with_y_overflow_auto.html"
    
    Set elem = driver.FindElement(By.TagName, "a")
    
    elem.ScrollIntoView(jump_smooth).Click
    
    Debug.Assert driver.FindElementByTagName("h1").GetText = "Clicked Successfully!"
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/scroll.html"
    
    Set elem = driver.FindElement(By.ID, "line9")
    
    elem.ScrollIntoView(jump_smooth).Click
    
    Debug.Assert driver.FindElement(By.ID, "clicked").GetText = "line9"
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/scroll4.html"
    
    Set elem = driver.FindElement(By.ID, "radio")
    
    elem.ScrollIntoView(jump_smooth).Click
    
    Debug.Assert elem.IsSelected
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/overflow/x_auto_y_auto.html"
    
    Set elem = driver.FindElement(By.ID, "right")
    elem.ScrollIntoView(jump_smooth).Click
    
    Set elem = driver.FindElement(By.ID, "bottom-right")
    elem.ScrollIntoView(jump_smooth).Click
    
    Set elem = driver.FindElement(By.ID, "bottom")
    elem.ScrollIntoView(jump_smooth).Click 'this one fails in FF - element is hidden by scroll bar
    
    Debug.Assert driver.FindElement(By.ID, "right-clicked").GetText = "ok"
    Debug.Assert driver.FindElement(By.ID, "bottom-right-clicked").GetText = "ok"
    Debug.Assert driver.FindElement(By.ID, "bottom-clicked").GetText = "ok"
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/overflow/x_hidden_y_hidden.html"
    
    Set elem = driver.FindElement(By.ID, "right")
    elem.ScrollIntoView(jump_auto).Click
    
    Set elem = driver.FindElement(By.ID, "bottom-right")
    elem.ScrollIntoView(jump_auto).Click
    
    Set elem = driver.FindElement(By.ID, "bottom")
    elem.ScrollIntoView(jump_auto).Click 'this one fails in FF - element is hidden by scroll bar
    
    Debug.Assert driver.FindElement(By.ID, "right-clicked").ScrollIntoView(jump_instant).GetText = "ok"
    Debug.Assert driver.FindElement(By.ID, "bottom-right-clicked").GetText = "ok"
    Debug.Assert driver.FindElement(By.ID, "bottom-clicked").GetText = "ok"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_long_scroll()
    Dim driver As SeleniumVBA.WebDriver
    Dim endElem As SeleniumVBA.WebElement
    Dim html As String
    Dim filePath As String
    Dim i As Long

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser

    'create the test html doc - set the style attribute scroll-behavior to "smooth"
    'this will affect the default "auto" scroll behavior in the scroll methods tested below
    html = "<!DOCTYPE html><html style='scroll-behavior:smooth;'><body>"
    For i = 1 To 10000: html = html & "<div><p>" & i & "</p></div>": Next i
    html = html & "<div id='end'><p>end</p></div>"
    html = html & "</body></html>"

    filePath = ".\snippet.html"
    driver.SaveStringToFile html, filePath

    driver.NavigateToFile filePath
    driver.ActiveWindow.Maximize
    driver.Wait 1000

    Set endElem = driver.FindElement(By.ID, "end")

    'this will smooth scroll because the default "jump_auto" scroll mode
    'takes its value from scrolling container's CSS
    driver.ScrollIntoView endElem
    
    driver.ScrollToTop jump_instant
    
    driver.Wait 1000
    
    driver.ScrollIntoView endElem
    
    driver.ScrollToTop
    
    driver.ScrollTo , 30000
    
    driver.ScrollBy , -30000
    
    driver.ScrollToBottom
    
    driver.Wait 1000
    
    driver.DeleteFiles filePath
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_element_scroll()
    Dim driver As SeleniumVBA.WebDriver
    Dim endElem As SeleniumVBA.WebElement
    Dim scrollContainer As SeleniumVBA.WebElement
    Dim html As String
    Dim filePath As String
    Dim i As Long

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser

    'create the test html doc - set the style attribute scroll-behavior to "smooth"
    'this will affect the default "auto" scroll behavior in the scroll methods tested below
    html = "<!DOCTYPE html><html><body>"
    html = html & "<div id='scroll' style='overflow-y:scroll; height:400px;scroll-behavior:smooth;'>"
    For i = 1 To 1000: html = html & "<div id='" & i & "'><p>" & i & "</p></div>": Next i
    html = html & "<div id='end'><p>end</p></div>"
    html = html & "</div>"
    html = html & "</body></html>"

    filePath = ".\snippet.html"
    driver.SaveStringToFile html, filePath

    driver.NavigateToFile filePath
    driver.ActiveWindow.Maximize
    driver.Wait 1000

    Set scrollContainer = driver.FindElement(By.ID, "scroll")
    Set endElem = driver.FindElement(By.ID, "end")

    'this will smooth scroll because the default "jump_auto" scroll mode
    'takes its value from scrolling container's CSS
    driver.ScrollIntoView endElem
    
    scrollContainer.ScrollToTop jump_instant
    
    driver.Wait 1000
    
    driver.ScrollIntoView endElem
    
    scrollContainer.ScrollToTop
    
    scrollContainer.ScrollTo , 30000
    
    scrollContainer.ScrollBy , -30000
    
    scrollContainer.ScrollToBottom
    
    driver.Wait 1000
    
    driver.DeleteFiles filePath
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_deep_scrollIntoView()
    Dim driver As SeleniumVBA.WebDriver
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser

    driver.NavigateTo "https://the-internet.herokuapp.com/large"
    driver.Wait 1000

    driver.FindElement(By.ID, "sibling-50.3").ScrollIntoView enSpeed:=jump_smooth, enAlign_horiz:=align_start
    
    driver.ScrollToTop enSpeed:=jump_smooth
    
    driver.FindElement(By.ID, "sibling-50.3").ScrollIntoView enSpeed:=jump_smooth, enAlign_horiz:=align_start, xOffset:=-200
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
