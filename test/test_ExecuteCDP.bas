Attribute VB_Name = "test_ExecuteCDP"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")
'The Chrome DevTools Protocol (CDP) via SeleniumVBA's ExecuteCDP method
'provides a low-level interface for interacting with Chrome\Edge.
'For more info, see https://chromedevtools.github.io/devtools-protocol

Sub test_cdp_enhanced_screenshot()
    'this demonstrates using the ExecuteCDP to perform a screenshot with enhanced user control
    Dim driver As SeleniumVBA.WebDriver
    Dim params As New Dictionary
    'Dim clipRect As New Dictionary
    Dim strB64 As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 500
    
    'see https://chromedevtools.github.io/devtools-protocol/tot/Page/#method-captureScreenshot

    params.Add "format", "jpeg" 'jpeg, png (default), webp
    params.Add "quality", 80 '0 to 100 (jpeg only), defaults to 80
    'clip parameter can be used to snapshot an element rectangle
    '(see GetRect method of the WebDriver and WebElement classes)
    'clipRect.Add "x", 200
    'clipRect.Add "y", 200
    'clipRect.Add "width", 400
    'clipRect.Add "height", 400
    'clipRect.Add "scale", 1
    'params.Add "clip", clipRect
    'the next 3 paramters are currently marked as experimental (as of 11 June, 2023)
    params.Add "captureBeyondViewport", True 'full screenshot
    params.Add "fromSurface", True 'defaults to true
    params.Add "optimizeForSpeed", False 'defaults to false
    
    Debug.Print SeleniumVBA.WebJsonConverter.ConvertToJson(params, 4)
    
    'send the cdp command to the WebDriver and return "data" key of the response dictionary
    strB64 = driver.ExecuteCDP("Page.captureScreenshot", params)("value")("data")
    
    'results in a base 64 encoded string which must be decoded into a bytearray before saving to file
    driver.SaveBase64StringToFile strB64, ".\screenshotfull.jpg"
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_cdp_enhanced_geolocation()
    'this demonstrates using the ExecuteCDP to manage geolocation with enhanced user control
    'even if the default profile is set to hide geolocation info, this will override that,
    'unlike SetGeolocation method of WebDriver class...
    Dim driver As SeleniumVBA.WebDriver
    Dim params As New Dictionary
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome 'Chrome and Edge only
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 2000
    
    'https://chromedevtools.github.io/devtools-protocol/tot/Emulation/#method-setGeolocationOverride
    'https://chromedevtools.github.io/devtools-protocol/tot/Emulation/#method-clearGeolocationOverride
    
    'set the override location
    params.Add "latitude", 41.1621429
    params.Add "longitude", -8.6219537
    params.Add "accuracy", 100
    
    driver.ExecuteCDP "Emulation.setGeolocationOverride", params
  
    driver.NavigateTo "https://the-internet.herokuapp.com/geolocation"
    
    driver.FindElementByXPath("//*[@id='content']/div/button").Click
    
    Debug.Print driver.FindElementByID("lat-value").GetText, driver.FindElementByID("long-value").GetText
    
    driver.Wait 1000
    
    'now clear the override...
    driver.ExecuteCDP "Emulation.clearGeolocationOverride"
    
    'refresh the page...
    driver.Refresh
    
    driver.FindElementByXPath("//*[@id='content']/div/button").Click
    
    Debug.Print driver.FindElementByID("lat-value").GetText, driver.FindElementByID("long-value").GetText
    
    driver.Wait 2000
    
    driver.FindElementByXPath("//*[@id='map-link']/a").Click
    
    driver.Wait 5000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_cdp_enhanced_file_download()
    'this demonstrates using the ExecuteCDP to redirect the default browser
    'download location AFTER capabilities have been set (post-OpenBrowser)
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    Dim params As New Dictionary
    
    Set driver = SeleniumVBA.New_WebDriver
   
    driver.StartChrome
    
    'set the directory path for saving download to
    Set caps = driver.CreateCapabilities()
    caps.SetDownloadPrefs "%USERPROFILE%\Desktop"
    driver.OpenBrowser caps
    
    'redirect the download location AFTER capabilities have been set!!
    'https://chromedevtools.github.io/devtools-protocol/tot/Page/#method-setDownloadBehavior
    params.Add "behavior", "allow" 'deny, allow, default
    params.Add "downloadPath", driver.ResolvePath(".\")
    driver.ExecuteCDP "Page.setDownloadBehavior", params
    
    'delete legacy copy if it exists
    driver.DeleteFiles ".\test.pdf"
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/raw/main/dev/test_files/test.pdf"
    
    driver.WaitForDownload ".\test.pdf"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_cdp_scripts()
    'this sub demonstrates the use of CDP commands Page.addScriptToEvaluateOnNewDocument
    'and Runtime.evaluate to run java scripts on a webpage
    Dim driver As SeleniumVBA.WebDriver
    Dim params As New Dictionary
    Dim resp As Dictionary
    Dim html As String
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser
    
    'create an html with a script that alerts after page load and a script that can alter element's text
    html = "<!DOCTYPE html>" & _
    "<html>" & _
    "<body onload='loadIt()'>" & _
    "<p id='text'>Original text here</p>" & _
    "<script>" & _
    "function loadIt(){alert('This embedded HTML script executes on document load, but only AFTER Page.addScriptToEvaluateOnNewDocument');}" & _
    "</script>" & _
    "<script>" & _
    "function changeText(txt){document.getElementById('text').innerText=txt;}" & _
    "</script>" & _
    "</body>" & _
    "</html>"

    driver.SaveStringToFile html, ".\snippet.html"
    
    'Use CDP to inject a script to run before HTML document scripts run
    'https://chromedevtools.github.io/devtools-protocol/tot/Page/#method-addScriptToEvaluateOnNewDocument
    params.Add "source", "alert('This injected CDP script executes on page load, BEFORE any other document scripts are executed');"
    driver.ExecuteCDP "Page.addScriptToEvaluateOnNewDocument", params

    driver.NavigateToFile ".\snippet.html"
    
    driver.Wait 3000
    
    driver.SwitchToAlert.Accept 'CDP injected alert
    
    driver.Wait 3000
    
    driver.SwitchToAlert.Accept 'embedded HTML load alert
    
    driver.Wait 2000
    
    'use CDP to call the HTML embedded script to alter element text
    'https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#method-evaluate
    params.RemoveAll
    params.Add "expression", "changeText('The text has been changed by a CDP call to embedded HTML script');"
    driver.ExecuteCDP "Runtime.evaluate", params
    
    driver.Wait 3000
    
    'use CDP to run a CDP script that returns the text value of the element
    params.RemoveAll
    params.Add "expression", "document.getElementById('text').outerText"
    Set resp = driver.ExecuteCDP("Runtime.evaluate", params)
    
    'print the result to debug window
    Debug.Print resp("value")("result")("type"), resp("value")("result")("value")
    
    'use CDP to run a script that uses DOM to change the text value
    params.RemoveAll
    params.Add "expression", "document.getElementById('text').innerText='The text has been changed by a CDP script'"
    driver.ExecuteCDP "Runtime.evaluate", params
    
    driver.Wait 2000
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_cdp_random_other_stuff()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    Dim resp As Dictionary

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    Set caps = driver.CreateCapabilities(initializeFromSettingsFile:=False)
    driver.OpenBrowser caps
    
    'create a browser navigation history
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 500
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 500
    driver.NavigateTo "https://www.wikipedia.org/"
    
    'use cdp get browser history
    'https://chromedevtools.github.io/devtools-protocol/tot/Page/#method-getNavigationHistory
    Set resp = driver.ExecuteCDP("Page.getNavigationHistory")
    Debug.Print SeleniumVBA.WebJsonConverter.ConvertToJson(resp, 4)
    
    'use cdp reset browser history
    'https://chromedevtools.github.io/devtools-protocol/tot/Page/#method-resetNavigationHistory
    driver.ExecuteCDP "Page.resetNavigationHistory"
    
    'use cdp to confirm browser history was erased
    Set resp = driver.ExecuteCDP("Page.getNavigationHistory")
    Debug.Print SeleniumVBA.WebJsonConverter.ConvertToJson(resp, 4)
    
    'use cdp to clear and disable browser cache
    'https://chromedevtools.github.io/devtools-protocol/tot/Network/#method-enable
    driver.ExecuteCDP "Network.enable"
    'https://chromedevtools.github.io/devtools-protocol/tot/Network/#method-clearBrowserCache
    driver.ExecuteCDP "Network.clearBrowserCache"
    'https://chromedevtools.github.io/devtools-protocol/tot/Network/#method-setCacheDisabled
    driver.ExecuteCDP "Network.setCacheDisabled", "{'cacheDisabled':true}"
    
    'use cdp to get cookies
    'https://chromedevtools.github.io/devtools-protocol/tot/Network/#method-getCookies
    Set resp = driver.ExecuteCDP("Network.getCookies")
    Debug.Print SeleniumVBA.WebJsonConverter.ConvertToJson(resp, 4)
    Debug.Print resp("value")("cookies").Count
    
    'use cdp to clear cookies
    'https://chromedevtools.github.io/devtools-protocol/tot/Network/#method-clearBrowserCookies
    driver.ExecuteCDP "Network.clearBrowserCookies"
    
    'use cdp to verify cookies are cleared
    Set resp = driver.ExecuteCDP("Network.getCookies")
    Debug.Print SeleniumVBA.WebJsonConverter.ConvertToJson(resp, 4)
    Debug.Print resp("value")("cookies").Count
    
    'use cdp to override UserAgent
    'note that with the conventional command warppers, this is done with capabilities
    'https://chromedevtools.github.io/devtools-protocol/tot/Network/#method-setUserAgentOverride
    driver.ExecuteCDP "Emulation.setUserAgentOverride", "{""userAgent"": 'Mozilla/5.1 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36'}"
    
    Debug.Print driver.GetUserAgent
    driver.Wait 1500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
