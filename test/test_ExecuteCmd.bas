Attribute VB_Name = "test_ExecuteCmd"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")
'The ExecuteCmd method can be useful for testing as yet unwrapped WebDriver commands
'see https://chromium.googlesource.com/chromium/src/+/master/chrome/test/chromedriver/server/http_handler.cc

Sub test_firefox_full_screenshot()
    'this test shows how to run a Selenium command (http end point) that is not currently wrapped in SeleniumVBA
    'the command below (Firefox only) takes a screenshot of the entire page (beyond the viewport)
    'similar to what can be done with Edge\Chrome using the CDP command "Page.captureScreenshot"
    Dim driver As SeleniumVBA.WebDriver
    Dim strB64 As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartFirefox
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    
    'all arguments in the command path string must be pefixed with "$", including sessionId
    'command parameters (if required) can be passed as either a dictionary object or valid JSON string
    strB64 = driver.ExecuteCmd("GET", "/session/$sessionId/moz/screenshot/full")("value")
    
    'results in a base 64 encoded string which must be decoded into a bytearray before saving to file
    driver.SaveBase64StringToFile strB64, ".\screenshotfull.png"

    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_chrome_edge_full_screenshot()
    'this test shows how to run a Selenium command (http end point) that is not currently wrapped in SeleniumVBA
    'the command below (Firefox only) takes a screenshot of the entire page (beyond the viewport)
    'similar to what can be done with Edge\Chrome using the CDP command "Page.captureScreenshot"
    Dim driver As SeleniumVBA.WebDriver
    Dim strB64 As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    
    'all arguments in the command path string must be pefixed with "$", including sessionId
    'command parameters (if required) can be passed as either a dictionary object or valid JSON string
    strB64 = driver.ExecuteCmd("GET", "/session/$sessionId/screenshot/full")("value")
    
    'results in a base 64 encoded string which must be decoded into a bytearray before saving to file
    driver.SaveBase64StringToFile strB64, ".\screenshotfull.png"

    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
