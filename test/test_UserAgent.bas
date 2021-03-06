Attribute VB_Name = "test_UserAgent"
Sub test_hide_headless()
    'some servers detect headless mode in the sent User Agent and then deny access, so
    'here we modify the user agent that gets sent to server
    Dim driver As New WebDriver
    Dim caps As WebCapabilities

    driver.StartChrome
    driver.OpenBrowser , True  'a way of running headless mode without explicitly adding --headless arg to Capabilities
    
    'get the user agent for this browser setup
    userAgent = driver.GetUserAgent
    
    Debug.Print "Original User Agent:  " & userAgent
    
    driver.CloseBrowser
    
    Set caps = driver.CreateCapabilities
    
    'now we modify the user agent string by tossing the "Headless" keyword and then
    'update Capabilities UserArgent argument
    caps.SetUserAgent = Replace(userAgent, "HeadlessChrome", "Chrome")
    
    driver.OpenBrowser caps, True
    
    'to see a full list of headers navigate to https://www.httpbin.org/headers
    driver.NavigateTo "https://www.whatismybrowser.com/detect/what-is-my-user-agent/"
    
    Debug.Print "Modfified User Agent: " & driver.FindElement(by.ID, "detected_value").GetText
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
