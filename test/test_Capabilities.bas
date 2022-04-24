Attribute VB_Name = "test_Capabilities"
Sub test_headless()
    'see also test_FileUpDownload for another example using Capabilities
    Dim driver As New WebDriver
    Dim caps As Capabilities
    
    driver.StartChrome
    
    'note that Capabilities object should be created after starting the driver (StartEdge or StartChrome methods)
    Set caps = driver.CreateCapabilities
    
    caps.AddArgument "--headless" 'makes browser run in invisible mode
    
    driver.OpenBrowser caps 'here is where caps is passed to driver
    
    'driver.OpenBrowser , True is another way of running headless above without invoking the Capabilities object
    
    driver.NavigateTo "https://www.whatismybrowser.com/detect/what-is-my-user-agent/"
    
    Debug.Print "User Agent Sent: " & driver.FindElement(by.ID, "detected_value").GetText
    
    driver.CloseBrowser
    
    'some servers detect headless mode in user agent and deny access, so here we modify the user agent that gets sent to server
    caps.AddArgument "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36"
    
    driver.OpenBrowser caps
    
    driver.NavigateTo "https://www.whatismybrowser.com/detect/what-is-my-user-agent/"
    
    Debug.Print "User Agent Sent: " & driver.FindElement(by.ID, "detected_value").GetText
    
    driver.CloseBrowser
    driver.Shutdown

End Sub
