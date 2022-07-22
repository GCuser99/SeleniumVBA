Attribute VB_Name = "test_Capabilities"
Sub test_headless()
    'see also test_FileUpDownload for another example using Capabilities
    Dim driver As New WebDriver
    Dim caps As WebCapabilities

    driver.StartChrome
    
    'note that Capabilities object should be created after starting the driver (StartEdge or StartChrome methods)
    Set caps = driver.CreateCapabilities
    
    caps.AddArgument "--headless" 'makes browser run in invisible mode
    
    driver.OpenBrowser caps 'here is where caps is passed to driver
    
    driver.NavigateTo "https://www.google.com/"
    
    Debug.Print "User Agent: " & driver.GetUserAgent

    driver.CloseBrowser
    driver.Shutdown
End Sub

