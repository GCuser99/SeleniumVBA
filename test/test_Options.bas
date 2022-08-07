Attribute VB_Name = "test_Options"
Sub test_headless()
    'see also test_FileUpDownload for another example using Options
    Dim driver As New WebDriver
    Dim caps As WebCapabilities

    driver.StartChrome
    
    'note that WebCapabilities object should be created after starting the driver (StartEdge, StartChrome, of StartFirefox methods)
    Set caps = driver.CreateOptions
    
    caps.AddArgument "--headless" 'makes browser run in invisible mode
    
    driver.OpenBrowser caps 'here is where caps is passed to driver
    
    driver.NavigateTo "https://www.google.com/"
    
    Debug.Print "User Agent: " & driver.GetUserAgent

    driver.CloseBrowser
    driver.Shutdown
End Sub

