Attribute VB_Name = "test_Options"
Sub test_headless()
    'see also test_FileUpDownload for another example using Options
    Dim driver As New WebDriver
    Dim options As WebOptions

    driver.StartChrome
    
    'note that WebOptions object should be created after starting the driver (StartEdge, StartChrome, of StartFirefox methods)
    Set options = driver.CreateOptions
    
    options.AddArgument "--headless" 'makes browser run in invisible mode
    
    driver.OpenBrowser options 'here is where options is passed to driver
    
    driver.NavigateTo "https://www.google.com/"
    
    Debug.Print "User Agent: " & driver.GetUserAgent

    driver.CloseBrowser
    driver.Shutdown
End Sub

