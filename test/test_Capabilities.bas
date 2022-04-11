Attribute VB_Name = "test_Capabilities"


Sub test_headless()
    'see also test_FileUpDownload for another example using Capabilities
    Dim driver As New WebDriver
    Dim caps As Capabilities
    
    driver.StartEdge
    
    'note that Capabilities object should be created after starting the driver (StartEdge or StartChrome methods)
    Set caps = driver.CreateCapabilities
    
    caps.AddArgument "--headless" 'makes browser run in invisible mode

    driver.OpenBrowser caps 'here is where caps is passed to driver
    
    'Driver.OpenBrowser ,  True 'this does same as above - set invisible parameter = true
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait
    
    Debug.Print "running in headless mode"

    driver.CloseBrowser
    driver.Shutdown

End Sub



