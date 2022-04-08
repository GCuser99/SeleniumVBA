Attribute VB_Name = "test_Capabilities"


Sub test_headless()
    'see also test_FileUpDownload for another example using Capabilities
    Dim Driver As New WebDriver
    Dim caps As Capabilities
    
    Driver.StartEdge
    
    'note that Capabilities object should be created after starting the driver (StartEdge or StartChrome methods)
    Set caps = Driver.CreateCapabilities
    
    caps.AddArgument "--headless" 'makes browser run in invisible mode

    Driver.OpenBrowser caps 'here is where caps is passed to driver
    
    'Driver.OpenBrowser ,  True 'this does same as above - set invisible parameter = true
    
    Driver.NavigateTo "https://www.google.com/"
    Driver.Wait
    
    Debug.Print "running in headless mode"

    Driver.CloseBrowser
    Driver.Shutdown

End Sub



