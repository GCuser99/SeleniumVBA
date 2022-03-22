Attribute VB_Name = "test_Capabilities"


Sub test_headless()
    Dim Driver As New WebDriver
    Dim cap As Capabilities
    
    Driver.Edge
    
    Set cap = Driver.CreateCapabilities
    
    cap.AddArgument "--headless"

    Driver.OpenBrowser cap
    
    'Driver.OpenBrowser ,  True 'this does same as above - set invisible parameter = true
    
    Driver.Navigate "https://www.google.com/"
    Driver.Wait 1000

    Driver.CloseBrowser
    Driver.Shutdown

End Sub



