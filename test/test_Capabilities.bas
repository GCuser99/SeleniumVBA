Attribute VB_Name = "test_Capabilities"


Sub test_headless()
    Dim driver As New WebDriver
    Dim cap As Capabilities
    
    driver.Edge
    
    Set cap = driver.CreateCapabilities
    
    cap.AddArgument "--headless"

    driver.OpenBrowser cap
    
    'Driver.OpenBrowser ,  True 'this does same as above - set invisible parameter = true
    
    driver.Navigate "https://www.google.com/"
    driver.Wait 250

    driver.CloseBrowser
    driver.Shutdown

End Sub



