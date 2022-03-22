Attribute VB_Name = "test_sendkeys"
Sub test_sendkeys()
    Dim Driver As New WebDriver
    Dim Keys As New Keyboard
    
    Driver.Chrome
    Driver.OpenBrowser
    
    Driver.Navigate "https://www.google.com/"
    Driver.Wait 1000
    
    Driver.FindElement(by.name, "q").SendKeys "This is COOKL!" & Keys.LeftKey & Keys.LeftKey & Keys.LeftKey & Keys.DeleteKey & Keys.ReturnKey

    Driver.Wait 2000
    
    Driver.CloseBrowser
    Driver.Shutdown
    
End Sub
