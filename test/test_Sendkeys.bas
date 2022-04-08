Attribute VB_Name = "test_Sendkeys"
Sub test_sendkeys()
    Dim Driver As New WebDriver
    Dim keys As New Keyboard
    
    Driver.StartChrome
    Driver.OpenBrowser
    
    Driver.NavigateTo "https://www.google.com/"
    Driver.Wait 1000
    
    Driver.FindElement(by.name, "q").SendKeys "This is COOKL!" & keys.LeftKey & keys.LeftKey & keys.LeftKey & keys.DeleteKey & keys.ReturnKey

    Driver.Wait 2000
    
    Driver.CloseBrowser
    Driver.Shutdown
    
End Sub

