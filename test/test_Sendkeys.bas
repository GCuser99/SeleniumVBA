Attribute VB_Name = "test_Sendkeys"
Sub test_sendkeys()
    Dim driver As New WebDriver
    Dim keys As New Keyboard
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    driver.FindElement(by.name, "q").SendKeys "This is COOKL!" & keys.LeftKey & keys.LeftKey & keys.LeftKey & keys.DeleteKey & keys.ReturnKey

    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
    
End Sub

