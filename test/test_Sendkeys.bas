Attribute VB_Name = "test_Sendkeys"
Sub test_sendkeys()
    Dim driver As New WebDriver
    Dim Keys As New Keyboard
    
    driver.Chrome
    driver.OpenBrowser
    
    driver.Navigate "https://www.google.com/"
    driver.Wait 1000
    
    driver.FindElement(by.name, "q").SendKeys "This is COOKL!" & Keys.LeftKey & Keys.LeftKey & Keys.LeftKey & Keys.DeleteKey & Keys.ReturnKey

    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
    
End Sub
