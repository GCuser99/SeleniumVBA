Attribute VB_Name = "test_Sendkeys"
Sub test_sendkeys()
    Dim driver As New WebDriver
    Dim keys As New WebKeyboard

    driver.StartEdge
    
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    keySeq = "This is COOKL!" & keys.LeftKey & keys.LeftKey & keys.LeftKey & keys.DeleteKey & keys.ReturnKey
    
    driver.FindElement(by.Name, "q").SendKeys keySeq

    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

