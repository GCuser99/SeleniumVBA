Attribute VB_Name = "test_Alerts"
Sub test_Alerts()
    'see https://www.guru99.com/alert-popup-handling-selenium.html
    Dim driver As New WebDriver

    driver.StartChrome
    driver.OpenBrowser

    driver.NavigateTo "http://demo.guru99.com/test/delete_customer.php"
    
    driver.Wait 1000
    
    Debug.Print "Is Alert Present: " & driver.IsAlertPresent
                                
    driver.FindElement(by.Name, "cusid").SendKeys "87654"
    
    driver.Wait 1000
    
    driver.FindElement(by.Name, "submit").Click
    
    driver.Wait 1000
    
    Debug.Print "Is Alert Present: " & driver.IsAlertPresent
    Debug.Print "Alert Text: " & driver.GetAlertText
    driver.AcceptAlert
    
    Debug.Print "Alert Text: " & driver.GetAlertText
    driver.AcceptAlert

    driver.Wait 1000
    driver.CloseBrowser
    driver.Shutdown
End Sub


