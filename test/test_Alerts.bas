Attribute VB_Name = "test_Alerts"
Sub test_Alerts()
    'see https://www.guru99.com/alert-popup-handling-selenium.html
    Dim Driver As New WebDriver
    
    Driver.StartChrome
    Driver.OpenBrowser

    Driver.NavigateTo "http://demo.guru99.com/test/delete_customer.php"
    
    Driver.Wait 1000
    
    Debug.Print "Is Alert Present: " & Driver.IsAlertPresent
                                
    Driver.FindElement(by.name, "cusid").SendKeys "87654"
    
    Driver.Wait 1000
    
    Driver.FindElement(by.name, "submit").Click
    
    Driver.Wait 1000
    
    Debug.Print "Is Alert Present: " & Driver.IsAlertPresent
    Debug.Print "Alert Text: " & Driver.GetAlertText
    Driver.AcceptAlert
    
    Debug.Print "Alert Text: " & Driver.GetAlertText
    Driver.AcceptAlert

    Driver.Wait 1000
    Driver.CloseBrowser
    Driver.Shutdown

End Sub


