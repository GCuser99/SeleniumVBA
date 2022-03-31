Attribute VB_Name = "test_AlertsWindows"
Sub test_alerts_and_windows()
    'see https://www.guru99.com/alert-popup-handling-selenium.html
    Dim Driver As WebDriver
    
    Set Driver = New WebDriver
    Driver.StartChrome
    Driver.OpenBrowser

    Driver.NavigateTo "http://demo.guru99.com/test/delete_customer.php"
    
    Driver.Wait 1000
    
    Debug.Print "IsAlertPresent:", Driver.IsAlertPresent
                                
    Driver.FindElement(by.name, "cusid").SendKeys "87654"
    
    Driver.Wait 1000
    
    Driver.FindElement(by.name, "submit").Click
    
    Debug.Print "IsAlertPresent:", Driver.IsAlertPresent
    Debug.Print Driver.GetAlertText
                
    Driver.Wait 1000
    
    Driver.AcceptAlert
    
    Debug.Print Driver.GetAlertText
    
    Driver.Wait 1000
    Driver.AcceptAlert
    
    Driver.Wait 1000
    
    Driver.NavigateTo "http://demo.guru99.com/popup.php"
    
    Driver.MaximizeWindow
    
    Driver.Wait 2000
    
    Driver.FindElement(by.XPath, "//*[contains(@href,'popup.php')]").Click
    
    mainWindow = Driver.GetCurrentWindowHandle
    whdls = Driver.GetWindowHandles
    
    For i = 1 To UBound(whdls)
        If whdls(i) <> mainWindow Then
            Driver.SwitchToWindow whdls(i)
            Driver.FindElement(by.name, "emailid").SendKeys "gaurav.3n@gmail.com"
            Driver.Wait 2000
            Driver.FindElement(by.name, "btnLogin").Click
            Driver.Wait 2000
            Driver.CloseWindow
            Exit For
        End If
    Next i
    
    ' Switching to Parent window i.e Main Window.
    Driver.SwitchToWindow mainWindow
    
    Driver.Wait 2000
    Driver.CloseBrowser
    Driver.Shutdown

End Sub
