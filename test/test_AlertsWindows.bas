Attribute VB_Name = "test_AlertsWindows"
Sub test_alerts_and_windows()
    'see https://www.guru99.com/alert-popup-handling-selenium.html
    Dim driver As WebDriver
    
    Set driver = New WebDriver
    driver.Chrome
    driver.OpenBrowser

    driver.Navigate "http://demo.guru99.com/test/delete_customer.php"
    
    driver.Wait 1000
    
    Debug.Print "IsAlertPresent:", driver.IsAlertPresent
                                
    driver.FindElement(by.name, "cusid").SendKeys "87654"
    
    driver.Wait 1000
    
    driver.FindElement(by.name, "submit").Click
    
    Debug.Print "IsAlertPresent:", driver.IsAlertPresent
    Debug.Print driver.GetAlertText
                
    driver.Wait 1000
    
    driver.AcceptAlert
    
    Debug.Print driver.GetAlertText
    
    driver.Wait 1000
    driver.AcceptAlert
    
    driver.Wait 1000
    
    driver.Navigate "http://demo.guru99.com/popup.php"
    
    driver.MaximizeWindow
    
    driver.Wait 2000
    
    driver.FindElement(by.XPath, "//*[contains(@href,'popup.php')]").Click
    
    MainWindow = driver.GetCurrentWindowHandle
    whdls = driver.GetWindowHandles
    
    For i = 0 To UBound(whdls)
        If whdls(i) <> MainWindow Then
            driver.SwitchToWindow whdls(i)
            driver.FindElement(by.name, "emailid").SendKeys ("gaurav.3n@gmail.com")
            driver.Wait 2000
            driver.FindElement(by.name, "btnLogin").Click
            driver.Wait 2000
            driver.CloseWindow
            Exit For
        End If
    Next i
    
    ' Switching to Parent window i.e Main Window.
    driver.SwitchToWindow MainWindow
    
    driver.Wait 2000
    driver.CloseBrowser
    driver.Shutdown

End Sub
