Attribute VB_Name = "test_Windows"
Sub test_windows1()
    Dim driver As New WebDriver
    
    driver.StartChrome
    driver.OpenBrowser

    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    hnd1 = driver.GetCurrentWindowHandle
    hnd2 = driver.SwitchToNewWindow(svbaTab) 'this will create a new browser tab
    'hnd2 = Driver.SwitchToNewWindow(svbaWindow) 'this will create a new browser window
    
    driver.NavigateTo "https://news.google.com/"
    driver.Wait 1000
    
    Debug.Print hnd2 & " is same as " & driver.GetCurrentWindowHandle
    
    driver.SwitchToWindow hnd1
    driver.Wait 1000
    driver.SwitchToWindow hnd2
    driver.Wait 1000
    
    Debug.Print "first window handle: " & driver.GetWindowHandles()(1)
    Debug.Print "second window handle: " & driver.GetWindowHandles()(2)
    
    'can switch based on index too
    For i = 1 To 5
        driver.SwitchToWindow 1
        driver.Wait 500
        driver.SwitchToWindow 2
        driver.Wait 500
    Next i
    
    driver.CloseWindow
    driver.Wait 1000

    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_windows2()
    Dim driver As New WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "http://demo.guru99.com/popup.php"
    
    driver.MaximizeWindow
    
    driver.Wait 2000
    
    driver.FindElement(by.XPath, "//*[contains(@href,'popup.php')]").Click
    
    mainWindow = driver.GetCurrentWindowHandle
    whdls = driver.GetWindowHandles
    
    For i = 1 To UBound(whdls)
        If whdls(i) <> mainWindow Then
            driver.SwitchToWindow whdls(i)
            driver.Wait
            driver.FindElement(by.Name, "emailid").SendKeys "gaurav.3n@gmail.com"
            driver.Wait 2000
            driver.FindElement(by.Name, "btnLogin").Click
            driver.Wait 2000
            driver.CloseWindow
            Exit For
        End If
    Next i
    
    ' Switching to Parent window i.e Main Window.
    driver.SwitchToWindow mainWindow
    
    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

