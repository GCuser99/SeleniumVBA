Attribute VB_Name = "test_Windows"
Sub test_windows1()
    Dim Driver As New WebDriver
    
    Driver.StartChrome
    Driver.OpenBrowser

    Driver.NavigateTo "https://www.google.com/"
    Driver.Wait 1000
    
    hnd1 = Driver.GetCurrentWindowHandle
    hnd2 = Driver.SwitchToNewWindow(svbaTab) 'this will create a new browser tab
    'hnd2 = Driver.SwitchToNewWindow(svbaWindow) 'this will create a new browser window
    
    Driver.NavigateTo "https://news.google.com/"
    Driver.Wait 1000
    
    Debug.Print hnd2 & " is same as " & Driver.GetCurrentWindowHandle
    
    Driver.SwitchToWindow hnd1
    Driver.Wait 1000
    Driver.SwitchToWindow hnd2
    Driver.Wait 1000
    
    Debug.Print "first window handle: " & Driver.GetWindowHandles()(1)
    Debug.Print "second window handle: " & Driver.GetWindowHandles()(2)
    
    'can switch based on index too
    For i = 1 To 5
        Driver.SwitchToWindow 1
        Driver.Wait 500
        Driver.SwitchToWindow 2
        Driver.Wait 500
    Next i
    
    Driver.CloseWindow
    Driver.Wait 1000

    Driver.CloseBrowser
    Driver.Shutdown

End Sub

Sub test_windows2()
    Dim Driver As New WebDriver
    
    Driver.StartChrome
    Driver.OpenBrowser
    
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
