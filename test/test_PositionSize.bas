Attribute VB_Name = "test_PositionSize"
Sub test_Position_Size()
    Dim Driver As New WebDriver, webElem As WebElement, rect As Dictionary
    
    Driver.StartChrome
    Driver.OpenBrowser
    
    url = "https://www.google.com/"

    Driver.NavigateTo url
    Set webElem = Driver.FindElement(by.name, "q")

    Driver.Wait 500
    
    'SeleniumVBA uses the dictionary object to represent rectangle position and size
    Set rect = webElem.GetRect
    
    Debug.Print rect("x"), rect("y"), rect("width"), rect("height")
    
    Set rect = Driver.GetWindowRect
    
    Debug.Print rect("x"), rect("y"), rect("width"), rect("height")
    
    'driver.SetWindowSize , rect("height") / 2
    'driver.SetWindowPosition , 200
    
    Set rect = Driver.SetWindowRect(, 200, , rect("height") / 2)
    
    Debug.Print rect("x"), rect("y"), rect("width"), rect("height")
    
    Driver.Wait 1000

    Driver.CloseBrowser
    Driver.Shutdown

End Sub
