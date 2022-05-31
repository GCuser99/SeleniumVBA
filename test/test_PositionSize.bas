Attribute VB_Name = "test_PositionSize"
Sub test_Position_Size()
    Dim driver As New WebDriver, webElem As WebElement, rect As Dictionary
    
    driver.StartEdge
    driver.OpenBrowser
    
    url = "https://www.google.com/"

    driver.NavigateTo url
    Set webElem = driver.FindElement(by.name, "q")

    driver.Wait 500
    
    'SeleniumVBA uses the dictionary object to represent rectangle position and size
    Set rect = webElem.GetRect
    
    Debug.Print rect("x"), rect("y"), rect("width"), rect("height")
    
    Set rect = driver.GetWindowRect
    
    Debug.Print rect("x"), rect("y"), rect("width"), rect("height")
    
    'driver.SetWindowSize , rect("height") / 2
    'driver.SetWindowPosition , 200
    
    Set rect = driver.SetWindowRect(, 200, , rect("height") / 2)
    
    Debug.Print rect("x"), rect("y"), rect("width"), rect("height")
    
    driver.Wait 1000

    driver.CloseBrowser
    driver.Shutdown
End Sub
