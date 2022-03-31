Attribute VB_Name = "test_Logging"

Sub test_logging()

    Dim Driver As New WebDriver, fruits As WebElement
    
    'The least troublesome way to get a combined driver and browser log is to enable logging at the driver command line.
    '(see https://chromedriver.chromium.org/logging). This method creates a readable log file to user's path of choice...

    'True enables verbose logging - default log file found in same directory as WebDriver executable
    Driver.StartEdge , , True

    Driver.OpenBrowser

    Driver.NavigateTo "https://jsbin.com/osebed/2"
    Driver.Wait
    
    Set fruits = Driver.FindElement(by.ID, "fruits")
    
    If fruits.IsMultiSelect Then
        fruits.SelectByVisibleText "Banana"
        Driver.Wait
        fruits.SelectByIndex 2 'Apple
        Driver.Wait
        fruits.SelectByValue "orange"
        Driver.Wait
        fruits.DeSelectAll
        Driver.Wait
        fruits.SelectAll
        Driver.Wait
        fruits.DeSelectByVisibleText "Banana"
        Driver.Wait
        fruits.DeSelectByIndex 2 'Apple
        Driver.Wait
        fruits.DeSelectByValue "orange"
        Driver.Wait
        Debug.Print fruits.GetSelectedOptionText
    End If
    
    Driver.CloseBrowser
    Driver.Shutdown
End Sub
