Attribute VB_Name = "test_logging"

Sub test_logging()

    Dim Driver As New WebDriver, fruits As WebElement
    
    'The least troublesome way to get a combined driver and browser log is to enable logging at the driver command line.
    '(see https://chromedriver.chromium.org/logging). This method creates a readable log file to user's path of choice...

    'True enables verbose logging - default log file found in same directory as WebDriver executable
    Driver.Edge , , True

    Driver.OpenBrowser

    Driver.Navigate "https://jsbin.com/osebed/2"
    Driver.Wait 250
    
    Set fruits = Driver.FindElement(by.ID, "fruits")
    
    If fruits.IsMultiSelect Then
        fruits.SelectByVisibleText ("Banana")
        Driver.Wait 250
        fruits.SelectByIndex (1)
        Driver.Wait 250
        fruits.SelectByValue ("orange")
        Driver.Wait 250
        fruits.DeSelectAll
        Driver.Wait 250
        fruits.SelectAll
        Driver.Wait 250
        fruits.DeSelectByVisibleText ("Banana")
        Driver.Wait 250
        fruits.DeSelectByIndex (1)
        Driver.Wait 250
        fruits.DeSelectByValue ("orange")
        Driver.Wait 250
        Debug.Print fruits.SelectedOptionText
    End If
    
    Driver.CloseBrowser
    Driver.Shutdown
End Sub
