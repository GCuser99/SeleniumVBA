Attribute VB_Name = "test_Logging"

Sub test_logging()

    Dim driver As New WebDriver, fruits As WebElement
    
    'The least troublesome way to get a combined driver and browser log is to enable logging at the driver command line.
    '(see https://chromedriver.chromium.org/logging). This method creates a readable log file to user's path of choice...

    'True enables verbose logging - default log file found in same directory as WebDriver executable
    driver.Edge , , True

    driver.OpenBrowser

    driver.Navigate "https://jsbin.com/osebed/2"
    driver.Wait 250
    
    Set fruits = driver.FindElement(by.ID, "fruits")
    
    If fruits.IsMultiSelect Then
        fruits.SelectByVisibleText ("Banana")
        driver.Wait 250
        fruits.SelectByIndex (1)
        driver.Wait 250
        fruits.SelectByValue ("orange")
        driver.Wait 250
        fruits.DeSelectAll
        driver.Wait 250
        fruits.SelectAll
        driver.Wait 250
        fruits.DeSelectByVisibleText ("Banana")
        driver.Wait 250
        fruits.DeSelectByIndex (1)
        driver.Wait 250
        fruits.DeSelectByValue ("orange")
        driver.Wait 250
        Debug.Print fruits.SelectedOptionText
    End If
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
