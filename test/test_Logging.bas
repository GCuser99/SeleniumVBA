Attribute VB_Name = "test_Logging"
Option Explicit
Option Private Module

Sub test_logging()
    Dim driver As SeleniumVBA.WebDriver, fruits As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    'The least troublesome way to get a combined driver and browser log is to enable logging at the driver command line.
    '(see https://chromedriver.chromium.org/logging). This method creates a readable log file to user's path of choice...

    'True enables verbose logging
    driver.StartChrome , , True
    driver.OpenBrowser

    driver.NavigateTo "https://jsbin.com/osebed/2"
    driver.Wait
    
    Set fruits = driver.FindElement(by.ID, "fruits")
    
    If fruits.IsMultiSelect Then
        fruits.SelectByVisibleText "Banana"
        driver.Wait
        fruits.SelectByIndex 2 'Apple
        driver.Wait
        fruits.SelectByValue "orange"
        driver.Wait
        fruits.DeSelectAll
        driver.Wait
        fruits.SelectAll
        driver.Wait
        fruits.DeSelectByVisibleText "Banana"
        driver.Wait
        fruits.DeSelectByIndex 2 'Apple
        driver.Wait
        fruits.DeSelectByValue "orange"
        driver.Wait
        Debug.Print fruits.GetSelectedOptionText
    End If
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
