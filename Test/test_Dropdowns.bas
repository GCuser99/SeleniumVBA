Attribute VB_Name = "test_dropdowns"
Sub test_select()
    'https://www.guru99.com/select-option-dropdown-selenium-webdriver.html
    Dim Driver As New WebDriver, fruits As WebElement
    
    Driver.Chrome
    Driver.OpenBrowser

    Driver.Navigate "https://jsbin.com/osebed/2"
    Driver.Wait 1000
    
    Set fruits = Driver.FindElement(by.ID, "fruits")
    
    fruits.SelectByVisibleText ("Banana")
    Driver.Wait 500
    fruits.SelectByIndex (1)
    Driver.Wait 500
    fruits.SelectByValue ("orange")
    Driver.Wait 500
    fruits.DeSelectAll
    Driver.Wait 500
    fruits.SelectAll
    Driver.Wait 500
    fruits.DeSelectByVisibleText ("Banana")
    Driver.Wait 500
    fruits.DeSelectByIndex (1)
    Driver.Wait 500
    fruits.DeSelectByValue ("orange")
    Driver.Wait 500
    
    Driver.CloseBrowser
    Driver.Shutdown

End Sub
