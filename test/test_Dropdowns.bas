Attribute VB_Name = "test_Dropdowns"
Sub test_select()
    'https://www.guru99.com/select-option-dropdown-selenium-webdriver.html
    Dim Driver As New WebDriver, fruits As WebElement
    
    Driver.StartChrome
    Driver.OpenBrowser

    Driver.NavigateTo "https://jsbin.com/osebed/2"
    Driver.Wait 1000
    
    Set fruits = Driver.FindElement(by.ID, "fruits")
    
    fruits.SelectByVisibleText "Banana"
    Driver.Wait 500
    fruits.SelectByIndex 2 'Apple
    Driver.Wait 500
    fruits.SelectByValue "orange"
    Driver.Wait 500
    fruits.DeSelectAll
    Driver.Wait 500
    fruits.SelectAll
    Driver.Wait 500
    fruits.DeSelectByVisibleText "Banana"
    Driver.Wait 500
    fruits.DeSelectByIndex 2 'Apple
    Driver.Wait 500
    fruits.DeSelectByValue "orange"
    Driver.Wait 500
    
    Debug.Print fruits.GetAllSelectedOptionsText()(1) 'Grape

    Driver.CloseBrowser
    Driver.Shutdown

End Sub
