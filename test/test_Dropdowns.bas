Attribute VB_Name = "test_Dropdowns"
Sub test_select()
    'https://www.guru99.com/select-option-dropdown-selenium-webdriver.html
    Dim driver As New WebDriver, fruits As WebElement
    
    driver.StartChrome
    driver.OpenBrowser

    driver.NavigateTo "https://jsbin.com/osebed/2"
    driver.Wait 1000
    
    Set fruits = driver.FindElement(by.ID, "fruits")
    
    fruits.SelectByVisibleText "Banana"
    driver.Wait 500
    fruits.SelectByIndex 2 'Apple
    driver.Wait 500
    fruits.SelectByValue "orange"
    driver.Wait 500
    fruits.DeSelectAll
    driver.Wait 500
    fruits.SelectAll
    driver.Wait 500
    fruits.DeSelectByVisibleText "Banana"
    driver.Wait 500
    fruits.DeSelectByIndex 2 'Apple
    driver.Wait 500
    fruits.DeSelectByValue "orange"
    driver.Wait 500
    
    Debug.Print fruits.GetAllSelectedOptionsText()(1) 'Grape

    driver.CloseBrowser
    driver.Shutdown

End Sub
