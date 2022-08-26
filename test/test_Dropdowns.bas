Attribute VB_Name = "test_Dropdowns"
Option Explicit
Option Private Module

Sub test_select()
    'https://www.guru99.com/select-option-dropdown-selenium-webdriver.html
    Dim driver As SeleniumVBA.WebDriver, fruits As SeleniumVBA.WebElement

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser

    driver.NavigateTo "https://jsbin.com/osebed/2"
    driver.Wait 1000
    
    Set fruits = driver.FindElement(by.ID, "fruits")
    
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
    
    Debug.Print fruits.GetAllSelectedOptionsText()(1) 'Grape

    driver.CloseBrowser
    driver.Shutdown
End Sub
