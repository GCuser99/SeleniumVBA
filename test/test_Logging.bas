Attribute VB_Name = "test_Logging"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_logging()
    Dim driver As SeleniumVBA.WebDriver
    Dim fruits As SeleniumVBA.WebElement
    Dim html As String

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome enableLogging:=True
    driver.OpenBrowser
    
    html = "<!DOCTYPE html><html><head><title>Test Select</title></head><body>"
    html = html & "<div>Select your preference:</div>"
    html = html & "<select multiple='' id='fruits'>"
    html = html & "<option value='banana'>Banana</option>"
    html = html & "<option value='apple'>Apple</option>"
    html = html & "<option value='orange'>Orange</option>"
    html = html & "<option value='grape'>Grape</option>"
    html = html & "</select>"
    html = html & "</body></html>"
    
    driver.NavigateToString html
    driver.Wait 1000
    
    Set fruits = driver.FindElement(By.ID, "fruits")
    
    fruits.SelectByVisibleText "Banana"
    driver.Wait
    fruits.SelectByIndex 2  'Apple
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
    
    Debug.Assert fruits.GetSelectedOption.GetText = "Grape"
    Debug.Assert driver.FindElementByCssSelector("option[value='grape']", fruits).IsSelected
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
