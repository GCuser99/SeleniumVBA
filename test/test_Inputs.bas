Attribute VB_Name = "test_Inputs"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_select()
    Dim driver As SeleniumVBA.WebDriver
    Dim fruits As SeleniumVBA.WebElement
    Dim html As String

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
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

Sub test_radio()
    Dim driver As SeleniumVBA.WebDriver
    Dim html As String
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    'create a radio button sample
    html = "<!DOCTYPE html><html><head><title>Test Radio Button</title></head><body>"
    html = html & "<h1>Display Radio Buttons</h1>"
    html = html & "<form action='/action_page.php'>"
    html = html & "  <p>Please select your favorite Web language:</p>"
    html = html & "  <input type='radio' id='html' name='fav_language' value='HTML'>"
    html = html & "  <label for='html'>HTML</label><br>"
    html = html & "  <input type='radio' id='css' name='fav_language' value='CSS'>"
    html = html & "  <label for='css'>CSS</label><br>"
    html = html & "  <input type='radio' id='javascript' name='fav_language' value='JavaScript'>"
    html = html & "  <label for='javascript'>JavaScript</label>"
    html = html & "</form>"
    html = html & "</body></html>"
    
    driver.NavigateToString html
    driver.ActiveWindow.Maximize
    
    driver.Wait 1000
    
    driver.FindElement(By.ID, "css").Click
    
    Debug.Assert driver.FindElement(By.ID, "css").IsSelected
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

