Attribute VB_Name = "test_Attributes"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

'https://stackoverflow.com/questions/6003819/what-is-the-difference-between-properties-and-attributes-in-html

Sub test_element_attributes_and_properties()
    Dim driver As SeleniumVBA.WebDriver, str As String, filePath As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    str = "<!DOCTYPE html><html><body><input id=""the-input"" type=""text"" value=""Sally""></body></html>"
    filePath = ".\snippet.html"
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.SaveStringToFile str, filePath
    
    driver.NavigateToFile filePath
    
    driver.Wait 1000
    
    'sends "John" to input box
    driver.FindElementByID("the-input").SendKeys "John", True
    
    'this gets the initial default attribute value "Sally"
    Debug.Print "value attribute:", driver.FindElementByID("the-input").GetAttribute("value")
    
    'this gets the current value of the input box, "John"
    Debug.Print "value property:", driver.FindElementByID("the-input").GetProperty("value")
    
    'Note that after browser parses html, new properties are created
    Debug.Print "defaultValue property:", driver.FindElementByID("the-input").GetProperty("defaultValue")
    
    driver.Wait 1000
    
    str = "<!DOCTYPE html><html><body><h1>Show Checkboxes</h1><form action='/action_page.php'><input type='checkbox' id='vehicle1' name='vehicle1' value='Bike'><label for='vehicle1'> I have a bike</label><br><input type='checkbox' id='vehicle2' name='vehicle2' value='Car'><label for='vehicle2'> I have a car</label><br><input type='checkbox' id='vehicle3' name='vehicle3' value='Boat' checked><label for='vehicle3'> I have a boat</label><br><br><input type='submit' value='Submit'></form></body></html>"
    driver.SaveStringToFile str, filePath
    
    driver.NavigateToFile filePath
    
    driver.Wait 1000
    
    driver.FindElementByID("vehicle1").Click
    
    'Note that after browser parses html, the checked property is created for vehicle1 and vehicle2 checkboxes
    Debug.Print "checked property for vehicle1:", driver.FindElementByID("vehicle1").GetProperty("checked")
    Debug.Print "checked property for vehicle2:", driver.FindElementByID("vehicle2").GetProperty("checked")
    Debug.Print "checked property for vehicle3:", driver.FindElementByID("vehicle3").GetProperty("checked")
    
    'the html for vehicle3 has a "checked" attribute so it gets returned by getAttribute, but vehicle1 and vehicle2 do not and thus return null string
    Debug.Print "checked attribute for vehicle1:", driver.FindElementByID("vehicle1").GetAttribute("checked")
    Debug.Print "checked attribute for vehicle2:", driver.FindElementByID("vehicle2").GetAttribute("checked")
    Debug.Print "checked attribute for vehicle3:", driver.FindElementByID("vehicle3").GetAttribute("checked")
    
    driver.Wait 1000

    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_css_property()
    Dim driver As SeleniumVBA.WebDriver, str As String, color As String

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://example.com"
    
    driver.Wait
    
    Debug.Print "Background color: " & driver.FindElementByTagName("html").GetCSSProperty("background-color")
    Debug.Print "Font family: " & driver.FindElementByTagName("html").GetCSSProperty("font-family")
    
    driver.Wait
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_element_aria()
    Dim driver As SeleniumVBA.WebDriver, str As String, filePath As String

    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.Path '(this is the default)
    
    str = "<!DOCTYPE html><html><body><div role='button' class='xyz' aria-label='Add food' aria-disabled='false' data-tooltip='Add food'><span class='abc' aria-hidden='true'>icon</span></body></html>"
    
    filePath = ".\snippet.html"

    driver.StartChrome
    driver.OpenBrowser
    
    driver.SaveStringToFile str, filePath
    
    driver.NavigateToFile filePath
    
    driver.Wait 1000
    
    Debug.Print "Label: " & driver.FindElementByClassName("xyz").GetAriaLabel
    Debug.Print "Role: " & driver.FindElementByClassName("xyz").GetAriaRole
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
