Attribute VB_Name = "test_AttributesProperties"
'https://stackoverflow.com/questions/6003819/what-is-the-difference-between-properties-and-attributes-in-html

Sub test_element_attributes_and_properties()
    Dim Driver As New WebDriver, str As String
    
    str = "<!DOCTYPE html><html><body><input id=""the-input"" type=""text"" value=""Sally""></body></html>"
    filepath = ".\snippet.html"
    
    Driver.StartChrome
    Driver.OpenBrowser
    
    Driver.SaveHTMLToFile str, filepath
    
    Driver.NavigateTo "file:///" & filepath
    
    Driver.Wait 1000
    
    'sends "John" to input box
    Driver.FindElementByID("the-input").SendKeys "John"
    
    'this gets the initial default attribute value "Sally"
    Debug.Print "value attribute:", Driver.FindElementByID("the-input").GetAttribute("value")
    
    'this gets the current value of the input box, "John"
    Debug.Print "value property:", Driver.FindElementByID("the-input").GetProperty("value")
    
    'Note that after browser parses html, new proprties are created
    Debug.Print "defaultValue property:", Driver.FindElementByID("the-input").GetProperty("defaultValue")
    
    Driver.Wait 1000
    
    str = "<!DOCTYPE html><html><body><h1>Show Checkboxes</h1><form action='/action_page.php'><input type='checkbox' id='vehicle1' name='vehicle1' value='Bike'><label for='vehicle1'> I have a bike</label><br><input type='checkbox' id='vehicle2' name='vehicle2' value='Car'><label for='vehicle2'> I have a car</label><br><input type='checkbox' id='vehicle3' name='vehicle3' value='Boat' checked><label for='vehicle3'> I have a boat</label><br><br><input type='submit' value='Submit'></form></body></html>"
    Driver.SaveHTMLToFile str, filepath
    
    Driver.NavigateTo "file:///" & filepath
    
    Driver.Wait 1000
    
    Driver.FindElementByID("vehicle1").Click
    
    'Note that after browser parses html, the checked property is created for vehicle1 and vehicle2 checkboxes
    Debug.Print "checked property for vehicle1:", Driver.FindElementByID("vehicle1").GetProperty("checked")
    Debug.Print "checked property for vehicle2:", Driver.FindElementByID("vehicle2").GetProperty("checked")
    Debug.Print "checked property for vehicle3:", Driver.FindElementByID("vehicle3").GetProperty("checked")
    
    'the html for vehicle3 has a "checked" attribute so it gets returned by getAttribute, but vehicle1 and vehicle2 do not and thus return null string
    Debug.Print "checked attribute for vehicle1:", Driver.FindElementByID("vehicle1").GetAttribute("checked")
    Debug.Print "checked attribute for vehicle2:", Driver.FindElementByID("vehicle2").GetAttribute("checked")
    Debug.Print "checked attribute for vehicle3:", Driver.FindElementByID("vehicle3").GetAttribute("checked")
    
    Driver.Wait 1000
    
    Driver.CloseBrowser
    Driver.Shutdown
    
End Sub

Sub test_css_property()
    Dim Driver As New WebDriver, str As String, color As String
    
    Driver.StartChrome
    Driver.OpenBrowser
    
    Driver.NavigateTo "https://example.com"
    
    Driver.Wait
    
    Debug.Print "Background color: " & Driver.FindElementByTagName("html").GetCSSProperty("background-color")
    Debug.Print "Font family: " & Driver.FindElementByTagName("html").GetCSSProperty("font-family")
    
    Driver.Wait
    
    Driver.CloseBrowser
    Driver.Shutdown
End Sub

Sub test_element_aria()
    Dim Driver As New WebDriver, str As String
    
    str = "<!DOCTYPE html><html><body><div role='button' class='xyz' aria-label='Add food' aria-disabled='false' data-tooltip='Add food'><span class='abc' aria-hidden='true'>icon</span></body></html>"
    
    filepath = ".\snippet.html"
    
    Driver.StartChrome
    Driver.OpenBrowser
    
    Driver.SaveHTMLToFile str, filepath
    
    Driver.NavigateTo "file:///" & filepath
    
    Driver.Wait 1000
    
    'sends "John" to input box
    Debug.Print "Label: " & Driver.FindElementByClassName("xyz").GetAriaLabel
    Debug.Print "Role: " & Driver.FindElementByClassName("xyz").GetAriaRole
    
    Driver.CloseBrowser
    Driver.Shutdown
    
End Sub
