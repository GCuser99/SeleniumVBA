Attribute VB_Name = "test_Inputs"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_select()
    Dim driver As SeleniumVBA.WebDriver
    Dim selectElem As SeleniumVBA.WebElement
    Dim textElem As SeleniumVBA.WebElement
    Dim html As String

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser

    html = vbNullString
    html = html & "<!DOCTYPE html>" & vbCrLf
    html = html & "<html>" & vbCrLf
    html = html & "<head>" & vbCrLf
    html = html & "    <title>Test Select</title>" & vbCrLf
    html = html & "    <script>" & vbCrLf
    html = html & "        function updateSelection(selectEl, outputId) {" & vbCrLf
    html = html & "            const selectedTexts = Array.from(selectEl.selectedOptions)" & vbCrLf
    html = html & "                .map(opt => opt.text);" & vbCrLf
    html = html & "            document.getElementById(outputId).textContent =" & vbCrLf
    html = html & "                selectedTexts.length" & vbCrLf
    html = html & "                    ? selectedTexts.join("", "")" & vbCrLf
    html = html & "                    : ""(nothing selected)"";" & vbCrLf
    html = html & "        }" & vbCrLf
    html = html & "    </script>" & vbCrLf
    html = html & "</head>" & vbCrLf
    html = html & "<body>" & vbCrLf
    html = html & vbCrLf
    html = html & "    <div>Select your preference:</div>" & vbCrLf
    html = html & "    <select id=""fruits""" & vbCrLf
    html = html & "            onchange=""updateSelection(this, 'out_fruits')"">" & vbCrLf
    html = html & "        <option value=""banana"">Banana</option>" & vbCrLf
    html = html & "        <option value=""apple"">Apple</option>" & vbCrLf
    html = html & "        <option value=""orange"">Orange</option>" & vbCrLf
    html = html & "        <option value=""grape"">Grape</option>" & vbCrLf
    html = html & "        <option value=""lcgrape"">grape</option>" & vbCrLf
    html = html & "    </select>" & vbCrLf
    html = html & "    <div id=""out_fruits""></div>" & vbCrLf
    html = html & vbCrLf
    html = html & "    <div>Select your preference:</div>" & vbCrLf
    html = html & "    <select multiple id=""fruits_multi""" & vbCrLf
    html = html & "            onchange=""updateSelection(this, 'out_fruits_multi')"">" & vbCrLf
    html = html & "        <option value=""default"" disabled>Choose which one</option>" & vbCrLf
    html = html & "        <option value=""banana"">Banana</option>" & vbCrLf
    html = html & "        <option value=""apple"">Apple</option>" & vbCrLf
    html = html & "        <option value=""orange"">Orange</option>" & vbCrLf
    html = html & "        <option value=""grape"">Grape</option>" & vbCrLf
    html = html & "    </select>" & vbCrLf
    html = html & "    <div id=""out_fruits_multi""></div>" & vbCrLf
    html = html & vbCrLf
    html = html & "    <div>Select your preference:</div>" & vbCrLf
    html = html & "    <select id=""fruits_default""" & vbCrLf
    html = html & "            onchange=""updateSelection(this, 'out_fruits_default')"">" & vbCrLf
    html = html & "        <option value=""default"" disabled selected>Choose which one</option>" & vbCrLf
    html = html & "        <option value=""banana"">Banana</option>" & vbCrLf
    html = html & "        <option value=""apple"">Apple</option>" & vbCrLf
    html = html & "        <option value=""orange"">Orange</option>" & vbCrLf
    html = html & "        <option value=""grape"">Grape</option>" & vbCrLf
    html = html & "    </select>" & vbCrLf
    html = html & "    <div id=""out_fruits_default""></div>" & vbCrLf
    html = html & vbCrLf
    html = html & "</body>" & vbCrLf
    html = html & "</html>" & vbCrLf
    html = html & vbCrLf
    
    driver.NavigateToString html
    driver.Wait (1000)
    
    Set selectElem = driver.FindElementByID("fruits")
    Set textElem = driver.FindElementByID("out_fruits")
    selectElem.SelectByValue ("orange")
    Debug.Assert textElem.GetText = "Orange"
    selectElem.SelectByVisibleText ("grape")
    Debug.Assert textElem.GetText = "grape"
    selectElem.SelectByVisibleText ("Grape")
    Debug.Assert textElem.GetText = "Grape"
    selectElem.SelectByIndex 1
    Debug.Assert textElem.GetText = "Banana"
    Debug.Assert selectElem.GetSelectedOption.IsSelected
    
    Set selectElem = driver.FindElementByID("fruits_multi")
    Set textElem = driver.FindElementByID("out_fruits_multi")
    selectElem.SelectByValue ("orange")
    selectElem.SelectByVisibleText ("Apple")
    Debug.Assert textElem.GetText = "Apple, Orange"

    selectElem.SelectAll
    Debug.Assert textElem.GetText = "Banana, Apple, Orange, Grape"
    selectElem.DeSelectByValue ("orange")
    Debug.Assert textElem.GetText = "Banana, Apple, Grape"
    selectElem.DeSelectAll
    Debug.Assert textElem.GetText = "(nothing selected)"
    selectElem.SelectByIndex 2
    selectElem.SelectByIndex 3
    selectElem.DeSelectByIndex 2
    Debug.Assert textElem.GetText = "Apple"
    
    Set selectElem = driver.FindElementByID("fruits_default")
    Set textElem = driver.FindElementByID("out_fruits_default")
    selectElem.SelectByValue ("orange")
    Debug.Assert textElem.GetText = "Orange"
    selectElem.SelectByIndex 2
    Debug.Assert textElem.GetText = "Banana"
    
    'make sure to set Tools->Options->General->Error Traaping->Break on Unhandled Errors
    On Error Resume Next
    selectElem.DeSelectByIndex (2)
    Debug.Assert Err.Description = "Error in WebDiver: Not allowed for a non-multi-select option element"
    On Error GoTo 0
    
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

