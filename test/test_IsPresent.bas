Attribute VB_Name = "test_IsPresent"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")
    
Sub test_IsPresent()
    Dim driver As SeleniumVBA.WebDriver
    Dim htmlStr As String
    Dim elem As SeleniumVBA.WebElement
    Dim elems As SeleniumVBA.WebElements
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    driver.StartChrome
    driver.OpenBrowser
    
    htmlStr = "<html><body>"
    htmlStr = htmlStr & "<div id='parent1'><div id='child1'><p>child1 from parent1</p></div><div id='child2'><p>child2 from parent1</p></div></div>"
    htmlStr = htmlStr & "<div id='parent2'><div id='child1'><p>child1 from parent2</p></div><div id='child2'><p>child2 from parent2</p></div></div>"
    htmlStr = htmlStr & "</body></html>"
    
    driver.SaveStringToFile htmlStr, ".\snippet.html"
    driver.NavigateToFile ".\snippet.html"
    
    driver.Wait 500
    
    Debug.Print "does any child1 exist:", driver.IsPresent(By.XPath, "//div[@id = 'child1']", , elem)
    Debug.Print "first found child1 text:", elem.GetText
    
    Debug.Print "child1 of parent2 present:", driver.FindElement(By.ID, "parent2").IsPresent(By.XPath, ".//div[@id = 'child1']")
    
    'waiting up to 3 secs for elem to be present
    Debug.Print "child2 of parent2 present:", driver.FindElement(By.ID, "parent2").IsPresent(By.XPath, ".//div[@id = 'child2']", 3000, elem)
    Debug.Print "child2 of parent2 text:", elem.GetText
    
    'waiting up to 3 secs for elem to be present
    Debug.Print "child3 of parent2 present:", driver.FindElement(By.ID, "parent2").IsPresent(By.XPath, ".//div[@id = 'child3']", 3000, elem)
    Debug.Print "child3 of parent2 reference is nothing:", elem Is Nothing
    
    Set elems = driver.FindElements(By.CssSelector, "[id^='parent']")
    
    For Each elem In elems
        Debug.Print "child1 of " & elem.GetAttribute("id") & " present:", elem.IsPresent(By.XPath, ".//div[@id = 'child1']")
        Debug.Print "child2 of " & elem.GetAttribute("id") & " present:", elem.IsPresent(By.XPath, ".//div[@id = 'child2']")
        Debug.Print "child3 of " & elem.GetAttribute("id") & " present:", elem.IsPresent(By.XPath, ".//div[@id = 'child3']")
    Next elem
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_IsPresent_wait()
    Dim driver As SeleniumVBA.WebDriver
    Dim html1 As String
    Dim html2 As String
    Dim timeDelay As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartChrome
    driver.OpenBrowser
    
    'create an html that waits to load a second html with a div element of interest
    timeDelay = 5000

    html1 = "<!DOCTYPE html>" & _
    "<html>" & _
    "<script>" & _
    "function loading(){" & _
    "      setTimeout(""location.replace('snippet2.html')""," & timeDelay & ");" & _
    "}" & _
    "</script>" & _
    "<body onLoad=""loading();"">" & _
    "<div>waiting for load...</div>" & _
    "</body>" & _
    "</html>"

    'create the second html to be loaded
    html2 = "<!DOCTYPE html><html><body><div id='testDiv'>I'm here now after " & timeDelay & " ms!</div></body></html>"
    
    driver.SaveStringToFile html1, ".\snippet1.html"
    driver.SaveStringToFile html2, ".\snippet2.html"

    driver.NavigateToFile ".\snippet1.html"

    'wait up to 20 secs for the div from the second html gets loaded
    Debug.Print driver.IsPresent(By.ID, "testDiv", 20000)
    
    driver.Wait 1500
        
    driver.CloseBrowser
    driver.Shutdown
End Sub
