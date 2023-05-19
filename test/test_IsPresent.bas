Attribute VB_Name = "test_IsPresent"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")
    
Sub test_isPresent_xpath()
    Dim driver As SeleniumVBA.WebDriver
    Dim htmlStr As String
    Dim elem As SeleniumVBA.WebElement
    Dim elems As SeleniumVBA.WebElements
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    htmlStr = "<html><body>"
    htmlStr = htmlStr & "<div id='parent1'><div id='child1'><p>child1 from parent1</p></div><div id='child2'><p>child2 from parent1</p></div></div>"
    htmlStr = htmlStr & "<div id='parent2'><div id='child1'><p>child1 from parent2</p></div><div id='child2'><p>child2 from parent2</p></div></div>"
    htmlStr = htmlStr & "</body></html>"
    
    driver.SaveStringToFile htmlStr, ".\snippet.html"
    driver.NavigateToFile ".\snippet.html"
    
    driver.Wait 500
    
    Debug.Print "does any child1 exist", driver.IsPresent(By.XPath, "//div[@id = 'child1']", , elem)
    Debug.Print "first found child1 text", elem.GetText
    
    Debug.Print "child1 of parent2 exists:", driver.FindElement(By.ID, "parent2").IsPresent(By.XPath, ".//div[@id = 'child1']")
    
    'waiting up to 3 secs for elem to be present
    Debug.Print "child2 of parent2 exists:", driver.FindElement(By.ID, "parent2").IsPresent(By.XPath, ".//div[@id = 'child2']", 3000, elem)
    Debug.Print "child2 of parent2 text", elem.GetText
    
    'waiting up to 3 secs for elem to be present
    Debug.Print "child3 of parent2 exists:", driver.FindElement(By.ID, "parent2").IsPresent(By.XPath, ".//div[@id = 'child3']", 3000, elem)
    Debug.Print "child3 of parent2 reference is nothing:", elem Is Nothing
    
    Set elems = driver.FindElements(By.cssSelector, "[id^='parent']")
    
    For Each elem In elems
        Debug.Print "child1 of " & elem.GetAttribute("id") & " exists:", elem.IsPresent(By.XPath, ".//div[@id = 'child1']")
        Debug.Print "child2 of " & elem.GetAttribute("id") & " exists:", elem.IsPresent(By.XPath, ".//div[@id = 'child2']")
        Debug.Print "child3 of " & elem.GetAttribute("id") & " exists:", elem.IsPresent(By.XPath, ".//div[@id = 'child3']")
    Next elem
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
