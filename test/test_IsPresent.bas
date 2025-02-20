Attribute VB_Name = "test_IsPresent"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_IsPresent()
    Dim driver As SeleniumVBA.WebDriver
    Dim html As String
    Dim elem As SeleniumVBA.WebElement
    Dim elems As SeleniumVBA.WebElements
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    html = "<html><head><title>Test Is Element Present</title></head><body>"
    html = html & "<div id='parent1'><div id='child1'><p>child1 from parent1</p></div><div id='child2'><p>child2 from parent1</p></div></div>"
    html = html & "<div id='parent2'><div id='child1'><p>child1 from parent2</p></div><div id='child2'><p>child2 from parent2</p></div></div>"
    html = html & "</body></html>"
    
    driver.NavigateToString html
    
    driver.Wait 500
    
    Debug.Assert driver.IsPresent(By.XPath, "//div[@id = 'child1']", elemFound:=elem) = True  'does any child1 exist?
    Debug.Assert elem.GetText = "child1 from parent1"
    
    Debug.Assert driver.FindElement(By.ID, "parent2").IsPresent(By.XPath, ".//div[@id = 'child1']") = True 'child1 of parent2 present?
    
    'waiting up to 3 secs for elem to be present
    Debug.Assert driver.FindElement(By.ID, "parent2").IsPresent(By.XPath, ".//div[@id = 'child2']", 3000, elemFound:=elem) = True 'child2 of parent2 present?
    Debug.Assert elem.GetText = "child2 from parent2"
    
    'waiting up to 3 secs for elem to be present
    Debug.Assert driver.FindElement(By.ID, "parent2").IsPresent(By.XPath, ".//div[@id = 'child3']", 3000, elemFound:=elem) = False 'child3 of parent2 present
    Debug.Assert elem Is Nothing 'child3 of parent2 reference is nothing
    
    Set elems = driver.FindElements(By.CssSelector, "[id^='parent']")
    
    For Each elem In elems
        Debug.Assert elem.IsPresent(By.XPath, ".//div[@id = 'child1']") = True
        Debug.Assert elem.IsPresent(By.XPath, ".//div[@id = 'child2']") = True
        Debug.Assert elem.IsPresent(By.XPath, ".//div[@id = 'child3']") = False
    Next elem
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_IsPresent_wait()
    Dim driver As SeleniumVBA.WebDriver
    Dim html As String
    Dim timeDelay As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    timeDelay = 3000

    'create an html that creates a new element after a specified time delay
    html = vbNullString
    html = html & "<!DOCTYPE html>" & vbCrLf
    html = html & "<html>" & vbCrLf
    html = html & "    <head>" & vbCrLf
    html = html & "        <title>Is Present?</title>" & vbCrLf
    html = html & "        <script>" & vbCrLf
    html = html & "            function insertDivWithDelay(delay, parentElementId) {" & vbCrLf
    html = html & "                setTimeout(function() {" & vbCrLf
    html = html & "                    const newDiv = document.createElement(""div"");" & vbCrLf
    html = html & "                    newDiv.id = 'new div';" & vbCrLf
    html = html & "                    newDiv.textContent = ""This div appeared after "" + delay / 1000 + "" seconds."";" & vbCrLf
    html = html & "                    newDiv.style.color = ""blue"";" & vbCrLf
    html = html & "                    const parentElement = document.getElementById(parentElementId);" & vbCrLf
    html = html & "                    if (parentElement) {" & vbCrLf
    html = html & "                        parentElement.append(newDiv);" & vbCrLf
    html = html & "                    } else {" & vbCrLf
    html = html & "                    }" & vbCrLf
    html = html & "                }, delay);" & vbCrLf
    html = html & "            }" & vbCrLf
    html = html & "        </script>" & vbCrLf
    html = html & "    </head>" & vbCrLf
    html = html & "    <body onLoad=""insertDivWithDelay(" & timeDelay & ", 'parent');"">" & vbCrLf
    html = html & "        <div id=""parent"">waiting for load...</div>" & vbCrLf
    html = html & "    </body>" & vbCrLf
    html = html & "</html>"

    driver.NavigateToString html

    'wait up to 20 secs for the new div is created
    Debug.Assert driver.IsPresent(By.ID, "new div", 20000)
    
    driver.Wait 1000

    driver.CloseBrowser
    driver.Shutdown
End Sub
