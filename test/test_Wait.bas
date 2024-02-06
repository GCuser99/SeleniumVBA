Attribute VB_Name = "test_Wait"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_ImplicitMaxWait()
    Dim driver As SeleniumVBA.WebDriver
    Dim html1 As String
    Dim html2 As String
    Dim timeDelay As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartChrome
    driver.OpenBrowser
    
    timeDelay = 5000
    
    'set the implicit wait for finding element(s)
    driver.ImplicitMaxWait = timeDelay + 100
    
    'create an html with an element of interest that waits to load a second html

    html1 = "<!DOCTYPE html>" & _
    "<html>" & _
    "<script>" & _
    "function calling(){" & _
    "      setTimeout(""location.replace('snippet2.html')""," & timeDelay & ");" & _
    "}" & _
    "</script>" & _
    "<body onLoad=""calling();"">" & _
    "<div>Not here yet...</div>" & _
    "</body>" & _
    "</html>"

    'create the second html to be loaded by first
    html2 = "<!DOCTYPE html><html><body><div id='here'>I'm here after " & timeDelay & " ms!</div></body></html>"
    
    driver.SaveStringToFile html1, ".\snippet1.html"
    driver.SaveStringToFile html2, ".\snippet2.html"

    driver.NavigateToFile ".\snippet1.html"
    
    'wait until the second html is loaded
    driver.FindElement By.ID, "here"

    driver.Wait 500
        
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitUntilDisplayed()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim html As String
    Dim timeDelay As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartChrome
    driver.OpenBrowser
    
    timeDelay = 5000
    
    'create an html with a script to hide the display of an element
    
    html = "<!DOCTYPE html>" & _
    "<html>" & _
    "<body>" & _
    "<div id='testDiv'>I'm ready now after " & timeDelay & " ms!</div>" & _
    "<script>" & _
    "  var content = document.getElementById('testDiv');" & _
    "  content.style.display='none';" & _
    "  setTimeout(function(){" & _
    "    content.style.display='inline';" & _
    "  }, " & timeDelay & ");" & _
    "</script>" & _
    "</body>" & _
    "</html>"
    
    driver.SaveStringToFile html, ".\snippet.html"

    driver.NavigateToFile ".\snippet.html"
    
    'find the "not displayed" element
    Set elem = driver.FindElement(By.ID, "testDiv")
    
    Debug.Print "Is displayed?:", driver.IsDisplayed(elem)
    
    'wait for it to display...
    driver.WaitUntilDisplayed elem
    
    Debug.Print "Is displayed?:", driver.IsDisplayed(elem)
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitUntilNotDisplayed()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim html As String
    Dim timeDelay As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartChrome
    driver.OpenBrowser
    
    timeDelay = 5000
    
    'create an html with a script to block the display of an element
    
    html = "<!DOCTYPE html>" & _
    "<html>" & _
    "<body>" & _
    "<div id='testDiv'>I'm displayed for " & timeDelay & " ms...</div>" & _
    "<script>" & _
    "  var content = document.getElementById('testDiv');" & _
    "  content.style.display='inline';" & _
    "  setTimeout(function(){" & _
    "    content.style.display='none';" & _
    "  }, " & timeDelay & ");" & _
    "</script>" & _
    "</body>" & _
    "</html>"
    
    driver.SaveStringToFile html, ".\snippet.html"

    driver.NavigateToFile ".\snippet.html"
    
    'find the "not displayed" element
    Set elem = driver.FindElement(By.ID, "testDiv")
    
    Debug.Print "Is displayed?:", driver.IsDisplayed(elem)
    
    'wait for it to disappear...
    driver.WaitUntilNotDisplayed elem
    
    Debug.Print "Is displayed?:", driver.IsDisplayed(elem)
    
    'WaitUntilNotDisplayed allows for method chaining too
    'Debug.Print "Is displayed?:", driver.WaitUntilNotDisplayed(elem).IsDisplayed
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitUntilNotPresent()
    Dim driver As SeleniumVBA.WebDriver
    Dim html1 As String
    Dim html2 As String
    Dim timeDelay As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartChrome
    driver.OpenBrowser
    
    timeDelay = 5000
    
    'create an html with an element of interest that waits to load a second html

    html1 = "<!DOCTYPE html>" & _
    "<html>" & _
    "<script>" & _
    "function calling(){" & _
    "      setTimeout(""location.replace('snippet2.html')""," & timeDelay & ");" & _
    "}" & _
    "</script>" & _
    "<body onLoad=""calling();"">" & _
    "<div id='testDiv'>I'm present!</div>" & _
    "</body>" & _
    "</html>"

    'create the second html to be loaded by first
    html2 = "<!DOCTYPE html><html><body><div>I'm gone after " & timeDelay & " ms!</div></body></html>"
    
    driver.SaveStringToFile html1, ".\snippet1.html"
    driver.SaveStringToFile html2, ".\snippet2.html"

    driver.NavigateToFile ".\snippet1.html"
    
    'wait until the second html is loaded
    driver.WaitUntilNotPresent By.ID, "testDiv"

    driver.Wait 500
        
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitForDownload()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
   
    driver.StartChrome
    
    'set the directory path for saving download to
    Set caps = driver.CreateCapabilities
    caps.SetDownloadPrefs ".\"
    driver.OpenBrowser caps
    
    'delete legacy copy if it exists
    driver.DeleteFiles ".\test.pdf"
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/raw/main/dev/test_files/test.pdf"
    
    'wait until the download is complete
    driver.WaitForDownload ".\test.pdf"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitUntilDisplayed2()
    Dim driver As SeleniumVBA.WebDriver
    Dim searchButton As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    
    Set searchButton = driver.FindElement(By.Name, "btnK")
    
    driver.Wait 500
    
    'search button is there, but not interactable...
    Debug.Print "Is search button interactable yet? " & searchButton.IsDisplayed
    
    driver.FindElement(By.Name, "q").SendKeys "Interactable"

    'searchButton.Click 'will often throw an error here because it takes some time
    'for search button to get ready after typing search phrase
    Debug.Print "Is search button interactable yet? " & searchButton.IsDisplayed
    
    'can place an explicit Wait here but another way is to use WaitUntilReady method
    'it returns the "ready" input element object so can use methods on same line
    searchButton.WaitUntilDisplayed().Click
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub


