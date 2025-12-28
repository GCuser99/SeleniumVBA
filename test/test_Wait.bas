Attribute VB_Name = "test_Wait"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_ImplicitMaxWait()
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 10000
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/ajaxy_page.html"
    
    driver.FindElementByName("typer").SendKeys "Hello New World!"
    driver.FindElementByID("red").Click
    driver.FindElementByName("submit").Click
    
    'wait for element creation...
    Debug.Assert driver.FindElementByClassName("label").GetText = "Hello New World!"
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitUntilDisplayed()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim html As String
    Dim timeDelay As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    timeDelay = 3000
    
    'create an html with a script to hide the display of an element
    
    html = "<!DOCTYPE html>" & _
    "<html>" & _
    "<head><title>Test Wait Until Displayed</title></head>" & _
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

    driver.NavigateToString html
    
    'find the "not displayed" element
    Set elem = driver.FindElement(By.ID, "testDiv")
    
    Debug.Assert driver.IsDisplayed(elem) = False
    
    'wait for it to display...
    driver.WaitUntilDisplayed elem
    
    Debug.Assert driver.IsDisplayed(elem) = True
    
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
    
    driver.StartChrome
    driver.OpenBrowser
    
    timeDelay = 3000
    
    'create an html with a script to block the display of an element
    
    html = "<!DOCTYPE html>" & _
    "<html>" & _
    "<head><title>Test Wait Until Not Displayed</title></head>" & _
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

    driver.NavigateToString html
    
    'find the "not displayed" element
    Set elem = driver.FindElement(By.ID, "testDiv")
    
    Debug.Assert driver.IsDisplayed(elem) = True
    
    'wait for it to disappear...
    driver.WaitUntilNotDisplayed elem
    
    Debug.Assert driver.IsDisplayed(elem) = False
    
    'WaitUntilNotDisplayed allows for method chaining too
    'Debug.Print "Is displayed?:", driver.WaitUntilNotDisplayed(elem).IsDisplayed
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitUntilNotPresent()
    Dim driver As SeleniumVBA.WebDriver
    Dim html As String
    Dim timeDelay As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    timeDelay = 3000
    
    'create an html with an element of interest that is removed after a delay
    html = vbNullString
    html = html & "<!DOCTYPE html>" & vbCrLf
    html = html & "<html>" & vbCrLf
    html = html & "    <head>" & vbCrLf
    html = html & "        <title>Is Not Present?</title>" & vbCrLf
    html = html & "        <script>" & vbCrLf
    html = html & "            function removeElementWithDelay(delay, elementId) {" & vbCrLf
    html = html & "                setTimeout(function() {" & vbCrLf
    html = html & "                const elementToRemove = document.getElementById(elementId);" & vbCrLf
    html = html & "                if (elementToRemove) {" & vbCrLf
    html = html & "                    elementToRemove.remove();" & vbCrLf
    html = html & "                }" & vbCrLf
    html = html & "                }, delay);" & vbCrLf
    html = html & "            }" & vbCrLf
    html = html & "        </script>" & vbCrLf
    html = html & "    </head>" & vbCrLf
    html = html & "    <body onLoad=""removeElementWithDelay(" & timeDelay & ", 'div here');"">" & vbCrLf
    html = html & "        <div id=""div here"">I'm here...</div>" & vbCrLf
    html = html & "    </body>" & vbCrLf
    html = html & "</html>"

    driver.NavigateToString html
    
    'wait until the div is removed
    driver.WaitUntilNotPresent By.ID, "div here"

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
    
    driver.DeleteFiles ".\test.pdf"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_wait_for_idle_state()
    Dim driver As New WebDriver
    Dim elms_title1 As WebElements ' List of article elements immediately after navigation
    Dim elms_title2 As WebElements ' List of article elements after waiting with WaitForIdelState
    Dim elms_title3 As WebElements ' List of article elements after waiting 3 seconds more

    driver.StartChrome
    driver.OpenBrowser

    driver.ImplicitMaxWait = 10000

    driver.NavigateTo "https://note.com/topic/novel"
    
    Set elms_title1 = driver.FindElementsByCssSelector(".a-link.m-largeNoteWrapper__link.fn")

    driver.WaitForIdleNetwork 1500
    
    Set elms_title2 = driver.FindElementsByCssSelector(".a-link.m-largeNoteWrapper__link.fn")
    
    driver.Wait 3000
    
    Set elms_title3 = driver.FindElementsByCssSelector(".a-link.m-largeNoteWrapper__link.fn")

    Debug.Assert elms_title1.Count < elms_title2.Count
    Debug.Assert elms_title2.Count = elms_title3.Count 'idle state achieved
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
