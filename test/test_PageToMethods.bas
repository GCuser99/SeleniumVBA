Attribute VB_Name = "test_PageToMethods"
Option Explicit
Option Private Module

Sub test_PageToHTMLMethods()
    Dim driver As SeleniumVBA.WebDriver
    Dim htmlDoc As HTMLDocument, url As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    url = "https://it.wikipedia.org/wiki/Pagina_principale"
    
    driver.NavigateTo url
    driver.Wait 1000
    
    'use DOM to parse htmlDocument here if desired....
    'html DOM can much faster than Selenium if complicated parse is needed
    Set htmlDoc = driver.PageToHTMLDoc(sanitize:=False)
    Debug.Print htmlDoc.body.ChildNodes.Length
    
    'save raw page to html file
    driver.PageToHTMLFile "source_raw.html", sanitize:=False
    
    'note that santization leaves DOM tree intact
    Set htmlDoc = driver.PageToHTMLDoc(sanitize:=True)
    Debug.Print htmlDoc.body.ChildNodes.Length
    
    'save sanitized page to html file
    driver.PageToHTMLFile "source_sanitized.html", sanitize:=True
    
    'this is much faster because santization disables "online" dynamic elements
    driver.NavigateToFile "source_sanitized.html"
    driver.Wait 1000
    
    'uncomment the following to see how long it takes to render unsanitized html file - be patient!
    'driver.NavigateToFile "source_raw.html"
    'driver.Wait 1000
    
    driver.Shutdown
End Sub

Sub test_PageToXMLMethods()
    Dim driver As SeleniumVBA.WebDriver
    Dim xmlDoc As DOMDocument60, url As String

    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser

    url = "https://www.w3schools.com/xml/note.xml"

    driver.NavigateTo url
    driver.Wait 500
    
    'save page to xml file
    driver.PageToXMLFile "test.xml"
    
    'load up an xml document for further processing
    Set xmlDoc = driver.PageToXMLDoc
    
    Debug.Print xmlDoc.SelectSingleNode("//heading").text
        
    'read the test file back into browser
    driver.NavigateToFile "test.xml"
    
    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_PageToJSONMethods()
    Dim driver As SeleniumVBA.WebDriver
    Dim json As Collection, url As String

    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser

    url = "https://api.github.com/repos/gcuser99/seleniumVBA/releases"

    driver.NavigateTo url
    driver.Wait 1000
    
    'save page to json file
    driver.PageToJSONFile "test.json"
    
    'load up a json object for further processing
    Set json = driver.PageToJSONObject
    Debug.Print json(1)("url")
        
    'read the test file back into browser
    driver.NavigateToFile "test.json"
    
    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
