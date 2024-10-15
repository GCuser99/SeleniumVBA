Attribute VB_Name = "test_Firefox"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

'To run in Geckodriver, you must have the Firefox browser installed, and then download the
'appropriate version of geckodriver.exe from the following link:
'
'https://github.com/mozilla/geckodriver/releases
'
'The Firefox Geckodriver is nearly as functional as the Chrome/Edge drivers. There are just
'a few limitations that need attention...
'
'Known limitations for of Geckodriver:
'
'- Shutdown Method not recognized (currrently using taskkill to shutdown)
'- Multi-sessions not supported
'- GetSessionsInfo not functional
'- PrintScale method of PrintSettings class does not seem to have effect
'
Sub test_InstallAddon()
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartFirefox
    driver.OpenBrowser
    
    'this is a Firefox browser only method - use AddExtensions method of WebCapabilities for Edge/Chrome
    driver.InstallAddon Environ("USERPROFILE") & "\Documents\SeleniumVBA\extensions\darkreader-4.9.94.xpi"
    
    driver.Wait 5000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_logging()
    Dim driver As SeleniumVBA.WebDriver, fruits As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.CommandWindowStyle = vbNormalFocus
    
    'True enables verbose logging
    driver.StartFirefox , , True
    
    driver.OpenBrowser

    driver.NavigateTo "https://jsbin.com/osebed/2"
    driver.Wait
    
    Set fruits = driver.FindElement(By.ID, "fruits")
    
    If fruits.IsMultiSelect Then
        fruits.SelectByVisibleText "Banana"
        driver.Wait
        fruits.SelectByIndex 2 'Apple
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
    End If
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_file_download()
    Dim driver As SeleniumVBA.WebDriver, caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartFirefox
    
    driver.DeleteFiles ".\BrowserStack - List of devices to test*"
    
    Set caps = driver.CreateCapabilities
    
    caps.SetDownloadPrefs downloadFolderPath:=".\", promptForDownload:=False, disablePDFViewer:=True
    
    Debug.Print caps.ToJson

    driver.OpenBrowser caps
    
    driver.NavigateTo "https://www.browserstack.com/test-on-the-right-mobile-devices"
    driver.Wait 500
    
    'driver.FindElementByID("accept-cookie-notification").Click
    'driver.Wait 500
    
    driver.FindElementByCssSelector(".icon-csv").ScrollIntoView yOffset:=-150
    driver.Wait 1000
    
    driver.FindElementByCssSelector(".icon-csv").Click
    
    driver.WaitForDownload ".\BrowserStack - List of devices to test on.csv"
            
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_print()
    Dim driver As SeleniumVBA.WebDriver
    Dim settings As SeleniumVBA.WebPrintSettings
    Dim keys As SeleniumVBA.WebKeyboard

    Set driver = SeleniumVBA.New_WebDriver
    Set settings = SeleniumVBA.New_WebPrintSettings
    Set keys = SeleniumVBA.New_WebKeyboard
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    driver.StartFirefox
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    
    driver.FindElement(By.ID, "searchInput").SendKeys "Leonardo da Vinci" & keys.EnterKey
    
    driver.Wait 1000
    
    settings.Units = svbaInches
    settings.MarginsAll = 0.4
    settings.Orientation = svbaPortrait
    settings.PrintScale = 0.25
    'settings.PageRanges "1-2"  'prints the first 2 pages
    'settings.PageRanges 1, 2   'prints the first 2 pages
    'settings.PageRanges 2       'prints only the 2nd page
    
    'prints pdf file to specified filePath parameter (defaults to .\printpage.pdf)
    driver.PrintToPDF , settings

    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_GetSessionInfo()
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartFirefox
    
    driver.OpenBrowser
    
    'firefox does not support "Get All Sessions" command
    
    Debug.Print SeleniumVBA.WebJsonConverter.ConvertToJson(driver.GetSessionsInfo, 4)
    
    driver.Wait 1000
    driver.CloseBrowser
    
    driver.Shutdown
End Sub

Sub test_firefox_json_viewer_bug()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    Dim jsonStr As String
    
    'see bug report https://bugzilla.mozilla.org/show_bug.cgi?id=1797871
    'this tests function fixFirefoxBug1797871 in WebDriver class to fix problem
    
    jsonStr = "{""key1"": ""simple json example"",""key2"": ""for firefox bug report"",""key3"": ""utf-16 encoding"",""key4"": ""this does not work with firefox json viewer""}"

    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartFirefox
    
    driver.SaveStringToFile jsonStr, "test.json"
    
    Set caps = driver.CreateCapabilities
    caps.SetPreference "devtools.jsonview.enabled", True '(this is the default)
    
    driver.OpenBrowser caps:=caps

    driver.NavigateToFile "test.json"
    driver.Wait 2000
    
    driver.FindElementByID("rawdata-tab").Click
    
    driver.Wait 3000
    
    Debug.Assert driver.PageToJSONObject()("key1") = "simple json example"
    
    driver.CloseBrowser
    
    caps.SetPreference "devtools.jsonview.enabled", False
    
    driver.OpenBrowser caps:=caps

    driver.NavigateToFile "test.json"
    driver.Wait 5000
    
    Debug.Assert driver.PageToJSONObject()("key1") = "simple json example"
   
    driver.Shutdown
End Sub

Sub test_geolocation()
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartFirefox
    driver.OpenBrowser
    
    'firefox does not support geolocation commands
    
    driver.SetGeolocation 41.1621429, -8.6219537
  
    driver.NavigateTo "https://www.gps-coordinates.net/my-location"
    driver.Wait 1000
    
    'print the name of the location
    Debug.Print driver.FindElementByXPath("//*[@id='addr']").GetText
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
