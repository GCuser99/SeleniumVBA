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

Sub test_file_download()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    driver.StartFirefox
    
    Set caps = driver.CreateCapabilities

    'caps.SetPreference "download.default_directory", ".\" 'download to same directory as this excel file
    'caps.SetPreference "download.prompt_for_download", False
    'caps.SetPreference "plugins.always_open_pdf_externally", True 'if its a pdf then bypass the pdf viewer
    
    'this does the above in one line
    caps.SetDownloadPrefs downloadFolderPath:=".\", promptForDownload:=False, disablePDFViewer:=True

    driver.OpenBrowser caps
        
    driver.NavigateTo "https://www.selenium.dev/selenium/web/downloads/download.html"
    driver.Wait 500
    
    driver.DeleteFiles ".\file_1.txt", ".\file_2.jpg"
    
    driver.FindElementByCssSelector("#file-1").Click
    driver.WaitForDownload ".\file_1.txt"

    driver.FindElementByCssSelector("#file-2").Click
    driver.WaitForDownload ".\file_2.jpg"

    driver.DeleteFiles ".\file_1.txt", ".\file_2.jpg"
            
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
    
    driver.DeleteFiles "printpage.pdf"
    
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
