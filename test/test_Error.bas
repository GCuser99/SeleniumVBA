Attribute VB_Name = "test_Error"
Sub test_error1()
    Dim driver As SeleniumVBA.WebDriver
    Dim html As String
    Dim elem As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    
    On Error Resume Next
    driver.ImplicitMaxWait = 2000
    Debug.Assert Err.Number = 404
    Debug.Assert Err.Description = "invalid session id"
    On Error GoTo 0
    
    driver.OpenBrowser
    
    On Error Resume Next
    driver.ImplicitMaxWait = -999
    Debug.Assert Err.Number = 400
    Debug.Assert Left(Err.Description, 54) = "invalid argument: value must be a non-negative integer"
    On Error GoTo 0
    
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
    
    On Error Resume Next
    Set elem = driver.FindElement(By.ID, "css1")
    Debug.Assert Err.Number = 404
    Debug.Assert Left(Err.Description, 93) = "no such element: Unable to locate element: {""method"":""css selector"",""selector"":""[id=""css1""]""}"
    On Error GoTo 0
    
    'this wrongly returns "automation error" as description for twinBASIC DLL as of SeleniumVBA v6.9
    On Error Resume Next
    Set elem = driver.FindElement(By.ID, "css")
    Set elem = elem.FindElementByID("xxxx")
    Debug.Assert Err.Number = 404
    Debug.Assert Left(Err.Description, 93) = "no such element: Unable to locate element: {""method"":""css selector"",""selector"":""[id=""xxxx""]""}"
    On Error GoTo 0
    
    'this wrongly returns "automation error" as description for twinBASIC DLL as of SeleniumVBA v6.9
    On Error Resume Next
    Set elem = driver.FindElement(By.ID, "css").FindElementByID("xxxx")
    Debug.Assert Err.Number = 404
    Debug.Assert Left(Err.Description, 93) = "no such element: Unable to locate element: {""method"":""css selector"",""selector"":""[id=""xxxx""]""}"
    On Error GoTo 0
    
    driver.Wait 1000
    
    'this wrongly returns "automation error" as description for twinBASIC DLL as of SeleniumVBA v6.9
    On Error Resume Next
    driver.FindElement(By.ID, "css").Clear
    Debug.Assert Err.Number = 400
    Debug.Assert Left(Err.Description, 21) = "invalid element state"
    On Error GoTo 0
    
    On Error Resume Next
    Set elem = driver.FindElement(By.ID, "css")
    driver.Clear elem
    Debug.Assert Err.Number = 400
    Debug.Assert Left(Err.Description, 21) = "invalid element state"
    On Error GoTo 0
    
    On Error Resume Next
    driver.NavigateTo "bad_url"
    Debug.Assert Err.Number = 400
    Debug.Assert Left(Err.Description, 16) = "invalid argument"
    On Error GoTo 0
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_chrome_download_error()
    'ChromeDriver ignores download.prompt_for_download preference when in Incognito mode
    'https://github.com/GCuser99/SeleniumVBA/issues/87
    'https://issues.chromium.org/issues/42323611
    'Note: running this sub will create a temp file in SeleniumVBA's code library folder
    
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
   
    driver.StartChrome
    
    'set the directory path for saving download to
    Set caps = driver.CreateCapabilities(initializeFromSettingsFile:=False)
    caps.SetDownloadPrefs downloadFolderPath:=".\", promptForDownload:=False, disablePDFViewer:=True
    
    driver.OpenBrowser caps, incognito:=True
    
    'delete legacy copy if it exists
    driver.DeleteFiles ".\test.pdf"
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/raw/main/dev/test_files/test.pdf"
    
    On Error Resume Next
    driver.WaitForDownload ".\test.pdf", 2000
    Debug.Assert Err.Description = "Error in WaitForDownload method: maximum wait time exceeded while waiting for file download."
    On Error GoTo 0
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
