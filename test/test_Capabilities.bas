Attribute VB_Name = "test_Capabilities"
Option Explicit
Option Private Module

' see also test_FileUpDownload for another example using Capabilities

Sub test_headless()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities

    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    
    'note that WebCapabilities object should be created after starting the driver (StartEdge, StartChrome, of StartFirefox methods)
    Set caps = driver.CreateCapabilities
    
    caps.AddArguments "--headless"  'makes browser run in invisible mode
    
    driver.OpenBrowser caps 'here is where caps is passed to driver
    
    driver.NavigateTo "https://www.google.com/"
    
    Debug.Print "User Agent: " & driver.GetUserAgent

    driver.CloseBrowser
    
    'now let's do it the easy way using optional OpenBrowser parameter...
    driver.OpenBrowser invisible:=True
    
    driver.NavigateTo "https://www.google.com/"
    
    Debug.Print "User Agent: " & driver.GetUserAgent
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_incognito()
    'in private or incognito mode helps keep your browsing private from other people who use your device
    'see https://www.wired.com/story/incognito-mode-explainer/
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    
    Set caps = driver.CreateCapabilities
    
    caps.RunIncognito
    
    driver.OpenBrowser caps  'here is where caps is passed to driver
    
    driver.NavigateTo "https://www.google.com/"
    
    driver.Wait 3000
    
    driver.CloseBrowser
    
    'now let's do it the easy way using optional OpenBrowser parameter...
    driver.OpenBrowser incognito:=True
    
    driver.NavigateTo "https://www.google.com/"
    
    driver.Wait 3000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_initialize_caps_from_file()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities

    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    
    Set caps = driver.CreateCapabilities
    
    'first lets set some preferred capabilities
    caps.RunIncognito
    caps.SetDownloadPrefs
    caps.RemoveControlNotification
    caps.SetUserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.5112.102 Safari/537.36"
    caps.SetProfile ".\User Data\Chrome\profile 1"
    
    'save to json file
    caps.SaveToFile "chrome.json"
    
    'shutdown driver
    driver.Shutdown
    
    'now let's start again
    driver.StartChrome
    
    Set caps = driver.CreateCapabilities
    
    'load the saved capabilities into new instance of caps
    caps.LoadFromFile "chrome.json"
    
    'pass caps to OpenBrowser
    driver.OpenBrowser caps
    
    driver.NavigateTo "https://www.google.com/"
    
    driver.Wait 3000
    
    driver.CloseBrowser
    
    'lastly, do above the easy way using optional OpenBrowser parameter...
    driver.OpenBrowser capabilitiesFilePath:="chrome.json"
    
    driver.NavigateTo "https://www.google.com/"
    
    driver.Wait 3000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_user_profile()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities

    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    
    Set caps = driver.CreateCapabilities
    
    'this will create and populate a profile if it doesn't yet exist,
    'otherwise will use a previously created profile
    'recommended to customize your Selenium profiles in a different location
    'than the profiles in AppData to avoid conflicts with manual browsing
    'must specify the path to profile, not just the profile name
    caps.SetProfile ".\User Data\Chrome\profile 1"
    
    driver.OpenBrowser caps
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_unhandledPrompts()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    
    Set caps = driver.CreateCapabilities
    
    'try different settings here to see what happens with flow below
    caps.SetUnhandledPromptBehavior svbaAccept
    
    driver.OpenBrowser caps

    driver.NavigateTo "https://www.google.com"
    
    driver.ExecuteScript "alert('HI');"
    
    driver.Wait 2000
    
    Debug.Print driver.GetTitle
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
