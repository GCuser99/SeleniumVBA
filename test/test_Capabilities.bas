Attribute VB_Name = "test_Capabilities"
Option Explicit
Option Private Module

'see also test_FileUpDownload for another example using Capabilities
Sub test_invisible()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities

    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    
    'note that WebCapabilities object should be created after starting the driver (StartEdge, StartChrome, of StartFirefox methods)
    Set caps = driver.CreateCapabilities
    
    caps.RunInvisible 'makes browser run in invisible mode
    
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
    
    'driver.CloseBrowser
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
    'caps.SetProfile ".\User Data\Chrome\profile 1"
    
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

Sub test_unhandled_prompts()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    
    Set caps = driver.CreateCapabilities
    
    'try different settings here to see what happens with flow below
    caps.SetUnhandledPromptBehavior svbaAccept
    
    driver.OpenBrowser caps

    driver.NavigateTo "https://www.google.com"
    
    driver.ExecuteScript "alert('Hi!');"
    
    driver.Wait 2000
    
    Debug.Print driver.GetTitle
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_detach_browser()
    'use this if you want browser to remain open after shutdown clean-up - only for Chrome/Edge
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.CommandWindowStyle = vbNormalFocus
    
    driver.StartEdge
    
    Set caps = driver.CreateCapabilities
    
    'this sets whether browser is closed (false) or left open (true)
    'when the driver is sent the shutdown command before browser is closed
    'defaults to false
    'only applicable to edge/chrome browsers
    caps.SetDetachBrowser True
    
    driver.OpenBrowser caps
    
    driver.NavigateTo "https://www.google.com/"
    
    driver.Wait 1000
    
    'driver.CloseBrowser 'detach does nothing if browser is closed properly by user
    driver.Shutdown
End Sub

Sub test_kiosk_printing()
    'this advanced test uses kiosk printing to save a webpage to pdf file (Chrome/Edge only)
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    Dim jc As New WebJSonConverter
    Dim settings As New Dictionary
    Dim appState As New Dictionary
    Dim recentDestination As New Dictionary
    Dim customMargins As New Dictionary
    Dim mediaSizeOptions As New Dictionary
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    
    Set caps = driver.CreateCapabilities
    
    caps.AddArguments "--kiosk-printing"
    
    'build the appState dictionary for controling the print settings
    
    'make a print destination
    recentDestination.Add "id", "Save as PDF"
    recentDestination.Add "origin", "local"
    recentDestination.Add "account", ""
    
    'add the destination array to recentDestinations key of the appState dictionary
    appState.Add "recentDestinations", Array(recentDestination)
    
    'now add more properties to appState
    appState.Add "selectedDestination", "Save as PDF" 'this selects from recentDestinations array
    appState.Add "version", 2 'this is required
    appState.Add "isLandscapeEnabled", False
    appState.Add "isHeaderFooterEnabled", True
    appState.Add "scalingType", 3 '0: default; 1: fit to page; 2: fit to paper; 3: custom
    appState.Add "scalingTypePdf", 3 '0: default; 1: fit to page; 2: fit to paper; 3: custom
    appState.Add "isCssBackgroundEnabled", True
    appState.Add "scaling", 100 'in percent
    
    'initalize margins object in pts (72 pts=1 inch)
    appState.Add "marginsType", 3 'Default=0, None=1, Minimum=2, Custom=3
    
    customMargins.Add "marginTop", Round(0.5 * 72)
    customMargins.Add "marginRight", Round(0.5 * 72)
    customMargins.Add "marginBottom", Round(0.5 * 72)
    customMargins.Add "marginLeft", Round(0.5 * 72)
    
    appState.Add "customMargins", customMargins
    
    'populate paper size options to choose from
    'for this to work, these size properties must match exactly (values and order specified) with chrome preference file in profile
    'C:\Users\[user]\AppData\Local\Google\Chrome\User Data\Default\Preferences
    mediaSizeOptions.Add "A0", jc.ParseJSON("{'height_microns':1189000,'name':'ISO_A0','width_microns':841000,'custom_display_name':'A0'}")
    mediaSizeOptions.Add "A1", jc.ParseJSON("{'height_microns':841000,'name':'ISO_A1','width_microns':594000,'custom_display_name':'A1'}")
    mediaSizeOptions.Add "A2", jc.ParseJSON("{'height_microns':594000,'name':'ISO_A2','width_microns':420000,'custom_display_name':'A2'}")
    mediaSizeOptions.Add "A3", jc.ParseJSON("{'height_microns':420000,'name':'ISO_A3','width_microns':297000,'custom_display_name':'A3'}")
    mediaSizeOptions.Add "A4", jc.ParseJSON("{'height_microns':297000,'name':'ISO_A4','width_microns':210000,'custom_display_name':'A4'}")
    mediaSizeOptions.Add "A5", jc.ParseJSON("{'height_microns':210000,'name':'ISO_A5','width_microns':148000,'custom_display_name':'A5'}")
    mediaSizeOptions.Add "Letter", jc.ParseJSON("{'height_microns':279400,'name':'NA_LETTER','width_microns':215900,'custom_display_name':'Letter'}")
    mediaSizeOptions.Add "Legal", jc.ParseJSON("{'height_microns':355600,'name':'NA_LEGAL','width_microns':215900,'custom_display_name':'Legal'}")
    mediaSizeOptions.Add "Tabloid", jc.ParseJSON("{'height_microns':431800,'name':'NA_LEDGER','width_microns':279400,'custom_display_name':'Tabloid'}")
    
    'add selected paper size defined above to appState object
    appState.Add "mediaSize", mediaSizeOptions("Legal")

    'print the appState object to immediate window for qc
    Debug.Print jc.ConvertToJson(appState, 4)
    
    'this is the tricky part - we need to assign "appState" key of the settings object
    'to a json string - not a dictionary!!! So convert appstate to a string value...
    settings.Add "appState", jc.ConvertToJson(appState)
    
    'finally, set print settings and location to save pdf to
    caps.SetPreference "printing.print_preview_sticky_settings", settings
    caps.SetPreference "savefile.default_directory", ".\"
    
    'this will send webpage to system default printer instead of pdf file
    'caps.SetPreference "printing.use_system_default_printer", True
    
    driver.OpenBrowser caps:=caps
    
    driver.NavigateTo "https://news.google.com"
    
    driver.Wait 1000
    
    'default print file name is based on webpage title
    driver.DeleteFiles ".\" & driver.GetTitle & ".pdf"

    'now print the page
    driver.ExecuteScript ("window.print();")
    
    driver.Wait 7000 'need to wait long enough for print preview to complete!
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
