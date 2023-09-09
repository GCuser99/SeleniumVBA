Attribute VB_Name = "test_IEMode"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

'this test module contains tests that fail in IE Mode
'see following discussion link for limitations:
'https://github.com/GCuser99/SeleniumVBA/discussions/10#discussion-4123927
'also see:
'https://jimevansmusic.blogspot.com/2014/09/screenshots-sendkeys-and-sixty-four.html

Sub test_action_chain()
    Dim driver As SeleniumVBA.WebDriver
    Dim actions As SeleniumVBA.WebActionChain
    Dim from1 As SeleniumVBA.WebElement, to1 As SeleniumVBA.WebElement
    Dim from2 As SeleniumVBA.WebElement, to2 As SeleniumVBA.WebElement
    Dim from3 As SeleniumVBA.WebElement, to3 As SeleniumVBA.WebElement
    Dim from4 As SeleniumVBA.WebElement, to4 As SeleniumVBA.WebElement
    Dim elem As SeleniumVBA.WebElement
    
    'IE mode does not support wheel-type actions so must avoid ScrollBy action
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartIE
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/drag_drop.html"
    
    driver.Wait 1000
    
    Set from1 = driver.FindElement(By.XPath, "//*[@id='credit2']/a")
    Set to1 = driver.FindElement(By.XPath, "//*[@id='bank']/li")
    
    Set from2 = driver.FindElement(By.XPath, "//*[@id='credit1']/a")
    Set to2 = driver.FindElement(By.XPath, "//*[@id='loan']/li")
    
    Set from3 = driver.FindElement(By.XPath, "//*[@id='fourth']/a")
    Set to3 = driver.FindElement(By.XPath, "//*[@id='amt7']/li")
    
    Set from4 = driver.FindElement(By.XPath, "//*[@id='fourth']/a")
    Set to4 = driver.FindElement(By.XPath, "//*[@id='amt8']/li")
    
    driver.Wait 500
    
    'driver.ScrollBy , 700 'in IE mode, must scroll before performing action chain
    
    Set actions = driver.ActionChain
    actions.ScrollBy , 500 'IE mode does not support actionchain scrolls
    actions.DragAndDrop from1, to1
    actions.DragAndDrop from2, to2
    actions.DragAndDrop from3, to3
    'an alternative method to Drag and Drop
    actions.ClickAndHold(from4).MoveToElement(to4).ReleaseButton
    actions.Perform
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_action_chain_sendkeys()
    'This works but must get focus on target element prior to sending keys
    Dim driver As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    Dim actions As SeleniumVBA.WebActionChain
    Dim searchBox As SeleniumVBA.WebElement
    
    Set keys = SeleniumVBA.New_WebKeyboard
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartIE
    
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 500
    
    Set searchBox = driver.FindElement(By.Name, "q")
    
    Set actions = driver.ActionChain
    
    'build the chain and then execute with Perform method
    'must get focus first!!
    'actions.MoveToElement(searchBox).Click 'this is not necessary with other browsers
    
    actions.KeyDown(keys.ShiftKey).SendKeys("upper case").KeyUp (keys.ShiftKey)
    actions.SendKeys(keys.ReturnKey).Perform

    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_shadowroot()
    'IE mode does not support Shadowroots
    Dim driver As SeleniumVBA.WebDriver, shadowHost As SeleniumVBA.WebElement
    Dim shadowContent As SeleniumVBA.WebElement, shadowRootelem As WebShadowRoot
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartIE
    driver.OpenBrowser
    driver.NavigateTo ("http://watir.com/examples/shadow_dom.html")
    
    Set shadowHost = driver.FindElement(By.CssSelector, "#shadow_host")
    
    'this returns "Command not found"
    Set shadowRootelem = shadowHost.GetShadowRoot()
    
    Set shadowContent = shadowRootelem.FindElement(By.ID, "shadow_content")
    
    Debug.Print shadowContent.GetText  'should return "some text"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_cookies()
    'with IE Mode, SetCookies method does not actually set the cookies
    Dim driver As SeleniumVBA.WebDriver, cks As SeleniumVBA.WebCookies
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartIE
    
    Set cks = driver.CreateCookies
    
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    driver.FindElement(By.Name, "username").SendKeys ("abc123")
    driver.FindElement(By.Name, "password").SendKeys ("123xyz")
    driver.FindElement(By.Name, "submit").Click
    
    driver.Wait 500
    
    'get all cookies for this domain and then save to file
    driver.GetAllCookies().SaveToFile ".\cookies.txt"
    
    driver.DeleteAllCookies
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    driver.Wait 1000
    
    'load and set saved cookies from file
    driver.SetCookies cks.LoadFromFile(".\cookies.txt")

    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_Windows()
    'SwitchToWindow works in IE mode but does not bring switched-to windows into foreground
    'see https://learn.microsoft.com/en-us/microsoft-edge/webdriver-chromium/ie-mode?tabs=c-sharp
    'https://titusfortner.com/2022/09/28/edge-ie-mode.html
    Dim driver As SeleniumVBA.WebDriver
    Dim hnd1 As String, hnd2 As String
    Dim i As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartIE
    driver.OpenBrowser

    driver.NavigateTo "https://www.google.com/"
    driver.Wait 500

    driver.Windows.SwitchToNew svbaTab 'this will create a new browser tab

    driver.NavigateTo "https://www.news.google.com/"
    driver.Wait 500
    
    'the following loop works, but does not bring switched-windows into foreground
    For i = 1 To 10
        Debug.Print driver.Windows(1).Activate.Title
        Debug.Print driver.Windows(2).Activate.Title
    Next i
    
    driver.ActiveWindow.CloseIt
    driver.ActiveWindow.CloseIt
    driver.Wait 1000

    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_file_download2()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
   
    driver.StartIE
    
    'IE seems to have no way of directing the download location
    '(defaults to downloads folder) and to specify "don't ask" on download
    Set caps = driver.CreateCapabilities
    caps.SetDownloadPrefs ".\"
    driver.OpenBrowser caps
    
    'delete legacy copy if it exists
    driver.DeleteFiles ".\test.pdf"
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/raw/main/dev/test_files/test.pdf"
    
    driver.WaitForDownload ".\test.pdf"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_element_aria()
    'IE mode does not support Aria attributes
    Dim driver As SeleniumVBA.WebDriver, str As String
    Dim filePath As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    str = "<!DOCTYPE html><html><body><div role='button' class='xyz' aria-label='Add food' aria-disabled='false' data-tooltip='Add food'><span class='abc' aria-hidden='true'>icon</span></body></html>"
    
    filePath = ".\snippet.html"
    
    driver.StartIE
    driver.OpenBrowser
    
    driver.SaveStringToFile str, filePath
    
    driver.NavigateToFile filePath
    
    driver.Wait 1000
    
    'these will throw error
    Debug.Print "Label: " & driver.FindElementByClassName("xyz").GetAriaLabel
    Debug.Print "Role: " & driver.FindElementByClassName("xyz").GetAriaRole
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_MultiSession_IE()
    Dim driver1 As SeleniumVBA.WebDriver
    Dim driver2 As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    Dim keySeq As String
    
    Set driver1 = SeleniumVBA.New_WebDriver
    Set driver2 = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard

    driver1.CommandWindowStyle = vbNormalFocus
    driver2.CommandWindowStyle = vbNormalFocus

    'it fails on same port
    'this seems to work fine if on different ports
    'logging works fine on different ports
    
    'driver1.StartIE , 5555, True, ".\ie1.log"
    'driver2.StartIE , 5556, True, ".\ie2.log"
    
    driver1.StartIE , 5555
    driver2.StartIE , 5556
    
    driver1.OpenBrowser
    driver2.OpenBrowser

    driver1.NavigateTo "http://demo.guru99.com/test/delete_customer.php"
    driver1.Wait 1000
    
    driver2.NavigateTo "https://www.google.com/"
    driver2.Wait 1000
    
    keySeq = "This is COOKL!" & keys.LeftKey & keys.LeftKey & keys.LeftKey & keys.DeleteKey & keys.ReturnKey
    
    driver2.FindElement(By.Name, "q").SendKeys keySeq
    driver2.Wait 1000
    
    Debug.Print "Is Alert Present: " & driver1.IsAlertPresent
                                
    driver1.FindElement(By.Name, "cusid").SendKeys "87654"
    driver1.Wait 1000
    
    driver1.FindElement(By.Name, "submit").Click
    driver1.Wait 1000
    
    Debug.Print "Is Alert Present: " & driver1.IsAlertPresent
    Debug.Print "Alert Text: " & driver1.SwitchToAlert.GetText
    driver1.SwitchToAlert.Accept
    
    Debug.Print "Alert Text: " & driver1.SwitchToAlert.GetText
    driver1.SwitchToAlert.Accept

    driver1.Wait 1000
    driver1.CloseBrowser
    driver2.CloseBrowser
    
    driver1.Shutdown 'shuts down all instances listening to same port
    driver2.Shutdown 'this is needed if different port
End Sub

Sub test_invisible()
    'headless mode does not work for IE mode
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities

    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartIE
    
    'note that WebCapabilities object should be created after starting the driver (StartEdge, StartChrome, of StartFirefox methods)
    Set caps = driver.CreateCapabilities
    
    caps.RunInvisible 'makes browser run in invisible mode
    
    driver.OpenBrowser caps 'here is where caps is passed to driver
    
    driver.NavigateTo "https://www.wikipedia.org/"
    
    Debug.Print "User Agent: " & driver.GetUserAgent

    driver.CloseBrowser
    
    'now let's do it the easy way using optional OpenBrowser parameter...
    driver.OpenBrowser invisible:=True
    
    driver.NavigateTo "https://www.wikipedia.org/"
    
    Debug.Print "User Agent: " & driver.GetUserAgent
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_incognito()
    'in private or incognito mode does not seem to work for IE Mode
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartIE
    
    Set caps = driver.CreateCapabilities
    
    caps.RunIncognito
    
    driver.OpenBrowser caps  'here is where caps is passed to driver
    
    driver.NavigateTo "https://www.wikipedia.org/"
    
    driver.Wait 3000
    
    driver.CloseBrowser
    
    'now let's do it the easy way using optional OpenBrowser parameter...
    driver.OpenBrowser incognito:=True
    
    driver.NavigateTo "https://www.wikipedia.org/"
    
    driver.Wait 3000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_user_profile()
    'set user profile does not seem to work
    'see https://github.com/MicrosoftEdge/EdgeWebDriver/issues/29
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities

    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartIE
    
    Set caps = driver.CreateCapabilities
    
    'this will create and populate a profile if it doesn't yet exist,
    'otherwise will use a previously created profile
    'recommended to customize your Selenium profiles in a different location
    'than the profiles in AppData to avoid conflicts with manual browsing
    'must specify the path to profile, not just the profile name
    caps.SetProfile ".\User Data\IE\profile 1"
    
    driver.OpenBrowser caps
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_detach_browser()
    'only for Chrome/Edge - not supported in IE mode
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartIE
    
    Set caps = driver.CreateCapabilities
    
    'this sets whether browser is closed (false) or left open (true)
    'when the driver is sent the shutdown command before browser is closed
    'defaults to false
    'only applicable to edge/chrome browsers
    caps.SetDetachBrowser True
    
    driver.OpenBrowser caps
    
    driver.NavigateTo "https://www.wikipedia.org/"
    
    driver.Wait 1000
    
    'driver.CloseBrowser 'detach does nothing if browser is closed properly by user
    driver.Shutdown
End Sub

Sub test_print()
    'Print method does not work for IE Mode
    
    Dim driver As SeleniumVBA.WebDriver
    Dim settings As SeleniumVBA.WebPrintSettings
    Dim keys As SeleniumVBA.WebKeyboard

    Set driver = SeleniumVBA.New_WebDriver
    Set settings = SeleniumVBA.New_WebPrintSettings
    Set keys = SeleniumVBA.New_WebKeyboard
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    driver.StartIE
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    
    driver.FindElement(By.ID, "searchInput").SendKeys "Leonardo da Vinci" & keys.EnterKey
    
    driver.Wait 1000
    
    settings.Units = svbaInches
    settings.MarginsAll = 0.4
    settings.Orientation = svbaPortrait
    settings.PrintScale = 1
    'settings.PageRanges "1-2"  'prints the first 2 pages
    'settings.PageRanges 1, 2   'prints the first 2 pages
    'settings.PageRanges 2       'prints only the 2nd page
    
    'prints pdf file to specified filePath parameter (defaults to .\printpage.pdf)
    driver.PrintToPDF , settings

    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_screenshot()
    'SaveScreenShot does not work properly in IE mode 32-bit
    Dim driver As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    Dim caps As SeleniumVBA.WebCapabilities
    Dim params As New Dictionary
    
    Set driver = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartIE
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    
    driver.SaveScreenshot

    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_GetSessionInfo()
    'this is provided to see a list of default capabilities for IE mode
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartIE
    
    driver.OpenBrowser
    
    Debug.Print SeleniumVBA.WebJsonConverter.ConvertToJson(driver.GetSessionsInfo, 4)
    
    driver.Wait 1000
    driver.CloseBrowser
    
    driver.Shutdown
End Sub

Sub test_geolocation()
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartIE
    driver.OpenBrowser
    
    'IE mode does not support geolocation commands
    
    driver.SetGeolocation 41.1621429, -8.6219537
  
    driver.NavigateTo "https://www.gps-coordinates.net/my-location"
    driver.Wait 1000
    
    'print the name of the location
    Debug.Print driver.FindElementByXPath("//*[@id='addr']").GetText
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
