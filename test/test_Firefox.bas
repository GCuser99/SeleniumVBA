Attribute VB_Name = "test_Firefox"
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
'- Wheel-type actions in action chains not supported (no action scrolls)
'- Aria methods not supported (GetAriaLabel & GetAriaRole)
'- Shutdown Method not recognized (currrently using taskkill to shutdown)
'- Multi-sessions not supported
'- GetSessionsInfo not functional
'- Shadowroots are only partially supported (CMD_GET_ELEMENT_SHADOW_ROOT works but
'  CMD_FIND_ELEMENT_FROM_SHADOW_ROOT & CMD_FIND_ELEMENTS_FROM_SHADOW_ROOT do not).
'  Apparently this support may be coming: see https://github.com/mozilla/geckodriver/issues/2005
'- PrintScale method of PrintSettings class does not seem to have effect
'
'
'Any suggested ideas, comments, fixes, and use cases are welcome!

Sub test_logging()
    Dim driver As New WebDriver, fruits As WebElement
    
    'driver.CommandWindowStyle = vbNormalFocus
    
    'True enables verbose logging - default log file found in same directory as WebDriver executable
    driver.StartFirefox , , True
    
    driver.OpenBrowser

    driver.NavigateTo "https://jsbin.com/osebed/2"
    driver.Wait
    
    Set fruits = driver.FindElement(by.ID, "fruits")
    
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
        Debug.Print fruits.GetSelectedOptionText
    End If
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_file_download2()
    Dim driver As New WebDriver, caps As WebCapabilities
    
    driver.DefaultIOFolder = ThisWorkbook.Path
    
    driver.StartFirefox
    
    driver.DeleteFiles ".\BrowserStack - List of devices to test*"
    
    Set caps = driver.CreateCapabilities
    
    caps.SetDownloadPrefs
    
    Debug.Print caps.ToJson

    driver.OpenBrowser caps
    
    driver.NavigateTo "https://www.browserstack.com/test-on-the-right-mobile-devices"
    driver.Wait 500
    
    driver.FindElementByID("accept-cookie-notification").Click
    driver.Wait 500
    
    driver.FindElementByCssSelector(".icon-csv").ScrollToElement , -150
    driver.Wait 1000
    
    driver.FindElementByCssSelector(".icon-csv").Click
    'driver.FindElementByCssSelector(".icon-pdf").Click
    
    driver.Wait 4000
            
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_action_chain()
    'wheel-type actions are not allowed in Firefox, so must remove ScrollBy action from chain
    'otherwise action chains work fine
    Dim driver As New WebDriver, actions As WebActionChain
    Dim from1 As WebElement, to1 As WebElement
    Dim from2 As WebElement, to2 As WebElement
    Dim from3 As WebElement, to3 As WebElement
    Dim from4 As WebElement, to4 As WebElement
    Dim elem As WebElement
    
    driver.StartFirefox
    
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/drag_drop.html"
    
    Set from1 = driver.FindElement(by.XPath, "//*[@id='credit2']/a")
    Set to1 = driver.FindElement(by.XPath, "//*[@id='bank']/li")
    
    Set from2 = driver.FindElement(by.XPath, "//*[@id='credit1']/a")
    Set to2 = driver.FindElement(by.XPath, "//*[@id='loan']/li")
    
    Set from3 = driver.FindElement(by.XPath, "//*[@id='fourth']/a")
    Set to3 = driver.FindElement(by.XPath, "//*[@id='amt7']/li")
    
    Set from4 = driver.FindElement(by.XPath, "//*[@id='fourth']/a")
    Set to4 = driver.FindElement(by.XPath, "//*[@id='amt8']/li")
    
    driver.Wait 1000
    
    driver.ScrollBy , 500
    
    Set actions = driver.ActionChain
    'scroll actions are not accepted by Firefox
    'actions.ScrollBy , 500
    actions.DragAndDrop(from1, to1).Wait
    actions.DragAndDrop(from2, to2).Wait
    actions.DragAndDrop(from3, to3).Wait
    'an alternative method to Drag and Drop
    actions.ClickAndHold(from4).MoveToElement(to4).ReleaseButton.Wait (1000)
    actions.Perform 'do all the actions defined above
    
    driver.Wait 1000
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_element_aria()
    'Firefox does not support Aria attributes
    Dim driver As New WebDriver, str As String
    
    str = "<!DOCTYPE html><html><body><div role='button' class='xyz' aria-label='Add food' aria-disabled='false' data-tooltip='Add food'><span class='abc' aria-hidden='true'>icon</span></body></html>"
    
    filePath = ".\snippet.html"
    
    driver.StartFirefox
    driver.OpenBrowser
    
    driver.SaveHTMLToFile str, filePath
    
    driver.NavigateTo "file:///" & filePath
    
    driver.Wait 1000
    
    Debug.Print "Label: " & driver.FindElementByClassName("xyz").GetAriaLabel
    Debug.Print "Role: " & driver.FindElementByClassName("xyz").GetAriaRole
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_shadowroot()
    'Firefox has partial support for Shadowroots
    Dim driver As New WebDriver, shadowHost As WebElement
    Dim shadowContent As WebElement, shadowRootelem As WebShadowRoot
    
    driver.StartFirefox
    driver.OpenBrowser
    driver.NavigateTo ("http://watir.com/examples/shadow_dom.html")
    
    Set shadowHost = driver.FindElement(by.ID, "shadow_host")
    
    'this works for Firefox
    Set shadowRootelem = shadowHost.GetShadowRoot()
    
    Debug.Print "got shadowroot element ok"
    
    'the following returns "HTTP method not allowed" for Firefox
    'apparently FindElement support for shadowroots may be coming:
    'https://github.com/mozilla/geckodriver/issues/2005
    Set shadowContent = shadowRootelem.FindElement(by.ID, "shadow_content")
    
    Debug.Print shadowContent.GetText 'should return "some text"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_Alerts()
    'see https://www.guru99.com/alert-popup-handling-selenium.html
    Dim driver As New WebDriver
    
    driver.StartFirefox
    driver.OpenBrowser

    driver.NavigateTo "http://demo.guru99.com/test/delete_customer.php"
    
    driver.Wait 1000
    
    Debug.Print "Is Alert Present: " & driver.IsAlertPresent
                                
    driver.FindElement(by.Name, "cusid").SendKeys "87654"
    
    driver.Wait 1000
    
    driver.FindElement(by.Name, "submit").Click
    
    driver.Wait 1000
    
    Debug.Print "Is Alert Present: " & driver.IsAlertPresent
    Debug.Print "Alert Text: " & driver.GetAlertText
    driver.AcceptAlert
    
    driver.Wait 'Firefox needs a nominal wait here - chrome and edge will fail with the wait
    
    Debug.Print "Alert Text: " & driver.GetAlertText
    driver.AcceptAlert

    driver.Wait 1000
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_GetSessionInfo()
    Dim driver As New WebDriver
    Dim col As Collection
    
    driver.StartFirefox
    
    driver.OpenBrowser
    
    'firefox does not support "Get All Sessions" command
    
    Set col = driver.GetSessionsInfo
    
    driver.Wait 1000
    driver.CloseBrowser
    
    driver.Shutdown
End Sub

