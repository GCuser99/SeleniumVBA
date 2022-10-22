Attribute VB_Name = "test_Firefox"
Option Explicit
Option Private Module

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
'- Aria methods not supported (GetAriaLabel & GetAriaRole)
'- Shutdown Method not recognized (currrently using taskkill to shutdown)
'- Multi-sessions not supported
'- GetSessionsInfo not functional
'- Shadowroots are only partially supported (CMD_GET_ELEMENT_SHADOW_ROOT works but
'  CMD_FIND_ELEMENT_FROM_SHADOW_ROOT & CMD_FIND_ELEMENTS_FROM_SHADOW_ROOT do not).
'  Apparently this support may be coming: see https://github.com/mozilla/geckodriver/issues/2005
'- PrintScale method of PrintSettings class does not seem to have effect
'
Sub test_logging()
    Dim driver As SeleniumVBA.WebDriver, fruits As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.CommandWindowStyle = vbNormalFocus
    
    'True enables verbose logging
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

Sub test_file_download()
    Dim driver As SeleniumVBA.WebDriver, caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
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

Sub test_element_aria()
    'Firefox does not support Aria attributes
    Dim driver As SeleniumVBA.WebDriver, str As String
    Dim filePath As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    str = "<!DOCTYPE html><html><body><div role='button' class='xyz' aria-label='Add food' aria-disabled='false' data-tooltip='Add food'><span class='abc' aria-hidden='true'>icon</span></body></html>"
    
    filePath = ".\snippet.html"
    
    driver.StartFirefox
    driver.OpenBrowser
    
    driver.SaveStringToFile str, filePath
    
    driver.NavigateToFile filePath
    
    driver.Wait 1000
    
    Debug.Print "Label: " & driver.FindElementByClassName("xyz").GetAriaLabel
    Debug.Print "Role: " & driver.FindElementByClassName("xyz").GetAriaRole
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_shadowroot()
    'Firefox has partial support for Shadowroots
    Dim driver As SeleniumVBA.WebDriver, shadowHost As SeleniumVBA.WebElement
    Dim shadowContent As SeleniumVBA.WebElement, shadowRootelem As WebShadowRoot
    
    Set driver = SeleniumVBA.New_WebDriver
    
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
    'https://wpt.fyi/results/webdriver/tests?label=experimental&label=master&aligned&view=subtest
    Set shadowContent = shadowRootelem.FindElement(by.ID, "shadow_content")
    
    Debug.Print shadowContent.GetText 'should return "some text"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_Alerts()
    'see https://www.guru99.com/alert-popup-handling-selenium.html
    Dim driver As SeleniumVBA.WebDriver

    Set driver = SeleniumVBA.New_WebDriver
    
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
    Dim driver As SeleniumVBA.WebDriver
    Dim col As Collection
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartFirefox
    
    driver.OpenBrowser
    
    'firefox does not support "Get All Sessions" command
    
    Set col = driver.GetSessionsInfo
    
    driver.Wait 1000
    driver.CloseBrowser
    
    driver.Shutdown
End Sub

