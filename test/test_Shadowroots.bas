Attribute VB_Name = "test_Shadowroots"
Option Explicit
Option Private Module

Sub test_shadowroot()
    Dim driver As SeleniumVBA.WebDriver, shadowHost As SeleniumVBA.WebElement
    Dim shadowContent As SeleniumVBA.WebElement, shadowRootelem As SeleniumVBA.WebShadowRoot
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "http://watir.com/examples/shadow_dom.html"
    
    Set shadowHost = driver.FindElement(by.ID, "shadow_host")
    
    Set shadowRootelem = shadowHost.GetShadowRoot()
    
    Set shadowContent = shadowRootelem.FindElement(by.ID, "shadow_content")
    
    Debug.Print shadowContent.GetText  'should return "some text"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_shadowroots_clear_browser_history()
    Dim driver As SeleniumVBA.WebDriver
    Dim webelem1 As SeleniumVBA.WebElement, webelem2 As SeleniumVBA.WebElement, webelem3 As SeleniumVBA.WebElement
    Dim webelem4 As SeleniumVBA.WebElement, webelem5 As SeleniumVBA.WebElement, webelem6 As SeleniumVBA.WebElement
    Dim clearData As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome 'this is a chome-only demo
    driver.OpenBrowser
    
    'make some browsing history
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    driver.NavigateTo "chrome://settings/clearBrowserData/"
    driver.Wait 1000
    
    'work way down the shadowroot tree until we find the clear history button
    Set webelem1 = driver.FindElement(by.cssSelector, "settings-ui")
    Set webelem2 = webelem1.GetShadowRoot.FindElement(by.cssSelector, "settings-main") 'belongs to shadow root under downloads-manager
    Set webelem3 = webelem2.GetShadowRoot.FindElement(by.cssSelector, "settings-basic-page")     'belongs to shadow root under downloads-manager
    Set webelem4 = webelem3.GetShadowRoot.FindElement(by.cssSelector, "settings-section > settings-privacy-page")
    Set webelem5 = webelem4.GetShadowRoot.FindElement(by.cssSelector, "settings-clear-browsing-data-dialog")
    Set webelem6 = webelem5.GetShadowRoot.FindElement(by.cssSelector, "#clearBrowsingDataDialog")
    
    Set clearData = webelem6.FindElement(by.cssSelector, "#clearBrowsingDataConfirm")
    clearData.Click 'to clear browsing history
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

