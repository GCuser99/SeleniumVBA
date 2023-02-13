Attribute VB_Name = "test_Shadowroots"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_shadowroot()
    Dim driver As SeleniumVBA.WebDriver, shadowHost As SeleniumVBA.WebElement
    Dim shadowContent As SeleniumVBA.WebElement, shadowRootelem As SeleniumVBA.WebShadowRoot
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "http://watir.com/examples/shadow_dom.html"
    
    Set shadowHost = driver.FindElement(By.ID, "shadow_host")
    
    Set shadowRootelem = shadowHost.GetShadowRoot()
    
    Set shadowContent = shadowRootelem.FindElement(By.ID, "shadow_content")
    
    Debug.Print shadowContent.GetText  'should return "some text"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_shadowroots_clear_browser_history()
    Dim driver As SeleniumVBA.WebDriver
    Dim webElem1 As SeleniumVBA.WebElement, webElem2 As SeleniumVBA.WebElement, webElem3 As SeleniumVBA.WebElement
    Dim webElem4 As SeleniumVBA.WebElement, webElem5 As SeleniumVBA.WebElement, webElem6 As SeleniumVBA.WebElement
    Dim clearData As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome 'this is a chrome-only demo
    driver.OpenBrowser
    
    'make some browsing history
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    driver.NavigateTo "chrome://settings/clearBrowserData/"
    driver.Wait 1000
    
    'work way down the shadowroot tree until we find the clear history button
    Set webElem1 = driver.FindElement(By.cssSelector, "settings-ui")
    Set webElem2 = webElem1.GetShadowRoot.FindElement(By.cssSelector, "settings-main") 'belongs to shadow root under downloads-manager
    Set webElem3 = webElem2.GetShadowRoot.FindElement(By.cssSelector, "settings-basic-page")     'belongs to shadow root under downloads-manager
    Set webElem4 = webElem3.GetShadowRoot.FindElement(By.cssSelector, "settings-section > settings-privacy-page")
    Set webElem5 = webElem4.GetShadowRoot.FindElement(By.cssSelector, "settings-clear-browsing-data-dialog")
    Set webElem6 = webElem5.GetShadowRoot.FindElement(By.cssSelector, "#clearBrowsingDataDialog")
    
    Set clearData = webElem6.FindElement(By.cssSelector, "#clearBrowsingDataConfirm")
    clearData.Click 'to clear browsing history
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

