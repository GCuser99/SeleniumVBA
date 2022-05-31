Attribute VB_Name = "test_Shadowroots"

Sub test_shadow_roots_clear_browser_history()

    Dim driver As New WebDriver
    Dim webelem1 As WebElement, webelem2 As WebElement, webelem3 As WebElement
    Dim webelem4 As WebElement, webelem5 As WebElement, webelem6 As WebElement
    Dim clearData As WebElement
        
    driver.StartChrome 'this is a chome-only demo
    driver.OpenBrowser
    
    'make some browsing history
    driver.NavigateTo "https://www.google.com/"
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

