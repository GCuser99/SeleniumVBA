Attribute VB_Name = "test_shadowroots"

Sub test_shadow_roots_clear_browser_history()

    Dim Driver As New WebDriver
    Dim webelem1 As WebElement, webelem2 As WebElement, webelem3 As WebElement
    Dim webelem4 As WebElement, webelem5 As WebElement, webelem6 As WebElement
    Dim clearData As WebElement
    'Dim jc As New JSonConverter
        
    Driver.Chrome
    Driver.OpenBrowser
    
    'make some browsing history
    Driver.Navigate "https://www.google.com/"
    Driver.Wait 1000
    Driver.Navigate "chrome://settings/clearBrowserData/"
    Driver.Wait 1000
    
    'work way down the shadowroot tree until we find the clear history button
    Set webelem1 = Driver.FindElement(by.CssSelector, "settings-ui")
    Set webelem2 = webelem1.GetShadowRoot.FindElement(by.CssSelector, "settings-main") 'belongs to shadow root under downloads-manager
    Set webelem3 = webelem2.GetShadowRoot.FindElement(by.CssSelector, "settings-basic-page")     'belongs to shadow root under downloads-manager
    Set webelem4 = webelem3.GetShadowRoot.FindElement(by.CssSelector, "settings-section > settings-privacy-page")
    Set webelem5 = webelem4.GetShadowRoot.FindElement(by.CssSelector, "settings-clear-browsing-data-dialog")
    Set webelem6 = webelem5.GetShadowRoot.FindElement(by.CssSelector, "#clearBrowsingDataDialog")
    
    Set clearData = webelem6.FindElement(by.CssSelector, "#clearBrowsingDataConfirm")
    clearData.Click 'to clear browsing history
    
    Driver.Wait 3000
    
    Driver.CloseBrowser
    Driver.Shutdown

End Sub




Sub test_shadow_roots()
'"chrome://downloads/"

    Dim Driver As New WebDriver
    Dim webelem As WebElement, webelem2 As WebElement, webelem3 As WebElement, webelem4 As WebElement, webelem5 As WebElement, webelem6 As WebElement, clearData As WebElement
    'Dim jc As New JSonConverter
    Dim sr1 As ShadowRoot, sr2 As ShadowRoot, sr3  As ShadowRoot, sr4  As ShadowRoot, sr5  As ShadowRoot, sr6  As ShadowRoot
    'see https://chromedriver.chromium.org/logging
    
    'serviceArgs = "--verbose --log-path=" & Chr(34) & "C:\Users\waite\Documents\finance\financial model\mike's code\scraping\driver_verbose.log" & Chr(34)
    
    Driver.Chrome , , True
    
    Driver.OpenBrowser

    Driver.Navigate "chrome://settings/clearBrowserData/"
    
    Driver.Wait 1000
    'Set webelem = Driver.FindElement(by.TagName, "downloads-manager")
    
    'Debug.Print Driver.ExecuteScript("return arguments[0].shadowRoot.innerHtml", , webelem) 'returns Null
    
    Set sr1 = Driver.FindElement(by.CssSelector, "settings-ui").GetShadowRoot
    
    Set webelem2 = sr1.FindElement(by.CssSelector, "settings-main") 'belongs to shadow root under downloads-manager
    
    Set sr2 = Driver.GetShadowRoot(webelem2)
    
    Set webelem3 = sr2.FindElement(by.CssSelector, "settings-basic-page")     'belongs to shadow root under downloads-manager
    
    Set sr3 = Driver.GetShadowRoot(webelem3)
    
    Set webelem4 = sr3.FindElement(by.CssSelector, "settings-section > settings-privacy-page")
    
    Set sr4 = Driver.GetShadowRoot(webelem4)
    
    Set webelem5 = sr4.FindElement(by.CssSelector, "settings-clear-browsing-data-dialog")
    
    Set sr5 = Driver.GetShadowRoot(webelem5)
    
    Set webelem6 = sr5.FindElement(by.CssSelector, "#clearBrowsingDataDialog")
    
    Set clearData = sr5.FindElement(by.CssSelector, "#clearBrowsingDataConfirm")
    clearData.Click
    
    Driver.CloseBrowser
    Driver.Shutdown

End Sub
