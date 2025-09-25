Attribute VB_Name = "test_Shadowroots"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_shadowroot()
    Dim driver As SeleniumVBA.WebDriver
    Dim shadowHost As SeleniumVBA.WebElement
    Dim shadowContent As SeleniumVBA.WebElement
    Dim shadowRootelem As SeleniumVBA.WebShadowRoot
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/shadowRootPage.html"
    
    'inspect the element you want to access with "copy full xpath":
    '/html/body/div[2]/custom-checkbox-element//div/input
    'note the double slash in the xpath above - that indicates a shadow root
    'first find the host element (in this case custom-checkbox-element)
    Set shadowHost = driver.FindElement(By.CssSelector, "body > div:nth-child(3) > custom-checkbox-element")
    
    'then return the shadow root from the host element
    Set shadowRootelem = shadowHost.GetShadowRoot()
    
    'now we can use find methods and other DOM operations as usual
    Set shadowContent = shadowRootelem.FindElement(By.CssSelector, "div > input[type=checkbox]")
    
    shadowContent.Click
    
    Debug.Assert shadowContent.GetProperty("checked")
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
