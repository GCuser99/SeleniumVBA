Attribute VB_Name = "test_Scroll"
Sub test_scroll_ops()
    'see https://www.guru99.com/scroll-up-down-selenium-webdriver.html
    'for more info and tutorials see https://www.guru99.com/selenium-tutorial.html
    Dim driver As WebDriver, webelem As WebElement
    
    Set driver = New WebDriver
    driver.StartChrome
    driver.OpenBrowser

    driver.NavigateTo "http://demo.guru99.com/test/guru99home/"
    driver.MaximizeWindow
    
    'scroll down in increments of 50 pixels
    For i = 1 To 40
        driver.ScrollBy , 50
        driver.Wait 25
    Next i
    
    'scroll to top of window
    driver.ScrollToTop
    driver.Wait 1000
    
    'scroll down half-way from top to bottom of window
    driver.ScrollTo 0, driver.GetScrollHeight / 2
    driver.Wait 1000
    
    'this uses the WebElement ScrollIntoView method to scroll vertically to a WebElement
    driver.FindElement(by.linkText, "Linux").ScrollIntoView
    
    driver.Wait 1000
    
    'This will scroll the web page to end.
    driver.ScrollToBottom
    driver.Wait 1000
    
    driver.NavigateTo "http://demo.guru99.com/test/guru99home/scrolling.html"
    driver.Wait 1000
    
    'this uses the WebElement ScrollIntoView method to scroll horizontally to a WebElement
    driver.FindElement(by.linkText, "VBScript").ScrollIntoView
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown

End Sub
