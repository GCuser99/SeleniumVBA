Attribute VB_Name = "test_scroll"
Sub test_scroll_ops()
    'see https://www.guru99.com/scroll-up-down-selenium-webdriver.html
    'for more info and tutorials see https://www.guru99.com/selenium-tutorial.html
    Dim Driver As WebDriver, webelem As WebElement
    
    Set Driver = New WebDriver
    Driver.Chrome
    Driver.OpenBrowser

    Driver.Navigate "http://demo.guru99.com/test/guru99home/"
    Driver.MaximizeWindow
    
    'scroll down in increments of 50 pixels
    For i = 1 To 40
        Driver.ScrollBy , 50
        Driver.Wait 25
    Next i
    
    'scroll to top of window
    Driver.ScrollToTop
    Driver.Wait 1000
    
    'scroll down half-way from top to bottom of window
    Driver.ScrollTo 0, Driver.GetScrollHeight / 2
    Driver.Wait 1000
    
    'this uses the WebElement ScrollIntoView method to scroll vertically to a WebElement
    Driver.FindElement(by.linkText, "Linux").ScrollIntoView
    
    Driver.Wait 1000
    
    'This will scroll the web page to end.
    Driver.ScrollToBottom
    Driver.Wait 1000
    
    Driver.Navigate "http://demo.guru99.com/test/guru99home/scrolling.html"
    Driver.Wait 1000
    
    'this uses the WebElement ScrollIntoView method to scroll horizontally to a WebElement
    Driver.FindElement(by.linkText, "VBScript").ScrollIntoView
    Driver.Wait 1000
    
    Driver.CloseBrowser
    Driver.Shutdown

End Sub
