Attribute VB_Name = "test_ExecuteScript"
Sub test_scroll_ops()
    'see https://www.guru99.com/scroll-up-down-selenium-webdriver.html
    'for more info and tutorials see https://www.guru99.com/selenium-tutorial.html
    Dim Driver As WebDriver, webelem As WebElement
    
    Set Driver = New WebDriver
    Driver.Chrome
    Driver.OpenBrowser

    url = "http://demo.guru99.com/test/guru99home/"
    
    'navigate to url
    'arguments are specified in ParamArray in the order in which they occur in script
    Driver.ExecuteScript "window.location=arguments[0]", url
    
    Driver.Wait 1000
    Driver.MaximizeWindow
    
    Set webelem = Driver.FindElement(by.linkText, "Linux")
    
    'arguments are specified in ParamArray in the order in which they occur in script
    Driver.ExecuteScript "arguments[0].scrollIntoView(arguments[1]);", webelem, True
    
    Driver.Wait 1000
    
    Driver.CloseBrowser
    Driver.Shutdown

End Sub

