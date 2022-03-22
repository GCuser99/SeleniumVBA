Attribute VB_Name = "test_ExecuteScript"
Sub test_executescript()
    Dim Driver As New WebDriver, webelem As WebElement
    
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

Sub test_executescriptasync()
    'see https://www.lambdatest.com/blog/how-to-use-javascriptexecutor-in-selenium-webdriver/
    Dim Driver As New WebDriver, webelem As WebElement, jc As New JSonConverter
    
    Driver.Chrome
    Driver.OpenBrowser
    
    url = "https://www.google.com/"

    waitTime = 3000
    
    Driver.SetScriptTimeout 2 * waitTime '30000 is the default, so this isn't needed unless waitTime > 30 secs is needed
    
    Driver.Navigate url
        
    'Driver.ExecuteScriptAsync "window.setTimeout(arguments[arguments.length - 1], arguments[0]);", waitTime
    'Driver.ExecuteScriptAsync "window.setTimeout(arguments[1], arguments[0]);", waitTime 'this is equivalent
    
    'here the callback sends an alert "wait is over!" after the desired waitTime
    Driver.ExecuteScriptAsync "var callback = arguments[arguments.length - 1]; setTimeout(function(){callback(alert('wait is over!'))}, arguments[0]);", waitTime
    
    Driver.Wait 1000
        
    Driver.CloseBrowser
    Driver.Shutdown

End Sub
