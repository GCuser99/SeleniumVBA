Attribute VB_Name = "test_ExecuteScript"
Sub test_executescript()
    Dim driver As New WebDriver, webelem As WebElement
    
    driver.StartChrome
    driver.OpenBrowser

    url = "http://demo.guru99.com/test/guru99home/"
    
    'navigate to url
    'arguments are specified in ParamArray in the order in which they occur in script
    driver.ExecuteScript "window.location=arguments[0]", url
    
    driver.Wait 1000
    driver.MaximizeWindow
    
    Set webelem = driver.FindElement(by.linkText, "Linux")
    
    'arguments are specified in ParamArray in the order in which they occur in script
    driver.ExecuteScript "arguments[0].scrollIntoView(arguments[1]);", webelem, True
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown

End Sub

Sub test_executescriptasync()
    'see https://www.lambdatest.com/blog/how-to-use-javascriptexecutor-in-selenium-webdriver/
    Dim driver As New WebDriver, webelem As WebElement, jc As New JSonConverter
    
    driver.StartChrome
    driver.OpenBrowser
    
    url = "https://www.google.com/"

    waitTime = 3000
    
    If waitTime > 30000 Then driver.SetScriptTimeout 2 * waitTime '30000 is the default, so this isn't needed unless waitTime > 30 secs is needed
    
    driver.Navigate url
        
    'Driver.ExecuteScriptAsync "window.setTimeout(arguments[arguments.length - 1], arguments[0]);", waitTime
    'Driver.ExecuteScriptAsync "window.setTimeout(arguments[1], arguments[0]);", waitTime 'this is equivalent
    
    'here the callback sends an alert "wait is over!" after the desired waitTime
    driver.ExecuteScriptAsync "var callback = arguments[arguments.length - 1]; setTimeout(function(){callback(alert('wait is over!'))}, arguments[0]);", waitTime
    
    driver.Wait 1000
        
    driver.CloseBrowser
    driver.Shutdown

End Sub
