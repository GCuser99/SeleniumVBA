Attribute VB_Name = "test_ExecuteScript"
Sub test_executescript()
    Dim Driver As New WebDriver, webElem As WebElement
    
    Driver.StartChrome
    Driver.OpenBrowser

    url = "http://demo.guru99.com/test/guru99home/"
    
    'Navigate to url
    'arguments are specified in ParamArray list where first parameter value is associated
    'with arguments[0], second parameter value is associated with arguments[1], etc
    Driver.ExecuteScript "window.location=arguments[0]", url
    
    Driver.Wait 1000
    Driver.MaximizeWindow
    
    'ExecuteScript returns a WebElement object if script results in a WebElement object
    Set webElem = Driver.ExecuteScript("return document.getElementById('philadelphia-field-submit')")
    
    'arguments are specified in ParamArray list where first parameter value is associated
    'with arguments[0], second parameter value is associated with arguments[1], etc
    Driver.ExecuteScript "arguments[0].scrollIntoView(arguments[1]);", webElem, True
    
    Driver.Wait 1000
    
    'ExecuteScript returns a single WebElements object if script results in a collection of WebElement objects
    Dim divElems As WebElements
    Set divElems = Driver.ExecuteScript("return document.getElementsByTagName(arguments[0])", "div")
    Debug.Print "Number of div elements: " & divElems.Count
    
    Driver.CloseBrowser
    Driver.Shutdown

End Sub

Sub test_executescriptasync()
    'see https://www.lambdatest.com/blog/how-to-use-javascriptexecutor-in-selenium-webdriver/
    Dim Driver As New WebDriver, webElem As WebElement, jc As New JSonConverter
    
    Driver.StartChrome
    Driver.OpenBrowser
    
    url = "https://www.google.com/"

    waitTime = 3000
    
    If waitTime > 30000 Then Driver.SetScriptTimeout 2 * waitTime '30000 is the default, so this isn't needed unless waitTime > 30 secs is needed
    
    Driver.NavigateTo url
        
    'Driver.ExecuteScriptAsync "window.setTimeout(arguments[arguments.length - 1], arguments[0]);", waitTime
    'Driver.ExecuteScriptAsync "window.setTimeout(arguments[1], arguments[0]);", waitTime 'this is equivalent
    
    'here the callback sends an alert "wait is over!" after the desired waitTime
    Driver.ExecuteScriptAsync "var callback = arguments[arguments.length - 1]; setTimeout(function(){callback(alert('WAIT IS OVER!'))}, arguments[0]);", waitTime
    Driver.Wait 2000
    
    Driver.AcceptAlert
    Driver.Wait 1000
        
    Driver.CloseBrowser
    Driver.Shutdown

End Sub

