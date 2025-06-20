Attribute VB_Name = "test_Alerts"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_Alerts()
    'NOTES: As of July 2023, Chrome and Edge have a ("WontFix") reported bug where sending text to
    'a prompt alert via SwitchToAlert.SendKeys does not display in the text input
    'field but otherwise does work as shown in this demo. It has been classified as a "display-only issue".
    'see https://bugs.chromium.org/p/chromedriver/issues/detail?id=1120#c11
    
    'Also be aware - the only WebDriver commands that should be executed between the show Alert event
    '(e.g. after Click) and SwitchToAlert.Accept/Dismiss are Wait, SwitchToAlert.GetAlertText, and
    'SwitchToAlert.SendKeys - other commands executed in the time interval while waiting for user
    'response could interfere with Alert interaction.
    '
    'The SwitchToAlert waits until the alert shows, up to a maximum time specified by the maxWaitTimeMS
    'argument (default 10000 ms). See slow alert test in this procedure below for an example.
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 10000
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/alerts.html"
    
    'standard alert 1
    driver.FindElement(By.ID, "alert").Click
    With driver.SwitchToAlert
        Debug.Assert .GetText = "cheese"
        .Accept
    End With
    
    'standard alert 2
    driver.FindElement(By.ID, "empty-alert").Click
    With driver.SwitchToAlert
        Debug.Assert .GetText = ""
        .Accept
    End With
    
    'input prompt alert 3
    driver.FindElement(By.ID, "prompt").Click
    With driver.SwitchToAlert
        Debug.Assert .GetText = "Enter something"
        .SendKeys "here is my response text to prompt"
        .Accept
    End With
    Debug.Assert driver.FindElement(By.ID, "text").GetText = "here is my response text to prompt"
    
    'input prompt alert 4
    driver.FindElement(By.ID, "prompt-with-default").Click
    With driver.SwitchToAlert
        Debug.Assert .GetText = "Enter something"
        .SendKeys "here is my response text to prompt with default"
        .Accept
    End With
    Debug.Assert driver.FindElement(By.ID, "text").GetText = "here is my response text to prompt with default"
    
    'input double prompt alerts 5 and 6
    driver.FindElement(By.ID, "double-prompt").Click
    With driver.SwitchToAlert
        Debug.Assert .GetText = "First"
        .SendKeys "here is my response text to first of double prompt"
        .Accept
    End With
    With driver.SwitchToAlert
        Debug.Assert .GetText = "Second"
        .SendKeys "here is my response text to second of double prompt"
        .Accept
    End With
    'note that this first GetText must be performed after the second alert above
    'so that it does not interfere with that alert!!
    Debug.Assert driver.FindElement(By.ID, "text1").GetText = "here is my response text to first of double prompt"
    Debug.Assert driver.FindElement(By.ID, "text2").GetText = "here is my response text to second of double prompt"
    
    'test for a delayed alert 7
    'without the non-zero max wait, this will throw an error
    driver.FindElement(By.ID, "slow-alert").Click
    With driver.SwitchToAlert
        Debug.Assert .GetText = "Slow"
        .Accept
    End With
    
    'a confirm alert 8
    driver.FindElement(By.ID, "confirm").Click
    With driver.SwitchToAlert
        Debug.Assert .GetText = "Are you sure?"
        .Dismiss
    End With
    
    driver.Wait 1000
    driver.GoBack
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_Alerts2()
    Dim driver As SeleniumVBA.WebDriver
    Dim html As String
    
    Set driver = SeleniumVBA.New_WebDriver
        
    driver.StartChrome
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 2000
    
    html = "<html lang=""en"">" & vbCrLf
    html = html & "    <head>" & vbCrLf
    html = html & "        <title>Test Alerts</title>" & vbCrLf
    html = html & "    </head>" & vbCrLf
    html = html & "    <body>" & vbCrLf
    html = html & "        <div><div></div></div>" & vbCrLf
    html = html & "        <div>" & vbCrLf
    html = html & "            <div id=""content"">" & vbCrLf
    html = html & "                <script>" & vbCrLf
    html = html & "                    function jsAlert() {" & vbCrLf
    html = html & "                        alert('I am a JS Alert');" & vbCrLf
    html = html & "                        log('You successfully clicked an alert');" & vbCrLf
    html = html & "                    }" & vbCrLf
    html = html & "                    function jsConfirm() {" & vbCrLf
    html = html & "                        var c = confirm('I am a JS Confirm');" & vbCrLf
    html = html & "                        var result = c === true ? 'Ok' : 'Cancel';" & vbCrLf
    html = html & "                        log('You clicked: ' + result);" & vbCrLf
    html = html & "                    }" & vbCrLf
    html = html & "                    function jsPrompt() {" & vbCrLf
    html = html & "                        var p = prompt('I am a JS prompt');" & vbCrLf
    html = html & "                        log('You entered: ' + p);" & vbCrLf
    html = html & "                    }" & vbCrLf
    html = html & "                    function log(msg) {" & vbCrLf
    html = html & "                        var result = document.getElementById('result');" & vbCrLf
    html = html & "                        result.innerHTML = msg;" & vbCrLf
    html = html & "                    }" & vbCrLf
    html = html & "                </script>" & vbCrLf
    html = html & "                <div class=""example"">" & vbCrLf
    html = html & "                    <h3>JavaScript Alerts</h3>" & vbCrLf
    html = html & "                    <p>Here are some examples of different JavaScript alerts which can be troublesome for automation</p>" & vbCrLf
    html = html & "                    <ul style=""list-style-type: none;"">" & vbCrLf
    html = html & "                        <li><button onclick=""jsAlert()"">Click for JS Alert</button></li>" & vbCrLf
    html = html & "                        <li><button onclick=""jsConfirm()"">Click for JS Confirm</button></li>" & vbCrLf
    html = html & "                        <li><button onclick=""jsPrompt()"">Click for JS Prompt</button></li>" & vbCrLf
    html = html & "                    </ul>" & vbCrLf
    html = html & "                    <h4>Result:</h4>" & vbCrLf
    html = html & "                    <p id=""result"" style=""color:green"">You entered:</p>" & vbCrLf
    html = html & "                </div>" & vbCrLf
    html = html & "            </div>" & vbCrLf
    html = html & "        </div>" & vbCrLf
    html = html & "    </body>" & vbCrLf
    html = html & "</html>"


    driver.NavigateToString html
    
    'find and then click on an element that throws a prompt-type alert
    driver.FindElement(By.XPath, "//*[@id='content']/div/ul/li[3]/button").Click
        
    'SwitchToAlert waits up to a user-specified max time (default = 10 secs)
    'for alert to show, and then returns a WebAlert object for interaction
    driver.SwitchToAlert.SendKeys("hola mi nombre es Jose").Accept
    
    Debug.Assert driver.FindElementByID("result").GetText = "You entered: hola mi nombre es Jose"
        
    driver.CloseBrowser
    driver.Shutdown
End Sub

