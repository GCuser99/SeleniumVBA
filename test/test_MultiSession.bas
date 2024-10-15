Attribute VB_Name = "test_MultiSession"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

'Chrome and Edge have multi-session support. Firefox does not.
'
'For Edge/Chrome, user must create multiple driver instances, as the WebDriver
'class is currently designed to only allow one session per driver instance.
'
'For Edge/Chrome, there doesn't seem to be a big problem with multiple driver instances sharing the same port.
'If on same port, only one command window is generated and one shutdown shuts down all instances of the drivers.
'
'If drivers are assigned to different ports, then multiple command windows are generated and must shutdown each
'separately.
'
'Some features may not work as expected in multi-session mode on same port, such as logging,
'which generates a single log file per port.
'
'The surest way to get multi-session working for Edge and Chrome without any interference is to start
'the drivers on different ports.
'
'Firefox multi-session will not function unless drivers are assigned to different ports.
'
Sub test_MultiSession_Edge()
    Dim driver1 As SeleniumVBA.WebDriver
    Dim driver2 As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    Dim keySeq As String
    
    Set driver1 = SeleniumVBA.New_WebDriver
    Set driver2 = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    
    'driver1.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    'driver2.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    'driver1.CommandWindowStyle = vbNormalFocus
    'driver2.CommandWindowStyle = vbNormalFocus
    
    'this should work  with only limited interference
    'however, for logging we only get one log (the first one)
    'driver1.StartEdge , , True, ".\edge1.log"
    'driver2.StartEdge , , True, ".\edge2.log"
    
    'this will work with no interferrence
    driver1.StartEdge , 9515, True, ".\edge1.log"
    driver2.StartEdge , 9516, True, ".\edge2.log"
    
    driver1.OpenBrowser
    driver2.OpenBrowser

    driver1.NavigateTo "http://demo.guru99.com/test/delete_customer.php"
    driver1.Wait 1000
    
    driver2.NavigateTo "https://www.wikipedia.org/"
    driver2.Wait 1000
    
    keySeq = "Leonardo da VinJci" & keys.LeftKey & keys.LeftKey & keys.LeftKey & keys.DeleteKey & keys.ReturnKey
    
    driver2.FindElement(By.ID, "searchInput").SendKeys keySeq
    driver2.Wait 1000
    
    Debug.Assert driver1.IsAlertPresent = False
                                
    driver1.FindElement(By.Name, "cusid").SendKeys "87654"
    driver1.Wait 1000
    
    driver1.FindElement(By.Name, "submit").Click
    driver1.Wait 1000
    
    Debug.Assert driver1.IsAlertPresent = True
    Debug.Assert driver1.SwitchToAlert.GetText = "Do you really want to delete this Customer?"
    driver1.SwitchToAlert.Accept
    
    Debug.Assert driver1.SwitchToAlert.GetText = "Customer Successfully Delete!"
    driver1.SwitchToAlert.Accept

    driver1.Wait 1000
    driver1.CloseBrowser
    driver2.CloseBrowser
    
    'shuts down all instances listening to same port
    driver1.Shutdown
    'if drivers are on same port, this will fail, but is needed if on different ports
    driver2.Shutdown
End Sub

Sub test_MultiSession_mix_Edge_Chrome()
    'mixing driver Edge and Chrome works similar to running two of same
    Dim driver1 As SeleniumVBA.WebDriver
    Dim driver2 As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    Dim keySeq As String
    
    Set driver1 = SeleniumVBA.New_WebDriver
    Set driver2 = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    
    'driver1.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    'driver2.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    'driver1.CommandWindowStyle = vbNormalFocus
    'driver2.CommandWindowStyle = vbNormalFocus
    
    'this should work with only limited interference
    'however, for logging we only get one log (the first one)
    'driver1.StartChrome , , True, ".\chrome1.log"
    'driver2.StartEdge , , True, ".\edge1.log"
    
    'this will work with no interferrence
    driver1.StartChrome , 9515, True, ".\chrome1.log"
    driver2.StartEdge , 9516, True, ".\edge1.log"
    
    driver1.OpenBrowser
    driver2.OpenBrowser

    driver1.NavigateTo "http://demo.guru99.com/test/delete_customer.php"
    driver1.Wait 1000
    
    driver2.NavigateTo "https://www.wikipedia.org/"
    driver2.Wait 1000
    
    keySeq = "Leonardo da VinJci" & keys.LeftKey & keys.LeftKey & keys.LeftKey & keys.DeleteKey & keys.ReturnKey
    
    driver2.FindElement(By.ID, "searchInput").SendKeys keySeq
    driver2.Wait 1000
    
    Debug.Assert driver1.IsAlertPresent = False
                                
    driver1.FindElement(By.Name, "cusid").SendKeys "87654"
    driver1.Wait 1000
    
    driver1.FindElement(By.Name, "submit").Click
    driver1.Wait 1000
    
    Debug.Assert driver1.IsAlertPresent = True
    Debug.Assert driver1.SwitchToAlert.GetText = "Do you really want to delete this Customer?"
    driver1.SwitchToAlert.Accept
    
    Debug.Assert driver1.SwitchToAlert.GetText = "Customer Successfully Delete!"
    driver1.SwitchToAlert.Accept

    driver1.Wait 1000
    driver1.CloseBrowser
    driver2.CloseBrowser
    
    'shuts down all instances listening to same port
    driver1.Shutdown
    'if drivers are on same port, this will fail, but is needed if on different ports
    driver2.Shutdown
End Sub

Sub test_MultiSession_Firefox()
    'Firefox does not support multi-session on same port, so must run on different ports
    Dim driver1 As SeleniumVBA.WebDriver
    Dim driver2 As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    Dim keySeq As String
    
    Set driver1 = SeleniumVBA.New_WebDriver
    Set driver2 = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    
    'driver1.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    'driver2.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    'driver1.CommandWindowStyle = vbNormalFocus
    'driver2.CommandWindowStyle = vbNormalFocus

    'this fails as driver2 kills driver1 if running on same port
    'driver1.StartFirefox
    'driver2.StartFirefox
    
    'this works fine
    driver1.StartFirefox , 4444, True, ".\firefox1.log"
    driver2.StartFirefox , 4445, True, ".\firefox2.log"
    
    driver1.OpenBrowser
    driver2.OpenBrowser

    driver1.NavigateTo "http://demo.guru99.com/test/delete_customer.php"
    driver1.Wait 1000
    
    driver2.NavigateTo "https://www.wikipedia.org/"
    driver2.Wait 1000

    keySeq = "Leonardo da VinJci" & keys.LeftKey & keys.LeftKey & keys.LeftKey & keys.DeleteKey & keys.ReturnKey
    
    driver2.FindElement(By.ID, "searchInput").SendKeys keySeq
    driver2.Wait 1000
    
    Debug.Print "Is Alert Present: " & driver1.IsAlertPresent
                                
    driver1.FindElement(By.Name, "cusid").SendKeys "87654"
    driver1.Wait 1000
    
    driver1.FindElement(By.Name, "submit").Click
    driver1.Wait 1000
    
    Debug.Print "Is Alert Present: " & driver1.IsAlertPresent
    Debug.Print "Alert Text: " & driver1.SwitchToAlert.GetText
    driver1.SwitchToAlert.Accept
    driver1.Wait 1000
    
    Debug.Print "Alert Text: " & driver1.SwitchToAlert.GetText
    driver1.SwitchToAlert.Accept

    driver1.Wait 1000
    driver1.CloseBrowser
    driver2.CloseBrowser
    
    driver1.Shutdown
    driver2.Shutdown
End Sub
