Attribute VB_Name = "test_Sendkeys"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_Sendkeys()
    Dim driver As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    Dim keySeq As String
    
    Set driver = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard

    driver.StartEdge
    
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 2000
    
    driver.NavigateTo "https://www.wikipedia.org/"
    
    keySeq = "Leonardo da VinJci" & keys.Repeat(keys.LeftKey, 3) & keys.DeleteKey & keys.ReturnKey
    
    driver.FindElement(By.ID, "searchInput").SendKeys keySeq

    driver.Wait 1500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_SendKeysToOS()
    'SendKeysToOS can be useful for sending keyboard input to non-browser OS windows,
    'as well as browser windows where the standard SendKeys method does not function
    'with the window of interest. For more info, see:
    'https://github.com/GCuser99/SeleniumVBA/wiki#browser-keyboard-interaction-with-sendkeys
    'https://github.com/GCuser99/SeleniumVBA/discussions/84.
    Dim driver As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    
    Set driver = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    
    driver.StartChrome
    driver.OpenBrowser
    
    'Sends input to the future InputBox(es) via separate thread(s).
    'If windowTitle is not specified, then sends input to active window
    driver.SendKeysToOS _
        keyString:="Here is my input to InputBox1!" & keys.EnterKey, _
        windowTitle:="InputBox1 Title", _
        runOnSeparateThread:=True, _
        waitForWindow:=True, _
        maxTimeToWaitMS:=5000, _
        timeDelayMS:=0
        
    'just for fun, let's launch another thread looking for a second InputBox
    'but this time we send the escape key to cancel the dialog after keys are sent
    driver.SendKeysToOS _
        keyString:="Here is my input to InputBox2!" & keys.EscapeKey, _
        windowTitle:="InputBox2 Title", _
        runOnSeparateThread:=True, _
        waitForWindow:=True, _
        maxTimeToWaitMS:=5000, _
        timeDelayMS:=0

    driver.NavigateTo "https://example.com/"
    driver.Wait 1000
    
    'these InputBox dialogs will block execution flow until they receive input
    'so must send input keys from separate threads
    ThisWorkbook.Activate
    'InputBox2 gets cancelled with escape key, so returns vbNullString
    Debug.Print InputBox("Please enter input keys:", "InputBox2 Title")
    'InputBox1 returns the user input keys
    Debug.Print InputBox("Please enter input keys:", "InputBox1 Title")
    
    driver.Wait 1000
        
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_Authentication()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim creds As String
    Dim keys As SeleniumVBA.WebKeyboard
    
    Set driver = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 10000
    
    driver.NavigateTo "http://the-internet.herokuapp.com/basic_auth"
    
    'no need to run on a separate thread in this case as the login popup does
    'not block execution while waiting for user response...
    creds = "admin" & keys.TabKey & "admin" & keys.EnterKey 'username and password
    driver.SendKeysToOS _
        keyString:=creds, _
        timeDelayMS:=0, _
        windowTitle:="", _
        runOnSeparateThread:=False, _
        waitForWindow:=False
    
    If driver.IsPresent(By.CssSelector, "#content > div > p", elemFound:=elem) Then
        Debug.Print elem.GetText
    End If
  
    driver.CloseBrowser
    driver.Shutdown
End Sub
