Attribute VB_Name = "test_Sendkeys"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_sendkeys()
    Dim driver As New WebDriver
    Dim input1 As WebElement
    Dim display_keys As WebElement
    Dim keys As New WebKeyboard
    Dim html As String

    driver.StartChrome
    driver.OpenBrowser
    
    html = "<html>" & vbCrLf
    html = html & "    <head>" & vbCrLf
    html = html & "        <title>Test SendKeys</title>" & vbCrLf
    html = html & "        <style>" & vbCrLf
    html = html & "            #display_keys {color: red;}" & vbCrLf
    html = html & "            #display_key {color: blue;}" & vbCrLf
    html = html & "            #display_keyCode {color: blue;}" & vbCrLf
    html = html & "            #display_code {color: blue;}" & vbCrLf
    html = html & "            #display_location {color: blue;}" & vbCrLf
    html = html & "            #display_ctrlKey {color: blue;}" & vbCrLf
    html = html & "            #display_shiftKey {color: blue;}" & vbCrLf
    html = html & "            #display_altKey {color: blue;}" & vbCrLf
    html = html & "            #display_metaKey {color: blue;}" & vbCrLf
    html = html & "        </style>" & vbCrLf
    html = html & "    </head>" & vbCrLf
    html = html & "    <body>" & vbCrLf
    html = html & "        <div><div>Type Keys and Hit Enter to Process:</div><input type='text' id='input1'>" & vbCrLf
    html = html & "        <div id='display_keys'></div></div>" & vbCrLf
    html = html & "        <p></p><p></p>" & vbCrLf
    html = html & "        <div><div>Type Single Key for Info:</div><input type='text' id='input2'></div>" & vbCrLf
    html = html & "        <div id='display_key'></div>" & vbCrLf
    html = html & "        <div id='display_code'></div>" & vbCrLf
    html = html & "        <div id='display_keyCode'></div>" & vbCrLf
    html = html & "        <div id='display_location'></div>" & vbCrLf
    html = html & "        <div id='display_ctrlKey'></div>" & vbCrLf
    html = html & "        <div id='display_shiftKey'></div>" & vbCrLf
    html = html & "        <div id='display_altKey'></div>" & vbCrLf
    html = html & "        <div id='display_metaKey'></div>" & vbCrLf
    html = html & "        <script>" & vbCrLf
    html = html & "            const input1 = document.getElementById('input1');" & vbCrLf
    html = html & "            const input2 = document.getElementById('input2');" & vbCrLf
    html = html & "            const display_keys = document.getElementById('display_keys');" & vbCrLf
    html = html & "            const display_key = document.getElementById('display_key');" & vbCrLf
    html = html & "            const display_code = document.getElementById('display_code');" & vbCrLf
    html = html & "            const display_keyCode = document.getElementById('display_keyCode');" & vbCrLf
    html = html & "            const display_location = document.getElementById('display_location');" & vbCrLf
    html = html & "            const display_ctrlKey = document.getElementById('display_ctrlKey');" & vbCrLf
    html = html & "            const display_shiftKey = document.getElementById('display_shiftKey');" & vbCrLf
    html = html & "            const display_altKey = document.getElementById('display_altKey');" & vbCrLf
    html = html & "            const display_metaKey = document.getElementById('display_metaKey');" & vbCrLf
    html = html & "            input1.addEventListener('keydown', function(event) {" & vbCrLf
    html = html & "                if (event.key === 'Enter') {" & vbCrLf
    html = html & "                    display_keys.textContent = input1.value;" & vbCrLf
    html = html & "                    input1.value = ''; // Clear the input field" & vbCrLf
    html = html & "                }" & vbCrLf
    html = html & "            });" & vbCrLf
    html = html & "            input2.addEventListener('keydown', function(event) {" & vbCrLf
    html = html & "                display_key.textContent = 'key: ' + event.key;" & vbCrLf
    html = html & "                display_code.textContent = 'code: ' + event.code;" & vbCrLf
    html = html & "                display_keyCode.textContent = 'keyCode: ' + event.keyCode;" & vbCrLf
    html = html & "                display_location.textContent = 'location: ' + event.location;" & vbCrLf
    html = html & "                display_ctrlKey.textContent = 'ctrlKey: ' + event.ctrlKey;" & vbCrLf
    html = html & "                display_shiftKey.textContent = 'shiftKey: ' + event.shiftKey;" & vbCrLf
    html = html & "                display_altKey.textContent = 'altKey: ' + event.altKey;" & vbCrLf
    html = html & "                display_metaKey.textContent = 'metaKey: ' + event.metaKey;" & vbCrLf
    html = html & "            });" & vbCrLf
    html = html & "        </script>" & vbCrLf
    html = html & "    </body>" & vbCrLf
    html = html & "</html>" & vbCrLf

    driver.NavigateToString html
    driver.Wait 1000
    Set input1 = driver.FindElementByID("input1")
    Set display_keys = driver.FindElementByID("display_keys")

    input1.SendKeys "abcdefghijklmnopqrstuvwxyz" & keys.EnterKey
    Debug.Assert display_keys.GetText = "abcdefghijklmnopqrstuvwxyz"
    
    driver.Wait
    
    input1.SendKeys "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & keys.EnterKey
    Debug.Assert display_keys.GetText = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    driver.Wait
    
    input1.SendKeys "`1234567890-=[]\;',./" & keys.EnterKey
    Debug.Assert display_keys.GetText = "`1234567890-=[]\;',./"
    
    driver.Wait
    
    input1.SendKeys "~!@#$%^&*()_+{}|:""<>?" & keys.EnterKey
    Debug.Assert display_keys.GetText = "~!@#$%^&*()_+{}|:""<>?"
    
    driver.Wait
    
    input1.SendKeys keys.ShiftKey & "abcdefghijklmnopqrstuvwxyz" & keys.ShiftKey & keys.EnterKey
    Debug.Assert display_keys.GetText = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    driver.Wait
    
    input1.SendKeys keys.ShiftKey & "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & keys.ShiftKey & keys.EnterKey
    Debug.Assert display_keys.GetText = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    driver.Wait

    input1.SendKeys keys.ShiftKey & "`1234567890-=[]\;',./" & keys.ShiftKey & keys.EnterKey
    Debug.Assert display_keys.GetText = "~!@#$%^&*()_+{}|:""<>?"
    
    driver.Wait
    
    input1.SendKeys keys.ShiftKey & "~!@#$%^&*()_+{}|:""<>?" & keys.ShiftKey & keys.EnterKey
    Debug.Assert display_keys.GetText = "~!@#$%^&*()_+{}|:""<>?"
    
    driver.Wait
    
    input1.SendKeys keys.ShiftKey & "this is" & keys.ShiftKey & " Mike" & keys.EnterKey
    Debug.Assert display_keys.GetText = "THIS IS Mike"
    
    driver.Wait
    
    input1.SendKeys "this is Mike" & keys.ShiftKey & keys.Repeat(keys.LeftKey, 5) & keys.ShiftKey & keys.DeleteKey & " Sally" & keys.EnterKey
    Debug.Assert display_keys.GetText = "this is Sally"
    
    driver.Wait
    
    input1.SendKeys "this is Mike" & keys.CtrlKey & "A" & keys.CtrlKey & keys.DeleteKey & "all clear" & keys.EnterKey
    Debug.Assert display_keys.GetText = "all clear"
    
    driver.Wait
    
    input1.SendKeys "this is Mike" & keys.CtrlKey & "a" & keys.CtrlKey & keys.DeleteKey & "all clear" & keys.EnterKey
    Debug.Assert display_keys.GetText = "all clear"
    
    driver.Wait
    
    input1.SendKeys keys.ShiftKey & "this is" & keys.NullKey & keys.ShiftKey & " Mike" & keys.EnterKey
    Debug.Assert display_keys.GetText = "THIS IS MIKE"
    
    driver.Wait
    
    input1.SendKeys "this is Mike" & keys.EnterKey
    Debug.Assert display_keys.GetText = "this is Mike"
    
    driver.Wait
    
    input1.SendKeys "this is Mike"
    input1.SendKeys keys.EnterKey, False
    Debug.Assert display_keys.GetText = "this is Mike"
    
    driver.Wait
    
    input1.SendKeys "this is Mike" & keys.CtrlKey & "a" & keys.CtrlKey & keys.CtrlKey & "c" & keys.CtrlKey & keys.DeleteKey & keys.EnterKey
    input1.SendKeys keys.CtrlKey & "v" & keys.CtrlKey & keys.EnterKey
    Debug.Assert display_keys.GetText = "this is Mike"
    
    driver.Wait
    
    input1.SendKeys "this is Mike" & keys.Chord(keys.CtrlKey, "a") & keys.Chord(keys.CtrlKey, "c") & keys.DeleteKey & keys.EnterKey()
    input1.SendKeys keys.Chord(keys.CtrlKey, "v") & keys.EnterKey
    Debug.Assert display_keys.GetText = "this is Mike"
    
    driver.Wait
    
    input1.SendKeys "this is Mike" & keys.Chord(keys.CtrlKey, "A") & keys.Chord(keys.CtrlKey, "C") & keys.DeleteKey & keys.EnterKey()
    input1.SendKeys keys.Chord(keys.CtrlKey, "V") & keys.EnterKey
    Debug.Assert display_keys.GetText = "this is Mike"
    
    driver.Wait
    
    Dim input2 As WebElement
    Dim display_key As WebElement
    Dim display_code As WebElement
    Dim display_keyCode As WebElement
    Dim display_location As WebElement
    Dim display_shiftKey As WebElement
    Dim display_ctrlKey As WebElement
    Dim display_altKey As WebElement
    Dim display_metaKey As WebElement
    
    Set input2 = driver.FindElementByID("input2")
    Set display_key = driver.FindElementByID("display_key")
    Set display_code = driver.FindElementByID("display_code")
    Set display_keyCode = driver.FindElementByID("display_keyCode")
    Set display_location = driver.FindElementByID("display_location")
    Set display_shiftKey = driver.FindElementByID("display_shiftKey")
    Set display_ctrlKey = driver.FindElementByID("display_ctrlKey")
    Set display_altKey = driver.FindElementByID("display_altKey")
    Set display_metaKey = driver.FindElementByID("display_metaKey")
    
    input2.SendKeys "a", True
    Debug.Assert display_key.GetText = "key: a"
    Debug.Assert display_code.GetText = "code: KeyA"
    Debug.Assert display_keyCode.GetText = "keyCode: 65"
    Debug.Assert display_location.GetText = "location: 0"
    Debug.Assert display_shiftKey.GetText = "shiftKey: false"
    Debug.Assert display_ctrlKey.GetText = "ctrlKey: false"
    Debug.Assert display_altKey.GetText = "altKey: false"
    Debug.Assert display_metaKey.GetText = "metaKey: false"
    
    driver.Wait
    
    input2.SendKeys "A", True
    Debug.Assert display_key.GetText = "key: A"
    Debug.Assert display_code.GetText = "code: KeyA"
    Debug.Assert display_keyCode.GetText = "keyCode: 65"
    Debug.Assert display_location.GetText = "location: 0"
    Debug.Assert display_shiftKey.GetText = "shiftKey: true"
    Debug.Assert display_ctrlKey.GetText = "ctrlKey: false"
    Debug.Assert display_altKey.GetText = "altKey: false"
    Debug.Assert display_metaKey.GetText = "metaKey: false"
    
    driver.Wait
    
    input2.SendKeys keys.ShiftKey & "a", True
    Debug.Assert display_key.GetText = "key: A"
    Debug.Assert display_code.GetText = "code: KeyA"
    Debug.Assert display_keyCode.GetText = "keyCode: 65"
    Debug.Assert display_location.GetText = "location: 0"
    Debug.Assert display_shiftKey.GetText = "shiftKey: true"
    Debug.Assert display_ctrlKey.GetText = "ctrlKey: false"
    Debug.Assert display_altKey.GetText = "altKey: false"
    Debug.Assert display_metaKey.GetText = "metaKey: false"
    
    driver.Wait
    
    input2.SendKeys keys.ShiftKey & "A", True
    Debug.Assert display_key.GetText = "key: A"
    Debug.Assert display_code.GetText = "code: KeyA"
    Debug.Assert display_keyCode.GetText = "keyCode: 65"
    Debug.Assert display_location.GetText = "location: 0"
    Debug.Assert display_shiftKey.GetText = "shiftKey: true"
    Debug.Assert display_ctrlKey.GetText = "ctrlKey: false"
    Debug.Assert display_altKey.GetText = "altKey: false"
    Debug.Assert display_metaKey.GetText = "metaKey: false"
    
    driver.Wait
    
    input2.SendKeys keys.CtrlKey & "a", True
    Debug.Assert display_key.GetText = "key: a"
    Debug.Assert display_code.GetText = "code: KeyA"
    Debug.Assert display_keyCode.GetText = "keyCode: 65"
    Debug.Assert display_location.GetText = "location: 0"
    Debug.Assert display_shiftKey.GetText = "shiftKey: false"
    Debug.Assert display_ctrlKey.GetText = "ctrlKey: true"
    Debug.Assert display_altKey.GetText = "altKey: false"
    Debug.Assert display_metaKey.GetText = "metaKey: false"
    
    driver.Wait
    
    input2.SendKeys keys.EnterKey, True
    Debug.Assert display_key.GetText = "key: Enter"
    Debug.Assert display_code.GetText = "code: Enter"
    Debug.Assert display_keyCode.GetText = "keyCode: 13"
    Debug.Assert display_location.GetText = "location: 0"
    Debug.Assert display_shiftKey.GetText = "shiftKey: false"
    Debug.Assert display_ctrlKey.GetText = "ctrlKey: false"
    Debug.Assert display_altKey.GetText = "altKey: false"
    Debug.Assert display_metaKey.GetText = "metaKey: false"
    
    driver.Wait
    
    input2.SendKeys keys.CtrlShiftKeys & keys.HomeKey, True
    Debug.Assert display_key.GetText = "key: Home"
    Debug.Assert display_code.GetText = "code: Home"
    Debug.Assert display_keyCode.GetText = "keyCode: 36"
    Debug.Assert display_location.GetText = "location: 0"
    Debug.Assert display_shiftKey.GetText = "shiftKey: true"
    Debug.Assert display_ctrlKey.GetText = "ctrlKey: true"
    Debug.Assert display_altKey.GetText = "altKey: false"
    Debug.Assert display_metaKey.GetText = "metaKey: false"

    driver.Wait 1000
    driver.Shutdown
End Sub

Sub test_SendKeysToOS()
    'WARNING: SendKeysToOS could crash host app with MalwareBytes AV real-time protection
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
        Debug.Assert elem.GetText = "Congratulations! You must have the proper credentials."
    End If
  
    driver.CloseBrowser
    driver.Shutdown
End Sub
