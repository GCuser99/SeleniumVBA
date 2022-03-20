# SeleniumVBA

A Selenium wrapper for Edge and Chrome written in Windows VBA based on JSon wire protocol.

Modified/extended from https://github.com/uezo/TinySeleniumVBA/

# Features

- Edge and Chrome browser support
- Wrappers for most of Selenium's JSon wire protocol
- Support for HTML DOM, Action Chains, SendKeys, Shadow Roots, Cookies, and Capabilities
- Optional Browser/WebDriver version alignment via WebDriverManager class
- Open spec: Basically this wrapper is just a HTTP client of WebDriver server. Learning this wrapper equals to learning WebDriver.
https://www.w3.org/TR/webdriver/


# Setup

1. Import standard modules and classes to Excel VBA
2. Set the following references:
<img src="https://user-images.githubusercontent.com/26237126/159189311-001ed714-0378-496f-abc7-5f26dfeb054e.png" width="300" height="200">`
3. OR... download Excel file seleniumvba_v0.0.1.xlsm
4. Download WebDriver (driver and browser should be the same version)
    - Edge: https://
    - Chrome: https://

# Usage

```vb
Sub Main()
    Dim Driver As New WebDriver
    Dim keys As New Keyboard
    
    Driver.Chrome
    Driver.OpenBrowser
    
    Driver.Navigate "https://www.google.com/"
    Driver.Wait 1000
    
    Driver.FindElement(by.name, "q").SendKeys "This is COOKL!" & keys.LeftKey & keys.LeftKey & keys.LeftKey & keys.DeleteKey & keys.ReturnKey

    Driver.Wait 2000
    
    Driver.CloseBrowser
    Driver.Shutdown
End Sub
```

# Credits

[TinySeleniumVBA](https://github.com/uezo/TinySeleniumVBA/) by Uezo and other contributors to that project

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) by Tim Hall, JSON converter for VBA
