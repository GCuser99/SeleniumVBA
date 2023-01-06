## SeleniumVBA v3.0

A comprehensive Selenium wrapper for automating Edge, Chrome, Firefox, and IE written in Windows Office VBA

Modified/extended from [TinySeleniumVBA](https://github.com/uezo/TinySeleniumVBA/)

## Features

- Edge, Chrome, Firefox, and IE Mode browser support
- A superset of Selenium's JSon Wire Protocol commands - [over 350 public methods and properties](https://github.com/GCuser99/SeleniumVBA/wiki/Object-Model-Overview)
- Support for HTML DOM, Action Chains, SendKeys, Shadow Roots, Cookies, ExecuteScript, and Capabilities
- Automated Browser/WebDriver version alignment - works out-of-the-box with no manual downloads necessary!
- Relative paths and OneDrive support
- This wrapper is an HTTP client of the Selenium WebDriver server, conforming closely to [W3C standards](https://www.w3.org/TR/webdriver/).
- Help documentation is available - see the [SeleniumVBA Wiki](https://github.com/GCuser99/SeleniumVBA/wiki)

## Setup

**SeleniumVBA will function right out-of-the-box**. Just download the [SeleniumVBA.xlam](https://github.com/GCuser99/SeleniumVBA/tree/main/dist) Excel Addin, open it, and run any one of the subs in the "test" Standard modules. If the Selenium WebDriver does not exist, or is out-of-date, SeleniumVBA will detect this automatically and download the appropriate driver to a desired location (currently defaults to user's download folder but that is easily configurable).

Driver updates can also be programmatically invoked via the WebDriverManager class.

## SendKeys Example

```vba
Sub doSendKeys()
    Dim driver As New WebDriver
    Dim keys As New WebKeyboard
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    keySeq = "This is COOKL!" & keys.LeftKey & keys.LeftKey & keys.LeftKey & keys.DeleteKey & keys.ReturnKey
    
    driver.FindElement(by.name, "q").SendKeys keySeq
    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
```

## File Download Example
```vba
Sub doFileDownload()
    Dim driver As New WebDriver
    Dim caps As WebCapabilities
   
    driver.StartChrome
    
    'set the directory path for saving download to
    Set caps = driver.CreateCapabilities
    caps.SetDownloadPrefs ".\"
    driver.OpenBrowser caps
    
    'delete legacy copy if it exists
    driver.DeleteFiles ".\test.pdf"
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/raw/main/dev/test_files/test.pdf"
    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
```

## Action Chain Example
```vba
Sub doActionChain()
    Dim driver As New WebDriver
    Dim keys As New WebKeyboard
    Dim actions As WebActionChain
    Dim elemSearch As WebElement
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    Set elemSearch = driver.FindElement(by.name, "btnK")
    
    Set actions = driver.ActionChain
    
    'build the chain and then execute with Perform method
    actions.KeyDown(keys.ShiftKey).SendKeys("upper case").KeyUp(keys.ShiftKey)
    actions.MoveToElement(elemSearch).Click().Perform

    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
```
## Collaborators

[@6DiegoDiego9](https://github.com/6DiegoDiego9)

## Credits

[TinySeleniumVBA](https://github.com/uezo/TinySeleniumVBA/) by Uezo and other contributors to that project

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) by Tim Hall, JSON converter for VBA

[RubberDuck](https://rubberduckvba.com/) by Mathieu Guindon
