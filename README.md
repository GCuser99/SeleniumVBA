<img src="https://github.com/GCuser99/SeleniumVBA/blob/main/dev/logo/logo.png" alt="SeleniumVBA" width=25% height=25%>

A comprehensive Selenium wrapper for browser automation written in Windows Office VBA (64-bit)

## Features

- Edge, Chrome, Firefox, and IE Mode browser automation support
- MS Excel Add-in, MS Access DB, and experimental [twinBasic](https://twinbasic.com/preview.html) ActiveX Dll solutions available
- A superset of Selenium's JSon Wire Protocol commands - [over 350 public methods and properties](https://github.com/GCuser99/SeleniumVBA/wiki/Object-Model-Overview)
- Support for HTML DOM, Action Chains, SendKeys, Shadow Roots, Cookies, ExecuteScript, and Capabilities
- Automated Browser/WebDriver version alignment - works out-of-the-box with no manual downloads necessary!
- Help documentation is available in the [SeleniumVBA Wiki](https://github.com/GCuser99/SeleniumVBA/wiki)

## Setup

**SeleniumVBA will function right out-of-the-box**. Just download the [SeleniumVBA.xlam](https://github.com/GCuser99/SeleniumVBA/tree/main/dist) Excel Addin, open it, and run any one of the subs in the "test" Standard modules. If the Selenium WebDriver does not exist, or is out-of-date, SeleniumVBA will detect this automatically and download the appropriate driver to a desired location (currently defaults to user's download folder but that is easily configurable).

Driver updates can also be programmatically invoked via the WebDriverManager class.

To try the experimental SeleniumVBA twinBasic Dll, see instructions in the [twinBasic folder](https://github.com/GCuser99/SeleniumVBA/tree/main/twinBasic).

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

[@GCUser99](https://github.com/GCUser99)

## Credits

This project is an extensively modified/extended version of uezo's [TinySeleniumVBA](https://github.com/uezo/TinySeleniumVBA/)

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) by Tim Hall, JSON converter for VBA

[RubberDuck](https://rubberduckvba.com/) by Mathieu Guindon

[twinBasic](https://twinbasic.com/preview.html) by Wayne Phillips

[Inno Setup](https://jrsoftware.org/isinfo.php) by Jordan Russell

[UninsIS](https://github.com/Bill-Stewart/UninsIS) by Bill Stewart
