## SeleniumVBA v2.4

A comprehensive Selenium wrapper for automating Edge, Chrome, and Firefox written in Windows Excel VBA

Modified/extended from [TinySeleniumVBA](https://github.com/uezo/TinySeleniumVBA/)

## Features

- Edge, Chrome, and Firefox browser support
- Wrappers for most of Selenium's JSon Wire Protocol
- Support for HTML DOM, Action Chains, SendKeys, Shadow Roots, Cookies, ExecuteScript, and Capabilities
- Automated Browser/WebDriver version alignment - works out-of-the-box with no manual driver downloads necessary!
- Relative paths and OneDrive support
- Open spec: This wrapper is an HTTP client of the Selenium WebDriver server, conforming closely to [W3C standards](https://www.w3.org/TR/webdriver/).

## Setup

SeleniumVBA has been designed to work out-of-the-box. Just download the [SeleniumVBA.xlam](https://github.com/GCuser99/SeleniumVBA/tree/main/dist) Excel Addin, open it, and run any one of the subs in the "test" Standard modules. If the Selenium WebDriver does not exist, or is out-of-date, SeleniumVBA will detect this automatically and download the appropriate driver to a desired location (currently defaults to user's download folder but that is easily configurable).

The user can also programmatically invoke driver updates via the WebDriverManager class (see example below).

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

## WebDriver/Browser Version Alignment

```vba
Sub updateDrivers()
    'This checks if driver is installed, or if installed driver is compatibile
    'with installed browser, and then if needed, installs an updated driver.

    'Note: SeleniumVBA automatically detects and updates (if needed) silently in the 
    'background everytime a WebDriver session is started - so running this sub is
    'not required to maintain driver/browser compatibility.
    Dim mngr As New WebDriverManager
    
    'mngr.DefaultDriverFolder = [your binary folder path here] 'defaults to Downloads dir
    
    'check/update the drivers and report the informative status messages
    MsgBox mngr.AlignEdgeDriverWithBrowser()
    MsgBox mngr.AlignChromeDriverWithBrowser()
    MsgBox mngr.AlignFirefoxDriverWithBrowser()
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
