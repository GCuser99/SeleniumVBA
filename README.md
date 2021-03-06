## SeleniumVBA v1.5

A comprehensive Selenium wrapper for automating Edge, Chrome, and Firefox written in Windows Excel VBA

Modified/extended from [TinySeleniumVBA](https://github.com/uezo/TinySeleniumVBA/)

## Features

- Edge, Chrome, and Firefox browser support
- Wrappers for most of Selenium's JSon Wire Protocol
- Support for HTML DOM, Action Chains, SendKeys, Shadow Roots, Cookies, ExecuteScript, and Capabilities
- Automated Browser/WebDriver version alignment via WebDriverManager class (see [test_UpdateDriver.bas](https://github.com/GCuser99/SeleniumVBA/tree/main/test))
- Open spec: This wrapper is an HTTP client of the Selenium WebDriver server, conforming to [W3C standards](https://www.w3.org/TR/webdriver/).

## Setup

1. Import class and standard modules from this repository into into Excel VBA
2. Set the following VBA references:

<img src="https://github.com/GCuser99/SeleniumVBA/blob/main/src/references.png" width="300" height="200">`

3. Or alternatively... download the zipped Excel file [seleniumvba_v1.5.zip](https://github.com/GCuser99/SeleniumVBA/tree/main/dist/) - it's ready to go...
4. Download WebDrivers into same directory as the Excel file (each driver should be same major version as corresponding browser)
   
   Edge: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
   
   Chrome: https://chromedriver.chromium.org/downloads

5. Or alternatively... let WebDriverManager class download and install drivers automatically (see [test_UpdateDriver.bas](https://github.com/GCuser99/SeleniumVBA/tree/main/test))

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
    'this checks if driver is installed, or if installed driver is compatibile
    'with installed browser, and then if needed, installs an updated driver
    Dim mngr As New WebDriverManager
    
    'update the drivers and report the informative status messages
    MsgBox mngr.AlignEdgeDriverWithBrowser(".\msedgedriver.exe")
    MsgBox mngr.AlignChromeDriverWithBrowser(".\chromedriver.exe")
    MsgBox mngr.AlignFirefoxDriverWithBrowser(".\geckodriver.exe")
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

## Credits

[TinySeleniumVBA](https://github.com/uezo/TinySeleniumVBA/) by Uezo and other contributors to that project

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) by Tim Hall, JSON converter for VBA
