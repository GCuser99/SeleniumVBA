<img src="https://github.com/GCuser99/SeleniumVBA/blob/main/dev/logo/logo.png" alt="SeleniumVBA" width=25%>

A comprehensive Selenium wrapper for browser automation developed for MS Office VBA running in Windows

## Features

- Edge, Chrome, and Firefox browser automation support
- MS Excel Workbook, MS Access DB, and [twinBASIC](https://twinbasic.com/preview.html) ActiveX DLL solutions available
- A superset of Selenium's [W3C WebDriver](https://w3c.github.io/webdriver/) commands - [over 400 public methods and properties](https://github.com/GCuser99/SeleniumVBA/wiki/Object-Model-Overview)
- Support for HTML DOM, Action Chains, SendKeys, Shadow Roots, Cookies, ExecuteScript, and Capabilities
- ExecuteCDP method exposing [Chrome DevTools Protocol](https://chromedevtools.github.io/devtools-protocol/) - a low-level interface for browser interaction.
- Automated Browser/WebDriver version alignment - works out-of-the-box with no manual downloads necessary!
- Help documentation is available in the [SeleniumVBA Wiki](https://github.com/GCuser99/SeleniumVBA/wiki)

**What's New?**

- Now [WebDriver BiDi](https://w3c.github.io/webdriver-bidi/) compatible - see hanamichi77777's [WebDriver BiDi for SeleniumVBA GitHub repo](https://github.com/hanamichi77777/WebDriverBiDi-via-VBA-test)
- Featured on [Sancarn](https://github.com/sancarn)'s curated [Awesome VBA list](https://github.com/sancarn/awesome-vba?tab=readme-ov-file#web-tools) [![Awesome](https://awesome.re/badge.svg)](https://awesome.re)

## Setup

**SeleniumVBA will function right out-of-the-box**. Just download/install any one of the provided [SeleniumVBA solutions](https://github.com/GCuser99/SeleniumVBA/tree/main/dist) and then run one of the subs in the "test" Standard modules. If the Selenium WebDriver does not exist, or is out-of-date, SeleniumVBA will detect this automatically and download the appropriate driver to a [configurable location](https://github.com/GCuser99/SeleniumVBA/wiki#advanced-customization) on your system.

Driver updates can also be programmatically invoked via the [WebDriverManager class](https://github.com/GCuser99/SeleniumVBA/wiki/Object-Model-Overview#webdrivermanager).

The [twinBASIC](https://twinbasic.com/preview.html) ActiveX DLL solution requires no dependencies (such as .Net Framework). To try it, download and run the installer in the [dist folder](https://github.com/GCuser99/SeleniumVBA/tree/main/dist).

## SendKeys Example

```vba
Sub doSendKeys()
    Dim driver As New WebDriver
    Dim keys As New WebKeyboard
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    keySeq = "This is COOKL!" & keys.Repeat(keys.LeftKey, 3) & keys.DeleteKey & keys.ReturnKey
    
    driver.FindElement(By.Name, "q").SendKeys keySeq
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
    caps.SetDownloadPrefs downloadFolderPath:=".\"
    driver.OpenBrowser caps
    
    'delete legacy copy if it exists
    driver.DeleteFiles ".\test.pdf"
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/raw/main/dev/test_files/test.pdf"

    'wait until the download is complete before closing browser
    driver.WaitForDownload ".\test.pdf"
    
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
    
    Set elemSearch = driver.FindElement(By.Name, "btnK")
    
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

[@6DiegoDiego9](https://github.com/6DiegoDiego9) and [@GCUser99](https://github.com/GCUser99)

## Credits

This project is an extensively modified/extended version of uezo's [TinySeleniumVBA](https://github.com/uezo/TinySeleniumVBA/)

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) by Tim Hall, JSON converter for VBA

[RubberDuck](https://rubberduckvba.com/) by Mathieu Guindon

[twinBASIC](https://twinbasic.com/preview.html) by Wayne Phillips

[Inno Setup](https://jrsoftware.org/isinfo.php) by Jordan Russell

[vba-regex](https://github.com/sihlfall/vba-regex) by sihlfall


