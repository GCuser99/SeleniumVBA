# NotSoTinySeleniumVBA

A Selenium wrapper written in Windows VBA based on JSon wire protocol.

Modified extensively from https://github.com/uezo/TinySeleniumVBA/

# Features

- No installation permissions required
- Wrappers for most of Selenium JSon wire protocol
- Optional Browser/WebDriver version alignment via WebDriverManager class
- Open spec: Basically this wrapper is just a HTTP client of WebDriver server. Learning this wrapper equals to learning WebDriver.
https://www.w3.org/TR/webdriver/


# Setup

1. Set reference to `Microsoft Scripting Runtime`

1. Add `WebDriver.cls`, `WebElement.cls` and `JsonConverter.bas` to your VBA Project
    - Latest (v0.1.1): https://

1. Download WebDriver (driver and browser should be the same version)
    - Edge: https://
    - Chrome: https://

# Usage

```vb
Public Sub main()

End Sub
```

# Credits

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) by Tim Hall, JSON converter for VBA helps me a lot to make HTTP client and this awesome library is included in the release under its license. Thank you!
