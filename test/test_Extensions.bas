Attribute VB_Name = "test_Extensions"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_addExtensions()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities

    Set driver = SeleniumVBA.New_WebDriver
    driver.StartChrome

    Set caps = driver.CreateCapabilities()
    
    'temporary work-around for this issue: https://github.com/SeleniumHQ/selenium/issues/15788
    caps.AddArguments "--disable-features=DisableLoadExtensionCommandLineSwitch"
    caps.AddArguments "--enable-unsafe-extension-debugging"
    'this will load a local crx file extension(s) (Chrome/Edge only)
    caps.AddExtensions Environ("USERPROFILE") & "\Documents\SeleniumVBA\extensions\" & "TickTick-Todo-Task-List-Chrome-Web-Store.crx"
    
    driver.OpenBrowser caps

    driver.NavigateTo "https://www.wikipedia.org/"
    
    driver.Wait 1000

    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_addExtensions2()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities

    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome

    Set caps = driver.CreateCapabilities()
    
    'temporary work-around for this issue: https://github.com/SeleniumHQ/selenium/issues/15788
    caps.AddArguments "--disable-features=DisableLoadExtensionCommandLineSwitch"
    caps.AddArguments "--enable-unsafe-extension-debugging"

    'this will load an unpacked extension from Chrome's User Data extensions directory
    caps.AddArguments "--load-extension=" & Environ("LOCALAPPDATA") & "\Google\Chrome\User Data\Default\Extensions\ajkhmmldknmfjnmeedkbkkojgobmljda\1.5.9_0"

    driver.OpenBrowser caps

    driver.NavigateTo "https://www.wikipedia.org/"
    
    driver.Wait 1000

    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_InstallAddon()
    Dim driver As SeleniumVBA.WebDriver
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartFirefox
    driver.OpenBrowser
    
    'this will install an xpi add-on (Firefox only) - use AddExtensions method of WebCapabilities for Edge/Chrome
    driver.InstallAddon Environ("USERPROFILE") & "\Documents\SeleniumVBA\extensions\darkreader-4.9.94.xpi"
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
