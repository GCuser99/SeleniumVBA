Attribute VB_Name = "ClassFactory"
Attribute VB_Description = "This class is used for object instantiation when referencing SeleniumVBA externally from another code project"
'@ModuleDescription "This class is used for object instantiation when referencing SeleniumVBA externally from another code project"
'@folder("SeleniumVBA.Source")
' ==========================================================================
' SeleniumVBA v3.0
' A Selenium wrapper for Edge, Chrome, Firefox, and IE written in Windows VBA based on JSon wire protocol.
'
' (c) GCUser99
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' ==========================================================================
' These methods are used to instantiate the objects in other projects that reference this Add-in project
' See https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class
'
' From the referencing vba project use the following syntax:
'
' Dim driver as WedDriver
' Set driver = SeleniumVBA.New_WebDriver
'
' Late binding looks like this:
'
' Set driver = Application.Run("'C:\path_to_addin\seleniumVBA.xlam'!New_WebDriver")'

Option Explicit

'@Description("Instantiates a WebDriver object")
Public Function New_WebDriver() As WebDriver
Attribute New_WebDriver.VB_Description = "Instantiates a WebDriver object"
    Set New_WebDriver = New WebDriver
End Function

'@Description("Instantiates a WebDriverManager object")
Public Function New_WebDriverManager() As WebDriverManager
Attribute New_WebDriverManager.VB_Description = "Instantiates a WebDriverManager object"
    Set New_WebDriverManager = New WebDriverManager
End Function

'@Description("Instantiates a WebElement object")
Public Function New_WebElement() As WebElement
Attribute New_WebElement.VB_Description = "Instantiates a WebElement object"
    Set New_WebElement = New WebElement
End Function

'@Description("Instantiates a WebElements object")
Public Function New_WebElements() As WebElements
Attribute New_WebElements.VB_Description = "Instantiates a WebElements object"
    Set New_WebElements = New WebElements
End Function

'new WebActionChain should be instantiated from WebDriver.Actions
'Public Function New_WebActionChain() As WebActionChain
'    Set New_WebActionChain = New WebActionChain
'End Function

'@Description("Instantiates a WebCookie object")
Public Function New_WebCookie() As WebCookie
Attribute New_WebCookie.VB_Description = "Instantiates a WebCookie object"
    Set New_WebCookie = New WebCookie
End Function

'new WebCookies should be instantiated from WebDriver.CreateCookies
'Public Function New_WebCookies() As WebCookies
'    Set New_WebCookies = New WebCookies
'End Function

'@Description("Instantiates a WebJsonConverter object - this is optional as this object is predeclared")
Public Function New_WebJSonConverter() As WebJsonConverter
Attribute New_WebJSonConverter.VB_Description = "Instantiates a WebJsonConverter object - this is optional as this object is predeclared"
    Set New_WebJSonConverter = New WebJsonConverter
End Function

'@Description("Instantiates a WebKeyboard object - this is optional as this object is predeclared")
Public Function New_WebKeyboard() As WebKeyboard
Attribute New_WebKeyboard.VB_Description = "Instantiates a WebKeyboard object - this is optional as this object is predeclared"
    Set New_WebKeyboard = New WebKeyboard
End Function

'new WebCapabilities should be instantiated from WebDriver.CreateCapabilities
'Public Function New_WebCapabilities() As WebCapabilities
'    Set New_WebCapabilities = New WebCapabilities
'End Function

'@Description("Instantiates a WebPrintSettings object")
Public Function New_WebPrintSettings() As WebPrintSettings
Attribute New_WebPrintSettings.VB_Description = "Instantiates a WebPrintSettings object"
    Set New_WebPrintSettings = New WebPrintSettings
End Function

'@Description("Instantiates a WebShadowRoot object")
Public Function New_WebShadowRoot() As WebShadowRoot
Attribute New_WebShadowRoot.VB_Description = "Instantiates a WebShadowRoot object"
    Set New_WebShadowRoot = New WebShadowRoot
End Function
