Attribute VB_Name = "NewClass"
'these are used to instantiate the objects in other projects that reference this Add-in project
'see https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class
'late binding looks like this Set driver = Application.Run("'C:\path_to_addin\seleniumVBA.xlam'!New_WebDriver")

Public Function New_WebDriver() As WebDriver
    Set New_WebDriver = New WebDriver
End Function

Public Function New_WebDriverManager() As WebDriverManager
    Set New_WebDriverManager = New WebDriverManager
End Function

Public Function New_WebElement() As WebElement
    Set New_WebElement = New WebElement
End Function

Public Function New_WebElements() As WebElements
    Set New_WebElements = New WebElements
End Function

Public Function New_WebActionChain() As WebActionChain
    Set New_WebActionChain = New WebActionChain
End Function

Public Function New_WebCookie() As WebCookie
    Set New_WebCookie = New WebCookie
End Function

'new WebCookies should be instantiated from WebDriver.CreateCookies
'Public Function New_WebCookies() As WebCookies
'    Set New_WebCookies = New WebCookies
'End Function

Public Function New_WebJSonConverter() As WebJSonConverter
    Set New_WebJSonConverter = New WebJSonConverter
End Function

Public Function New_WebKeyboard() As WebKeyboard
    Set New_WebKeyboard = New WebKeyboard
End Function

'new WebCapabilities should be instantiated from WebDriver.CreateCapabilities
'Public Function New_WebCapabilities() As WebCapabilities
'    Set New_WebCapabilities = New WebCapabilities
'End Function

Public Function New_WebPrintSettings() As WebPrintSettings
    Set New_WebPrintSettings = New WebPrintSettings
End Function

Public Function New_WebShadowRoot() As WebShadowRoot
    Set New_WebShadowRoot = New WebShadowRoot
End Function
