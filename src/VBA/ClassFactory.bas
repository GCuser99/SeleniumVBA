Attribute VB_Name = "ClassFactory"
Attribute VB_Description = "This class is used for object instantiation when referencing SeleniumVBA externally from another code project"
'@ModuleDescription "This class is used for object instantiation when referencing SeleniumVBA externally from another code project"
'@folder("SeleniumVBA.Source")
' ==========================================================================
' SeleniumVBA v6.3
'
' A Selenium wrapper for browser automation developed for MS Office VBA
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' Contact Info:
'
' https://github.com/6DiegoDiego9
' https://github.com/GCUser99
' ==========================================================================
' MIT License
'
' Copyright (c) 2023-2024, GCUser99 and 6DiegoDiego9 (https://github.com/GCuser99/SeleniumVBA)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
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

'new WebAlert should be instantiated from WebDriver.SwitchToAlert
'new WebCapabilities should be instantiated from WebDriver.CreateCapabilities
'new WebActionChain should be instantiated from WebDriver.Actions
'new WebCookies should be instantiated from WebDriver.CreateCookies
'new WebWindow should be instantiated from WebDriver.ActiveWindow or from one of the methods in WebWindows object
'new WebWindows should be instantiated from WebDriver.Windows
'new WebElement should be instantiated from WebDriver.FindElement*
'new WebShadowRoot should be instantiated from WebDriver.GetWebShadowRoot

'@Description("Instantiates a WebDriver object")
Public Function New_WebDriver() As WebDriver
Attribute New_WebDriver.VB_Description = "Instantiates a WebDriver object"
    Set New_WebDriver = New WebDriver
End Function

'@Description("Instantiates a WebElements object")
Public Function New_WebElements() As WebElements
Attribute New_WebElements.VB_Description = "Instantiates a WebElements object"
    Set New_WebElements = New WebElements
End Function

'@Description("Instantiates a WebDriverManager object")
Public Function New_WebDriverManager() As WebDriverManager
Attribute New_WebDriverManager.VB_Description = "Instantiates a WebDriverManager object"
    Set New_WebDriverManager = New WebDriverManager
End Function

'@Description("Instantiates a WebCookie object")
Public Function New_WebCookie() As WebCookie
Attribute New_WebCookie.VB_Description = "Instantiates a WebCookie object"
    Set New_WebCookie = New WebCookie
End Function

'@Description("Instantiates a WebJsonConverter object - this is optional as this object is predeclared")
Public Function New_WebJsonConverter() As WebJsonConverter
Attribute New_WebJsonConverter.VB_Description = "Instantiates a WebJsonConverter object - this is optional as this object is predeclared"
    Set New_WebJsonConverter = New WebJsonConverter
End Function

'@Description("Instantiates a WebKeyboard object - this is optional as this object is predeclared")
Public Function New_WebKeyboard() As WebKeyboard
Attribute New_WebKeyboard.VB_Description = "Instantiates a WebKeyboard object - this is optional as this object is predeclared"
    Set New_WebKeyboard = New WebKeyboard
End Function

'@Description("Instantiates a WebPrintSettings object")
Public Function New_WebPrintSettings() As WebPrintSettings
Attribute New_WebPrintSettings.VB_Description = "Instantiates a WebPrintSettings object"
    Set New_WebPrintSettings = New WebPrintSettings
End Function
