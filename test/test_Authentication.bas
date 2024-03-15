Attribute VB_Name = "test_Authentication"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_LoginWithSendKeys()
    'an example of how to authenticate using SendKeys
    Dim driver As SeleniumVBA.WebDriver
    Dim userName As String
    Dim pw As String
    
    'substitute your own User Name and Password for this example to work
    userName = "MyUserName"
    pw = "MyPassword"
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 2000
    
    driver.NavigateTo "https://www.vbforums.com"
    
    driver.FindElement(By.CssSelector, "#navbar_username").SendKeys userName
    driver.FindElement(By.CssSelector, "#navbar_password_hint").Click
    driver.FindElement(By.CssSelector, "#navbar_password").SendKeys pw
    driver.FindElement(By.CssSelector, "#logindetails > div > div > input.loginbutton").Click

    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_LoginWithInjectedScript()
    'an example of how to authenticate using a hidden form submital with an injected script
    Dim driver As SeleniumVBA.WebDriver
    Dim params As New Dictionary
    Dim javaScript As String
    Dim userName As String
    Dim pw As String
    
    'substitute your own User Name and Password for this example to work
    userName = "MyUserName"
    pw = "MyPassword"
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser

    javaScript = "function submitLoginForm(usr, pwd){" & vbCrLf
    javaScript = javaScript & "   var f = document.forms[0]" & vbCrLf
    javaScript = javaScript & "       f.elements['vb_login_username'].value= usr" & vbCrLf
    javaScript = javaScript & "       f.elements['vb_login_password'].value= pwd" & vbCrLf
    javaScript = javaScript & "       f.submit()" & vbCrLf
    javaScript = javaScript & "}"
    
    params.Add "source", javaScript
    driver.ExecuteCDP "Page.addScriptToEvaluateOnNewDocument", params
    
    driver.NavigateTo "https://www.vbforums.com/"
    
    params.RemoveAll
    params.Add "expression", "submitLoginForm('" & userName & "','" & pw & "');"
    driver.ExecuteCDP "Runtime.evaluate", params
    
    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_BasicAuthentication()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim userName As String
    Dim pw As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    userName = "admin"
    pw = "admin"
    
    driver.NavigateTo "http://" & userName & ":" & pw & "@the-internet.herokuapp.com/basic_auth"
  
    If driver.IsPresent(By.CssSelector, "#content > div > p", elemFound:=elem) Then
        Debug.Print elem.GetText
    End If
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

'https://www.selenium.dev/documentation/webdriver/bidirectional/chrome_devtools/cdp_endpoint/#basic-authentication
Sub test_CDP_BasicAuthentication()
    Dim driver As SeleniumVBA.WebDriver
    Dim userName As String
    Dim pw As String
    Dim params As Dictionary
    Dim authString As String
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 10000
    
    driver.NavigateTo "https://jigsaw.w3.org/HTTP/"
    
    userName = "guest"
    pw = "guest"
    
    'https://chromedevtools.github.io/devtools-protocol/tot/Network/#method-enable
    driver.ExecuteCDP "Network.enable"

    'build authorization string
    authString = "Basic " & EncodeBase64(userName & ":" & pw)
    
    'build the CDP parameter dictionary
    Set params = New Dictionary
    params.Add "headers", New Dictionary
    params("headers").Add "authorization", authString
    
    'https://chromedevtools.github.io/devtools-protocol/tot/Network/#method-setExtraHTTPHeaders
    driver.ExecuteCDP "Network.setExtraHTTPHeaders", params
    
    driver.FindElement(By.LinkText, "Basic Authentication test").Click
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

'https://stackoverflow.com/questions/169907/how-do-i-base64-encode-a-string-efficiently-using-excel-vba
Private Function EncodeBase64(text As String) As String
    Dim bytes() As Byte
    Dim domDoc As Object
    Dim domElem As Object
    bytes = StrConv(text, vbFromUnicode)
    Set domDoc = CreateObject("MSXML2.DOMDocument")
    Set domElem = domDoc.createElement("b64")
    domElem.DataType = "bin.base64"
    domElem.nodeTypedValue = bytes
    EncodeBase64 = Replace(domElem.text, vbLf, "")
End Function
