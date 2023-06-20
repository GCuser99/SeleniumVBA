Attribute VB_Name = "test_Alerts"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_Alerts()
    'see https://www.guru99.com/alert-popup-handling-selenium.html
    Dim driver As SeleniumVBA.WebDriver

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser

    driver.NavigateTo "http://demo.guru99.com/test/delete_customer.php"
    
    driver.Wait 1000
    
    Debug.Print "Is Alert Present: " & driver.IsAlertPresent
                                
    driver.FindElement(By.Name, "cusid").SendKeys "87654"
    
    driver.Wait 1000
    
    driver.FindElement(By.Name, "submit").Click
    
    driver.Wait 1000
    
    Debug.Print "Is Alert Present: " & driver.IsAlertPresent
    Debug.Print "Alert Text: " & driver.GetAlertText
    If driver.IsAlertPresent Then driver.AcceptAlert
    
    Debug.Print "Alert Text: " & driver.GetAlertText
    If driver.IsAlertPresent Then driver.AcceptAlert

    driver.Wait 1000
    driver.CloseBrowser
    driver.Shutdown
End Sub


