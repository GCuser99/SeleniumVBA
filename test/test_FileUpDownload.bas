Attribute VB_Name = "test_FileUpDownload"
Sub test_file_upload()
    'see https://www.guru99.com/upload-download-file-selenium-webdriver.html
    Dim driver As New WebDriver, str
    
    driver.StartChrome
    driver.OpenBrowser
    
    str = "<!DOCTYPE html><html><body><div role='button' class='xyz' aria-label='Add food' aria-disabled='false' data-tooltip='Add food'><span class='abc' aria-hidden='true'>icon</span></body></html>"
    filePath = ".\snippet.html"
    
    driver.SaveHTMLToFile str, filePath
    
    driver.NavigateTo "https://demo.guru99.com/test/upload/"

    driver.Wait 1000
    
    'enter the file path onto the file-selection input field
    driver.FindElement(by.ID, "uploadfile_0").UploadFile filePath 'this is just a special wrapper for sendkeys
    
    driver.Wait 1000

    'check the "I accept the terms of service" check box
    driver.FindElement(by.ID, "terms").Click

    'click the "Submit File" button
    driver.FindElement(by.name, "send").Click
    
    driver.Wait 1000
            
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_file_download()
    'see https://www.browserstack.com/guide/download-file-using-selenium-python
    Dim driver As New WebDriver, caps As Capabilities
    
    dirPath = ".\" 'download to same directory as this excel file
    
    driver.StartChrome
    
    Set caps = driver.CreateCapabilities

    caps.AddPref "download.default_directory", dirPath
    caps.AddPref "download.prompt_for_download", False
    
    'caps.SetDownloadPrefs filepath 'this does the above in one line

    driver.OpenBrowser caps
    
    driver.NavigateTo "https://www.browserstack.com/test-on-the-right-mobile-devices"

    driver.Wait 500
    
    driver.FindElementByID("accept-cookie-notification").Click
    
    driver.Wait 500
    
    driver.FindElementByCssSelector(".icon-csv").ScrollToElement , -150
    
    driver.Wait 1000
    
    driver.FindElementByCssSelector(".icon-csv").Click
        
    driver.Wait 2000
            
    driver.CloseBrowser
    driver.Shutdown
End Sub


