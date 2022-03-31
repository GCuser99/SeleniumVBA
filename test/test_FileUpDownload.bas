Attribute VB_Name = "test_FileUpDownload"
Sub test_file_upload()
    'see https://www.guru99.com/upload-download-file-selenium-webdriver.html
    Dim Driver As New WebDriver, str
    
    str = "<!DOCTYPE html><html><body><div role='button' class='xyz' aria-label='Add food' aria-disabled='false' data-tooltip='Add food'><span class='abc' aria-hidden='true'>icon</span></body></html>"
    
    filepath = ".\snippet.html"
    
    Driver.StartChrome
    Driver.OpenBrowser
    
    Driver.SaveHTMLToFile str, filepath
    
    Driver.NavigateTo "https://demo.guru99.com/test/upload/"

    Driver.Wait 1000
    
    'enter the file path onto the file-selection input field
    Driver.FindElement(by.ID, "uploadfile_0").UploadFile filepath 'this is just a special wrapper for sendkeys
    
    Driver.Wait 1000

    'check the "I accept the terms of service" check box
    Driver.FindElement(by.ID, "terms").Click

    'click the "Submit File" button
    Driver.FindElement(by.name, "send").Click
    
    Driver.Wait 1000
            
    Driver.CloseBrowser
    Driver.Shutdown
End Sub

Sub test_file_download()
    'see https://www.browserstack.com/guide/download-file-using-selenium-python
    Dim Driver As New WebDriver, caps As Capabilities
    
    filepath = ".\" 'download to same directory as this excel file
    
    Driver.StartChrome
    
    Set caps = Driver.CreateCapabilities

    caps.AddPref "download.default_directory", filepath
    caps.AddPref "download.prompt_for_download", False
    
    'caps.SetDownloadPrefs filepath 'this does the above in one line

    Driver.OpenBrowser caps
    
    Driver.NavigateTo "https://www.browserstack.com/test-on-the-right-mobile-devices"

    Driver.Wait 500
    
    Driver.FindElementByID("accept-cookie-notification").Click
    
    Driver.Wait 500
    
    Driver.FindElementByCssSelector(".icon-csv").ScrollToElement , -150
    
    Driver.Wait 1000
    
    Driver.FindElementByCssSelector(".icon-csv").Click
        
    Driver.Wait 2000
            
    Driver.CloseBrowser
    Driver.Shutdown
End Sub


