Attribute VB_Name = "test_FileUpload"
'see https://www.guru99.com/upload-download-file-selenium-webdriver.html
Sub test_file_upload()
    Dim driver As New WebDriver
    
    filepath = ".\snippet1.html"
    
    driver.Chrome
    driver.OpenBrowser
    
    driver.Navigate "https://demo.guru99.com/test/upload/"

    driver.Wait 1000
    
    'enter the file path onto the file-selection input field
    driver.FindElement(by.ID, "uploadfile_0").UploadFile filepath 'this is just a special wrapper for sendkeys
    
    driver.Wait 1000

    'check the "I accept the terms of service" check box
    driver.FindElement(by.ID, "terms").Click

    'click the "Submit File" button
    driver.FindElement(by.name, "send").Click
    
    driver.Wait 1000
            
    driver.CloseBrowser
    driver.Shutdown
End Sub
