Attribute VB_Name = "test_FileUpDownload"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_file_upload()
    'see https://www.guru99.com/upload-download-file-selenium-webdriver.html
    Dim driver As SeleniumVBA.WebDriver, str As String

    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    driver.StartEdge
    driver.OpenBrowser
    
    str = "<!DOCTYPE html><html><body><div role='button' class='xyz' aria-label='Add food' aria-disabled='false' data-tooltip='Add food'><span class='abc' aria-hidden='true'>icon</span></body></html>"
    
    driver.SaveStringToFile str, ".\snippet.html"
    
    driver.NavigateTo "https://demo.guru99.com/test/upload/"

    driver.Wait 1000
    
    'enter the file path onto the file-selection input field
    driver.FindElement(By.ID, "uploadfile_0").UploadFile ".\snippet.html" 'this is just a special wrapper for sendkeys
    
    driver.Wait 1000

    'check the "I accept the terms of service" check box
    driver.FindElement(By.ID, "terms").Click

    'click the "Submit File" button
    driver.FindElement(By.Name, "send").Click
    
    driver.Wait 1000
            
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_file_download()
    'see https://www.browserstack.com/guide/download-file-using-selenium-python
    Dim driver As SeleniumVBA.WebDriver, caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    driver.StartChrome
    
    driver.DeleteFiles ".\BrowserStack - List of devices to test on*.csv"
    
    Set caps = driver.CreateCapabilities

    'caps.SetPreference "download.default_directory", ".\" 'download to same directory as this excel file
    'caps.SetPreference "download.prompt_for_download", False
    'caps.SetPreference "plugins.always_open_pdf_externally", True 'if its a pdf then bypass the pdf viewer
    
    caps.SetDownloadPrefs ".\"  'this does the above in one line

    driver.OpenBrowser caps
    
    driver.NavigateTo "https://www.browserstack.com/test-on-the-right-mobile-devices"
    driver.Wait 500
    
    driver.FindElementByID("accept-cookie-notification").Click
    driver.Wait 500
    
    driver.FindElementByCssSelector(".icon-csv").ScrollToElement , -150
    driver.Wait 1000
    
    driver.FindElementByCssSelector(".icon-csv").Click
    
    driver.WaitForDownload ".\BrowserStack - List of devices to test on.csv"
            
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_file_download2()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
   
    driver.StartChrome
    
    'set the directory path for saving download to
    Set caps = driver.CreateCapabilities
    caps.SetDownloadPrefs ".\"
    driver.OpenBrowser caps
    
    'delete legacy copy if it exists
    driver.DeleteFiles ".\test.pdf"
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/raw/main/dev/test_files/test.pdf"
    
    driver.WaitForDownload ".\test.pdf"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_download_resource()
    'this test uses the DownloadResource method of the WebElement class to download the src to an img element
    Dim driver As SeleniumVBA.WebDriver
    Dim element As WebElement

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/wiki"
    driver.Wait 1000

    Set element = driver.FindElement(By.cssSelector, "img[alt='SeleniumVBA'")
    
    'if a folder path is specified for fileOrFolderPath, then the saved file inherits the name of the source
    element.DownloadResource srcAttribute:="src", fileOrFolderPath:=".\"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
