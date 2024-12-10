Attribute VB_Name = "test_FileUpDownload"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_file_upload()
    Dim driver As SeleniumVBA.WebDriver

    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    driver.StartEdge
    driver.OpenBrowser
    
    driver.SaveStringToFile "Hello World", ".\file_1.txt"
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/upload.html"

    driver.Wait 1000
    
    'enter the file path onto the file-selection input field
    driver.FindElement(By.CssSelector, "#upload").UploadFile ".\file_1.txt" 'this is just a special wrapper for sendkeys
    
    driver.Wait 1000

    'click the "Go" submit button
    driver.FindElement(By.CssSelector, "#go").Click
    
    driver.Wait 1000
            
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_file_download()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)

    driver.StartChrome
    
    Set caps = driver.CreateCapabilities

    'caps.SetPreference "download.default_directory", ".\" 'download to same directory as this excel file
    'caps.SetPreference "download.prompt_for_download", False
    'caps.SetPreference "plugins.always_open_pdf_externally", True 'if its a pdf then bypass the pdf viewer
    
    'this does the above in one line
    caps.SetDownloadPrefs downloadFolderPath:=".\", promptForDownload:=False, disablePDFViewer:=True

    driver.OpenBrowser caps
    
    'driver.SetDownloadFolder ".\" 'for Edge and Chrome only - no need to set in capabilities
        
    driver.NavigateTo "https://www.selenium.dev/selenium/web/downloads/download.html"
    driver.Wait 500
    
    'driver.FindElementByID("accept-cookie-notification").Click
    'driver.Wait 500
    
    driver.DeleteFiles ".\file_1.txt", ".\file_2.jpg"
    
    driver.FindElementByCssSelector("#file-1").Click
    driver.WaitForDownload ".\file_1.txt"
    
    driver.FindElementByCssSelector("#file-2").Click
    driver.WaitForDownload ".\file_2.jpg"
    
    driver.DeleteFiles ".\file_1.txt", ".\file_2.jpg"
            
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
    caps.SetDownloadPrefs downloadFolderPath:=".\", promptForDownload:=False, disablePDFViewer:=True
    driver.OpenBrowser caps
    
    'delete legacy copy if it exists
    driver.DeleteFiles ".\test.pdf"
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/raw/main/dev/test_files/test.pdf"
    
    driver.WaitForDownload ".\test.pdf"
    
    driver.DeleteFiles ".\test.pdf"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_download_resource()
    'this test uses the DownloadResource method of the WebElement class to download the src to an img element
    Dim driver As SeleniumVBA.WebDriver
    Dim element As SeleniumVBA.WebElement

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/wiki"
    driver.Wait 1000

    Set element = driver.FindElement(By.CssSelector, "img[alt='SeleniumVBA'")
    
    'if a folder path is specified for fileOrFolderPath, then the saved file inherits the name of the source
    element.DownloadResource srcAttribute:="src", fileOrFolderPath:=".\"
    
    driver.DeleteFiles ".\logo.png"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
