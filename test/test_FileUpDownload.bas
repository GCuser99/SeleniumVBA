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
    
    driver.DeleteFiles ".\file_1.txt"
            
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
    Dim driver As SeleniumVBA.WebDriver
    Dim element As SeleniumVBA.WebElement
    Dim fso As New FileSystemObject

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser

    driver.ImplicitMaxWait = 2000

    'img element with only src attribute
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/wiki"
    Set element = driver.FindElement(By.CssSelector, "img[alt='SeleniumVBA'")
    element.DownloadResource ("src")
    Debug.Assert fso.FileExists(driver.ResolvePath("logo.png"))
    driver.DeleteFiles "logo.png"

    driver.ActiveWindow.Maximize

    'img element with srcset attribute
    driver.NavigateTo "https://mdn.github.io/learning-area/html/multimedia-and-embedding/responsive-images/responsive.html"
    Set element = driver.FindElement(By.CssSelector, "body > main > section > img")
    element.DownloadResource ("srcset")
    Debug.Assert fso.FileExists(driver.ResolvePath("elva-fairy-800w.jpg"))
    driver.DeleteFiles "elva-fairy-800w.jpg"

    'download the default src attribute
    element.DownloadResource ("src")
    Debug.Assert fso.FileExists(driver.ResolvePath("elva-fairy-800w.jpg"))
    driver.DeleteFiles "elva-fairy-800w.jpg"

    'source elements with srcset attributes
    Set element = driver.FindElement(By.CssSelector, "body > main > section > picture > source:nth-child(1)")
    element.DownloadResource ("srcset")
    Debug.Assert fso.FileExists(driver.ResolvePath("elva-480w-close-portrait.jpg"))
    driver.DeleteFiles "elva-480w-close-portrait.jpg"
    
    'source elements with srcset attributes
    Set element = driver.FindElement(By.CssSelector, "body > main > section > picture > source:nth-child(2)")
    element.DownloadResource ("srcset")
    Debug.Assert fso.FileExists(driver.ResolvePath("elva-800w.jpg"))
    driver.DeleteFiles "elva-800w.jpg"

    'img element with only src attribute
    driver.NavigateTo "https://html.com/attributes/img-srcset/"
    Set element = driver.FindElement(By.CssSelector, "#post-1847 > div > div:nth-child(46) > img")
    element.DownloadResource ("src")
    Debug.Assert fso.FileExists(driver.ResolvePath("flamingo-fallback.jpg"))
    driver.DeleteFiles "flamingo-fallback.jpg"
    
    'img element with srcset attribute
    driver.NavigateTo "https://webkit.org/demos/srcset/"
    Set element = driver.FindElement(By.CssSelector, "body > img")
    element.DownloadResource ("srcset")
    Debug.Assert fso.FileExists(driver.ResolvePath("image-4x.png"))
    driver.DeleteFiles "image-4x.png"
    
    'video source element with src attribute
    driver.NavigateTo "https://www.w3schools.com/html/tryit.asp?filename=tryhtml5_video"
    driver.SwitchToFrame driver.FindElement(By.ID, "iframeResult")
    Set element = driver.FindElement(By.CssSelector, "body > video > source:nth-child(1)")
    element.DownloadResource
    Debug.Assert fso.FileExists(driver.ResolvePath("mov_bbb.mp4"))
    driver.DeleteFiles "mov_bbb.mp4"

    driver.CloseBrowser
    driver.Shutdown
End Sub
