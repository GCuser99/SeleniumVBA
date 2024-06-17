'this checks if driver is installed, or if installed driver is compatibile
'with installed browser, and then if needed, installs an updated driver
    
Set mngr = CreateObject("SeleniumVBA.WebDriverManager")

'defaults to specification in SeleniumVBA.ini settings file, or in absence of settings,
'file specification, then the %USERPROFILE%\Downloads dir
'mngr.DefaultDriverFolder = [your binary folder path here]

MsgBox mngr.AlignEdgeDriverWithBrowser()
MsgBox mngr.AlignChromeDriverWithBrowser()
MsgBox mngr.AlignFirefoxDriverWithBrowser()
MsgBox mngr.AlignIEDriverWithBrowser()