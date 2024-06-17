Attribute VB_Name = "test_Settings"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_settings()
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver

    'this creates a new SeleniumVBA.ini file if one does not exist
    'or refreshes/updates while keeping valid entries of an existing one
    'to set the ini file entries to system default values, use keepExistingValues:=False
    driver.CreateSettingsFile keepExistingValues:=True
End Sub
