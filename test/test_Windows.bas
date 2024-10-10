Attribute VB_Name = "test_Windows"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_Selenium_way()
    'this test follows the strategy recommended in Selenium's documentation, using window handles
    'https://www.selenium.dev/documentation/webdriver/interactions/windows/#switching-windows-or-tabs
    Dim driver As SeleniumVBA.WebDriver
    Dim mainHandle As String
    Dim allHandles As Collection
    Dim childHandle As String
    Dim i As Long
    
    Set driver = SeleniumVBA.New_WebDriver
        
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "http://the-internet.herokuapp.com/windows"
        
    'get the handle to the current active window
    mainHandle = driver.ActiveWindow.Handle
        
    'spawn new window
    driver.FindElementByCssSelector("#content > div > a").Click
        
    'note here that main window is still the active one from Selenium's perspective!!
    Debug.Print driver.ActiveWindow.Title 'prints "The Internet"
        
    'now get the collection of all open window handles
    Set allHandles = driver.Windows.Handles
        
    'find the child window's handle by elimination
    For i = 1 To allHandles.Count
        If allHandles(i) <> mainHandle Then
            childHandle = allHandles(i)
            Exit For
        End If
    Next i
        
    'activate the child window and print title
    driver.Windows(childHandle).Activate
    Debug.Print driver.ActiveWindow.Title 'prints "New Window"
    
    driver.Shutdown
End Sub

Sub test_windows_Selenium_way_with_oop_approach()
    'this test follows the strategy recommended in Selenium's documentation, using window objects
    'https://www.selenium.dev/documentation/webdriver/interactions/windows/#switching-windows-or-tabs
    Dim driver As SeleniumVBA.WebDriver
    Dim mainWindow As SeleniumVBA.WebWindow
    Dim window As SeleniumVBA.WebWindow
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "http://the-internet.herokuapp.com/windows"
    
    'get the current active window
    Set mainWindow = driver.ActiveWindow
    
    'spawn a new window
    driver.FindElementByCssSelector("#content > div > a").Click
    
    'note here that main window is still the active one from Selenium's perspective!!
    Debug.Print driver.ActiveWindow.Title 'prints "The Internet"
    
    'find and activate the child window and then print its title
    For Each window In driver.Windows
        If window.IsNotSameAs(mainWindow) Then
            Debug.Print window.Activate.Title 'prints "New Window"
            Exit For
        End If
    Next window
    
    driver.Shutdown
End Sub

Sub test_windows_SwitchToByTitle()
    'this test uses SwitchToTitle to shortcut the finding of the child window,
    'without having to enumerate the windows collection
    Dim driver As SeleniumVBA.WebDriver
    Dim mainWindow As SeleniumVBA.WebWindow
    Dim childWindow As SeleniumVBA.WebWindow
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://the-internet.herokuapp.com/windows"
    
    'get the current active window
    Set mainWindow = driver.ActiveWindow
    
    'spawn a new window
    driver.FindElementByCssSelector("#content > div > a").Click
    
    'note here that main window is still the active one from Selenium's perspective!!
    Debug.Print driver.ActiveWindow.Title 'prints "The Internet"
    
    Set childWindow = driver.Windows.SwitchToByTitle("New Window")
    
    Debug.Print driver.ActiveWindow.Title 'prints "New Window"
    Debug.Print childWindow.Title 'prints "New Window"
    
    driver.Shutdown
End Sub

Sub test_windows_SwitchToByUrl()
    'this test uses SwitchToUrl to shortcut the finding of the child window,
    'without having to enumerate the windows collection
    Dim driver As SeleniumVBA.WebDriver
    Dim mainWindow As SeleniumVBA.WebWindow
    Dim childWindow As SeleniumVBA.WebWindow
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://the-internet.herokuapp.com/windows"
    
    'get the current active window
    Set mainWindow = driver.ActiveWindow
    
    'spawn a new window
    driver.FindElementByCssSelector("#content > div > a").Click

    'note here that main window is still the active one from Selenium's perspective!!
    Debug.Print driver.ActiveWindow.Url 'prints "https://the-internet.herokuapp.com/windows"
    
    Set childWindow = driver.Windows.SwitchToByUrl("https://the-internet.herokuapp.com/windows/new")
    Debug.Print driver.ActiveWindow.Url 'prints "https://the-internet.herokuapp.com/windows/new"
    Debug.Print childWindow.Url 'prints "https://the-internet.herokuapp.com/windows/new"
    
    driver.Shutdown
End Sub

Sub test_windows_SwitchToNext()
    'this test uses SwitchToNext to shortcut the finding of the child window,
    'without having to enumerate the windows collection
    Dim driver As SeleniumVBA.WebDriver
    Dim mainWindow As SeleniumVBA.WebWindow
    Dim childWindow As SeleniumVBA.WebWindow
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "http://the-internet.herokuapp.com/windows"
    
    'get the current active window
    Set mainWindow = driver.ActiveWindow
    
    'spawn a new window
    driver.FindElementByCssSelector("#content > div > a").Click
    
    'note here that main window is still the active one from Selenium's perspective!!
    Debug.Print driver.ActiveWindow.Title 'prints "The Internet"
    
    'switch to the next open window in the collection AFTER the current active window
    Set childWindow = driver.Windows.SwitchToNext
    Debug.Print driver.ActiveWindow.Title 'prints "New Window"
    Debug.Print childWindow.Title 'prints "New Window"
    
    driver.Shutdown
End Sub

Sub test_windows_SwitchToNew()
    Dim driver As SeleniumVBA.WebDriver
    Dim win1 As SeleniumVBA.WebWindow
    Dim win2 As SeleniumVBA.WebWindow
    Dim i As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "http://the-internet.herokuapp.com/windows"
    
    'get the current active window
    Set win1 = driver.ActiveWindow
    
    'open and activate a new Window-type window
    Set win2 = driver.Windows.SwitchToNew(windowType:=svbaWindow)
    'Set win2 = driver.Windows.SwitchToNew(windowType:=svbaTab) 'for Tab-type window
    
    'Note: creating new windowType:=svbaTab while using incognito mode will throw an error
    'see issue https://github.com/GCuser99/SeleniumVBA/issues/56
    'a work-around is to use the following instead:
    'driver.ExecuteScript "window.open('', '_blank')" 'creates the new tab
    'Set win2 = driver.Windows.SwitchToByUrl("about:blank") 'switches to new tab
    
    driver.NavigateTo "http://google.com"
    
    For i = 1 To 5
        Debug.Print win1.Activate.Title
        Debug.Print win2.Activate.Title
    Next i
    
    driver.Shutdown
End Sub

Sub test_windows_CloseIt()
    'SeleniumVBA CloseIt method solves the activate-after-close problem
    'see https://www.selenium.dev/documentation/webdriver/interactions/windows/#closing-a-window-or-tab
    Dim driver As SeleniumVBA.WebDriver
    Dim mainWindow As SeleniumVBA.WebWindow
    Dim childWindow As SeleniumVBA.WebWindow
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "http://the-internet.herokuapp.com/windows"
    
    'get the current active window
    Set mainWindow = driver.ActiveWindow
    
    'spawn a new window
    driver.FindElementByCssSelector("#content > div > a").Click
    
    'note here that main window is still the active one from Selenium's perspective!!
    Debug.Print driver.ActiveWindow.Title 'prints "The Internet"
    
    Set childWindow = driver.Windows.SwitchToNext
    Debug.Print driver.ActiveWindow.Title 'prints "New Window"
    
    childWindow.CloseIt 'this automatically activates the mainWindow upon close
    
    Debug.Print driver.ActiveWindow.Title 'prints "The Internet"
    Debug.Print mainWindow.Title 'prints "The Internet"
    
    driver.Shutdown
End Sub

Sub test_windows_state()
    Dim driver As SeleniumVBA.WebDriver
    Dim winBounds As Dictionary
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    
    'get the current bounds dictionary object
    Set winBounds = driver.ActiveWindow.Bounds
    
    Debug.Print "current window position/size", winBounds("x"), winBounds("y"), winBounds("width"), winBounds("height")
    
    'maximize the window state
    driver.ActiveWindow.Maximize
    
    'get the maximized bounds dictionary object
    Set winBounds = driver.ActiveWindow.Bounds
    
    Debug.Print "maximized window position/size", winBounds("x"), winBounds("y"), winBounds("width"), winBounds("height")
    
    'modify the position and size
    winBounds("y") = 200
    winBounds("height") = winBounds("height") / 2
    
    'set the new window bounds
    Set driver.ActiveWindow.Bounds = winBounds
    
    'these shortcut methods can be used to do above as well
    'driver.ActiveWindow.SetSize height:=winBounds("height") / 2
    'driver.ActiveWindow.SetPosition y:=200
    
    'get the modified bounds dictionary object
    Set winBounds = driver.ActiveWindow.Bounds
    
    Debug.Print "modified window position/size", winBounds("x"), winBounds("y"), winBounds("width"), winBounds("height")
    
    driver.Shutdown
End Sub

Sub test_url_encoding()
    Dim driver As WebDriver
    Dim urlEncoded As String
    Dim urlDecoded As String
    
    Set driver = New WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.mozilla.org/?x=%D1%88%D0%B5%D0%BB%D0%BB%D1%8B"
    
    '****************************************************************************************************
    'test retrieving both the encoded and decoded version of the current url
    urlEncoded = driver.GetCurrentUrl()
    urlDecoded = driver.GetCurrentUrl(decode:=True)
    
    '****************************************************************************************************
    'test if IsPageFound is encoding agnostic
    Debug.Print "is page found using decoded url: " & driver.IsPageFound(urlDecoded)
    Debug.Print "is page found using encoded url: " & driver.IsPageFound(urlEncoded)
    
    '****************************************************************************************************
    'spawn a new window
    driver.Windows.SwitchToNew svbaTab

    Debug.Print "the active window's encoded url: " & driver.ActiveWindow.Url
    
    '****************************************************************************************************
    'test if SwitchToByUrl is encoding agnostic and test Window.Url method
    driver.Windows.SwitchToByUrl urlDecoded
    
    Debug.Print "the active window's encoded url: " & driver.ActiveWindow.Url()
    Debug.Print "the active window's decoded url: " & driver.ActiveWindow.Url(decode:=True)
    
    driver.Windows.SwitchToByUrl "about:blank"
    driver.Windows.SwitchToByUrl urlEncoded
    
    Debug.Print "the active window's encoded url: " & driver.ActiveWindow.Url()
    Debug.Print "the active window's decoded url: " & driver.ActiveWindow.Url(decode:=True)
    
    '****************************************************************************************************
    'test Windows.Urls method
    Dim urlCol As Collection, urlString As Variant
    
    Set urlCol = driver.Windows.Urls()
    For Each urlString In urlCol
        Debug.Print "encoded window url: " & urlString
    Next urlString
    
    Set urlCol = driver.Windows.Urls(decode:=True)
    For Each urlString In urlCol
        Debug.Print "decoded window url: " & urlString
    Next urlString
    
    driver.Shutdown
End Sub
