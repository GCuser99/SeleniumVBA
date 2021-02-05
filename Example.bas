Attribute VB_Name = "Example"
' TinySeleniumVBA
' A tiny Selenium wrapper written in pure VBA
'
' (c)2021 uezo
'
' Mail: uezo@uezo.net
' Twitter: @uezochan
' https://github.com/uezo/TinySeleniumVBA
'
' ==========================================================================
' セットアップ
'
' 1. ツール＞参照設定で`Microsoft Scripting Runtime`をオンにする
'
' 2. WebDriver.cls, WebElement.cls JsonConverter.bas をプロジェクトに追加
'
' 3. WebDriverをダウンロード（ブラウザのメジャーバージョンと同じもの）
'   - Edge: https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
'   - Chrome: https://chromedriver.chromium.org/downloads
'
' 使い方
'    `WebDriver`のインスタンスをダウンロードしたWebDriverを使って生成します。
'    そこから先は下のExampleを参照ください。
' ==========================================================================

' ==========================================================================
' Setup
'
' 1. Set reference to `Microsoft Scripting Runtime`
'
' 2. Add WebDriver.cls, WebElement.cls and JsonConverter.bas to your VBA Project
'
' 3. Download WebDriver (driver and browser should be the same version)
'   - Edge: https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
'   - Chrome: https://chromedriver.chromium.org/downloads
'
' Usase
'    Create instance of `WebDriver` with the path to the driver you download.
'    See also the example below.
' ==========================================================================


' ==========================================================================
' Example
' ==========================================================================
Option Explicit

Public Sub main()
    ' Start WebDriver (Edge)
    Dim Driver As New WebDriver
    Driver.Edge "path\to\msedgedriver.exe"
    
    ' Open browser
    Driver.OpenBrowser
    
    ' Navigate to Google
    Driver.Navigate "https://www.google.co.jp/?q=selenium"

    ' Get search textbox
    Dim searchInput
    Set searchInput = Driver.FindElement(By.Name, "q")
    
    ' Get value from textbox
    Debug.Print searchInput.GetValue
    
    ' Set value to textbox
    searchInput.SetValue "yomoda soba"
    
    ' Click search button
    Driver.FindElement(By.Name, "btnK").Click
    
    ' Refresh - you can use Execute with driver command even if the method is not provided
    Driver.Execute Driver.CMD_REFRESH
End Sub





