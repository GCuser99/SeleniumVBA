# TinySeleniumVBA

A tiny Selenium wrapper written in pure VBA.

[🇬🇧English version is here](https://github.com/uezo/TinySeleniumVBA/blob/main/README.md) : 

[🇧🇷Versão em Português](https://github.com/tdmsoares/TinySeleniumVBA/blob/ReadmeInPortuguese/README.pt.md)


# ✨ 特長

- インストール不要: VBAのみで書かれているので、インストール権限のない人でもすぐにブラウザ自動操作に取り掛かることができます
- 便利なヘルパー機能: FindElement(s)By*、フォームへの値の入出力、クリックやその他便利な機能を提供しています
- オープンな仕様: 基本的にこのラッパーはWebDriverのHTTPクライアントですので、ラッパーの使い方を学ぶことはWebDriverの仕様を知ることと同義です。無駄になるものはありません
https://www.w3.org/TR/webdriver/


# 📦 セットアップ

1. ツール＞参照設定から `Microsoft Scripting Runtime` に参照を通してください

1. `WebDriver.cls`、`WebElement.cls`、`JsonConverter.bas`をプロジェクトに追加してください
    - 最新版 (v0.1.0): https://github.com/uezo/TinySeleniumVBA/archive/v0.1.0.zip

1. WebDriverをダウンロードしてください（ブラウザのメジャーバージョンと同じもの）
    - Edge: https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
    - Chrome: https://chromedriver.chromium.org/downloads

# 🪄 使い方

```vb
Public Sub main()
    ' WebDriverの開始 (Edge)
    Dim Driver As New WebDriver
    Driver.Edge "path\to\msedgedriver.exe"
    
    ' ブラウザを開く
    Driver.OpenBrowser
    
    ' Googleへ移動
    Driver.Navigate "https://www.google.co.jp/?q=selenium"

    ' 検索テキストボックスを取得
    Dim searchInput
    Set searchInput = Driver.FindElement(By.Name, "q")
    
    ' テキストボックスの値を取得
    Debug.Print searchInput.GetValue
    
    ' テキストボックスに値を入力
    searchInput.SetValue "yomoda soba"
    
    ' 検索ボタンのクリック
    Driver.FindElement(By.Name, "btnK").Click
    
    ' 再読み込み - ヘルパーメソッドを提供していない場合でも、ドライバーコマンドを直接実行することができます
    Driver.Execute Driver.CMD_REFRESH
End Sub
```

# ❤️ 謝辞

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) という Tim Hall さんが開発したVBA用JSONコンバーターはHTTPクライアントを作る上でとても役に立ちました。このすばらしいライブラリは当該ライブラリのライセンスのもとでリリースに含まれています。ありがとうございます！
