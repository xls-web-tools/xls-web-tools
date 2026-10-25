# WebDriver によるブラウザ取得方針

Web情報取得では、Selenium を導入せず、VBA から Edge WebDriver の HTTP API を直接呼び出してブラウザ取得を行う。可能な限り既定でインストールされているソフトウェアと、利用者が所定パスへ配置する `msedgedriver.exe` だけで実現したいからである。

対象システムは URL 直指定ではなく onClick の JavaScript 呼び出しで画面遷移する作りも想定されるため、詳細ページ、一覧復帰、ページングなどの画面遷移は URL を組み立てて直接開かず、WebDriver で利用者操作と同等の画面遷移操作を再現する。

認証情報は Excel ブック、VBA、設定シート、ログに保存しない。認証遷移は Edge 側の Windows 統合認証、Cookie、または利用者の手動ログインに任せ、Web情報取得は専用ブラウザプロファイルを使って通常利用の Edge プロファイルと状態を分離する。

WebDriver 実行ファイルは配布物に同梱しない。入手が容易でインストール不要であり、利用者が `Web情報取得/bin/msedgedriver.exe` に配置すればよいため、既定パスは `ThisWorkbook.Path\bin\msedgedriver.exe` とし、必要に応じて `settings` シートで上書きできるようにする。
