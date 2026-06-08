# Web情報取得 Product Spec

## Problem

Web情報取得では、Windows 11 上の Excel 2024 マクロから、認証を伴う Web システムの一覧画面と詳細ページを巡回し、詳細ページ 1 件をワークシート 1 行として出力したい。
対象環境にはインストール制限があり、Excel、Microsoft Edge、利用者が配置する Edge WebDriver 以外のソフトウェアやライブラリ導入は避ける。

対象システムは認証遷移後に一覧画面へ進み、一覧項目から詳細ページへ移動する。画面遷移は URL 直指定ではなく onClick の JavaScript 呼び出しで行われ、詳細ページから一覧画面へ戻る場合もサイト提供の一覧復帰リンクを使う必要がある。

## Goals

- VBA から Edge WebDriver の HTTP API を直接呼び出してブラウザ取得を行う。
- Selenium などの追加ライブラリを使わない。
- 認証情報を Excel ブック、VBA、settings シート、ログに保存しない。
- 専用ブラウザプロファイルを使い、通常利用の Edge プロファイルと状態を分離する。
- 詳細ページ 1 件から 1 つの出力行を作る。
- 初期実装では診断モードを先に作り、先頭 1 件の詳細ページ取得と一覧復帰までを確認する。

## Non-Goals

- WebDriver 実行ファイルを配布物へ同梱しない。
- WebDriver をツールからダウンロードしない。
- Selenium を導入しない。
- 詳細ページ URL を組み立てて直接開く方式を採用しない。
- 認証用の ID、password、loginconfirm URL の値を保存、解析、再利用しない。
- 初期診断モードでは全件巡回とページングを実装しない。

## Constraints

- 利用者は各自で `msedgedriver.exe` を入手し、既定では `Web情報取得/bin/msedgedriver.exe` に配置する。
- WebDriver パスは settings シートで上書きできる。
- 専用ブラウザプロファイルの既定パスは `ThisWorkbook.Path\browser-profile` とする。
- `Web情報取得/bin/` と `Web情報取得/browser-profile/` は Git 追跡対象にしない。
- Edge は settings の `Headless` で表示/不可視を切り替える。既定値は `True` とし、不可視ブラウザで起動する。
- 初回認証や調査では `Headless=False` に切り替えて可視実行できるようにする。

## Domain Flow

1. 取得開始 URL に遷移する。
2. 認証遷移を経て、認証後の起点ページに到達したことを確認する。
3. 画面遷移操作で一覧画面へ進む。一覧画面と認証後の起点ページが同一になる場合も許容する。
4. 一覧画面上の一覧項目を表示順に扱う。
5. 一覧項目の詳細遷移操作を WebDriver の click で再現し、詳細ページへ入る。
6. 詳細ページで対象IDを取得し、設定された列定義に基づいて出力行を作る。
7. 一覧復帰リンクの画面遷移操作で一覧画面へ戻る。
8. 全件収集では、現在ページの全一覧項目を処理後、次ページ操作でページングする。

## Settings Sheet

シート名は `settings` とする。
Web情報取得の VBA では、他ツールと同様に `IToolSettings`、`ToolSettings`、`ToolSettingsTestDouble` の構造を使う。

単一値設定の候補:

- `WebDriverPath`: WebDriver 実行ファイルの上書きパス。空欄の場合は `ThisWorkbook.Path\bin\msedgedriver.exe` を使う。
- `WebDriverPort`: Edge WebDriver を起動する port。
- `BrowserProfilePath`: 専用ブラウザプロファイルの上書きパス。空欄の場合は `ThisWorkbook.Path\browser-profile` を使う。
- `Headless`: `True` の場合は不可視ブラウザ、`False` の場合は可視ブラウザで起動する。既定値は `True`。
- `StartUrl`: 取得開始 URL。
- `AuthenticatedStartSelector`: 認証後の起点ページ到達を判定する selector。
- `ListPageSelector`: 一覧画面到達を判定する selector。
- `ListTransitionOperationName`: 一覧画面へ進む画面遷移操作名。
- `ListItemSelector`: 現在ページの一覧項目を数える selector。
- `ListItemTargetIdSelector`: 一覧上の対象IDを読み取る selector。本番収集では `{{index}}` と `{{rowNumber}}` の置換トークンを使える。
- `DetailTransitionOperationName`: 一覧項目から詳細ページへ入る画面遷移操作名。
- `TargetIdSelector`: 詳細ページ上の対象IDを読み取る selector。
- `OutputSheetName`: 取得結果の出力先シート名。
- `ExistingRowMode`: 再実行時の既存行扱い。既定値は `SkipExisting`。`SkipExisting` は既存 OK 行をスキップし、`RetryError` は既存 ERROR 行だけ再試行し、`Overwrite` は既存状態に関係なく全更新する。
- `TimeoutSeconds`: 画面到達や selector 出現を待つ秒数。

画面遷移操作定義の候補:

- 操作名。
- locator 種別。
- locator 値。
- JavaScript 関数名または script 本文。click が難しい場合だけ使う。
- 待機条件名。

CSS selector で iframe / frame 内の要素を指定する場合は、`frame selector >> target selector` の形式を使う。具体的な selector 値は取得対象の環境に依存するため、ドキュメントや配布テンプレートには実サイトの値を残さず、利用者が settings に入力する。settings テンプレートには `OpenList`、`OpenDetail`、`ReturnToList` の操作名だけを用意し、locator 値や selector 値は空欄にする。

本番収集で一覧項目ごとに変わる selector / script には、0-based の `{{index}}` と 1-based の `{{rowNumber}}` を使える。`{{index}}` は `querySelectorAll(...)[{{index}}]` などに使い、`{{rowNumber}}` は CSS の `nth-child({{rowNumber}})` などに使う。

詳細ページ列定義の候補:

- 出力列名。
- selector。
- 取得種別。`TextContent`、`InnerText`、`Attribute` など。
- attribute 名。
- 必須かどうか。
- 空欄時の扱い。

## Output Sheet

固定管理列:

- `詳細ページID`: 一覧画面の JavaScript 呼び出しなどに渡される、画面遷移操作用の識別値。取得できる場合だけ保持する補助識別子とし、取得できない場合は対象IDを処理基準にする。
- `対象ID`: 詳細ページ上に表示される、出力行の対象を業務上識別する値。出力行の正となる識別値として扱う。
- `取得状態`: `OK` または `ERROR`。
- `エラー内容`: 詳細ページ単位の取得失敗や必須項目欠落の内容。

固定管理列の後ろに、settings の詳細ページ列定義に従う取得列を並べる。

## Diagnostic Mode

`RUN_DIAGNOSTIC` は、一覧画面の先頭 1 件だけを処理する。
診断モードでは、認証後の起点ページ到達、一覧画面到達、詳細遷移操作、対象ID取得、DOM 抽出、一覧復帰リンクによる復帰を確認する。

診断モードの出力先は本番収集と同じ出力先シートとする。
診断モードの再実行時は、同じ `対象ID` の既存行を上書きする。`詳細ページID` が取得できる場合は補助的に照合または記録する。
ページングは診断モードでは扱わない。

## Collection Mode

`RUN_COLLECT` は、全件収集を行う。
一覧画面の表示順に一覧項目を処理し、各詳細ページを 1 行として出力する。
Issue #8 の slice では現在ページだけを処理し、ページングは扱わない。後続 slice では、現在ページの未処理項目がなくなったら次ページ操作の画面遷移操作を実行し、次ページがなくなるまで繰り返す。

現在ページ到達時に `ListItemSelector` の件数を使って、一覧上の対象IDを表示順に読み取り、現在ページ開始時点の対象ID一覧として保持する。`ListItemSelector` の件数はこの初期一覧を作るための loop bound とし、一覧復帰後のズレ検知には使わない。ページサイズが固定される場合、項目数だけでは新規項目挿入や表示順変化を検知できないため。

再実行時の既定動作は `SkipExisting` とする。
`対象ID` を主キーとして既存行を判定し、既存行がある場合はその行を更新し、未登録の場合は末尾へ追加する。
`取得状態=OK` の既存 `対象ID` は `SkipExisting` ではスキップする。`RetryError` では `取得状態=ERROR` の既存行を再試行し、`Overwrite` では既存状態に関係なく全更新する。`詳細ページID` が取得できる場合は補助的な再開候補として扱う。

## Waiting And Error Handling

画面到達判定は URL より selector を優先する。
詳細ページ到達判定では、対象ID selector の出現を必須条件とする。一覧上の対象IDと詳細ページ上の対象IDが両方取れる場合、診断モードでは一致を必須とし、不一致は失敗扱いにする。本番収集では不一致を詳細ページ単位の ERROR として出力し、一覧復帰できる場合は次の一覧項目へ進む。
ページング後は一覧画面 selector の出現に加え、可能であればページ番号または一覧先頭の詳細ページID変化を確認する。

詳細ページの必須列が取れない場合は、その件を `ERROR` の出力行として残し、一覧復帰リンクで戻れる場合は次の一覧項目へ進む。ERROR 行の `対象ID` は、詳細ページ上の対象IDを読めた場合はその値を使い、読めない場合は一覧上の対象IDを使う。
本番収集では、各一覧項目を処理する直前に現在 index の一覧上の対象IDを再評価し、現在ページ開始時点で保持した同じ index の対象IDと一致することを確認する。一致しない場合は、一覧画面が収集中に変化したものとして本番収集を中断する。
一覧復帰に失敗した場合、WebDriver session が壊れた場合、認証切れで復帰不能な場合は処理を中断する。

## Implementation Sequence

1. `settings` シート読み取りと `IToolSettings`、`ToolSettings`、`ToolSettingsTestDouble` を用意する。
2. Edge WebDriver の起動、session 作成、session 終了を実装する。
3. 認証後の起点ページ到達と一覧画面到達を確認する。
4. 一覧画面の先頭 1 件に対して詳細遷移操作を再現する。
5. 対象IDと列定義に基づく 1 行分の DOM 抽出を実装する。
6. 一覧復帰リンクで一覧画面へ戻る。
7. `RUN_DIAGNOSTIC` を完成させる。
8. `RUN_COLLECT`、全件巡回、ページング、既存行スキップを後続実装する。

## References

- `docs/adr/0001-webdriver-browser-collection-policy.md`
- `CONTEXT.md`
