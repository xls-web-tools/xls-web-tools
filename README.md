# Web情報取得ツール

`Web情報取得.xlsm` は、Microsoft Edge を自動操作して Web ページの一覧画面から詳細ページを巡回し、取得結果を Excel の `output` シートへ書き出すツールです。ログインが必要なサイトや、JavaScript で画面遷移するサイトの取得を想定しています。

認証情報は Excel ブック、VBA、`settings` シート、ログには保存しません。ただし、サイト側のログイン状態や Cookie は、通常の Edge とは別の専用ブラウザプロファイルに残る場合があります。

## 使う前に

- Windows 版 Excel で `Web情報取得.xlsm` を開き、マクロを有効化します。
- Microsoft Edge を利用できる状態にします。
- Edge WebDriver の `msedgedriver.exe` を用意します。通常は `Web情報取得\bin\msedgedriver.exe` に配置します。
- Edge WebDriver は、利用している Microsoft Edge に対応するものを使います。起動できない場合は、Edge と WebDriver の組み合わせを確認してください。
- 取得先サイトの利用規約、アクセス頻度、認証条件を確認してから実行してください。

## ブックのシート構成

| シート | 用途 |
| --- | --- |
| `output` | 取得結果を書き出します。1 行目はツールが見出しとして使います。 |
| `settings` | 取得先 URL、画面遷移、抽出項目、出力条件を設定します。 |

## 基本の流れ

1. `settings` シートの `Value` 列へ必要な値を入力します。
2. 詳細ページから取得したい項目を、`OutputColumnName` から `BlankMode` までの表へ入力します。
3. まず診断実行ボタンを使い、最初の一覧項目から詳細ページへ移動できることを確認します。内部の実行入口名は `RunDiagnostic_Click` です。
4. 診断が成功したら収集実行ボタンを使い、一覧項目を収集します。内部の実行入口名は `Collect_Click` です。
5. `output` シートで `取得状態`、`エラー内容`、出力列を確認します。

実行ボタンの表示名や配置は、配布時のブック表示に従ってください。問題調査時に入口名を確認する場合は、診断が `RunDiagnostic_Click`、本番収集が `Collect_Click` です。

## settings の基本設定

`settings` シート左側の `SettingName` / `Value` 表に入力します。空欄のままにできる項目もありますが、取得先に合わせて最低限の selector と操作名は設定してください。

| SettingName | 説明 |
| --- | --- |
| `WebDriverPath` | `msedgedriver.exe` のパス。空欄なら `bin\msedgedriver.exe` を使います。 |
| `WebDriverPort` | Edge WebDriver の待ち受けポート。通常は `9515` のままで構いません。 |
| `BrowserProfilePath` | 専用ブラウザプロファイルの保存先。空欄ならブック横の `browser-profile` を使います。 |
| `Headless` | `1` / `True` ならブラウザを非表示で実行します。初回確認やログイン確認では `0` / `False` が便利です。 |
| `StartUrl` | 最初に開く URL。 |
| `OutputSheetName` | 出力先シート名。通常は `output` です。 |
| `AuthenticatedStartSelector` | ログイン後または取得開始画面に到達したことを判定する CSS selector。 |
| `ListPageSelector` | 一覧画面に到達したことを判定する CSS selector。 |
| `ListTransitionOperationName` | 一覧画面へ移動する操作名。右側の画面遷移操作表の `OperationName` と一致させます。 |
| `ListItemSelector` | 一覧画面上の一覧項目を数える CSS selector。 |
| `ListItemTargetIdSelector` | 一覧項目ごとの対象IDを読む CSS selector。`{{index}}` または `{{rowNumber}}` を使えます。 |
| `DetailTransitionOperationName` | 一覧項目から詳細ページへ移動する操作名。 |
| `TargetIdSelector` | 詳細ページ上の対象IDを読む CSS selector。 |
| `ReturnToListOperationName` | 詳細ページから一覧画面へ戻る操作名。 |
| `NextPageOperationName` | 次ページへ移動する操作名。通常は `NextPage` です。 |
| `NextPageAvailableSelector` | 次ページがある場合だけ見つかる CSS selector。空欄なら 1 ページだけ処理します。 |
| `ExistingRowMode` | `SkipExisting` なら既存の `OK` 行をスキップします。`Overwrite` なら既存行も再取得します。 |
| `TimeoutSeconds` | 画面や selector の出現を待つ秒数。通常は `30` です。 |
| `OutputConditionExpression` | 出力対象を絞り込む条件式。不要なら空欄にします。 |

## 画面遷移操作

`settings` シートでは C 列をメモ、D 列を空列、E:J の `OperationName` から `WaitConditionName` までを画面遷移操作表、K 列を空列、L:R を詳細ページ列定義表として使います。

| 列 | 説明 |
| --- | --- |
| `OperationName` | 操作名。基本設定の各 `...OperationName` から参照します。 |
| `LocatorType` | 通常は `css selector` を指定します。`LocatorInnerText` を使う場合は `css selector` が必須です。 |
| `LocatorValue` | クリックする要素の selector。詳細ページへ入る操作では `{{index}}` や `{{rowNumber}}` を使えます。 |
| `LocatorInnerText` | selector の候補が複数ある場合に、画面上の `innerText` でクリック対象を 1 件に絞り込む任意列です。空白正規化後の完全一致で比較します。 |
| `Script` | クリックだけで動かない画面で使う JavaScript。通常は空欄です。 |
| `WaitConditionName` | 操作後の待機条件名。メモ用の列です。 |

frame 内の要素を指定する場合は、`frame selector >> target selector` の形で selector を書きます。`LocatorInnerText` は最後の target selector に一致する候補へ適用されます。

## 詳細ページ列定義

`settings` シート右側の `OutputColumnName`, `Selector`, `ExtractType`, `AttributeName`, `IsRequired`, `ValueExpression`, `BlankMode` の固定順テーブルに、詳細ページから抽出または派生したい列を入力します。`output` へ出したい列は、`output` シート 1 行目の E 列以降に `OutputColumnName` と一致するヘッダーとして入力します。

| 列 | 説明 |
| --- | --- |
| `OutputColumnName` | 詳細ページ列名。重複しない名前にします。 |
| `Selector` | 通常の抽出列で使う、詳細ページ上の取得対象 CSS selector。派生列では空欄にします。 |
| `ExtractType` | `InnerText`、`TextContent`、`Attribute` のいずれか。空欄なら `InnerText` です。 |
| `AttributeName` | `ExtractType` が `Attribute` の場合に読む属性名。 |
| `IsRequired` | 通常の抽出列で `1` / `True` なら必須項目として扱います。派生列では `True` にできません。 |
| `ValueExpression` | 派生列で使う値式。単純な別名は `[元列名]` の形式で指定します。通常の抽出列では空欄にします。 |
| `BlankMode` | 必須項目または派生値が空欄のときにエラーにする場合は `ErrorIfBlank`、許容する場合は `AllowBlank`。 |

通常の抽出列は `Selector` を指定して `ValueExpression` を空欄にします。派生列は `Selector` を空欄にして `ValueExpression` を指定します。両方空欄、または両方指定された非空行は settings エラーです。

条件付きの派生列は `IF([列2] == "", [列1], [列2])` のように指定します。`ValueExpression` に使える値式は、列参照、ダブルクォート文字列リテラル、`IF(条件, 値式, 値式)` です。

## 出力対象条件

`OutputConditionExpression` に条件式を書くと、条件に一致する詳細ページだけを `output` に出します。条件に一致しない場合は、取得失敗ではなく「出力対象外」として扱われ、既存行も削除しません。

使用できる形式は次のとおりです。

```text
[OutputColumnName] == "値"
[OutputColumnName] != "値"
[列1] == "対象" AND [列2] != "完了"
([列1] == "A" OR [列1] == "B") AND NOT [列2] == "除外"
```

列名は、詳細ページ列定義の `OutputColumnName` と完全に一致させます。`output` シートの出力列として指定していない抽出列も条件式から参照できます。

## ダウンロードを使う場合

ファイルのダウンロードも行う場合は、`settings` の左側表へ次の行を追加します。

| SettingName | 説明 |
| --- | --- |
| `DownloadEnabled` | `1` / `True` なら詳細ページ上のファイルをダウンロードします。 |
| `DownloadRequired` | `1` / `True` なら、ダウンロード失敗やファイルなしを `取得状態=ERROR` にします。 |
| `DownloadRootPath` | ダウンロード結果を保存する親フォルダー。 |
| `DownloadLinkSelector` | 詳細ページ上のダウンロードリンクまたはボタンの CSS selector。 |

ダウンロード結果は `DownloadRootPath\対象ID` の下へ保存されます。同名のファイルやフォルダーがすでにある場合は上書きせず、`ダウンロード状態=SKIPPED` とします。ZIP ファイルは展開して保存します。

## output の見方

`output` シートでは、A:D がツール管理の固定管理列です。E 列以降は、1 行目に `settings` の `OutputColumnName` と一致する非空ヘッダーを置いた列だけへ抽出値を書き込みます。空白ヘッダー列と未定義ヘッダー列は変更しません。

| 列 | 説明 |
| --- | --- |
| `対象ID` | 詳細ページ上で確認した対象ID。収集時の主キーです。 |
| `取得状態` | `OK` または `ERROR`。 |
| `エラー内容` | `ERROR` の原因。正常行では空欄です。 |
| `ダウンロード状態` | `DOWNLOADED`、`NO_FILE`、`SKIPPED`、`ERROR` のいずれか。ダウンロード未使用時は空欄です。 |

`Collect_Click` の完了時には、処理ページ数、正常件数、スキップ件数、出力対象外件数、エラー件数が表示されます。

## よくあるエラー

| メッセージ | 確認すること |
| --- | --- |
| `Edge WebDriver 実行ファイルが配置されていません` | `WebDriverPath` のパス、または `bin\msedgedriver.exe` の配置を確認します。 |
| `WebDriver port が使用中です` | 同じポートの WebDriver が起動済みです。実行中の WebDriver を終了するか、`WebDriverPort` を変更します。 |
| `selector が見つかりません` | 対象画面が開いているか、selector が現在の画面構造と一致しているか、`TimeoutSeconds` が短すぎないかを確認します。 |
| `一覧上の対象IDと詳細ページ上の対象IDが一致しません` | 一覧側と詳細側の対象ID selector、または詳細ページへの遷移操作を確認します。 |
| `既存行の取得状態が不正です` | `output` の `取得状態` は `OK` または `ERROR` にします。 |
| `ダウンロードリンクが複数あります` | `DownloadLinkSelector` が 1 件だけに一致するよう調整します。 |

## 注意事項

- 実行中は対象サイトの画面構成やデータが変わらない状態で使ってください。
- `Headless=False` で診断中にエラーが起きた場合、確認用にブラウザが残ることがあります。確認後に閉じてください。
- `output` の 1 行目はツールが見出しとして使うため、削除や並べ替えをしないでください。
- 取得対象サイトへの高頻度アクセスにならないよう、実行前に対象件数と利用ルールを確認してください。
