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
- 診断モードでは、先頭 1 件の詳細ページ取得、出力判定、ダウンロード処理、一覧復帰までを確認する。

## Non-Goals

- WebDriver 実行ファイルを配布物へ同梱しない。
- WebDriver をツールからダウンロードしない。
- Selenium を導入しない。
- 詳細ページ URL を組み立てて直接開く方式を採用しない。
- 認証用の ID、password、loginconfirm URL の値を保存、解析、再利用しない。

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
3. `ListTransitionOperationName` の画面遷移操作で一覧画面へ進む。一覧画面と認証後の起点ページが同一になる場合は、副作用のない script 操作を設定し、一覧到達判定へ進める。
4. 一覧画面上の一覧項目を表示順に扱う。
5. 一覧項目の詳細遷移操作を WebDriver の click で再現し、詳細ページへ入る。
6. 詳細ページで対象IDを取得し、設定された列定義に基づいて出力行を作る。
7. 一覧復帰リンクの画面遷移操作で一覧画面へ戻る。
8. 全件収集では、現在ページの全一覧項目を処理後、次ページ操作でページングする。

## Settings Sheet

シート名は `settings` とする。
Web情報取得の VBA では、他ツールと同様に `IToolSettings`、`ToolSettings`、`ToolSettingsTestDouble` の構造を使う。

単一値設定:

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
- `ReturnToListOperationName`: 詳細ページから一覧画面へ戻る画面遷移操作名。診断モードと本番収集の両方で必須とする。
- `NextPageOperationName`: 現在ページの処理後に次ページへ進む画面遷移操作名。未指定または空欄の場合は `NextPage` を使う。
- `NextPageAvailableSelector`: 次ページ操作が有効な場合だけ一致する selector。未指定または空欄の場合はページングせず、現在ページだけで正常終了する。0 件の場合は次ページなしとして正常終了する。
- `OutputSheetName`: 取得結果の出力先シート名。
- `ExistingRowMode`: 再実行時の既存行扱い。既定値は `SkipExisting`。有効値は `SkipExisting` と `Overwrite` とする。`SkipExisting` は既存 OK 行だけをスキップし、未登録行と既存 ERROR 行を取得する。`Overwrite` は既存状態に関係なく全更新する。`RetryError` は旧称であり、有効な設定値として扱わない。
- `OutputConditionExpression`: 詳細ページから取得した値に基づき、出力行を作る対象かどうかを判定する任意の出力対象条件。空欄または未定義の場合は全件を出力対象とする。
- `DownloadEnabled`: `True` の場合だけ、詳細ページ上のダウンロードリンクからダウンロード対象ファイルを取得する。既定値は `False`。
- `DownloadRequired`: `True` の場合は、`ダウンロード状態=NO_FILE` または `ダウンロード状態=ERROR` を詳細ページ単位の `ERROR` とする。既定値は `False`。
- `DownloadRootPath`: ダウンロード保存先ルート。`DownloadEnabled=True` の場合は必須とし、この配下に対象ID名のフォルダーを作成する。
- `DownloadLinkSelector`: 詳細ページ上のダウンロードリンクを特定する selector。`DownloadEnabled=True` の場合は必須とし、`frame selector >> target selector` 形式を使える。
- `TimeoutSeconds`: 画面到達、selector 出現、ページング後の先頭対象ID変化、ダウンロード完了を待つ秒数。

画面遷移操作定義:

- 操作名。
- locator 種別。
- locator 値。
- JavaScript 関数名または script 本文。click が難しい場合、または画面遷移が不要な場合に使う。
- 待機条件名。実行時の待機には使わず、設定上の識別またはメモとして扱う。

画面遷移操作は `OperationName` で参照する。`LocatorType` と `LocatorValue` がある場合は WebDriver click を先に実行し、click が失敗して `Script` が指定されている場合だけ script 実行へ切り替える。`LocatorType` が空で `Script` が指定されている場合は script だけを実行する。`OperationName` は必須であり、`LocatorType` または `Script` の少なくとも一方を指定する。`LocatorType` を指定する場合は `LocatorValue` も必須とする。`OperationName` の照合は大文字小文字を区別しない。

CSS selector で iframe / frame 内の要素を指定する場合は、`frame selector >> target selector` の形式を使う。`>>` で区切った途中要素は frame selector、最後の要素は target selector として扱う。frame selector と target selector は空にできない。具体的な selector 値は取得対象の環境に依存するため、ドキュメントや配布テンプレートには実サイトの値を残さず、利用者が settings に入力する。settings テンプレートには `OpenList`、`OpenDetail`、`ReturnToList`、`NextPage` の操作名だけを用意し、locator 値や selector 値は空欄にする。

本番収集で一覧項目ごとに変わる selector / script には、0-based の `{{index}}` と 1-based の `{{rowNumber}}` を使える。`{{index}}` は `querySelectorAll(...)[{{index}}]` などに使い、`{{rowNumber}}` は CSS の `nth-child({{rowNumber}})` などに使う。

詳細ページ列定義:

- 出力列名。
- selector。
- 取得種別。`TextContent`、`InnerText`、`Attribute` など。
- attribute 名。
- 必須かどうか。
- 空欄時の扱い。
- 詳細ページ列定義は全列が空欄の行だけを空行として無視する。`OutputColumnName` と `Selector` は必須とし、どちらかが空欄の非空行は settings エラーとして扱う。

## Output Sheet

固定管理列:

- `対象ID`: 詳細ページ上に表示される、出力行の対象を業務上識別する値。出力行の正となる識別値として扱う。
- `取得状態`: `OK` または `ERROR`。
- `エラー内容`: 詳細ページ単位の取得失敗や必須項目欠落の内容。
- `ダウンロード状態`: ダウンロード対象ファイルの取得結果。`DOWNLOADED`、`NO_FILE`、`SKIPPED`、`ERROR` のいずれかとする。`DownloadEnabled=False` の場合は空欄とする。

固定管理列の後ろでは、output シート 1 行目の E 列以降について、空文字列を無視した 1 行目の used range を取得し、その範囲内の非空ヘッダーのうち、settings の詳細ページ列定義に存在する `OutputColumnName` と大文字小文字を区別して一致する列を出力列として扱う。E 列以降の非空ヘッダーの並び順を出力列順とし、settings の詳細ページ列定義に存在しても非空ヘッダーに指定されていない抽出列は output シートへ出力しない。E 列以降に出力列がない場合も有効な設定とし、output シートには固定管理列だけを出力する。1 行目が空白の列、または settings の詳細ページ列定義に存在しない未定義ヘッダー列は処理対象外とし、値も数式も入っていない場合は空のまま維持し、値または数式が入っている場合も変更しない。非空ヘッダーが重複している場合は、同じ抽出値を該当するすべての列へ出力する。非空ヘッダーが固定管理列名と同じ場合も、固定管理列と E 列以降の出力列の両方へ出力する。出力列ではない抽出列が必須列エラーなどの原因になった場合も、その列自体は出力せず、固定管理列の `取得状態=ERROR` と `エラー内容` で原因を示す。抽出自体は出力列指定の有無では省略しない。実行時に A-D の固定管理列は正規化するが、E 列以降の出力列ヘッダーは settings から自動生成しない。

出力対象条件に一致しない詳細ページは出力対象外とし、output シートには新規行を書かず、ERROR 行にもせず、既存行があっても削除しない。条件式内の列参照は `[OutputColumnName]` の形式とし、論理演算子は `AND`、`OR`、`NOT` を使う。
サポートする比較演算子は `==` と `!=` のみとする。
条件式の評価は、詳細ページ列定義の抽出値を取得した後、必須項目エラー判定より先に行う。条件式に必要な値が取れて条件不一致になった場合は、他の必須列が空でも出力対象外とし、`ERROR` 行を作らない。条件式で参照している列自体を selector 欠落、frame 欠落、attribute 欠落などで読めない場合は、その列の抽出値を空文字列として条件判定する。ただし WebDriver session や script 実行そのものの失敗は、従来どおり詳細ページ単位の `ERROR` とする。条件が true になった後の通常の詳細列取得では、出力列として選ばれているかどうかに関係なく、元の `IsRequired` と `BlankMode` に従う。
条件式で参照できる列は、詳細ページ列定義の `OutputColumnName` に限定する。列参照の照合は大文字小文字を区別する。出力列として選ばれていない抽出列も条件式から参照できる。固定管理列の `対象ID`、`取得状態`、`エラー内容`、`ダウンロード状態` 自体は条件式から参照しない。同名の `OutputColumnName` を持つ詳細ページ列定義がある場合、`[対象ID]` などの列参照は固定管理列ではなく、その抽出列を参照する。
文字列比較は case-sensitive とする。比較時は、抽出値と条件式内の文字列リテラルの前後空白を除去してから `==` または `!=` を評価する。
文字列リテラルはダブルクォートのみを正式構文とする。シングルクォートによる文字列リテラルは扱わない。文字列リテラル内のエスケープシーケンスは、既存の `pJsonUnescape` と同じく `\"`、`\\`、`\n`、`\r`、`\t` を扱う。未知のエスケープシーケンス、末尾の `\`、`\uXXXX` は設定誤りとして扱う。
論理演算子の優先順位は、括弧、`NOT`、`AND`、`OR` の順とする。迷う条件式では括弧で明示する。
`OutputConditionExpression` は診断モードと本番収集の両方に適用する。診断モードの先頭 1 件が出力対象外になった場合も output シートには行を書かず、診断結果として出力対象外であることを利用者へ示す。
本番収集の完了メッセージには、出力対象外件数を含める。output シートには出力対象外を表す固定管理列や状態値を追加しない。
既存行がある対象が出力対象外になった場合は、`ExistingRowMode` が `Overwrite` であっても既存行を残し、更新も削除もしない。既存 `取得状態=ERROR` の行を `SkipExisting` で再取得した結果、出力対象条件に一致しなかった場合も、既存 ERROR 行を更新も削除もせず残し、その実行では出力対象外件数として数える。
詳細ページ列定義で同じ文字列の `OutputColumnName` が複数定義されている場合は、`OutputConditionExpression` の有無や output シートの出力列指定に関係なく settings エラーとする。大文字小文字だけが異なる `OutputColumnName` は別の抽出列として扱い、settings エラーにしない。すべての抽出列で `OutputColumnName` は必須とする。`OutputColumnName` が固定管理列名の `対象ID`、`取得状態`、`エラー内容`、`ダウンロード状態` と一致する場合も settings エラーにはしない。
条件式が不正な場合は、収集開始前の settings エラーとして中断する。未定義の `[OutputColumnName]`、括弧不整合、演算子不足、未知のエスケープシーケンスは、詳細ページ処理中の ERROR 行ではなく設定誤りとして扱う。`OutputColumnName` や `Selector` が空の詳細ページ列定義行も settings エラーとして中断する。

## Download File Handling

`DownloadEnabled=True` の場合、Web情報取得は詳細ページ 1 件につき 0 または 1 つのダウンロード対象ファイルを扱う。ダウンロードリンクは `DownloadLinkSelector` で特定し、URL を VBA 側で組み立てない。A 要素などでブラウザ解決済みの `http://` または `https://` の `href` が取得できる場合は認証済み Edge session 上でその URL へ遷移してダウンロードを開始し、それ以外の場合は認証済み Edge session 上のクリックとして実行する。selector が 0 件の場合は `ダウンロード状態=NO_FILE` とする。selector が 2 件以上に一致した場合は設定誤りまたはページ構造の想定外として `ダウンロード状態=ERROR` とする。

Edge の実ダウンロード先は、`FileSystemService` が返す OS のユーザー一時フォルダー配下に実行ごとの一時ダウンロード領域を作って指定する。利用者が settings で指定するのは最終保存先の `DownloadRootPath` だけとする。一時フォルダー場所取得は `xls-common-devtools` 側の `FileSystemService` API を前提とする。

ダウンロード完了待ちは、一時ダウンロード領域に新しいファイルが現れ、`.crdownload` が消えたことを `TimeoutSeconds` まで待つ。完了ファイルが 1 件に確定しない場合、複数ファイルが増えた場合、または待機 timeout の場合は `ダウンロード状態=ERROR` とする。

最終保存先は `{DownloadRootPath}\{対象ID}\` とする。対象IDは詳細ページ上の表示値を `Trim` した文字列をそのままフォルダー名に使う。対象IDが空、Windows フォルダー名で使えない文字または制御文字を含む、末尾がピリオドまたは空白である、または Windows 予約名である場合は、その詳細ページのダウンロード処理を `ERROR` とする。対象IDの自動置換やサニタイズは行わない。

通常ファイルの場合、ブラウザが決めた既定ファイル名の先頭に、対象ID文字列の末尾 3 文字と `_` を付けて保存する。対象IDが 3 文字未満の場合は対象ID全体を接頭辞に使う。同名ファイルが既に存在する場合は上書きせず、`ダウンロード状態=SKIPPED` とする。`ExistingRowMode=Overwrite` でも既存ファイルは上書きしない。

ダウンロードしたファイルの拡張子が `.zip` の場合は、最終出力先に ZIP ファイル自体を残さず、Windows 標準の `Shell.Application` ZIP namespace を使って展開結果を出力する。ZIP 直下が 1 フォルダーだけの場合は、そのフォルダー自体は出力せず、中身を `{DownloadRootPath}\{対象ID}\` 直下へ出力する。ZIP 直下が複数ファイルまたは複数フォルダーの場合は、その構造を対象IDフォルダー直下へ出力する。

ZIP 展開後、最終的に `{DownloadRootPath}\{対象ID}\` 直下へ置かれるファイルだけに、対象ID末尾 3 文字と `_` の接頭辞を付ける。フォルダー名とサブフォルダー内のファイル名には接頭辞を付けない。ZIP が空、または flatten 後に出力対象となるファイルやフォルダーが 0 件の場合は `ダウンロード状態=ERROR` とする。`__MACOSX` や `.DS_Store` などの除外設定は扱わない。

ZIP 展開結果に同名ファイルまたは同名フォルダーが既に存在する場合は、最終出力先へ反映せず `ダウンロード状態=SKIPPED` とする。通常ファイルと同じく既存ファイルや既存フォルダーを上書きしない。

`DownloadRequired=False` の場合、ダウンロードリンクなしやダウンロード失敗は詳細ページ全体の `取得状態` を `ERROR` にしない。`エラー内容` もダウンロード失敗理由では更新しない。`DownloadRequired=True` の場合だけ、`ダウンロード状態=NO_FILE` または `ダウンロード状態=ERROR` を詳細ページ単位の `ERROR` とし、`エラー内容` に理由を記録する。`ダウンロード状態=SKIPPED` は既存ファイルまたは既存フォルダー保護の結果であり、`DownloadRequired=True` でも詳細ページ単位の `ERROR` にはしない。

## Diagnostic Mode

`RunDiagnostic_Click` は、一覧画面の先頭 1 件だけを処理する。
診断モードでは、認証後の起点ページ到達、一覧画面到達、詳細遷移操作、対象ID取得、DOM 抽出、ダウンロード対象ファイルの取得と保存、一覧復帰リンクによる復帰を確認する。`DownloadEnabled=True` の場合、診断モードでも対象IDフォルダーと保存ファイルまたは ZIP 展開結果が実際に作られる。

診断モードの出力先は本番収集と同じ出力先シートとし、output シートのヘッダーによる列構成も本番収集と同じにする。詳細ページ列定義が 0 件の場合も、固定管理列だけの診断出力とダウンロード確認を有効とする。
診断モードの再実行時は、同じ `対象ID` の既存行を上書きする。
ページングは診断モードでは扱わない。

## Collection Mode

`Collect_Click` は、全件収集を行う。
一覧画面の表示順に一覧項目を処理し、各詳細ページを 1 行として出力する。`DownloadEnabled=True` の場合、各詳細ページの処理内でダウンロード対象ファイルの取得と保存も行う。
現在ページの未処理項目がなくなったら `NextPageAvailableSelector` で次ページ操作の有効性を確認し、有効な場合だけ `NextPageOperationName` の画面遷移操作を実行する。`NextPageAvailableSelector` が空欄または 0 件の場合は次ページなしとして正常終了する。
全ページ収集の完了メッセージでは、正常件数、スキップ件数、エラー件数に加えて、処理ページ数を表示する。スキップ件数は `SkipExisting` により再取得しなかった既存 OK 行だけを数える。既存 ERROR 行を再取得した件は、再取得後の結果に応じて正常件数またはエラー件数に数える。初回ページが 0 件でも、一覧画面に到達して確認した実行事実として処理ページ数は 1 とする。
収集進捗は、`Collect_Click` の本番収集で表示し、診断モードでは表示しない。表示名は `取得中` 固定とする。初回ページで一覧項目が 1 件以上あることを確認してから表示し、開始時は 0% をステータスバーへ表示する。終了値はページ到達ごとに判明した一覧項目数を加算した値とし、新しいページの一覧項目数を加算した結果、表示パーセントが下がることを許容する。進捗値は一覧項目の処理完了数とし、詳細ページ処理後に必要な一覧復帰まで完了した時点で 1 件進める。既存 OK 行としてスキップした一覧項目、出力対象外になった一覧項目、ERROR 行として出力した一覧項目も処理完了数に含める。初回ページの一覧項目が 0 件の場合は進捗表示を開始せず、正常終了の完了メッセージだけを表示する。収集中にエラーで中断する場合は、完了扱いにせず、ステータスバーの進捗表示を解除してからエラーメッセージを表示する。

各ページ到達時に `ListItemSelector` の件数を使って、一覧上の対象IDを表示順に読み取り、そのページ開始時点の対象ID一覧として保持する。ページごとに snapshot を作り直すため、最終ページの件数が少ない場合も現在ページの件数だけを処理する。`ListItemSelector` の件数はそのページ到達時の一覧を作るための loop bound とし、一覧復帰後のズレ検知には使わない。ページサイズが固定される場合、項目数だけでは新規項目挿入や表示順変化を検知できないため。

再実行時の既定動作は `SkipExisting` とする。
収集開始時は、output シートの固定管理列を利用者指定に基づかないツール管理領域として A:D に順序固定で整える。E 列以降は空文字列を無視した 1 行目の used range 内の非空ヘッダーを読む。settings の詳細ページ列定義に存在する `OutputColumnName` と大文字小文字を区別して一致するヘッダーだけを出力列として扱い、各出力列ヘッダーが置かれた列へ書き込む。空白ヘッダー列と未定義ヘッダー列は処理対象外とし、値または数式が入っていても変更しない。非空ヘッダーが重複している場合は、同じ抽出値を該当するすべての列へ出力する。非空ヘッダーが固定管理列名と同じ場合も、固定管理列と E 列以降の出力列の両方へ出力する。出力行を書き込むときは、A:D の固定管理列と、E 列以降の定義済み出力列だけを書き込み、それ以外の列は既存値や数式を変更しない。E 列以降の既存列は削除せず、出力列を左詰めに正規化しない。新しく増えた出力列は空欄のままとし、既存 OK 行の値を埋める必要がある場合は `Overwrite` で再取得する。
`対象ID` を主キーとして既存行を判定し、既存行がある場合はその行を更新し、未登録の場合は末尾へ追加する。ページ内の全一覧項目が既存 OK 行としてスキップされた場合でも、そのページの処理は完了扱いとし、次ページ操作が有効であれば次ページへ進む。
`SkipExisting` では、`取得状態=OK` の既存 `対象ID` だけをスキップし、`取得状態=ERROR` の既存行は取得エラー行として再取得する。既存行の `取得状態` が `OK` / `ERROR` 以外の場合は、出力シート不整合として中断する。`ダウンロード状態=ERROR` でも `取得状態=OK` の既存行は、`SkipExisting` のスキップ対象とする。`Overwrite` では既存状態に関係なく全更新する。`RetryError` は旧称であり、settings に指定されている場合は設定誤りとして収集開始前に中断する。`ExistingRowMode` が不正な場合は、指定値と有効値 `SkipExisting`, `Overwrite` が分かるエラーメッセージを表示する。

## Waiting And Error Handling

画面到達判定は URL より selector を優先する。
詳細ページ到達判定では、対象ID selector の出現を必須条件とする。一覧上の対象IDと詳細ページ上の対象IDが両方取れる場合、診断モードでは一致を必須とし、不一致は失敗扱いにする。本番収集では不一致を詳細ページ単位の ERROR として出力し、一覧復帰できる場合は次の一覧項目へ進む。
ページング前は `NextPageAvailableSelector` で次ページ操作の有効性を確認する。ページング後は一覧画面 selector の出現を待ったうえで、`TimeoutSeconds` の期限内に一覧先頭の対象IDが前ページの先頭対象IDから変化することを確認する。期限内に変化しない場合は、次ページ操作後にページが進まないものとして本番収集を中断する。ページごとの先頭対象IDを記録し、同じ先頭対象IDが再登場した場合はページングが循環したものとして本番収集を中断する。初回ページの一覧項目が 0 件の場合は収集対象なしとして正常終了する。次ページ操作後に一覧項目が 0 件の場合、または先頭対象IDが読めない場合は、selector 誤り、画面崩れ、認証切れ、ページング失敗の可能性があるため本番収集を中断する。

詳細ページの必須列が取れない場合は、その件を `ERROR` の出力行として残し、一覧復帰リンクで戻れる場合は次の一覧項目へ進む。ERROR 行の `対象ID` は、詳細ページ上の対象IDを読めた場合はその値を使い、読めない場合は一覧上の対象IDを使う。
本番収集では、各一覧項目を処理する直前に現在 index の一覧上の対象IDを再評価し、現在ページ開始時点で保持した同じ index の対象IDと一致することを確認する。一致しない場合は、一覧画面が収集中に変化したものとして本番収集を中断する。
一覧復帰に失敗した場合、WebDriver session が壊れた場合、認証切れで復帰不能な場合は処理を中断する。

## References

- `docs/adr/0001-webdriver-browser-collection-policy.md`
- `CONTEXT.md`
