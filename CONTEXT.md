# xls-web-tools

xls-web-tools は、Web 上の情報取得・整形・検証を支援する Excel マクロ ツール群です。
各ツールは `.xlsm` と `modules` を同じフォルダーに置き、VBA ソースを `modules` で管理します。

## フォルダー構成

- `common_modules_repo`: xls-web-tools に配布された共通モジュール。変更時は `xls-common-devtools/CommonModules/modules` を原本として扱います。
- `Web情報取得`: Web 上の情報を取得し、Excel で扱える形へ整形・確認するツール。
- `Web情報取得/modules`: Web 情報取得ツール固有の VBA ソースとユニットテストを管理する場所。
- `docs/agents`: Agent skills が参照する Issue tracker と domain docs の運用ルール。
- `docs/adr`: Web ツール群の architecture decision records。

## Language

**Web 関連ツール**:
Web ページ、HTTP API、公開データなどの取得・整形・確認を Excel から支援するツール群。
_Avoid_: 一般的な Web サイトや Web アプリ全般

**Web情報取得**:
Web 上の情報を取得し、Excel ブック上で利用・確認できる形式へ変換するツール。
_Avoid_: xls-web-tools 内の全ツール

**取得対象**:
Web情報取得が読み取り対象とする URL、HTTP API、HTML、JSON、CSV などの外部情報源。
_Avoid_: 取得後に Excel 上へ書き込まれた結果

**取得結果**:
取得対象から読み取って、Excel 上で確認または後続処理に使える形へ整えたデータ。
_Avoid_: 生のレスポンスだけ

**取得設定**:
取得先、取得方法、認証やヘッダー、出力先シートなど、Web 情報取得の動作を決める設定。
_Avoid_: ユーザーが一時的に入力する検索語だけ

**共通モジュール**:
`common_modules_repo` から各ツールの `modules` へ配布される VBA 基盤。変更時は `xls-common-devtools/CommonModules/modules` を原本として扱います。
_Avoid_: xls-web-tools 固有モジュール

## Development Notes

- Web ツール固有の業務処理は対象ツールの `modules` に置きます。
- `common_modules_repo` に含まれるファイルを直接原本として編集しません。共通モジュール変更は `xls-common-devtools/CommonModules/modules` から始めます。
- 取得対象、取得設定、シート名、列番号、業務文言を共通モジュールへ混入させません。
- 外部 Web へのアクセスを伴う処理は、通常の単体テストではテストダブルまたは固定テストデータへ切り出し、ネットワーク状態に依存しない検証を優先します。
- Excel への import/export や UnitTestMain 実行は Excel COM と VBIDE アクセスを使うため、実行前に対象 `.xlsm` が閉じていることを確認します。

## Example Dialogue

開発者: `取得結果` は HTTP の生レスポンス文字列だけを指しますか？
ドメイン専門家: いいえ。取得対象から読み取り、Excel 上で確認や後続処理に使える形へ整えたデータを指します。
