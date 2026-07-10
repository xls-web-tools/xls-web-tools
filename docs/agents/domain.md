# Domain Docs

Engineering skills がこのリポジトリの domain documentation を読むためのルール。

## Layout

このリポジトリは single-context として扱う。

- `CONTEXT.md`: xls-web-tools の domain glossary と前提
- `docs/adr/`: architecture decision records

## Related Context

- ワークスペース全体の前提が必要な場合は、親リポジトリの `CONTEXT.md` を確認する。
- 共通モジュールを変更する場合は、`xls-common-modules/CONTEXT.md` と `xls-common-modules/docs/agents/domain.md` を確認する。
- トリアージラベルは親リポジトリの `docs/agents/triage-labels.md` を参照する。
- `common_modules_repo` は配布結果であり、共通モジュールの原本ではない。

## Reading Rules

- 作業前に、関連する `CONTEXT.md` と `docs/adr/` を確認する。
- ファイルが存在しない場合は、欠落を問題として扱わず、そのまま作業を進める。
- Issue、設計案、テスト名、リファクタリング案では `CONTEXT.md` の用語を優先する。
- 既存 ADR と矛盾する提案をする場合は、矛盾する ADR を明示する。
- Web 情報取得、HTTP/API 取得、Excel シートへの出力を扱う場合は、`docs/product-spec.md`、該当ツールの vba-dev source set (`project.json` の `sourcePath`)、TODO / README があれば確認する。
