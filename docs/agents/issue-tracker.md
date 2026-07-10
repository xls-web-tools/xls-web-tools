# Issue tracker: GitHub

このリポジトリの issue と PRD は GitHub Issues で管理する。対象リポジトリは `xls-web-tools/xls-web-tools`。
mattpocock/skills が「issue tracker に publish する」と言う場合は、`gh` CLI で GitHub issue を作成する。

## Repository

- GitHub remote: `https://github.com/xls-web-tools/xls-web-tools.git`
- `gh` は、この clone 内で実行し、対象 repository は remote から推定する。

## Conventions

- issue 作成: `gh issue create --repo xls-web-tools/xls-web-tools --title "..." --body "..."`
- issue 参照: `gh issue view <number> --repo xls-web-tools/xls-web-tools --comments`
- issue 一覧: `gh issue list --repo xls-web-tools/xls-web-tools --state open --json number,title,body,labels,comments`
- comment 追加: `gh issue comment <number> --repo xls-web-tools/xls-web-tools --body "..."`
- label 更新: `gh issue edit <number> --repo xls-web-tools/xls-web-tools --add-label "..."`
- label 追加/削除: `gh issue edit <number> --add-label "..."` / `gh issue edit <number> --remove-label "..."`
- close: `gh issue close <number> --repo xls-web-tools/xls-web-tools --comment "..."`

## Relationships

Issue 作成時または PRD からの issue 分割時にブロック関係や親子関係がある場合は、GitHub Issues の Relationships を設定する。本文だけの記載で済ませない。

- Relationships 設定は `gh api` を標準とし、`curl` などの外部 HTTP クライアントではなく、まず `gh` CLI で実行する。将来 `gh issue` サブコマンドで同等の設定が可能になった場合は、`gh issue` を優先する。
- REST API の整数 `id` は `gh api repos/xls-web-tools/xls-web-tools/issues/NUMBER --jq .id` で取得する。
- 親子関係は sub-issues を使う。`PARENT_NUMBER` の子として `CHILD_NUMBER` を追加する場合は、`CHILD_NUMBER` の整数 `id` を取得し、`gh api --method POST repos/xls-web-tools/xls-web-tools/issues/PARENT_NUMBER/sub_issues -f sub_issue_id=CHILD_ISSUE_ID` を実行する。
- ブロック関係は issue dependencies の `blocked_by` を使う。`BLOCKING_NUMBER` が `BLOCKED_NUMBER` を block する場合は、`BLOCKING_NUMBER` の整数 `id` を取得し、`gh api --method POST repos/xls-web-tools/xls-web-tools/issues/BLOCKED_NUMBER/dependencies/blocked_by -f issue_id=BLOCKING_ISSUE_ID` を実行する。
- root、`xls-bfw-tools`、`xls-ces-timesheet` は `akhs-tkmr` owner、`xls-web-tools` は `xls-web-tools` owner、`xls-common-devtools`、`DoxyVB6`、`VBA-LanguageServer` は `modern-vba` owner で管理する。GitHub API が許可する範囲では、owner をまたぐ Relationships も設定する。
- API 権限不足、GitHub 側の制限、対象 Issue 未作成などで Relationships を設定できない場合は、本文またはコメントに関係を記載し、未設定の関係を作業結果で報告する。

## When a skill says "publish to the issue tracker"

GitHub Issue を作成する。

## When a skill says "fetch the relevant ticket"

`gh issue view <number> --repo xls-web-tools/xls-web-tools --comments` を実行する。

## GitHub Projects

この repository の issue は GitHub Issues と GitHub Projects v2 の両方で管理する。
Project は `xls-web-tools` organization project #1、`xls-web-tools main project` を使う。
Project field は `Status` を使い、issue lifecycle と同期する。

| Issue state / label | Project `Status` | Agent action |
| --- | --- | --- |
| `ready-for-agent` | `Ready` | Agent が着手できる issue として配置する |
| 実装中 | `In progress` | Agent が作業中であることを示す |
| 実装完了 + review 待ち + `ready-for-human` | `In review` | Human が Review 可能な状態として配置する |
| レビュー通過 + acceptance criteria 検証済み | `Done` | Agent が `ready-for-human` を外し、issue close と Project `Done` 更新を完了した状態として配置する |

Rules:

- `ready-for-agent` label を付けた issue は、Project `Status` を `Ready` にする。
- 実装を開始した issue は、Project `Status` を `In progress` にする。
- 実装を完了した issue は、`ready-for-agent` label を外して `ready-for-human` label を付け、Project `Status` を `In review` にする。
- レビューを通過した issue は、acceptance criteria を検証する。すべて pass した場合、または未達成項目について maintainer から明示的に対象外扱いの指示があった場合は、`ready-for-human` を外し、Issue を close して Project `Status` を `Done` にする。
- `ready-for-human` は「人間がレビューできる状態」として使う。実装担当待ちの意味だけに限定しない。
- acceptance criteria に未達成項目があり maintainer の明示指示がない場合は、close / Done へ進めない。
- `gh api graphql` で Projects v2 を操作する。scope 不足の場合は `gh auth refresh -h github.com -s project` で Projects scope を追加する。
