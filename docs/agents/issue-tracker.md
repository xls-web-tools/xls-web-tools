# Issue tracker: GitHub

このリポジトリの issue と PRD は GitHub Issues で管理する。対象リポジトリは `tkmr-akhs/xls-web-tools`。
現時点では GitHub repository 作成前のため、Issue/PRD 作成や Project 操作は repository 作成後に実行する。
mattpocock/skills が「issue tracker に publish する」と言う場合は、`gh` CLI で GitHub issue を作成する。

## Repository

- GitHub remote: `https://github.com/tkmr-akhs/xls-web-tools.git`
- `gh` は、この clone 内で実行し、対象 repository は remote から推定する。repository 作成前や remote 推定ができない場合は `--repo tkmr-akhs/xls-web-tools` を明示する。

## Conventions

- issue 作成: `gh issue create --repo tkmr-akhs/xls-web-tools --title "..." --body "..."`
- issue 参照: `gh issue view <number> --repo tkmr-akhs/xls-web-tools --comments`
- issue 一覧: `gh issue list --repo tkmr-akhs/xls-web-tools --state open --json number,title,body,labels,comments`
- comment 追加: `gh issue comment <number> --repo tkmr-akhs/xls-web-tools --body "..."`
- label 更新: `gh issue edit <number> --repo tkmr-akhs/xls-web-tools --add-label "..."`
- label 追加/削除: `gh issue edit <number> --add-label "..."` / `gh issue edit <number> --remove-label "..."`
- close: `gh issue close <number> --repo tkmr-akhs/xls-web-tools --comment "..."`

## Relationships

Issue 作成時または PRD からの issue 分割時にブロック関係や親子関係がある場合は、GitHub Issues の Relationships を設定する。本文だけの記載で済ませない。

- Relationships 設定は `gh api` を標準とし、`curl` などの外部 HTTP クライアントではなく、まず `gh` CLI で実行する。将来 `gh issue` サブコマンドで同等の設定が可能になった場合は、`gh issue` を優先する。
- REST API の整数 `id` は `gh api repos/tkmr-akhs/xls-web-tools/issues/NUMBER --jq .id` で取得する。
- 親子関係は sub-issues を使う。`PARENT_NUMBER` の子として `CHILD_NUMBER` を追加する場合は、`CHILD_NUMBER` の整数 `id` を取得し、`gh api --method POST repos/tkmr-akhs/xls-web-tools/issues/PARENT_NUMBER/sub_issues -f sub_issue_id=CHILD_ISSUE_ID` を実行する。
- ブロック関係は issue dependencies の `blocked_by` を使う。`BLOCKING_NUMBER` が `BLOCKED_NUMBER` を block する場合は、`BLOCKING_NUMBER` の整数 `id` を取得し、`gh api --method POST repos/tkmr-akhs/xls-web-tools/issues/BLOCKED_NUMBER/dependencies/blocked_by -f issue_id=BLOCKING_ISSUE_ID` を実行する。
- root、`xls-common-devtools`、`xls-web-tools`、`xls-bfw-tools`、`xls-ces-timesheet`、`DoxyVB6` は同じ `tkmr-akhs` owner なので、GitHub API が許可する範囲ではリポジトリをまたぐ Relationships も設定する。
- API 権限不足、GitHub 側の制限、対象 Issue 未作成などで Relationships を設定できない場合は、本文またはコメントに関係を記載し、未設定の関係を作業結果で報告する。

## When a skill says "publish to the issue tracker"

GitHub Issue を作成する。

## When a skill says "fetch the relevant ticket"

`gh issue view <number> --repo tkmr-akhs/xls-web-tools --comments` を実行する。

## GitHub Projects

この repository の GitHub Projects v2 は現時点で未作成または未確認である。
Project 操作を求められた場合は、作業前に対象 Project を確認する。Project 作成後は、この節を既存サブリポジトリと同じ形式へ更新する。
Issue lifecycle のラベル運用は root の `AGENTS.md` に従い、Project `Status` は対象 Project が明示された場合だけ同期する。
