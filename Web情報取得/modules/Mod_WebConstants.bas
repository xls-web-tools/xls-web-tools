Attribute VB_Name = "Mod_WebConstants"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! Web情報取得で共有する定数です。
'!
' #############################################################################

'* settings シート名。
Public Const G_WEB_SETTINGS_SHEET As String = "settings"

'* 出力シートの対象 ID 列名。
Public Const G_WEB_OUTPUT_COL_TARGET_ID As String = "対象ID"

'* 出力シートの取得状態列名。
Public Const G_WEB_OUTPUT_COL_STATUS As String = "取得状態"

'* 出力シートのエラー内容列名。
Public Const G_WEB_OUTPUT_COL_ERROR As String = "エラー内容"

'* 出力シートのダウンロード状態列名。
Public Const G_WEB_OUTPUT_COL_DOWNLOAD_STATUS As String = "ダウンロード状態"

'* 正常取得を表す状態値。
Public Const G_WEB_STATUS_OK As String = "OK"

'* エラー取得を表す状態値。
Public Const G_WEB_STATUS_ERROR As String = "ERROR"

'* ダウンロード対象ファイルを取得したことを表す状態値。
Public Const G_WEB_DOWNLOAD_STATUS_DOWNLOADED As String = "DOWNLOADED"

'* ダウンロード対象ファイルが存在しないことを表す状態値。
Public Const G_WEB_DOWNLOAD_STATUS_NO_FILE As String = "NO_FILE"

'* 既存ファイルやフォルダー保護のためダウンロード反映をスキップしたことを表す状態値。
Public Const G_WEB_DOWNLOAD_STATUS_SKIPPED As String = "SKIPPED"

'* ダウンロード対象ファイルの取得や保存に失敗したことを表す状態値。
Public Const G_WEB_DOWNLOAD_STATUS_ERROR As String = "ERROR"

'* 既存行を再取得せずにスキップする既存行モード。
Public Const G_WEB_ROW_MODE_SKIP_EXISTING As String = "SkipExisting"

'* 既存行を上書きして再取得する既存行モード。
Public Const G_WEB_ROW_MODE_OVERWRITE As String = "Overwrite"

'* 診断用の実行モード。
Public Const G_WEB_RUN_MODE_DIAGNOSTIC As String = "Diagnostic"

'* 通常取得用の実行モード。
Public Const G_WEB_RUN_MODE_COLLECT As String = "Collect"

'* 条件として評価する列参照記法種別。
Public Const G_COLUMN_EXPRESSION_KIND_CONDITION As String = "Condition"

'* 値として評価する列参照記法種別。
Public Const G_COLUMN_EXPRESSION_KIND_VALUE As String = "Value"
