Attribute VB_Name = "Mod_WebConstants"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Web情報取得で共有する定数です。
'!
' #############################################################################

'* settings シート名。
Public Const G_WEB_SETTINGS_SHEET As String = "settings"

'* 出力シートの詳細ページ ID 列名。
Public Const G_WEB_OUTPUT_COL_DETAIL_PAGE_ID As String = "詳細ページID"

'* 出力シートの対象 ID 列名。
Public Const G_WEB_OUTPUT_COL_TARGET_ID As String = "対象ID"

'* 出力シートの取得状態列名。
Public Const G_WEB_OUTPUT_COL_STATUS As String = "取得状態"

'* 出力シートのエラー内容列名。
Public Const G_WEB_OUTPUT_COL_ERROR As String = "エラー内容"

'* 正常取得を表す状態値。
Public Const G_WEB_STATUS_OK As String = "OK"

'* エラー取得を表す状態値。
Public Const G_WEB_STATUS_ERROR As String = "ERROR"

'* 既存行を再取得せずにスキップする既存行モード。
Public Const G_WEB_ROW_MODE_SKIP_EXISTING As String = "SkipExisting"

'* エラー行だけを再取得する既存行モード。
Public Const G_WEB_ROW_MODE_RETRY_ERROR As String = "RetryError"

'* 既存行を上書きして再取得する既存行モード。
Public Const G_WEB_ROW_MODE_OVERWRITE As String = "Overwrite"

'* 診断用の実行モード。
Public Const G_WEB_RUN_MODE_DIAGNOSTIC As String = "Diagnostic"

'* 通常取得用の実行モード。
Public Const G_WEB_RUN_MODE_COLLECT As String = "Collect"
