Attribute VB_Name = "GUIHandler"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! Web情報取得ツールのワークシート UI 入口をまとめた標準モジュールです。
'!
' #############################################################################

'* 取得開始 URL から詳細ページの対象ID取得までの診断を実行します。
'*
'* @details
'* settings の StartUrl、AuthenticatedStartSelector、ListTransitionOperationName、ListPageSelector、ListItemTargetIdSelector、DetailTransitionOperationName、TargetIdSelector、ReturnToListOperationName を使って詳細ページ出力と一覧復帰までを確認します。
'* 認証情報は Excel ブック、VBA、settings、ログに保存しません。
Public Sub RunDiagnostic_Click()
    Call InitializeCommonService(Force:=True)

    Dim app_state As ApplicationScreenUpdateManager
    Set app_state = New ApplicationScreenUpdateManager
    Call app_state.DisableUpdates(StopEvents:=False)

    On Error GoTo ON_ERROR

    Dim settings As IToolSettings
    Set settings = New ToolSettings

    Dim web_driver_client As WebDriverClient
    Set web_driver_client = New_WebDriverClient(settings)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(web_driver_client, settings)

    Dim process As WebDriverProcess
    Set process = New_WebDriverProcess(New_WebDriverPortProbe())

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, settings, KeepVisibleBrowserOnError:=True)

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(lifecycle, settings)

    Call runner.Run

    If runner.IsOutputExcluded Then
        Call MsgBoxPage("詳細ページ対象ID取得診断が完了しました。" & vbCrLf & "対象ID: " & runner.TargetId & vbCrLf & "出力対象外のため output シートへ行を書き込みませんでした。")
    Else
        Call MsgBoxPage("詳細ページ対象ID取得診断が完了しました。" & vbCrLf & "対象ID: " & runner.TargetId)
    End If

ON_EXIT:
    On Error Resume Next
    If Not app_state Is Nothing Then Call app_state.Restore
    On Error GoTo 0
    Exit Sub

ON_ERROR:
    Dim err_desc As String
    Dim err_num As Long
    Dim err_source As Variant
    Dim debug_lines As String

    err_desc = Err.Description
    err_num = Err.Number
    err_source = Err.Source

    On Error Resume Next
    If Not app_state Is Nothing Then Call app_state.Restore
    If Not DbgInfo Is Nothing Then debug_lines = vbCrLf & DbgInfo.BuildMessageLines()
    On Error GoTo 0

    Call MsgBoxPage(err_desc & " (0x" & Hex(err_num) & ") @" & err_source & debug_lines)
End Sub
'* 一覧項目を対象ID主キーで全ページ収集します。
'*
'* @details
'* settings の ListItemSelector と ListItemTargetIdSelector でページごとの対象IDを snapshot 化し、DetailTransitionOperationName、ReturnToListOperationName、NextPageOperationName で各ページを巡回します。
'* セレクタの実値や認証情報はログへ出力しません。
Public Sub Collect_Click()
    Call InitializeCommonService(Force:=True)

    Dim run_state As CommonRunStateManager
    Set run_state = New CommonRunStateManager

    Dim app_state As ApplicationScreenUpdateManager
    Set app_state = New ApplicationScreenUpdateManager
    Call app_state.DisableUpdates(StopEvents:=False)

    On Error GoTo ON_ERROR

    ' ==== 実処理 ========

    Dim settings As IToolSettings
    Set settings = New ToolSettings

    Dim web_driver_client As WebDriverClient
    Set web_driver_client = New_WebDriverClient(settings)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(web_driver_client, settings)

    Dim process As WebDriverProcess
    Set process = New_WebDriverProcess(New_WebDriverPortProbe())

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, settings)

    Dim runner As WebCollectionRunner
    Set runner = New_WebCollectionRunner(lifecycle, settings)

    Call runner.Run

    ' ==== 実処理ここまで ========

    Call MsgBoxPage("全ページの収集が完了しました。" _
            & vbCrLf & "処理ページ数: " & CStr(runner.PageCount) _
            & vbCrLf & "正常: " & CStr(runner.SucceededCount) _
            & vbCrLf & "スキップ: " & CStr(runner.SkippedCount) _
            & vbCrLf & "出力対象外: " & CStr(runner.OutputExcludedCount) _
            & vbCrLf & "エラー: " & CStr(runner.ErrorCount))

ON_EXIT:
    On Error Resume Next
    If Not app_state Is Nothing Then Call app_state.Restore
    On Error GoTo 0
    Exit Sub

ON_ERROR:
    Dim err_desc As String
    Dim err_num As Long
    Dim err_source As Variant
    Dim debug_lines As String

    err_desc = Err.Description
    err_num = Err.Number
    err_source = Err.Source

    On Error Resume Next
    If Not app_state Is Nothing Then Call app_state.Restore
    If Not DbgInfo Is Nothing Then debug_lines = vbCrLf & DbgInfo.BuildMessageLines()
    On Error GoTo 0

    Call MsgBoxPage(err_desc & " (0x" & Hex(err_num) & ") @" & err_source & debug_lines)
End Sub
