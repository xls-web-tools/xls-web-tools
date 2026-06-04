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

'* 取得開始 URL から一覧画面到達までの診断を実行します。
'*
'* @details
'* settings の StartUrl、AuthenticatedStartSelector、ListTransitionOperationName、ListPageSelector を使って一覧画面到達までを確認します。
'* 認証情報は Excel ブック、VBA、settings、ログに保存しません。
Public Sub RUN_DIAGNOSTIC()
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
    Set process = New_WebDriverProcess(FsSrv, New_WebDriverPortProbe())

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, settings)

    Call runner.Run

    Call MsgBoxPage("一覧画面到達診断が完了しました。")

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
