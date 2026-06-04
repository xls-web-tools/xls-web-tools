VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebNavDiagnosticRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 取得開始 URL から一覧画面到達までの診断を実行します。
'!
' #############################################################################

Private pProcess As IWebDriverProcess
Private pSessionClient As WebDriverSessionClient
Private pSettings As IToolSettings
Private pIsInitialized As Boolean

'* WebDriver navigation diagnostic runner を初期化します。
'*
'* @param Process WebDriver process 操作。
'* @param SessionClient WebDriver session client。
'* @param Settings Web情報取得の設定。
'*
'* @details
'* process 起動、session 操作、settings 読み取りの依存を外部から渡します。
Public Sub Initialize( _
        ByVal Process As IWebDriverProcess, _
        ByVal SessionClient As WebDriverSessionClient, _
        ByVal Settings As IToolSettings)

    If Process Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebNavDiagnosticRunner", Description:="Process は必須です。"
    End If
    If SessionClient Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebNavDiagnosticRunner", Description:="SessionClient は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebNavDiagnosticRunner", Description:="Settings は必須です。"
    End If

    Set pProcess = Process
    Set pSessionClient = SessionClient
    Set pSettings = Settings
    pIsInitialized = True
End Sub

'* 一覧画面到達までの診断を実行します。
'*
'* @return 作成された session id。
'*
'* @details
'* WebDriver process 起動、session 作成、取得開始 URL 遷移、認証後起点 selector 待機、一覧遷移操作、一覧 selector 待機、quit、process 終了を順に実行します。
Public Function Run() As String
    Call pCheckInit

    Dim result_value As String

    On Error GoTo ON_ERROR

    Call pProcess.Start(pSettings)

    result_value = pSessionClient.CreateSession()
    Call pSessionClient.NavigateToStartUrl
    Call pSessionClient.WaitForSelector(pSettings.AuthenticatedStartSelector)

    Dim list_operation As TransitionOperation
    Set list_operation = pFindTransitionOperation(pSettings.ListTransitionOperationName)
    Call pSessionClient.RunTransitionOperation(list_operation)
    Call pSessionClient.WaitForSelector(pSettings.ListPageSelector)

    Call pSessionClient.QuitSession(result_value)
    Call pProcess.StopProcess

    On Error GoTo 0
    Run = result_value
    Exit Function

ON_ERROR:
    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    Dim err_help_file As String
    Dim err_help_context As Long
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description
    err_help_file = Err.HelpFile
    err_help_context = Err.HelpContext

    If pShouldKeepBrowserForDiagnostic(result_value) Then
        err_desc = err_desc & vbCrLf & "Headless=False のため、診断用にブラウザを残しました。確認後、ブラウザを閉じてください。"
    Else
        On Error Resume Next
        If result_value <> "" Then Call pSessionClient.QuitSession(result_value)
        Call pProcess.StopProcess
        On Error GoTo 0
    End If

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebNavDiagnosticRunner", Description:="初期化されていません。"
    End If
End Sub

Private Function pShouldKeepBrowserForDiagnostic(ByVal SessionId As String) As Boolean
    pShouldKeepBrowserForDiagnostic = ((Not pSettings.Headless) And SessionId <> "")
End Function

Private Function pFindTransitionOperation(ByVal OperationName As String) As TransitionOperation
    Dim operations As ObjectList
    Set operations = pSettings.TransitionOperations

    Dim operation_idx As Long
    For operation_idx = 0 To operations.Count - 1
        Dim target_operation As TransitionOperation
        Set target_operation = operations.Item(operation_idx)
        If StrComp(target_operation.OperationName, OperationName, vbTextCompare) = 0 Then
            Set pFindTransitionOperation = target_operation
            Exit Function
        End If
    Next operation_idx

    Err.Raise Number:=vbObjectError + 1, Source:="Class WebNavDiagnosticRunner", Description:="一覧画面への画面遷移操作が見つかりません。(" & OperationName & ")"
End Function
