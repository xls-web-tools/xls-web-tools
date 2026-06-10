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
'! 取得開始 URL から詳細ページの対象ID取得までの診断を実行します。
'!
' #############################################################################

Private pLifecycle As WebDriverSessionLifecycle
Private pSessionClient As WebDriverSessionClient
Private pSettings As IToolSettings
Private pTransitionResolver As TransitionOperationResolver
Private pIndexTemplate As ListItemIndexTemplate
Private pListItemTargetId As String
Private pTargetId As String
Private pIsOutputExcluded As Boolean
Private pIsInitialized As Boolean

'* 一覧上の対象ID。
Public Property Get ListItemTargetId() As String
    Call pCheckInit
    ListItemTargetId = pListItemTargetId
End Property

'* 詳細ページ上の対象ID。
Public Property Get TargetId() As String
    Call pCheckInit
    TargetId = pTargetId
End Property

'* 出力対象条件により output シートへ書かなかったか否か。
Public Property Get IsOutputExcluded() As Boolean
    Call pCheckInit
    IsOutputExcluded = pIsOutputExcluded
End Property

'* WebDriver navigation diagnostic runner を初期化します。
'*
'* @param Lifecycle WebDriver session lifecycle。
'* @param Settings Web情報取得の設定。
'*
'* @details
'* session lifecycle の開始、正常終了、エラー時 cleanup は lifecycle に委譲します。
Public Sub Initialize( _
        ByVal Lifecycle As WebDriverSessionLifecycle, _
        ByVal Settings As IToolSettings)

    If Lifecycle Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebNavDiagnosticRunner", Description:="Lifecycle は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebNavDiagnosticRunner", Description:="Settings は必須です。"
    End If

    Set pLifecycle = Lifecycle
    Set pSessionClient = Lifecycle.SessionClient
    Set pSettings = Settings
    Set pTransitionResolver = New_TransitionOperationResolver(Settings)
    Set pIndexTemplate = New_ListItemIndexTemplate()
    pListItemTargetId = ""
    pTargetId = ""
    pIsOutputExcluded = False
    pIsInitialized = True
End Sub

'* 詳細ページの対象ID取得までの診断を実行します。
'*
'* @return 作成された session id。
'*
'* @details
'* WebDriver session lifecycle を開始し、取得開始 URL 遷移、認証後起点 selector 待機、一覧遷移操作、一覧 selector 待機、詳細遷移操作、対象ID取得を順に実行します。
Public Function Run() As String
    Call pCheckInit

    Dim result_value As String
    pListItemTargetId = ""
    pTargetId = ""
    pIsOutputExcluded = False

    Dim detail_runner As DetailPageRunner
    Set detail_runner = New_DetailPageRunner(pSessionClient, pSettings)
    On Error GoTo ON_ERROR

    Dim preparation As WebCollectionPreparation
    Set preparation = New_WebCollectionPreparation(pLifecycle, pSettings)
    result_value = preparation.Run()
    Call pRunFirstDetail(detail_runner)

    Call pLifecycle.FinishSession

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

    On Error Resume Next
    If Not preparation Is Nothing Then Call preparation.ClearProgress
    On Error GoTo 0
    Call pLifecycle.CleanupAfterError(err_desc)

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebNavDiagnosticRunner", Description:="初期化されていません。"
    End If
End Sub

Private Sub pRunFirstDetail(ByVal DetailRunner As DetailPageRunner)
    If Not pTransitionResolver.HasDetailTransitionOperation Then Exit Sub

    pListItemTargetId = pReadFirstListItemTargetId()

    Dim run_result As DetailPageRunResult
    Set run_result = DetailRunner.Run( _
            pListItemTargetId, _
            ItemIndex:=0, _
            IsDiagnosticMode:=True)

    pTargetId = run_result.TargetId
    pIsOutputExcluded = run_result.IsOutputExcluded

    Dim output_writer As OutputSheetWriter
    Set output_writer = New_OutputSheetWriter(pSettings)
    Call output_writer.EnsureHeaders

    If run_result.IsOutputExcluded Then Exit Sub

    Call output_writer.WriteDiagnosticRow( _
            run_result.TargetId, _
            run_result.Status, _
            run_result.ErrorDescription, _
            run_result.DetailValues, _
            run_result.DownloadStatus)
End Sub


Private Function pReadFirstListItemTargetId() As String
    Dim selector_value As String
    selector_value = pIndexTemplate.Apply(pSettings.ListItemTargetIdSelector, 0)

    Dim result_value As String
    result_value = Trim$(pSessionClient.ReadElementText(selector_value))
    If result_value = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebNavDiagnosticRunner", Description:="一覧上の対象IDを取得できません。"
    End If

    pReadFirstListItemTargetId = result_value
End Function
