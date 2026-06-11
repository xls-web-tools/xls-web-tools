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

Implements IWebCollectionRunBody

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
Private pDetailRunner As DetailPageRunner
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

    pListItemTargetId = ""
    pTargetId = ""
    pIsOutputExcluded = False
    Set pDetailRunner = New_DetailPageRunner(pSessionClient, pSettings)

    Dim run_lifecycle As WebCollectionRunLifecycle
    Set run_lifecycle = New_WebCollectionRunLifecycle(pLifecycle, pSettings)

    Run = run_lifecycle.Run(Me)
End Function

'* 一覧画面到達後に診断対象の詳細ページを処理します。
Public Sub RunAfterPreparation()
    Call pCheckInit
    Call pRunFirstDetail(pDetailRunner)
End Sub
Private Sub IWebCollectionRunBody_RunAfterPreparation()
    Call RunAfterPreparation
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

'* エラー時に診断 runner 側の進捗表示を解除します。
Public Sub ClearProgressAfterError()
    Call pCheckInit
End Sub
Private Sub IWebCollectionRunBody_ClearProgressAfterError()
    Call ClearProgressAfterError
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebNavDiagnosticRunner", Description:="初期化されていません。"
    End If
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
