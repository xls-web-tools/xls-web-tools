VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebCollectionPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Web情報取得の一覧画面到達までの準備を実行します。
'!
' #############################################################################

Private Const C_PREPARATION_PROGRESS_TASK_NAME As String = "取得準備中"
Private Const C_PREPARATION_PROGRESS_TOTAL As Long = 5

Private pLifecycle As WebDriverSessionLifecycle
Private pSessionClient As WebDriverSessionClient
Private pSettings As IToolSettings
Private pTransitionResolver As TransitionOperationResolver
Private pIsInitialized As Boolean

'* Web collection preparation を初期化します。
'*
'* @param Lifecycle WebDriver session lifecycle。
'* @param Settings Web情報取得の設定。
'*
'* @details
'* session 開始、取得開始 URL 遷移、認証後起点確認、一覧遷移、一覧画面到達確認をまとめます。
Public Sub Initialize( _
        ByVal Lifecycle As WebDriverSessionLifecycle, _
        ByVal Settings As IToolSettings)

    If Lifecycle Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionPreparation.Initialize", Description:="Lifecycle は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionPreparation.Initialize", Description:="Settings は必須です。"
    End If

    Set pLifecycle = Lifecycle
    Set pSessionClient = Lifecycle.SessionClient
    Set pSettings = Settings
    Set pTransitionResolver = New_TransitionOperationResolver(Settings)
    pIsInitialized = True
End Sub

'* 一覧画面到達までの準備を実行します。
'*
'* @return 作成された session id。
'*
'* @details
'* 一覧画面到達までの準備を実行します。
Public Function Run() As String
    Call pCheckInit

    Dim result_value As String
    Call pStartProgress

    result_value = pLifecycle.StartSession()
    If result_value = "" Then
        Run = result_value
        Exit Function
    End If
    Call pAdvanceProgress

    Call pSessionClient.NavigateToStartUrl
    Call pAdvanceProgress

    Call pSessionClient.WaitForSelector(pSettings.AuthenticatedStartSelector)
    Call pAdvanceProgress

    Dim list_operation As TransitionOperation
    Set list_operation = pTransitionResolver.ResolveListTransitionOperation()
    Call pSessionClient.RunTransitionOperation(list_operation)
    Call pAdvanceProgress

    Call pSessionClient.WaitForSelector(pSettings.ListPageSelector)
    Call pAdvanceProgress

    Run = result_value
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionPreparation.pCheckInit", Description:="初期化されていません。"
    End If
End Sub

Private Sub pStartProgress()
    If ProgStat Is Nothing Then Exit Sub

    Call ProgStat.Reset
    ProgStat.TaskName = C_PREPARATION_PROGRESS_TASK_NAME
    ProgStat.TotalValue = CDbl(C_PREPARATION_PROGRESS_TOTAL)
    ProgStat.ProcessedValue = 0
End Sub

Private Sub pAdvanceProgress()
    If ProgStat Is Nothing Then Exit Sub

    ProgStat.ProcessedValue = ProgStat.ProcessedValue + 1
End Sub

'* 準備中に中断した進捗表示を解除します。
'*
'* @details
'* 準備中に中断した進捗表示を解除します。
Public Sub ClearProgress()
    If ProgStat Is Nothing Then Exit Sub
    If StrComp(ProgStat.TaskName, C_PREPARATION_PROGRESS_TASK_NAME, vbBinaryCompare) <> 0 Then Exit Sub
    If ProgStat.IsComplete Then Exit Sub

    Call ProgStat.Reset(ResetTaskName:=False, ResetStartAndTotalValue:=False)
End Sub
