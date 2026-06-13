VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebCollectionRunner"
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
'! 一覧項目を詳細ページ経由で収集します。
'!
' #############################################################################

Private Const C_PROGRESS_TASK_NAME As String = "取得中"

Private pLifecycle As WebDriverSessionLifecycle
Private pSessionClient As WebDriverSessionClient
Private pSettings As IToolSettings
Private pTransitionResolver As TransitionOperationResolver
Private pIndexTemplate As ListItemIndexTemplate
Private pDetailRunner As DetailPageRunner
Private pOutputWriter As OutputSheetWriter
Private pSucceededCount As Long
Private pSkippedCount As Long
Private pOutputExcludedCount As Long
Private pErrorCount As Long
Private pPageCount As Long
Private pProgressStarted As Boolean
Private pIsInitialized As Boolean

'* 正常に出力した件数。
Public Property Get SucceededCount() As Long
    Call pCheckInit
    SucceededCount = pSucceededCount
End Property

'* 既存行扱いによりスキップした件数。
Public Property Get SkippedCount() As Long
    Call pCheckInit
    SkippedCount = pSkippedCount
End Property

'* 出力対象条件により output シートへ書かなかった件数。
Public Property Get OutputExcludedCount() As Long
    Call pCheckInit
    OutputExcludedCount = pOutputExcludedCount
End Property

'* ERROR 行として出力した件数。
Public Property Get ErrorCount() As Long
    Call pCheckInit
    ErrorCount = pErrorCount
End Property

'* 処理した一覧ページ数。
Public Property Get PageCount() As Long
    Call pCheckInit
    PageCount = pPageCount
End Property

'* Web collection runner を初期化します。
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
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner.Initialize", Description:="Lifecycle は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner.Initialize", Description:="Settings は必須です。"
    End If

    Set pLifecycle = Lifecycle
    Set pSessionClient = Lifecycle.SessionClient
    Set pSettings = Settings
    Set pTransitionResolver = New_TransitionOperationResolver(Settings)
    Set pIndexTemplate = New_ListItemIndexTemplate()
    pSucceededCount = 0
    pSkippedCount = 0
    pOutputExcludedCount = 0
    pErrorCount = 0
    pPageCount = 0
    pProgressStarted = False
    pIsInitialized = True
End Sub

'* 本番収集を実行します。
'*
'* @return 作成された session id。
'*
'* @details
'* 取得開始 URL から一覧画面へ移動し、一覧項目を表示順に巡回します。
'* 各ページ到達時に対象IDを snapshot 化し、各項目処理直前に同じ位置の対象IDが変化していないことを確認します。
'* 次ページ操作が有効な場合は、先頭対象IDの変化を確認してから次ページを処理します。
Public Function Run() As String
    Call pCheckInit

    pSucceededCount = 0
    pSkippedCount = 0
    pOutputExcludedCount = 0
    pErrorCount = 0
    pPageCount = 0
    pProgressStarted = False
    Set pDetailRunner = New_DetailPageRunner(pSessionClient, pSettings)
    Set pOutputWriter = New_OutputSheetWriter(pSettings)

    Dim run_lifecycle As WebCollectionRunLifecycle
    Set run_lifecycle = New_WebCollectionRunLifecycle(pLifecycle, pSettings)

    Run = run_lifecycle.Run(Me)
End Function

'* 一覧画面到達後に本番収集を実行します。
'*
'* @details
'* 一覧画面到達後に本番収集を実行します。
Public Sub RunAfterPreparation()
    Call pCheckInit
    Call pOutputWriter.EnsureHeaders
    Call pCollectPages(pOutputWriter, pDetailRunner)
    Call pMarkProgressComplete
End Sub
Private Sub IWebCollectionRunBody_RunAfterPreparation()
    Call RunAfterPreparation
End Sub

Private Sub pCollectPages(ByVal OutputWriter As OutputSheetWriter, ByVal DetailRunner As DetailPageRunner)
    Dim first_target_ids As ObjectSet
    Set first_target_ids = New ObjectSet

    Do
        pPageCount = pPageCount + 1

        Dim item_count As Long
        item_count = pSessionClient.CountElements(pSettings.ListItemSelector)
        If item_count = 0 Then
            If pPageCount = 1 Then Exit Sub
            Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner.pCollectPages", Description:="次ページ操作後に一覧項目がありません。"
        End If
        Call pAddProgressTotal(item_count)

        Dim target_snapshot As ArrayObject
        Set target_snapshot = pBuildListTargetSnapshot(item_count)

        Dim first_target_id As String
        first_target_id = CStr(target_snapshot.Item(0))
        If first_target_ids.Exists(first_target_id) Then
            Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner.pCollectPages", Description:="ページングが循環しました。(" & first_target_id & ")"
        End If
        Call first_target_ids.Add(first_target_id)

        Dim item_idx As Long
        For item_idx = 0 To item_count - 1
            Call pCollectListItem(OutputWriter, target_snapshot, item_idx, DetailRunner)
            Call pIncrementProgress
        Next item_idx

        If Not pHasNextPage() Then Exit Do
        Call pMoveToNextPage(first_target_id)
    Loop
End Sub

Private Sub pMarkProgressComplete()
    If ProgStat Is Nothing Then Exit Sub
    If Not pProgressStarted Then Exit Sub

    Call ProgStat.MarkAsComplete
End Sub

'* エラー時に本番収集側の進捗表示を解除します。
'*
'* @details
'* エラー時に本番収集側の進捗表示を解除します。
Public Sub ClearProgressAfterError()
    Call pCheckInit
    Call pClearProgress
End Sub
Private Sub IWebCollectionRunBody_ClearProgressAfterError()
    Call ClearProgressAfterError
End Sub

Private Sub pClearProgress()
    If ProgStat Is Nothing Then Exit Sub
    If Not pProgressStarted Then Exit Sub

    Call ProgStat.Reset(ResetTaskName:=False, ResetStartAndTotalValue:=False)
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner.pCheckInit", Description:="初期化されていません。"
    End If
End Sub

Private Sub pAddProgressTotal(ByVal ItemCount As Long)
    If ProgStat Is Nothing Then Exit Sub

    If pProgressStarted Then
        ProgStat.TotalValue = ProgStat.TotalValue + CDbl(ItemCount)
        ProgStat.ProcessedValue = ProgStat.ProcessedValue
        Call ProgStat.DisplayProgress
    Else
        Call ProgStat.Reset
        ProgStat.TaskName = C_PROGRESS_TASK_NAME
        ProgStat.TotalValue = CDbl(ItemCount)
        ProgStat.ProcessedValue = 0
        pProgressStarted = True
    End If
End Sub

Private Sub pIncrementProgress()
    If ProgStat Is Nothing Then Exit Sub
    If Not pProgressStarted Then Exit Sub

    ProgStat.ProcessedValue = ProgStat.ProcessedValue + 1
End Sub

Private Function pBuildListTargetSnapshot(ByVal ItemCount As Long) As ArrayObject
    Dim result_values As ArrayObject
    Set result_values = New ArrayObject

    If 0 < ItemCount Then
        Call result_values.ReDimArray(0, ItemCount - 1)
    End If

    Dim item_idx As Long
    For item_idx = 0 To ItemCount - 1
        Dim target_id As String
        target_id = pReadListItemTargetId(item_idx)
        Call result_values.Update(item_idx, target_id)
    Next item_idx

    Set pBuildListTargetSnapshot = result_values
End Function

Private Sub pCollectListItem( _
        ByVal OutputWriter As OutputSheetWriter, _
        ByVal TargetSnapshot As ArrayObject, _
        ByVal ItemIndex As Long, _
        ByVal DetailRunner As DetailPageRunner)

    Dim expected_target_id As String
    expected_target_id = CStr(TargetSnapshot.Item(ItemIndex))

    Dim current_target_id As String
    current_target_id = pReadListItemTargetId(ItemIndex)
    Call pRaiseIfListItemChanged(ItemIndex, expected_target_id, current_target_id)

    If Not OutputWriter.ShouldCollectTarget(current_target_id) Then
        pSkippedCount = pSkippedCount + 1
        Exit Sub
    End If

    Dim run_result As DetailPageRunResult
    Set run_result = DetailRunner.Run(current_target_id, ItemIndex:=ItemIndex)

    If run_result.IsOutputExcluded Then
        pOutputExcludedCount = pOutputExcludedCount + 1
        Exit Sub
    End If

    Call OutputWriter.WriteCollectionRow( _
            run_result.TargetId, _
            run_result.Status, _
            run_result.ErrorDescription, _
            run_result.DetailValues, _
            run_result.DownloadStatus)

    If run_result.Status = G_WEB_STATUS_ERROR Then
        pErrorCount = pErrorCount + 1
    Else
        pSucceededCount = pSucceededCount + 1
    End If
End Sub

Private Function pReadListItemTargetId(ByVal ItemIndex As Long) As String
    Dim selector_value As String
    selector_value = pIndexTemplate.Apply(pSettings.ListItemTargetIdSelector, ItemIndex)

    Dim result_value As String
    result_value = Trim$(pSessionClient.ReadElementText(selector_value))
    If result_value = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner.pReadListItemTargetId", Description:="一覧上の対象IDを取得できません。(index=" & CStr(ItemIndex) & ")"
    End If

    pReadListItemTargetId = result_value
End Function

Private Sub pRaiseIfListItemChanged( _
        ByVal ItemIndex As Long, _
        ByVal ExpectedTargetId As String, _
        ByVal ActualTargetId As String)

    If StrComp(ExpectedTargetId, ActualTargetId, vbBinaryCompare) = 0 Then Exit Sub

    Err.Raise Number:=vbObjectError + 1, _
            Source:="Class WebCollectionRunner.pRaiseIfListItemChanged", _
            Description:="一覧項目が収集中に変化しました。(index=" & CStr(ItemIndex) & ", expected=" & ExpectedTargetId & ", actual=" & ActualTargetId & ")"
End Sub

Private Function pHasNextPage() As Boolean
    If Trim$(pSettings.NextPageAvailableSelector) = "" Then
        pHasNextPage = False
        Exit Function
    End If

    pHasNextPage = (0 < pSessionClient.CountElements(pSettings.NextPageAvailableSelector))
End Function

Private Sub pMoveToNextPage(ByVal PreviousFirstTargetId As String)
    Dim next_operation As TransitionOperation
    Set next_operation = pTransitionResolver.ResolveNextPageOperation()
    Call pSessionClient.RunTransitionOperation(next_operation)
    Call pSessionClient.WaitForSelector(pSettings.ListPageSelector)
    Call pWaitForFirstTargetChange(PreviousFirstTargetId)
End Sub

Private Sub pWaitForFirstTargetChange(ByVal PreviousFirstTargetId As String)
    Dim deadline_time As Date
    deadline_time = DateAdd("s", pSettings.TimeoutSeconds, Now)

    Dim last_error_description As String
    Dim has_read_target As Boolean
    Dim current_target_id As String

    Do
        On Error Resume Next
        current_target_id = pReadListItemTargetId(0)
        If Err.Number = 0 Then
            has_read_target = True
            If StrComp(current_target_id, PreviousFirstTargetId, vbBinaryCompare) <> 0 Then
                On Error GoTo 0
                Exit Sub
            End If
        Else
            last_error_description = Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        DoEvents
    Loop While Now < deadline_time

    If Not has_read_target Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner.pWaitForFirstTargetChange", Description:="次ページ操作後に先頭対象IDを取得できません。" & vbCrLf & last_error_description
    End If

    Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner.pWaitForFirstTargetChange", Description:="次ページ操作後にページが進みません。(" & PreviousFirstTargetId & ")"
End Sub
