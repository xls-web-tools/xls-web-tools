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

' #############################################################################
'!
'! @brief
'! 一覧項目を詳細ページ経由で収集します。
'!
' #############################################################################

Private Const C_PROGRESS_TASK_NAME As String = "取得中"

Private pProcess As IWebDriverProcess
Private pSessionClient As WebDriverSessionClient
Private pSettings As IToolSettings
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
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="Process は必須です。"
    End If
    If SessionClient Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="SessionClient は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="Settings は必須です。"
    End If

    Set pProcess = Process
    Set pSessionClient = SessionClient
    Set pSettings = Settings
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

    Dim result_value As String
    pSucceededCount = 0
    pSkippedCount = 0
    pOutputExcludedCount = 0
    pErrorCount = 0
    pPageCount = 0
    pProgressStarted = False

    Dim detail_run As DetailPageRun
    Set detail_run = New_DetailPageRun(pSessionClient, pSettings)
    On Error GoTo ON_ERROR

    Call pProcess.Start(pSettings)

    result_value = pSessionClient.CreateSession()
    Call pSessionClient.NavigateToStartUrl
    Call pSessionClient.WaitForSelector(pSettings.AuthenticatedStartSelector)
    Call pOpenListPage

    Dim output_writer As OutputSheetWriter
    Set output_writer = New_OutputSheetWriter(pSettings)
    Call output_writer.EnsureHeaders
    Call pCollectPages(output_writer, detail_run)
    Call pMarkProgressComplete

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

    On Error Resume Next
    Call pClearProgress
    If result_value <> "" Then Call pSessionClient.QuitSession(result_value)
    Call pProcess.StopProcess
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function

Private Sub pOpenListPage()
    Dim list_operation As TransitionOperation
    Set list_operation = pFindTransitionOperation(pSettings.ListTransitionOperationName)
    Call pSessionClient.RunTransitionOperation(list_operation)
    Call pSessionClient.WaitForSelector(pSettings.ListPageSelector)
End Sub

Private Sub pCollectPages(ByVal OutputWriter As OutputSheetWriter, ByVal DetailRun As DetailPageRun)
    Dim first_target_ids As ObjectSet
    Set first_target_ids = New ObjectSet

    Do
        pPageCount = pPageCount + 1

        Dim item_count As Long
        item_count = pSessionClient.CountElements(pSettings.ListItemSelector)
        If item_count = 0 Then
            If pPageCount = 1 Then Exit Sub
            Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="次ページ操作後に一覧項目がありません。"
        End If
        Call pAddProgressTotal(item_count)

        Dim target_snapshot As ArrayObject
        Set target_snapshot = pBuildListTargetSnapshot(item_count)

        Dim first_target_id As String
        first_target_id = CStr(target_snapshot.Item(0))
        If first_target_ids.Exists(first_target_id) Then
            Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="ページングが循環しました。(" & first_target_id & ")"
        End If
        Call first_target_ids.Add(first_target_id)

        Dim item_idx As Long
        For item_idx = 0 To item_count - 1
            Call pCollectListItem(OutputWriter, target_snapshot, item_idx, DetailRun)
            Call pIncrementProgress
        Next item_idx

        If Not pHasNextPage() Then Exit Do
        Call pMoveToNextPage(first_target_id)
    Loop
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

Private Sub pMarkProgressComplete()
    If ProgStat Is Nothing Then Exit Sub
    If Not pProgressStarted Then Exit Sub

    Call ProgStat.MarkAsComplete
End Sub

Private Sub pClearProgress()
    If ProgStat Is Nothing Then Exit Sub
    If Not pProgressStarted Then Exit Sub

    Call ProgStat.Reset(ResetTaskName:=False, ResetStartAndTotalValue:=False)
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
        ByVal DetailRun As DetailPageRun)

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
    Set run_result = DetailRun.Run(current_target_id, pSettings.DetailTransitionOperationName, ItemIndex:=ItemIndex)

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
    selector_value = pApplyIndexTemplate(pSettings.ListItemTargetIdSelector, ItemIndex)

    Dim result_value As String
    result_value = Trim$(pSessionClient.ReadElementText(selector_value))
    If result_value = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="一覧上の対象IDを取得できません。(index=" & CStr(ItemIndex) & ")"
    End If

    pReadListItemTargetId = result_value
End Function

Private Sub pRaiseIfListItemChanged( _
        ByVal ItemIndex As Long, _
        ByVal ExpectedTargetId As String, _
        ByVal ActualTargetId As String)

    If StrComp(ExpectedTargetId, ActualTargetId, vbBinaryCompare) = 0 Then Exit Sub

    Err.Raise Number:=vbObjectError + 1, _
            Source:="Class WebCollectionRunner", _
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
    Set next_operation = pFindTransitionOperation(pSettings.NextPageOperationName)
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
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="次ページ操作後に先頭対象IDを取得できません。" & vbCrLf & last_error_description
    End If

    Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="次ページ操作後にページが進みません。(" & PreviousFirstTargetId & ")"
End Sub

Private Function pApplyIndexTemplate(ByVal TargetText As String, ByVal ItemIndex As Long) As String
    Dim result_value As String
    result_value = TargetText
    result_value = Replace(result_value, "{{index}}", CStr(ItemIndex))
    result_value = Replace(result_value, "{{rowNumber}}", CStr(ItemIndex + 1))

    pApplyIndexTemplate = result_value
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

    Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="画面遷移操作が見つかりません。(" & OperationName & ")"
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="初期化されていません。"
    End If
End Sub
