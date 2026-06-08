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
'! 現在ページの一覧項目を詳細ページ経由で収集します。
'!
' #############################################################################

Private pProcess As IWebDriverProcess
Private pSessionClient As WebDriverSessionClient
Private pSettings As IToolSettings
Private pSucceededCount As Long
Private pSkippedCount As Long
Private pErrorCount As Long
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

'* ERROR 行として出力した件数。
Public Property Get ErrorCount() As Long
    Call pCheckInit
    ErrorCount = pErrorCount
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
    pErrorCount = 0
    pIsInitialized = True
End Sub

'* 現在ページの本番収集を実行します。
'*
'* @return 作成された session id。
'*
'* @details
'* 取得開始 URL から一覧画面へ移動し、現在ページの一覧項目だけを表示順に巡回します。
'* 対象IDは一覧表示時点で snapshot 化し、各項目処理直前に同じ位置の対象IDが変化していないことを確認します。
'* 既存 OK 行のスキップや ERROR 行の再取得は OutputSheetWriter の ExistingRowMode 判定に委譲します。
Public Function Run() As String
    Call pCheckInit

    Dim result_value As String
    pSucceededCount = 0
    pSkippedCount = 0
    pErrorCount = 0

    On Error GoTo ON_ERROR

    Call pProcess.Start(pSettings)

    result_value = pSessionClient.CreateSession()
    Call pSessionClient.NavigateToStartUrl
    Call pSessionClient.WaitForSelector(pSettings.AuthenticatedStartSelector)
    Call pOpenListPage

    Dim output_writer As OutputSheetWriter
    Set output_writer = New_OutputSheetWriter(pSettings)
    Call output_writer.EnsureHeaders

    Dim item_count As Long
    item_count = pSessionClient.CountElements(pSettings.ListItemSelector)

    Dim target_snapshot As ArrayObject
    Set target_snapshot = pBuildListTargetSnapshot(item_count)

    Dim item_idx As Long
    For item_idx = 0 To item_count - 1
        Call pCollectListItem(output_writer, target_snapshot, item_idx)
    Next item_idx

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
        ByVal ItemIndex As Long)

    Dim expected_target_id As String
    expected_target_id = CStr(TargetSnapshot.Item(ItemIndex))

    Dim current_target_id As String
    current_target_id = pReadListItemTargetId(ItemIndex)
    Call pRaiseIfListItemChanged(ItemIndex, expected_target_id, current_target_id)

    If Not OutputWriter.ShouldCollectTarget(current_target_id) Then
        pSkippedCount = pSkippedCount + 1
        Exit Sub
    End If

    Dim output_target_id As String
    output_target_id = current_target_id

    Dim needs_return As Boolean
    needs_return = False

    On Error GoTo ITEM_ERROR

    Dim detail_operation As TransitionOperation
    Set detail_operation = pBuildIndexedTransitionOperation(pSettings.DetailTransitionOperationName, ItemIndex)
    Call pSessionClient.RunTransitionOperation(detail_operation)
    needs_return = True

    Call pSessionClient.WaitForSelector(pSettings.TargetIdSelector)

    Dim detail_target_id As String
    detail_target_id = Trim$(pSessionClient.ReadElementText(pSettings.TargetIdSelector))
    If detail_target_id = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="詳細ページ上の対象IDを取得できません。"
    End If

    output_target_id = detail_target_id
    If StrComp(current_target_id, detail_target_id, vbBinaryCompare) <> 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunner", Description:="一覧上の対象IDと詳細ページ上の対象IDが一致しません。(" & current_target_id & " / " & detail_target_id & ")"
    End If

    Dim detail_values As ArrayObject
    Set detail_values = pSessionClient.ReadDetailColumnValues(pSettings.DetailColumnDefinitions)
    Call OutputWriter.WriteCollectionRow("", output_target_id, G_WEB_STATUS_OK, "", detail_values)
    pSucceededCount = pSucceededCount + 1

    On Error GoTo 0
    Call pReturnToList
    Exit Sub

ITEM_ERROR:
    Dim err_desc As String
    err_desc = Err.Description
    Err.Clear
    On Error GoTo 0

    Call pWriteErrorCollectionRow(OutputWriter, output_target_id, err_desc)
    pErrorCount = pErrorCount + 1
    If needs_return Then Call pReturnToList
End Sub

Private Sub pWriteErrorCollectionRow( _
        ByVal OutputWriter As OutputSheetWriter, _
        ByVal TargetId As String, _
        ByVal ErrorDescription As String)

    Dim detail_values As ArrayObject
    Set detail_values = pCreateBlankDetailValues(pSettings.DetailColumnDefinitions)

    Call OutputWriter.WriteCollectionRow("", TargetId, G_WEB_STATUS_ERROR, ErrorDescription, detail_values)
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

Private Function pBuildIndexedTransitionOperation( _
        ByVal OperationName As String, _
        ByVal ItemIndex As Long) As TransitionOperation

    Dim base_operation As TransitionOperation
    Set base_operation = pFindTransitionOperation(OperationName)

    Set pBuildIndexedTransitionOperation = New_TransitionOperation( _
            base_operation.OperationName, _
            base_operation.LocatorType, _
            pApplyIndexTemplate(base_operation.LocatorValue, ItemIndex), _
            Script:=pApplyIndexTemplate(base_operation.Script, ItemIndex), _
            WaitConditionName:=base_operation.WaitConditionName)
End Function

Private Function pApplyIndexTemplate(ByVal TargetText As String, ByVal ItemIndex As Long) As String
    Dim result_value As String
    result_value = TargetText
    result_value = Replace(result_value, "{{index}}", CStr(ItemIndex))
    result_value = Replace(result_value, "{{rowNumber}}", CStr(ItemIndex + 1))

    pApplyIndexTemplate = result_value
End Function

Private Sub pReturnToList()
    On Error GoTo RETURN_ERROR

    Dim return_operation As TransitionOperation
    Set return_operation = pFindTransitionOperation(pSettings.ReturnToListOperationName)
    Call pSessionClient.RunTransitionOperation(return_operation)
    Call pSessionClient.WaitForSelector(pSettings.ListPageSelector)

    On Error GoTo 0
    Exit Sub

RETURN_ERROR:
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

    Err.Raise err_num, err_source, "一覧画面へ復帰できません。" & vbCrLf & err_desc, err_help_file, err_help_context
End Sub

Private Function pCreateBlankDetailValues(ByVal DetailDefinitions As ObjectList) As ArrayObject
    Dim result_values As ArrayObject
    Set result_values = New ArrayObject

    If 0 < DetailDefinitions.Count Then
        Call result_values.ReDimArray(0, DetailDefinitions.Count - 1)
    End If

    Dim def_idx As Long
    For def_idx = 0 To DetailDefinitions.Count - 1
        Call result_values.Update(def_idx, "")
    Next def_idx

    Set pCreateBlankDetailValues = result_values
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