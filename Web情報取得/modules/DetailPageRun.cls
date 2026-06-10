VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailPageRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 一覧項目から詳細ページ 1 件分を処理します。
'!
' #############################################################################

Private pSessionClient As WebDriverSessionClient
Private pSettings As IToolSettings
Private pConditionEvaluator As OutputConditionEvaluator
Private pIsInitialized As Boolean

'* 詳細ページ run を初期化します。
Public Sub Initialize( _
        ByVal SessionClient As WebDriverSessionClient, _
        ByVal Settings As IToolSettings)

    If SessionClient Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRun", Description:="SessionClient は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRun", Description:="Settings は必須です。"
    End If

    Set pSessionClient = SessionClient
    Set pSettings = Settings
    Set pConditionEvaluator = New OutputConditionEvaluator
    Call pConditionEvaluator.Initialize(pSettings.OutputConditionExpression, pSettings.DetailColumnDefinitions)
    pIsInitialized = True
End Sub

'* 詳細ページ 1 件分を処理します。
Public Function Run( _
        ByVal ListTargetId As String, _
        ByVal DetailOperationName As String, _
        Optional ByVal ItemIndex As Long = 0, _
        Optional ByVal IsDiagnosticMode As Boolean = False) As DetailPageRunResult

    Call pCheckInit

    Dim output_target_id As String
    output_target_id = ListTargetId

    Dim needs_return As Boolean
    needs_return = False

    Dim has_valid_detail_target As Boolean
    has_valid_detail_target = False

    On Error GoTo DETAIL_ERROR

    Dim detail_operation As TransitionOperation
    Set detail_operation = pBuildIndexedTransitionOperation(DetailOperationName, ItemIndex)
    Call pSessionClient.RunTransitionOperation(detail_operation)
    needs_return = True

    Call pSessionClient.WaitForSelector(pSettings.TargetIdSelector)

    Dim detail_target_id As String
    detail_target_id = Trim$(pSessionClient.ReadElementText(pSettings.TargetIdSelector))
    If detail_target_id = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRun", Description:="詳細ページ上の対象IDを取得できません。"
    End If

    output_target_id = detail_target_id
    If StrComp(ListTargetId, detail_target_id, vbBinaryCompare) <> 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRun", Description:="一覧上の対象IDと詳細ページ上の対象IDが一致しません。(" & ListTargetId & " / " & detail_target_id & ")"
    End If
    has_valid_detail_target = True

    If Not pShouldOutputByCondition() Then
        Dim excluded_values As ArrayObject
        Set excluded_values = pCreateBlankDetailValues(pSettings.DetailColumnDefinitions)

        Dim excluded_result As DetailPageRunResult
        Set excluded_result = New_DetailPageRunResult(output_target_id, "", "", excluded_values, IsOutputExcluded:=True)

        On Error GoTo 0
        Call pReturnToListIfNeeded(needs_return)
        Set Run = excluded_result
        Exit Function
    End If

    Dim detail_values As ArrayObject
    Set detail_values = pReadDetailValues()

    Dim download_error_description As String
    Dim download_status As String
    download_status = pDownloadCurrentDetailFile(output_target_id, download_error_description)

    Dim row_status As String
    row_status = G_WEB_STATUS_OK

    Dim row_error_description As String
    row_error_description = ""
    If pDownloadStatusRequiresRowError(download_status) Then
        row_status = G_WEB_STATUS_ERROR
        row_error_description = download_error_description
    End If

    Dim result_value As DetailPageRunResult
    Set result_value = New_DetailPageRunResult(output_target_id, row_status, row_error_description, detail_values, download_status)

    On Error GoTo 0
    Call pReturnToListIfNeeded(needs_return)

    Set Run = result_value
    Exit Function

DETAIL_ERROR:
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
    Err.Clear
    On Error GoTo 0

    If IsDiagnosticMode And Not has_valid_detail_target Then
        If needs_return Then Call pReturnToListIfNeeded(needs_return)
        Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
    End If

    Dim error_values As ArrayObject
    Set error_values = pCreateBlankDetailValues(pSettings.DetailColumnDefinitions)

    Dim error_result As DetailPageRunResult
    Set error_result = New_DetailPageRunResult(output_target_id, G_WEB_STATUS_ERROR, err_desc, error_values)

    If needs_return Then Call pReturnToListIfNeeded(needs_return)

    Set Run = error_result
End Function

Private Function pReadDetailValues() As ArrayObject
    Dim detail_defs As ObjectList
    Set detail_defs = pSettings.DetailColumnDefinitions

    If detail_defs.Count = 0 Then
        Set pReadDetailValues = pCreateBlankDetailValues(detail_defs)
        Exit Function
    End If

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver
    Call resolver.Initialize(detail_defs)

    Dim extraction_defs As ObjectList
    Set extraction_defs = resolver.ExtractionColumnDefinitions

    Dim extraction_values As ArrayObject
    If extraction_defs.Count = 0 Then
        Set extraction_values = pCreateBlankDetailValues(extraction_defs)
    Else
        Set extraction_values = pSessionClient.ReadDetailColumnValues(extraction_defs)
    End If

    Set pReadDetailValues = resolver.ResolveValues(extraction_values)
End Function

Private Function pDownloadCurrentDetailFile( _
        ByVal TargetId As String, _
        ByRef ErrorDescription As String) As String

    ErrorDescription = ""
    If Not pSettings.DownloadEnabled Then
        pDownloadCurrentDetailFile = ""
        Exit Function
    End If

    On Error GoTo DOWNLOAD_ERROR

    Dim downloaded_file_path As String
    Dim result_status As String
    result_status = pSessionClient.DownloadLinkedFile(downloaded_file_path)

    If result_status = G_WEB_DOWNLOAD_STATUS_DOWNLOADED Then
        Dim file_saver As DownloadedFileSaver
        Set file_saver = New_DownloadedFileSaver(pSettings, FsSrv)
        result_status = file_saver.SaveDownloadedFile(downloaded_file_path, TargetId)
        If result_status = G_WEB_DOWNLOAD_STATUS_ERROR Then ErrorDescription = file_saver.LastErrorDescription
    ElseIf result_status = G_WEB_DOWNLOAD_STATUS_ERROR Then
        ErrorDescription = pSessionClient.LastDownloadErrorDescription
    ElseIf result_status = G_WEB_DOWNLOAD_STATUS_NO_FILE Then
        ErrorDescription = "ダウンロードリンクがありません。"
    End If

    On Error GoTo 0
    pDownloadCurrentDetailFile = result_status
    Exit Function

DOWNLOAD_ERROR:
    ErrorDescription = Err.Description
    Err.Clear
    On Error GoTo 0

    pDownloadCurrentDetailFile = G_WEB_DOWNLOAD_STATUS_ERROR
End Function

Private Function pDownloadStatusRequiresRowError(ByVal DownloadStatus As String) As Boolean
    pDownloadStatusRequiresRowError = (pSettings.DownloadRequired _
            And (DownloadStatus = G_WEB_DOWNLOAD_STATUS_ERROR Or DownloadStatus = G_WEB_DOWNLOAD_STATUS_NO_FILE))
End Function

Private Function pShouldOutputByCondition() As Boolean
    If Not pConditionEvaluator.IsEnabled Then
        pShouldOutputByCondition = True
        Exit Function
    End If

    Dim referenced_values As ArrayObject
    Set referenced_values = pSessionClient.ReadDetailColumnValues(pConditionEvaluator.ReferencedColumnDefinitions)

    pShouldOutputByCondition = pConditionEvaluator.ShouldOutputByReferencedValues(referenced_values)
End Function

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

Private Sub pReturnToListIfNeeded(ByVal NeedsReturn As Boolean)
    If Not NeedsReturn Then Exit Sub
    If Trim$(pSettings.ReturnToListOperationName) = "" Then Exit Sub

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

    Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRun", Description:="画面遷移操作が見つかりません。(" & OperationName & ")"
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRun", Description:="初期化されていません。"
    End If
End Sub
