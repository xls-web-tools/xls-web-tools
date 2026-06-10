VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailPageRunner"
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
Private pTransitionResolver As TransitionOperationResolver
Private pConditionEvaluator As OutputConditionEvaluator
Private pValueResolver As DetailValueResolver
Private pFileDownloader As DetailFileDownloader
Private pIsInitialized As Boolean

'* 詳細ページ runner を初期化します。
'*
'* @param SessionClient WebDriver session client。
'* @param Settings Web情報取得の設定。
'* @param FileDownloader 詳細ページ上のファイルダウンロード処理。DownloadEnabled=True の場合は必須。
'*
'* @details
'* 一覧収集とナビゲーション診断で同じ詳細ページ 1 件分の処理を共有するため、browser session と設定を外部から受け取ります。
Public Sub Initialize( _
        ByVal SessionClient As WebDriverSessionClient, _
        ByVal Settings As IToolSettings, _
        Optional ByVal FileDownloader As DetailFileDownloader = Nothing)

    If SessionClient Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRunner", Description:="SessionClient は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRunner", Description:="Settings は必須です。"
    End If
    If Settings.DownloadEnabled And (FileDownloader Is Nothing) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRunner", Description:="FileDownloader は必須です。"
    End If

    Set pSessionClient = SessionClient
    Set pSettings = Settings
    Set pTransitionResolver = New_TransitionOperationResolver(Settings)
    Set pFileDownloader = FileDownloader
    Set pConditionEvaluator = New OutputConditionEvaluator
    Call pConditionEvaluator.Initialize(pSettings.OutputConditionExpression, pSettings.DetailColumnDefinitions)

    Set pValueResolver = New DetailValueResolver
    Call pValueResolver.Initialize(pSettings.DetailColumnDefinitions)

    pIsInitialized = True
End Sub

'* 詳細ページ 1 件分を処理します。
'*
'* @param ListTargetId 一覧上で取得した対象ID。
'* @param ItemIndex 一覧上の項目インデックス。locator template の {{index}} と {{rowNumber}} に使用します。
'* @param IsDiagnosticMode True の場合、対象ID確認前のエラーは診断 runner へ再送出します。
'* @return 詳細ページ 1 件分の処理結果。
'*
'* @details
'* 詳細ページへ遷移し、対象ID照合、出力条件判定、詳細値抽出、ダウンロード処理を実行してから一覧へ戻ります。
Public Function Run( _
        ByVal ListTargetId As String, _
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
    Set detail_operation = pTransitionResolver.ResolveDetailTransitionOperation(ItemIndex)
    Call pSessionClient.RunTransitionOperation(detail_operation)
    needs_return = True

    Call pSessionClient.WaitForSelector(pSettings.TargetIdSelector)

    Dim detail_target_id As String
    detail_target_id = Trim$(pSessionClient.ReadElementText(pSettings.TargetIdSelector))
    If detail_target_id = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRunner", Description:="詳細ページ上の対象IDを取得できません。"
    End If

    output_target_id = detail_target_id
    If StrComp(ListTargetId, detail_target_id, vbBinaryCompare) <> 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRunner", Description:="一覧上の対象IDと詳細ページ上の対象IDが一致しません。(" & ListTargetId & " / " & detail_target_id & ")"
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

    Dim download_result As DetailFileDownloadResult
    Set download_result = pDownloadCurrentDetailFile(output_target_id)

    Dim row_status As String
    row_status = G_WEB_STATUS_OK

    Dim row_error_description As String
    row_error_description = ""
    If download_result.RequiresDetailError Then
        row_status = G_WEB_STATUS_ERROR
        row_error_description = download_result.ErrorDescription
    End If

    Dim result_value As DetailPageRunResult
    Set result_value = New_DetailPageRunResult(output_target_id, row_status, row_error_description, detail_values, download_result.DownloadStatus)

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

    Dim extraction_defs As ObjectList
    Set extraction_defs = pValueResolver.ExtractionColumnDefinitions

    Dim extraction_values As ArrayObject
    If extraction_defs.Count = 0 Then
        Set extraction_values = pCreateBlankDetailValues(extraction_defs)
    Else
        Set extraction_values = pSessionClient.ReadDetailColumnValues(extraction_defs)
    End If

    Set pReadDetailValues = pValueResolver.ResolveValues(extraction_values)
End Function

Private Function pDownloadCurrentDetailFile(ByVal TargetId As String) As DetailFileDownloadResult
    If Not pSettings.DownloadEnabled Then
        Set pDownloadCurrentDetailFile = New_DetailFileDownloadResult("", "", False)
        Exit Function
    End If

    If pFileDownloader Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRunner", Description:="FileDownloader は必須です。"
    End If

    Set pDownloadCurrentDetailFile = pFileDownloader.DownloadForDetail(TargetId)
End Function

Private Function pShouldOutputByCondition() As Boolean
    If Not pConditionEvaluator.IsEnabled Then
        pShouldOutputByCondition = True
        Exit Function
    End If

    Dim condition_defs As ObjectList
    Set condition_defs = pValueResolver.BuildDependentDefinitions( _
            pConditionEvaluator.ReferencedColumnDefinitions, _
            OptionalExtraction:=True, _
            IgnoreDerivedBlankMode:=True)

    Dim condition_resolver As DetailValueResolver
    Set condition_resolver = New DetailValueResolver
    Call condition_resolver.Initialize(condition_defs)

    Dim extraction_defs As ObjectList
    Set extraction_defs = condition_resolver.ExtractionColumnDefinitions

    Dim extraction_values As ArrayObject
    If extraction_defs.Count = 0 Then
        Set extraction_values = pCreateBlankDetailValues(extraction_defs)
    Else
        Set extraction_values = pSessionClient.ReadDetailColumnValues(extraction_defs)
    End If

    Dim referenced_values As ArrayObject
    Set referenced_values = condition_resolver.ResolveValues(extraction_values, ValidateDerivedBlankMode:=False)

    pShouldOutputByCondition = pConditionEvaluator.ShouldOutputByColumnValues(referenced_values, condition_defs)
End Function

Private Sub pReturnToListIfNeeded(ByVal NeedsReturn As Boolean)
    If Not NeedsReturn Then Exit Sub
    If Not pTransitionResolver.HasReturnToListOperation Then Exit Sub

    On Error GoTo RETURN_ERROR

    Dim return_operation As TransitionOperation
    Set return_operation = pTransitionResolver.ResolveReturnToListOperation()
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

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRunner", Description:="初期化されていません。"
    End If
End Sub
