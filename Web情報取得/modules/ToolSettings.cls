VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IToolSettings

' #############################################################################
'!
'! @brief
'! settings シート由来の Web情報取得設定です。
'!
' #############################################################################

Private Const C_PARAM_COL_NAME As Long = 1
Private Const C_PARAM_COL_VALUE As Long = 2
Private Const C_PARAM_START_ROW As Long = 2
Private Const C_TRANSITION_COL_START As Long = 5
Private Const C_TRANSITION_COL_FINISH As Long = 10
Private Const C_DETAIL_COL_START As Long = 12
Private Const C_DETAIL_COL_FINISH As Long = 18
Private Const C_DEFAULT_WEB_DRIVER_PORT As Long = 9515
Private Const C_DEFAULT_TIMEOUT_SECONDS As Long = 30

'* Edge WebDriver 実行ファイルのパス。
Public Property Get WebDriverPath() As String
    WebDriverPath = pResolvePath(pGetOptionalParam("WebDriverPath", "bin" & G_FS_PATH_SEP & "msedgedriver.exe"))
End Property
Private Property Get IToolSettings_WebDriverPath() As String
    IToolSettings_WebDriverPath = WebDriverPath
End Property

'* Edge WebDriver の待ち受け port。
Public Property Get WebDriverPort() As Long
    WebDriverPort = pGetOptionalLongParam("WebDriverPort", C_DEFAULT_WEB_DRIVER_PORT)
End Property
Private Property Get IToolSettings_WebDriverPort() As Long
    IToolSettings_WebDriverPort = WebDriverPort
End Property

'* 専用ブラウザプロファイルのパス。
Public Property Get BrowserProfilePath() As String
    BrowserProfilePath = pResolvePath(pGetOptionalParam("BrowserProfilePath", "browser-profile"))
End Property
Private Property Get IToolSettings_BrowserProfilePath() As String
    IToolSettings_BrowserProfilePath = BrowserProfilePath
End Property

'* True の場合は不可視ブラウザで起動します。
Public Property Get Headless() As Boolean
    Headless = pGetOptionalBooleanParam("Headless", True)
End Property
Private Property Get IToolSettings_Headless() As Boolean
    IToolSettings_Headless = Headless
End Property

'* 取得開始 URL。
Public Property Get StartUrl() As String
    StartUrl = pGetRequiredParam("StartUrl")
End Property
Private Property Get IToolSettings_StartUrl() As String
    IToolSettings_StartUrl = StartUrl
End Property

'* 取得結果の出力先シート名。
Public Property Get OutputSheetName() As String
    OutputSheetName = pGetRequiredParam("OutputSheetName")
End Property
Private Property Get IToolSettings_OutputSheetName() As String
    IToolSettings_OutputSheetName = OutputSheetName
End Property

'* 認証後の起点ページ到達を判定する selector。
Public Property Get AuthenticatedStartSelector() As String
    AuthenticatedStartSelector = pGetRequiredParam("AuthenticatedStartSelector")
End Property
Private Property Get IToolSettings_AuthenticatedStartSelector() As String
    IToolSettings_AuthenticatedStartSelector = AuthenticatedStartSelector
End Property

'* 一覧画面到達を判定する selector。
Public Property Get ListPageSelector() As String
    ListPageSelector = pGetRequiredParam("ListPageSelector")
End Property
Private Property Get IToolSettings_ListPageSelector() As String
    IToolSettings_ListPageSelector = ListPageSelector
End Property

'* 一覧画面への画面遷移操作名。
Public Property Get ListTransitionOperationName() As String
    ListTransitionOperationName = pGetRequiredParam("ListTransitionOperationName")
End Property
Private Property Get IToolSettings_ListTransitionOperationName() As String
    IToolSettings_ListTransitionOperationName = ListTransitionOperationName
End Property

'* 現在ページの一覧項目を数える selector。
Public Property Get ListItemSelector() As String
    ListItemSelector = pGetRequiredParam("ListItemSelector")
End Property
Private Property Get IToolSettings_ListItemSelector() As String
    IToolSettings_ListItemSelector = ListItemSelector
End Property

'* 一覧上の対象IDを読み取る selector。
Public Property Get ListItemTargetIdSelector() As String
    ListItemTargetIdSelector = pGetRequiredParam("ListItemTargetIdSelector")
End Property
Private Property Get IToolSettings_ListItemTargetIdSelector() As String
    IToolSettings_ListItemTargetIdSelector = ListItemTargetIdSelector
End Property

'* 詳細ページへの画面遷移操作名。
Public Property Get DetailTransitionOperationName() As String
    DetailTransitionOperationName = pGetRequiredParam("DetailTransitionOperationName")
End Property
Private Property Get IToolSettings_DetailTransitionOperationName() As String
    IToolSettings_DetailTransitionOperationName = DetailTransitionOperationName
End Property

'* 詳細ページ上の対象IDを読み取る selector。
Public Property Get TargetIdSelector() As String
    TargetIdSelector = pGetRequiredParam("TargetIdSelector")
End Property
Private Property Get IToolSettings_TargetIdSelector() As String
    IToolSettings_TargetIdSelector = TargetIdSelector
End Property

'* 一覧画面へ戻る画面遷移操作名。
Public Property Get ReturnToListOperationName() As String
    ReturnToListOperationName = pGetRequiredParam("ReturnToListOperationName")
End Property
Private Property Get IToolSettings_ReturnToListOperationName() As String
    IToolSettings_ReturnToListOperationName = ReturnToListOperationName
End Property

'* 次ページへ進む画面遷移操作名。
Public Property Get NextPageOperationName() As String
    NextPageOperationName = pGetOptionalParam("NextPageOperationName", "NextPage")
End Property
Private Property Get IToolSettings_NextPageOperationName() As String
    IToolSettings_NextPageOperationName = NextPageOperationName
End Property

'* 次ページ操作が有効な場合だけ一致する selector。
Public Property Get NextPageAvailableSelector() As String
    NextPageAvailableSelector = pGetOptionalParam("NextPageAvailableSelector", "")
End Property
Private Property Get IToolSettings_NextPageAvailableSelector() As String
    IToolSettings_NextPageAvailableSelector = NextPageAvailableSelector
End Property

'* 既存行の扱い。
Public Property Get ExistingRowMode() As String
    ExistingRowMode = pNormalizeExistingRowMode(pGetOptionalParam("ExistingRowMode", G_WEB_ROW_MODE_SKIP_EXISTING))
End Property
Private Property Get IToolSettings_ExistingRowMode() As String
    IToolSettings_ExistingRowMode = ExistingRowMode
End Property

'* 出力対象条件式。
Public Property Get OutputConditionExpression() As String
    OutputConditionExpression = pGetOptionalParam("OutputConditionExpression", "")
End Property
Private Property Get IToolSettings_OutputConditionExpression() As String
    IToolSettings_OutputConditionExpression = OutputConditionExpression
End Property

'* True の場合はダウンロード対象ファイルを取得します。
Public Property Get DownloadEnabled() As Boolean
    DownloadEnabled = pGetOptionalBooleanParam("DownloadEnabled", False)
End Property
Private Property Get IToolSettings_DownloadEnabled() As Boolean
    IToolSettings_DownloadEnabled = DownloadEnabled
End Property

'* True の場合はダウンロード失敗を詳細ページ単位の ERROR にします。
Public Property Get DownloadRequired() As Boolean
    DownloadRequired = pGetOptionalBooleanParam("DownloadRequired", False)
End Property
Private Property Get IToolSettings_DownloadRequired() As Boolean
    IToolSettings_DownloadRequired = DownloadRequired
End Property

'* ダウンロード保存先ルート。
Public Property Get DownloadRootPath() As String
    If DownloadEnabled Then
        DownloadRootPath = pGetRequiredParam("DownloadRootPath")
    Else
        DownloadRootPath = pGetOptionalParam("DownloadRootPath", "")
    End If
End Property
Private Property Get IToolSettings_DownloadRootPath() As String
    IToolSettings_DownloadRootPath = DownloadRootPath
End Property

'* 詳細ページ上のダウンロードリンクを特定する selector。
Public Property Get DownloadLinkSelector() As String
    If DownloadEnabled Then
        DownloadLinkSelector = pGetRequiredParam("DownloadLinkSelector")
    Else
        DownloadLinkSelector = pGetOptionalParam("DownloadLinkSelector", "")
    End If
End Property
Private Property Get IToolSettings_DownloadLinkSelector() As String
    IToolSettings_DownloadLinkSelector = DownloadLinkSelector
End Property

'* 画面到達や selector 出現を待つ秒数。
Public Property Get TimeoutSeconds() As Long
    TimeoutSeconds = pGetOptionalLongParam("TimeoutSeconds", C_DEFAULT_TIMEOUT_SECONDS)
End Property
Private Property Get IToolSettings_TimeoutSeconds() As Long
    IToolSettings_TimeoutSeconds = TimeoutSeconds
End Property

'* 画面遷移操作定義の一覧。
Public Property Get TransitionOperations() As ObjectList
    Set TransitionOperations = pReadTransitionOperations()
End Property
Private Property Get IToolSettings_TransitionOperations() As ObjectList
    Set IToolSettings_TransitionOperations = TransitionOperations
End Property

'* 詳細ページ列定義の一覧。
Public Property Get DetailColumnDefinitions() As ObjectList
    Set DetailColumnDefinitions = pReadDetailColumnDefinitions()
End Property
Private Property Get IToolSettings_DetailColumnDefinitions() As ObjectList
    Set IToolSettings_DetailColumnDefinitions = DetailColumnDefinitions
End Property

Private Function pGetRequiredParam(ByVal Name As String) As String
    Dim result_value As String
    result_value = pGetOptionalParam(Name, "")
    If Trim$(result_value) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings", Description:=Name & " は必須です。"
    End If

    pGetRequiredParam = result_value
End Function

Private Function pGetOptionalParam(ByVal Name As String, ByVal DefaultValue As String) As String
    Dim param_range As WorksheetRangeBounds
    Set param_range = WsSrv.GetUsedRangeBounds(New_RangeBounds( _
            Row:=C_PARAM_START_ROW, _
            Column:=C_PARAM_COL_NAME, _
            FinishRow:=G_ROW_MAX, _
            FinishColumn:=C_PARAM_COL_NAME, _
            Sheet:=G_WEB_SETTINGS_SHEET))

    Dim found_cells As ObjectList
    Set found_cells = WsSrv.Find(Name, param_range)
    If found_cells.Count = 0 Then
        pGetOptionalParam = DefaultValue
        Exit Function
    End If

    Dim found_bounds As WorksheetRangeBounds
    Set found_bounds = found_cells.Item(0)

    Dim result_value As String
    result_value = "" & WsSrv.ReadCell(New_RangeBounds(Row:=found_bounds.Row, Column:=C_PARAM_COL_VALUE, Sheet:=G_WEB_SETTINGS_SHEET))
    If Trim$(result_value) = "" Then result_value = DefaultValue

    pGetOptionalParam = result_value
End Function

Private Function pGetOptionalLongParam(ByVal Name As String, ByVal DefaultValue As Long) As Long
    Dim param_value As String
    param_value = pGetOptionalParam(Name, CStr(DefaultValue))

    If Trim$(param_value) = "" Or Not IsNumeric(param_value) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings", Description:=Name & " は数値である必要があります。(" & param_value & ")"
    End If

    Dim result_value As Long
    result_value = CLng(param_value)
    If result_value <= 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings", Description:=Name & " は 1 以上である必要があります。(" & param_value & ")"
    End If

    pGetOptionalLongParam = result_value
End Function

Private Function pGetOptionalBooleanParam(ByVal Name As String, ByVal DefaultValue As Boolean) As Boolean
    Dim default_text As String
    If DefaultValue Then
        default_text = "True"
    Else
        default_text = "False"
    End If

    Dim param_value As String
    param_value = UCase$(Trim$(pGetOptionalParam(Name, default_text)))

    Select Case param_value
        Case "TRUE", "1", "YES"
            pGetOptionalBooleanParam = True
        Case "FALSE", "0", "NO"
            pGetOptionalBooleanParam = False
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings", Description:=Name & " は True または False である必要があります。(" & param_value & ")"
    End Select
End Function

Private Function pNormalizeExistingRowMode(ByVal ModeText As String) As String
    Select Case UCase$(Trim$(ModeText))
        Case UCase$(G_WEB_ROW_MODE_SKIP_EXISTING)
            pNormalizeExistingRowMode = G_WEB_ROW_MODE_SKIP_EXISTING
        Case UCase$(G_WEB_ROW_MODE_OVERWRITE)
            pNormalizeExistingRowMode = G_WEB_ROW_MODE_OVERWRITE
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings", Description:="ExistingRowMode が不正です。(値: " & ModeText & "、有効値: " & G_WEB_ROW_MODE_SKIP_EXISTING & ", " & G_WEB_ROW_MODE_OVERWRITE & ")"
    End Select
End Function

Private Function pResolvePath(ByVal PathLikeString As String) As String
    If FsSrv Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings", Description:="InitializeCommonService が呼ばれていません。"
    End If

    pResolvePath = FsSrv.GetAbsolutePath(PathLikeString)
End Function

Private Function pReadTransitionOperations() As ObjectList
    Dim result_value As ObjectList
    Set result_value = New_ObjectList("TransitionOperation")

    Dim table_values() As Variant
    table_values = pReadTableValues(C_TRANSITION_COL_START, C_TRANSITION_COL_FINISH)

    If Not IsEmptyArray(table_values) Then
        Dim row_idx As Long
        For row_idx = LBound(table_values, 1) To UBound(table_values, 1)
            If Not pIsEmptyTableRow(table_values, row_idx, 1, 6) Then
                Dim operation_name As String
                operation_name = pTableText(table_values, row_idx, 1)

                Dim locator_type As String
                locator_type = pTableText(table_values, row_idx, 2)

                Dim locator_inner_text As String
                locator_inner_text = pTableText(table_values, row_idx, 4)
                If Trim$(locator_inner_text) <> "" And LCase$(Trim$(locator_type)) <> "css selector" Then
                    Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings", Description:="LocatorInnerText を指定する場合、LocatorType は css selector である必要があります。(OperationName=" & operation_name & ")"
                End If

                Call result_value.Add(New_TransitionOperation( _
                        OperationName:=operation_name, _
                        LocatorType:=locator_type, _
                        LocatorValue:=pTableText(table_values, row_idx, 3), _
                        LocatorInnerText:=locator_inner_text, _
                        Script:=pTableText(table_values, row_idx, 5), _
                        WaitConditionName:=pTableText(table_values, row_idx, 6)))
            End If
        Next row_idx
    End If

    Set pReadTransitionOperations = result_value
End Function

Private Function pReadDetailColumnDefinitions() As ObjectList
    Dim result_value As ObjectList
    Set result_value = New_ObjectList("DetailColumnDefinition")

    Dim table_values() As Variant
    table_values = pReadTableValues(C_DETAIL_COL_START, C_DETAIL_COL_FINISH)

    If Not IsEmptyArray(table_values) Then
        Dim row_idx As Long
        For row_idx = LBound(table_values, 1) To UBound(table_values, 1)
            If Not pIsEmptyTableRow(table_values, row_idx, 1, 7) Then
                Call result_value.Add(New_DetailColumnDefinition( _
                        OutputColumnName:=pTableText(table_values, row_idx, 1), _
                        Selector:=pTableText(table_values, row_idx, 2), _
                        ExtractType:=pDefaultIfBlank(pTableText(table_values, row_idx, 3), "InnerText"), _
                        AttributeName:=pTableText(table_values, row_idx, 4), _
                        IsRequired:=pParseBoolean(pTableText(table_values, row_idx, 5), False, "IsRequired"), _
                        ValueExpression:=pTableText(table_values, row_idx, 6), _
                        BlankMode:=pDefaultIfBlank(pTableText(table_values, row_idx, 7), "AllowBlank")))
            End If
        Next row_idx
    End If

    Set pReadDetailColumnDefinitions = result_value
End Function

Private Function pReadTableValues(ByVal StartColumn As Long, ByVal FinishColumn As Long) As Variant()
    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds( _
            Row:=C_PARAM_START_ROW, _
            Column:=StartColumn, _
            FinishRow:=G_ROW_MAX, _
            FinishColumn:=FinishColumn, _
            Sheet:=G_WEB_SETTINGS_SHEET)

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = WsSrv.GetUsedRangeBounds(search_bounds)

    pReadTableValues = WsSrv.ReadRange(used_bounds)
End Function

Private Function pIsEmptyTableRow(ByRef TableValues() As Variant, ByVal RowIndex As Long, ByVal FirstColumnIndex As Long, ByVal LastColumnIndex As Long) As Boolean
    Dim last_existing_column As Long
    last_existing_column = LastColumnIndex
    If UBound(TableValues, 2) < last_existing_column Then last_existing_column = UBound(TableValues, 2)

    Dim col_idx As Long
    For col_idx = FirstColumnIndex To last_existing_column
        If Trim$(pTableText(TableValues, RowIndex, col_idx)) <> "" Then
            pIsEmptyTableRow = False
            Exit Function
        End If
    Next col_idx

    pIsEmptyTableRow = True
End Function

Private Function pTableText(ByRef TableValues() As Variant, ByVal RowIndex As Long, ByVal ColumnIndex As Long) As String
    If ColumnIndex < LBound(TableValues, 2) Or UBound(TableValues, 2) < ColumnIndex Then
        pTableText = ""
        Exit Function
    End If

    pTableText = CStr(TableValues(RowIndex, ColumnIndex))
End Function

Private Function pDefaultIfBlank(ByVal TargetText As String, ByVal DefaultValue As String) As String
    If Trim$(TargetText) = "" Then
        pDefaultIfBlank = DefaultValue
    Else
        pDefaultIfBlank = TargetText
    End If
End Function

Private Function pParseBoolean(ByVal TargetText As String, ByVal DefaultValue As Boolean, ByVal SourceName As String) As Boolean
    Dim normalized_text As String
    normalized_text = UCase$(Trim$(TargetText))

    If normalized_text = "" Then
        pParseBoolean = DefaultValue
        Exit Function
    End If

    Select Case normalized_text
        Case "TRUE", "1", "YES"
            pParseBoolean = True
        Case "FALSE", "0", "NO"
            pParseBoolean = False
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings", Description:=SourceName & " は True / False / 1 / 0 / Yes / No のいずれかである必要があります。(" & TargetText & ")"
    End Select
End Function
