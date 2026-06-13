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
Private Const C_DETAIL_COL_FINISH As Long = 19
Private Const C_DEFAULT_WEB_DRIVER_PORT As Long = 9515
Private Const C_DEFAULT_TIMEOUT_SECONDS As Long = 30

Private pSettingsSheet As IUserInputSheet
Private pHasReadSettings As Boolean

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
    Dim operation_name As String
    operation_name = pGetOptionalParam("NextPageOperationName", "")

    Dim available_selector As String
    available_selector = pGetOptionalParam("NextPageAvailableSelector", "")

    Call pValidateNextPageSettings(operation_name, available_selector)
    NextPageOperationName = operation_name
End Property
Private Property Get IToolSettings_NextPageOperationName() As String
    IToolSettings_NextPageOperationName = NextPageOperationName
End Property

'* 次ページ操作が有効な場合だけ一致する selector。
Public Property Get NextPageAvailableSelector() As String
    Dim operation_name As String
    operation_name = pGetOptionalParam("NextPageOperationName", "")

    Dim available_selector As String
    available_selector = pGetOptionalParam("NextPageAvailableSelector", "")

    Call pValidateNextPageSettings(operation_name, available_selector)
    NextPageAvailableSelector = available_selector
End Property
Private Property Get IToolSettings_NextPageAvailableSelector() As String
    IToolSettings_NextPageAvailableSelector = NextPageAvailableSelector
End Property

Private Sub pValidateNextPageSettings( _
        ByVal NextPageOperationNameText As String, _
        ByVal NextPageAvailableSelectorText As String)

    If (Trim$(NextPageOperationNameText) = "") <> (Trim$(NextPageAvailableSelectorText) = "") Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pValidateNextPageSettings", Description:="NextPageOperationName と NextPageAvailableSelector は両方指定するか、両方空にしてください。"
    End If
End Sub

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

'* settings シート読み取り元を差し替えます。
'*
'* @param SettingsSheet ユニットテスト用途で差し替える入力シート。
'*
'* @details
'* ユニットテスト用途の初期化 API です。通常実行では呼び出さず、settings 読み取り時に production の UserInputSheet を lazy 作成します。
'* 最初の settings 読み取りより前に 1 回だけ呼び出せます。
Public Sub Initialize(ByVal SettingsSheet As IUserInputSheet)
    If SettingsSheet Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.Initialize", Description:="SettingsSheet は必須です。"
    End If
    If pHasReadSettings Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.Initialize", Description:="settings 読み取り開始後は初期化できません。"
    End If
    If Not pSettingsSheet Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.Initialize", Description:="すでに初期化済みです。"
    End If

    Set pSettingsSheet = SettingsSheet
End Sub

Private Function pGetRequiredParam(ByVal Name As String) As String
    Dim result_value As String
    result_value = pGetOptionalParam(Name, "")
    If Trim$(result_value) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pGetRequiredParam", Description:=Name & " は必須です。"
    End If

    pGetRequiredParam = result_value
End Function

Private Function pGetOptionalParam(ByVal Name As String, ByVal DefaultValue As String) As String
    Dim settings_sheet As IUserInputSheet
    Set settings_sheet = pGetSettingsSheet()

    Dim item_range As WorksheetRangeBounds
    Set item_range = settings_sheet.GetItemRange(Name)
    If item_range Is Nothing Then
        pGetOptionalParam = DefaultValue
        Exit Function
    End If

    Dim result_value As String
    result_value = "" & WsSrv.ReadCell(item_range.GetCell(1, 1))
    If Trim$(result_value) = "" Then result_value = DefaultValue

    pGetOptionalParam = result_value
End Function

Private Function pGetOptionalLongParam(ByVal Name As String, ByVal DefaultValue As Long) As Long
    Dim param_value As String
    param_value = pGetOptionalParam(Name, CStr(DefaultValue))

    If Trim$(param_value) = "" Or Not IsNumeric(param_value) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pGetOptionalLongParam", Description:=Name & " は数値である必要があります。(" & param_value & ")"
    End If

    Dim result_value As Long
    result_value = CLng(param_value)
    If result_value <= 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pGetOptionalLongParam", Description:=Name & " は 1 以上である必要があります。(" & param_value & ")"
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
            Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pGetOptionalBooleanParam", Description:=Name & " は True または False である必要があります。(" & param_value & ")"
    End Select
End Function

Private Function pNormalizeExistingRowMode(ByVal ModeText As String) As String
    Select Case UCase$(Trim$(ModeText))
        Case UCase$(G_WEB_ROW_MODE_SKIP_EXISTING)
            pNormalizeExistingRowMode = G_WEB_ROW_MODE_SKIP_EXISTING
        Case UCase$(G_WEB_ROW_MODE_OVERWRITE)
            pNormalizeExistingRowMode = G_WEB_ROW_MODE_OVERWRITE
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pNormalizeExistingRowMode", Description:="ExistingRowMode が不正です。(値: " & ModeText & "、有効値: " & G_WEB_ROW_MODE_SKIP_EXISTING & ", " & G_WEB_ROW_MODE_OVERWRITE & ")"
    End Select
End Function

Private Function pResolvePath(ByVal PathLikeString As String) As String
    If FsSrv Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pResolvePath", Description:="InitializeCommonService が呼ばれていません。"
    End If

    pResolvePath = FsSrv.GetAbsolutePath(PathLikeString)
End Function

Private Function pGetSettingsSheet() As IUserInputSheet
    pHasReadSettings = True
    If pSettingsSheet Is Nothing Then
        Set pSettingsSheet = New_InputSheet(New_RangeBounds(Sheet:=G_WEB_SETTINGS_SHEET))
    End If

    Set pGetSettingsSheet = pSettingsSheet
End Function

Private Function pReadTransitionOperations() As ObjectList
    Dim result_value As ObjectList
    Set result_value = New_ObjectList("TransitionOperation")

    Dim header_names() As String
    header_names = pTransitionOperationHeaders()

    Dim table As WorksheetVirtualTable
    Set table = pReadVirtualTable("TransitionOperations", header_names)

    Dim row_idx As Long
    For row_idx = 0 To table.Count - 1
        Dim row_values As ObjectDictionary
        Set row_values = table.Item(row_idx)
        If Not pIsEmptyVirtualTableRow(row_values, header_names) Then
            Dim operation_name As String
            operation_name = pVirtualTableText(row_values, "OperationName")
            If Trim$(operation_name) = "" Then
                Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pReadTransitionOperations", Description:="TransitionOperations.OperationName は必須です。"
            End If

            Call result_value.Add(New_TransitionOperation( _
                    OperationName:=operation_name, _
                    ActionSelector:=pVirtualTableText(row_values, "ActionSelector"), _
                    ActionInnerText:=pVirtualTableText(row_values, "ActionInnerText"), _
                    ActionScript:=pVirtualTableText(row_values, "ActionScript"), _
                    WaitSelector:=pVirtualTableText(row_values, "WaitSelector")))
        End If
    Next row_idx

    Set pReadTransitionOperations = result_value
End Function

Private Function pTransitionOperationHeaders() As String()
    Dim result_value(0 To 4) As String
    result_value(0) = "OperationName"
    result_value(1) = "ActionSelector"
    result_value(2) = "ActionInnerText"
    result_value(3) = "ActionScript"
    result_value(4) = "WaitSelector"

    pTransitionOperationHeaders = result_value
End Function

Private Function pReadDetailColumnDefinitions() As ObjectList
    Dim result_value As ObjectList
    Set result_value = New_ObjectList("DetailColumnDefinition")

    Dim header_names() As String
    header_names = pDetailColumnDefinitionHeaders()

    Dim table As WorksheetVirtualTable
    Set table = pReadVirtualTable("DetailColumnDefinitions", header_names)

    Dim row_idx As Long
    For row_idx = 0 To table.Count - 1
        Dim row_values As ObjectDictionary
        Set row_values = table.Item(row_idx)
        If Not pIsEmptyVirtualTableRow(row_values, header_names) Then
            Dim output_column_name As String
            output_column_name = pVirtualTableText(row_values, "OutputColumnName")
            If Trim$(output_column_name) = "" Then
                Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pReadDetailColumnDefinitions", Description:="DetailColumnDefinitions.OutputColumnName は必須です。"
            End If

            Call result_value.Add(New_DetailColumnDefinition( _
                    OutputColumnName:=output_column_name, _
                    Selector:=pVirtualTableText(row_values, "Selector"), _
                    ExtractType:=pDefaultIfBlank(pVirtualTableText(row_values, "ExtractType"), "InnerText"), _
                    AttributeName:=pVirtualTableText(row_values, "AttributeName"), _
                    IsRequired:=pParseBoolean(pVirtualTableText(row_values, "IsRequired"), False, "IsRequired"), _
                    ValueExpression:=pVirtualTableText(row_values, "ValueExpression"), _
                    OutputValueType:=pDefaultIfBlank(pVirtualTableText(row_values, "OutputValueType"), "String"), _
                    BlankMode:=pDefaultIfBlank(pVirtualTableText(row_values, "BlankMode"), "AllowBlank")))
        End If
    Next row_idx

    Set pReadDetailColumnDefinitions = result_value
End Function

Private Function pDetailColumnDefinitionHeaders() As String()
    Dim result_value(0 To 7) As String
    result_value(0) = "OutputColumnName"
    result_value(1) = "Selector"
    result_value(2) = "ExtractType"
    result_value(3) = "AttributeName"
    result_value(4) = "IsRequired"
    result_value(5) = "ValueExpression"
    result_value(6) = "OutputValueType"
    result_value(7) = "BlankMode"

    pDetailColumnDefinitionHeaders = result_value
End Function

Private Function pReadVirtualTable(ByVal SectionName As String, ByRef HeaderNames() As String) As WorksheetVirtualTable
    Dim settings_sheet As IUserInputSheet
    Set settings_sheet = pGetSettingsSheet()

    Dim header_ranges As ObjectList
    Set header_ranges = New_ObjectList("WorksheetRangeBounds")

    Dim header_idx As Long
    For header_idx = LBound(HeaderNames) To UBound(HeaderNames)
        Dim header_range As WorksheetRangeBounds
        Set header_range = settings_sheet.GetItemRange(SectionName, HeaderNames(header_idx))
        If header_range Is Nothing Then
            Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pReadVirtualTable", Description:=SectionName & "." & HeaderNames(header_idx) & " header が見つかりません。"
        End If
        Call header_ranges.Add(header_range)
    Next header_idx

    Dim expanded_ranges As ObjectList
    Set expanded_ranges = pExpandHeaderRangesToMaxRows(header_ranges)

    Set pReadVirtualTable = New_WorksheetVirtualTable(expanded_ranges, HeaderNames, HeaderCompareMode:=vbBinaryCompare)
End Function

Private Function pExpandHeaderRangesToMaxRows(ByVal HeaderRanges As ObjectList) As ObjectList
    Select Case HeaderRanges.Count
        Case 5
            Set pExpandHeaderRangesToMaxRows = ExpandRangeBoundsToMax( _
                    True, False, _
                    HeaderRanges.Item(0), HeaderRanges.Item(1), HeaderRanges.Item(2), HeaderRanges.Item(3), HeaderRanges.Item(4))
        Case 8
            Set pExpandHeaderRangesToMaxRows = ExpandRangeBoundsToMax( _
                    True, False, _
                    HeaderRanges.Item(0), HeaderRanges.Item(1), HeaderRanges.Item(2), HeaderRanges.Item(3), _
                    HeaderRanges.Item(4), HeaderRanges.Item(5), HeaderRanges.Item(6), HeaderRanges.Item(7))
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pExpandHeaderRangesToMaxRows", Description:="未対応の header 件数です。(" & CStr(HeaderRanges.Count) & ")"
    End Select
End Function

Private Function pIsEmptyVirtualTableRow(ByVal RowValues As ObjectDictionary, ByRef HeaderNames() As String) As Boolean
    Dim header_idx As Long
    For header_idx = LBound(HeaderNames) To UBound(HeaderNames)
        If Trim$(pVirtualTableText(RowValues, HeaderNames(header_idx))) <> "" Then
            pIsEmptyVirtualTableRow = False
            Exit Function
        End If
    Next header_idx

    pIsEmptyVirtualTableRow = True
End Function

Private Function pVirtualTableText(ByVal RowValues As ObjectDictionary, ByVal HeaderName As String) As String
    Dim value_range As WorksheetRangeBounds
    Set value_range = RowValues.Item(HeaderName)
    If value_range Is Nothing Then
        pVirtualTableText = ""
        Exit Function
    End If
    If value_range.IsEmpty Then
        pVirtualTableText = ""
        Exit Function
    End If

    pVirtualTableText = CStr(WsSrv.ReadCell(value_range.GetCell(1, 1)))
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
            Err.Raise Number:=vbObjectError + 1, Source:="Class ToolSettings.pParseBoolean", Description:=SourceName & " は True / False / 1 / 0 / Yes / No のいずれかである必要があります。(" & TargetText & ")"
    End Select
End Function
