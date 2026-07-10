Attribute VB_Name = "Test_ToolSettings"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! ToolSettings のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Private pCurrentSettingsSheet As UserInputSheetTestDouble

Public Sub Test_ToolSettings_Settingsシート_UserInputSheetで単一値設定を読み取る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")

    Dim settings_sheet As UserInputSheetTestDouble
    Set settings_sheet = New UserInputSheetTestDouble

    Dim item_range As WorksheetRangeBounds
    Set item_range = New_RangeBounds(Row:=3, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:=G_WEB_SETTINGS_SHEET)
    Call settings_sheet.Store.SetReturn("GetItemRange", item_range, "StartUrl", "")

    Dim value_cell As WorksheetRangeBounds
    Set value_cell = item_range.GetCell(1, 1)
    Call ws_stub.Store.SetReturn("ReadCell", "https://example.test/start", value_cell, False)

    Dim tool_settings As ToolSettings
    Set tool_settings = New ToolSettings
    Call tool_settings.Initialize(settings_sheet)

    ' --- Act ---
    Dim actual_start_url As String
    actual_start_url = tool_settings.StartUrl

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.test/start", actual_start_url
    Assert.EqualsNumeric 1, settings_sheet.Store.GetCallCount("GetItemRange", "StartUrl", "")
    Assert.EqualsNumeric 1, ws_stub.Store.GetCallCount("ReadCell", value_cell, False)
End Sub

Public Sub Test_ToolSettings_Settingsシート_単一値設定を読み取る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetSetting(ws_stub, "WebDriverPath", "C:\Driver\msedgedriver.exe", 2)
    Call pSetSetting(ws_stub, "WebDriverPort", "9516", 3)
    Call pSetSetting(ws_stub, "BrowserProfilePath", "C:\Profile", 4)
    Call pSetSetting(ws_stub, "Headless", "False", 5)
    Call pSetSetting(ws_stub, "StartUrl", "https://example.test/start", 6)
    Call pSetSetting(ws_stub, "OutputSheetName", "result", 7)
    Call pSetSetting(ws_stub, "AuthenticatedStartSelector", "#top-ready", 8)
    Call pSetSetting(ws_stub, "ListTransitionOperationName", "OpenList", 10)
    Call pSetSetting(ws_stub, "ListItemSelector", "#list tbody tr", 11)
    Call pSetSetting(ws_stub, "ListItemTargetIdSelector", "#list-item-target-id", 12)
    Call pSetSetting(ws_stub, "DetailTransitionOperationName", "OpenDetail", 13)
    Call pSetSetting(ws_stub, "TargetIdSelector", "#target-id", 14)
    Call pSetSetting(ws_stub, "ReturnToListOperationName", "ReturnToList", 15)
    Call pSetSetting(ws_stub, "NextPageOperationName", "NextPage", 16)
    Call pSetSetting(ws_stub, "NextPageAvailableSelector", "#next-page", 17)
    Call pSetSetting(ws_stub, "ExistingRowMode", G_WEB_ROW_MODE_OVERWRITE, 18)
    Call pSetSetting(ws_stub, "OutputConditionExpression", "[申請者] != """"", 19)
    Call pSetSetting(ws_stub, "TimeoutSeconds", "45", 20)
    Call pSetSetting(ws_stub, "DownloadEnabled", "True", 21)
    Call pSetSetting(ws_stub, "DownloadRequired", "True", 22)
    Call pSetSetting(ws_stub, "DownloadRootPath", "C:\Downloads", 23)
    Call pSetSetting(ws_stub, "DownloadLinkSelector", "#download", 24)

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim actual_driver_path As String
    actual_driver_path = tool_settings.WebDriverPath

    Dim actual_port As Long
    actual_port = tool_settings.WebDriverPort

    Dim actual_profile_path As String
    actual_profile_path = tool_settings.BrowserProfilePath

    Dim actual_headless As Boolean
    actual_headless = tool_settings.Headless

    Dim actual_start_url As String
    actual_start_url = tool_settings.StartUrl

    Dim actual_output_sheet As String
    actual_output_sheet = tool_settings.OutputSheetName

    Dim actual_authenticated_selector As String
    actual_authenticated_selector = tool_settings.AuthenticatedStartSelector


    Dim actual_list_operation_name As String
    actual_list_operation_name = tool_settings.ListTransitionOperationName

    Dim actual_list_item_selector As String
    actual_list_item_selector = tool_settings.ListItemSelector

    Dim actual_list_target_selector As String
    actual_list_target_selector = tool_settings.ListItemTargetIdSelector

    Dim actual_detail_operation_name As String
    actual_detail_operation_name = tool_settings.DetailTransitionOperationName

    Dim actual_target_selector As String
    actual_target_selector = tool_settings.TargetIdSelector

    Dim actual_return_operation_name As String
    actual_return_operation_name = tool_settings.ReturnToListOperationName

    Dim actual_next_operation_name As String
    actual_next_operation_name = tool_settings.NextPageOperationName

    Dim actual_next_available_selector As String
    actual_next_available_selector = tool_settings.NextPageAvailableSelector

    Dim actual_existing_mode As String
    actual_existing_mode = tool_settings.ExistingRowMode

    Dim actual_condition_expression As String
    actual_condition_expression = tool_settings.OutputConditionExpression

    Dim actual_timeout As Long
    actual_timeout = tool_settings.TimeoutSeconds

    Dim actual_download_enabled As Boolean
    actual_download_enabled = tool_settings.DownloadEnabled

    Dim actual_download_required As Boolean
    actual_download_required = tool_settings.DownloadRequired

    Dim actual_download_root_path As String
    actual_download_root_path = tool_settings.DownloadRootPath

    Dim actual_download_link_selector As String
    actual_download_link_selector = tool_settings.DownloadLinkSelector

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Driver\msedgedriver.exe", actual_driver_path
    Assert.EqualsNumeric 9516, actual_port
    Assert.Equals "C:\Profile", actual_profile_path
    Assert.IsFalse actual_headless
    Assert.Equals "https://example.test/start", actual_start_url
    Assert.Equals "result", actual_output_sheet
    Assert.Equals "#top-ready", actual_authenticated_selector
    Assert.Equals "OpenList", actual_list_operation_name
    Assert.Equals "#list tbody tr", actual_list_item_selector
    Assert.Equals "#list-item-target-id", actual_list_target_selector
    Assert.Equals "OpenDetail", actual_detail_operation_name
    Assert.Equals "#target-id", actual_target_selector
    Assert.Equals "ReturnToList", actual_return_operation_name
    Assert.Equals "NextPage", actual_next_operation_name
    Assert.Equals "#next-page", actual_next_available_selector
    Assert.Equals G_WEB_ROW_MODE_OVERWRITE, actual_existing_mode
    Assert.Equals "[申請者] != """"", actual_condition_expression
    Assert.EqualsNumeric 45, actual_timeout
    Assert.IsTrue actual_download_enabled
    Assert.IsTrue actual_download_required
    Assert.Equals "C:\Downloads", actual_download_root_path
    Assert.Equals "#download", actual_download_link_selector
End Sub

Public Sub Test_ToolSettings_Settingsシート_ExistingRowModeでFillBlankCellsを読み取る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetSetting(ws_stub, "ExistingRowMode", G_WEB_ROW_MODE_FILL_BLANK_CELLS, 18)

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim actual_existing_mode As String
    actual_existing_mode = tool_settings.ExistingRowMode

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_ROW_MODE_FILL_BLANK_CELLS, actual_existing_mode
End Sub

Public Sub Test_ToolSettings_Settingsシート_パス設定は絶対パスへ解決して返す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetSetting(ws_stub, "WebDriverPath", "%LOCALAPPDATA%\Driver\msedgedriver.exe", 2)
    Call pSetSetting(ws_stub, "BrowserProfilePath", "%LOCALAPPDATA%\Profile", 4)

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Set FsSrv = fs_stub
    Call fs_stub.Store.SetReturn("GetAbsolutePath", "C:\Users\tester\AppData\Local\Driver\msedgedriver.exe", "%LOCALAPPDATA%\Driver\msedgedriver.exe")
    Call fs_stub.Store.SetReturn("GetAbsolutePath", "C:\Users\tester\AppData\Local\Profile", "%LOCALAPPDATA%\Profile")

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim actual_driver_path As String
    actual_driver_path = tool_settings.WebDriverPath

    Dim actual_profile_path As String
    actual_profile_path = tool_settings.BrowserProfilePath

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Users\tester\AppData\Local\Driver\msedgedriver.exe", actual_driver_path
    Assert.Equals "C:\Users\tester\AppData\Local\Profile", actual_profile_path
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetAbsolutePath", "%LOCALAPPDATA%\Driver\msedgedriver.exe")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetAbsolutePath", "%LOCALAPPDATA%\Profile")
End Sub

Public Sub Test_ToolSettings_Settingsシート_既定パスと既定値を補う(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetMissingSetting(ws_stub, "WebDriverPath")
    Call pSetMissingSetting(ws_stub, "WebDriverPort")
    Call pSetMissingSetting(ws_stub, "BrowserProfilePath")
    Call pSetMissingSetting(ws_stub, "Headless")
    Call pSetSetting(ws_stub, "StartUrl", "https://example.test/start", 6)
    Call pSetSetting(ws_stub, "OutputSheetName", "result", 7)
    Call pSetSetting(ws_stub, "AuthenticatedStartSelector", "#top-ready", 8)
    Call pSetSetting(ws_stub, "ListTransitionOperationName", "OpenList", 10)
    Call pSetSetting(ws_stub, "ListItemSelector", "#list tbody tr", 11)
    Call pSetMissingSetting(ws_stub, "NextPageOperationName")
    Call pSetMissingSetting(ws_stub, "NextPageAvailableSelector")
    Call pSetMissingSetting(ws_stub, "ExistingRowMode")
    Call pSetMissingSetting(ws_stub, "OutputConditionExpression")
    Call pSetMissingSetting(ws_stub, "TimeoutSeconds")
    Call pSetMissingSetting(ws_stub, "DownloadEnabled")
    Call pSetMissingSetting(ws_stub, "DownloadRequired")
    Call pSetMissingSetting(ws_stub, "DownloadRootPath")
    Call pSetMissingSetting(ws_stub, "DownloadLinkSelector")

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim actual_driver_path As String
    actual_driver_path = tool_settings.WebDriverPath

    Dim actual_port As Long
    actual_port = tool_settings.WebDriverPort

    Dim actual_profile_path As String
    actual_profile_path = tool_settings.BrowserProfilePath

    Dim actual_headless As Boolean
    actual_headless = tool_settings.Headless

    Dim actual_next_operation_name As String
    actual_next_operation_name = tool_settings.NextPageOperationName

    Dim actual_next_available_selector As String
    actual_next_available_selector = tool_settings.NextPageAvailableSelector

    Dim actual_existing_mode As String
    actual_existing_mode = tool_settings.ExistingRowMode

    Dim actual_condition_expression As String
    actual_condition_expression = tool_settings.OutputConditionExpression

    Dim actual_timeout As Long
    actual_timeout = tool_settings.TimeoutSeconds

    Dim actual_download_enabled As Boolean
    actual_download_enabled = tool_settings.DownloadEnabled

    Dim actual_download_required As Boolean
    actual_download_required = tool_settings.DownloadRequired

    Dim actual_download_root_path As String
    actual_download_root_path = tool_settings.DownloadRootPath

    Dim actual_download_link_selector As String
    actual_download_link_selector = tool_settings.DownloadLinkSelector

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Workbook\bin\msedgedriver.exe", actual_driver_path
    Assert.EqualsNumeric 9515, actual_port
    Assert.Equals "C:\Workbook\browser-profile", actual_profile_path
    Assert.IsTrue actual_headless
    Assert.Equals "", actual_next_operation_name
    Assert.Equals "", actual_next_available_selector
    Assert.Equals G_WEB_ROW_MODE_SKIP_EXISTING, actual_existing_mode
    Assert.Equals "", actual_condition_expression
    Assert.EqualsNumeric 30, actual_timeout
    Assert.IsFalse actual_download_enabled
    Assert.IsFalse actual_download_required
    Assert.Equals "", actual_download_root_path
    Assert.Equals "", actual_download_link_selector
End Sub

Public Sub Test_ToolSettings_Settingsシート_NextPageOperationNameのみ指定はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetSetting(ws_stub, "NextPageOperationName", "NextPage", 16)
    Call pSetMissingSetting(ws_stub, "NextPageAvailableSelector")

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim actual_next_operation_name As String
    actual_next_operation_name = tool_settings.NextPageOperationName

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "NextPageOperationName", vbTextCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "NextPageAvailableSelector", vbTextCompare)
End Sub

Public Sub Test_ToolSettings_Settingsシート_NextPageAvailableSelectorのみ指定はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetMissingSetting(ws_stub, "NextPageOperationName")
    Call pSetSetting(ws_stub, "NextPageAvailableSelector", "#next-page", 17)

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim actual_next_available_selector As String
    actual_next_available_selector = tool_settings.NextPageAvailableSelector

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "NextPageOperationName", vbTextCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "NextPageAvailableSelector", vbTextCompare)
End Sub

Public Sub Test_ToolSettings_Settingsシート_DownloadEnabledTrueで保存先ルート未設定はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetSetting(ws_stub, "DownloadEnabled", "True", 21)
    Call pSetMissingSetting(ws_stub, "DownloadRootPath")

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim actual_download_root_path As String
    actual_download_root_path = tool_settings.DownloadRootPath

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_ToolSettings_Settingsシート_DownloadEnabledTrueでリンクSelector未設定はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetSetting(ws_stub, "DownloadEnabled", "True", 21)
    Call pSetSetting(ws_stub, "DownloadRootPath", "C:\Downloads", 22)
    Call pSetMissingSetting(ws_stub, "DownloadLinkSelector")

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim actual_download_link_selector As String
    actual_download_link_selector = tool_settings.DownloadLinkSelector

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_ToolSettings_Settingsシート_操作定義と列定義を読み取る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetTransitionTable(ws_stub)
    Call pSetDetailColumnTable(ws_stub)

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim operations As ObjectList
    Set operations = tool_settings.TransitionOperations

    Dim detail_columns As ObjectList
    Set detail_columns = tool_settings.DetailColumnDefinitions

    Dim first_operation As TransitionOperation
    Set first_operation = operations.Item(0)

    Dim second_operation As TransitionOperation
    Set second_operation = operations.Item(1)

    Dim first_column As DetailColumnDefinition
    Set first_column = detail_columns.Item(0)

    Dim second_column As DetailColumnDefinition
    Set second_column = detail_columns.Item(1)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, operations.Count
    Assert.Equals "OpenList", first_operation.OperationName
    Assert.Equals "#open-list", first_operation.ActionSelector
    Assert.Equals "一覧を開く", first_operation.ActionInnerText
    Assert.Equals "#list-ready", first_operation.WaitSelector
    Assert.Equals "", second_operation.ActionInnerText
    Assert.EqualsNumeric 2, detail_columns.Count
    Assert.Equals "対象ID", first_column.OutputColumnName
    Assert.Equals "#target-id", first_column.Selector
    Assert.Equals "InnerText", first_column.ExtractType
    Assert.IsTrue first_column.IsRequired
    Assert.Equals "ErrorIfBlank", first_column.BlankMode
    Assert.Equals "件名", second_column.OutputColumnName
    Assert.Equals "TextContent", second_column.ExtractType
    Assert.IsFalse second_column.IsRequired
    Assert.Equals "AllowBlank", second_column.BlankMode
End Sub





Public Sub Test_ToolSettings_Settingsシート_ValueExpression付き列定義を読み取る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetDetailColumnTableWithValueExpression(ws_stub)

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim detail_columns As ObjectList
    Set detail_columns = tool_settings.DetailColumnDefinitions

    Dim normal_column As DetailColumnDefinition
    Set normal_column = detail_columns.Item(0)

    Dim derived_column As DetailColumnDefinition
    Set derived_column = detail_columns.Item(1)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, detail_columns.Count
    Assert.Equals "元列", normal_column.OutputColumnName
    Assert.Equals "#source", normal_column.Selector
    Assert.Equals "", CStr(CallByName(normal_column, "ValueExpression", VbGet))
    Assert.Equals "別名", derived_column.OutputColumnName
    Assert.Equals "", derived_column.Selector
    Assert.Equals "[元列]", CStr(CallByName(derived_column, "ValueExpression", VbGet))
    Assert.IsFalse derived_column.IsRequired
    Assert.Equals "AllowBlank", derived_column.BlankMode
End Sub

Public Sub Test_ToolSettings_Settingsシート_OutputValueTypeを読み取る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetDetailColumnTableWithOutputValueType(ws_stub)

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim detail_columns As ObjectList
    Set detail_columns = tool_settings.DetailColumnDefinitions

    Dim string_column As DetailColumnDefinition
    Set string_column = detail_columns.Item(0)

    Dim auto_column As DetailColumnDefinition
    Set auto_column = detail_columns.Item(1)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, detail_columns.Count
    Assert.Equals "String", CStr(CallByName(string_column, "OutputValueType", VbGet))
    Assert.Equals "Auto", CStr(CallByName(auto_column, "OutputValueType", VbGet))
    Assert.Equals "ErrorIfBlank", string_column.BlankMode
    Assert.Equals "AllowBlank", auto_column.BlankMode
End Sub

Public Sub Test_ToolSettings_Settingsシート_OutputValueType空欄はStringとして読み取る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetDetailColumnTable(ws_stub)

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim detail_columns As ObjectList
    Set detail_columns = tool_settings.DetailColumnDefinitions

    Dim first_column As DetailColumnDefinition
    Set first_column = detail_columns.Item(0)

    Dim second_column As DetailColumnDefinition
    Set second_column = detail_columns.Item(1)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "String", CStr(CallByName(first_column, "OutputValueType", VbGet))
    Assert.Equals "String", CStr(CallByName(second_column, "OutputValueType", VbGet))
End Sub

Public Sub Test_ToolSettings_Settingsシート_OutputValueType不正値はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetDetailColumnTableWithInvalidOutputValueType(ws_stub)

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim detail_columns As ObjectList
    Set detail_columns = tool_settings.DetailColumnDefinitions

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "OutputValueType", vbBinaryCompare)
End Sub

Public Sub Test_ToolSettings_Settingsシート_OutputColumnNameのLFを保持する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetDetailColumnTableWithLfOutputColumnName(ws_stub)

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim detail_columns As ObjectList
    Set detail_columns = tool_settings.DetailColumnDefinitions

    Dim detail_column As DetailColumnDefinition
    Set detail_column = detail_columns.Item(0)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, detail_columns.Count
    Assert.Equals "件名" & vbLf & "詳細", detail_column.OutputColumnName
End Sub

Public Sub Test_ToolSettings_Settingsシート_列定義のSelectorとValueExpressionが両方空ならエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetInvalidDetailColumnTable(ws_stub, "", "False", "")

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim detail_columns As ObjectList
    Set detail_columns = tool_settings.DetailColumnDefinitions

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_ToolSettings_Settingsシート_ActionInnerTextでActionSelector未設定はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetInvalidTransitionInnerTextTable(ws_stub)

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim operations As ObjectList
    Set operations = tool_settings.TransitionOperations

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "ActionInnerText", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "ActionSelector", vbTextCompare)
End Sub

Public Sub Test_ToolSettings_Settingsシート_列定義のSelectorとValueExpressionが両方指定ならエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetInvalidDetailColumnTable(ws_stub, "#alias", "False", "[元列]")

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim detail_columns As ObjectList
    Set detail_columns = tool_settings.DetailColumnDefinitions

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_ToolSettings_Settingsシート_派生列IsRequiredTrueはエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetInvalidDetailColumnTable(ws_stub, "", "True", "[元列]")

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim detail_columns As ObjectList
    Set detail_columns = tool_settings.DetailColumnDefinitions

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_ToolSettings_Settingsシート_必須設定不足はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetMissingSetting(ws_stub, "StartUrl")

    Dim tool_settings As IToolSettings
    Set tool_settings = pNewToolSettings()

    ' --- Act ---
    Dim actual_start_url As String
    actual_start_url = tool_settings.StartUrl

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Private Function pUseSettingsStubs(ByVal WorkbookDirectoryPath As String) As WorksheetServiceTestDouble
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Call wb_stub.Store.SetReturn("GetThisWorkbookDirectoryPath", WorkbookDirectoryPath)
    Set WbSrv = wb_stub

    Set FsSrv = New FileSystemService

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Set pCurrentSettingsSheet = New UserInputSheetTestDouble

    Set pUseSettingsStubs = ws_stub
End Function

Private Function pNewToolSettings() As ToolSettings
    If pCurrentSettingsSheet Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Module Test_ToolSettings.pNewToolSettings", Description:="settings sheet stub が準備されていません。"
    End If

    Dim result_value As ToolSettings
    Set result_value = New ToolSettings
    Call result_value.Initialize(pCurrentSettingsSheet)

    Set pNewToolSettings = result_value
End Function

Private Sub pSetSetting(ByVal WsStub As WorksheetServiceTestDouble, ByVal SettingName As String, ByVal SettingValue As String, ByVal RowIndex As Long)
    Dim item_range As WorksheetRangeBounds
    Set item_range = New_RangeBounds(Row:=RowIndex, Column:=2, FinishRow:=RowIndex, FinishColumn:=2, Sheet:=G_WEB_SETTINGS_SHEET)

    Call pCurrentSettingsSheet.Store.SetReturn("GetItemRange", item_range, SettingName, "")
    Call WsStub.Store.SetReturn("ReadCell", SettingValue, item_range.GetCell(1, 1), False)
End Sub

Private Sub pSetMissingSetting(ByVal WsStub As WorksheetServiceTestDouble, ByVal SettingName As String)
    Dim item_range As WorksheetRangeBounds
    Set item_range = Nothing

    Call pCurrentSettingsSheet.Store.SetReturn("GetItemRange", item_range, SettingName, "")
End Sub

Private Sub pSetTransitionTable(ByVal WsStub As WorksheetServiceTestDouble)
    Dim table_values() As Variant
    ReDim table_values(1 To 2, 1 To 5)
    table_values(1, 1) = "OpenList"
    table_values(1, 2) = "#open-list"
    table_values(1, 3) = "一覧を開く"
    table_values(1, 4) = ""
    table_values(1, 5) = "#list-ready"
    table_values(2, 1) = "OpenDetail"
    table_values(2, 2) = ".detail-link"
    table_values(2, 3) = ""
    table_values(2, 4) = ""
    table_values(2, 5) = "#target-id"

    Dim header_names() As String
    header_names = pTransitionOperationHeaders()
    Call pSetVirtualTable(WsStub, "TransitionOperations", header_names, table_values, 26, 2)
End Sub

Private Sub pSetInvalidTransitionInnerTextTable(ByVal WsStub As WorksheetServiceTestDouble)
    Dim table_values() As Variant
    ReDim table_values(1 To 1, 1 To 5)
    table_values(1, 1) = "OpenList"
    table_values(1, 2) = ""
    table_values(1, 3) = "一覧を開く"
    table_values(1, 4) = ""
    table_values(1, 5) = "#list-ready"

    Dim header_names() As String
    header_names = pTransitionOperationHeaders()
    Call pSetVirtualTable(WsStub, "TransitionOperations", header_names, table_values, 26, 2)
End Sub

Private Function pTransitionOperationHeaders() As String()
    Dim result_value(0 To 4) As String
    result_value(0) = "OperationName"
    result_value(1) = "ActionSelector"
    result_value(2) = "ActionInnerText"
    result_value(3) = "ActionScript"
    result_value(4) = "WaitSelector"

    pTransitionOperationHeaders = result_value
End Function

Private Sub pSetInvalidDetailColumnTable( _
        ByVal WsStub As WorksheetServiceTestDouble, _
        ByVal Selector As String, _
        ByVal IsRequiredText As String, _
        ByVal ValueExpression As String)

    Dim table_values() As Variant
    ReDim table_values(1 To 1, 1 To 8)
    table_values(1, 1) = "不正列"
    table_values(1, 2) = Selector
    table_values(1, 3) = "InnerText"
    table_values(1, 4) = ""
    table_values(1, 5) = IsRequiredText
    table_values(1, 6) = ValueExpression
    table_values(1, 7) = ""
    table_values(1, 8) = "AllowBlank"

    Call pSetDetailTableValues(WsStub, table_values)
End Sub

Private Sub pSetDetailColumnTableWithValueExpression(ByVal WsStub As WorksheetServiceTestDouble)
    Dim table_values() As Variant
    ReDim table_values(1 To 2, 1 To 8)
    table_values(1, 1) = "元列"
    table_values(1, 2) = "#source"
    table_values(1, 3) = "InnerText"
    table_values(1, 4) = ""
    table_values(1, 5) = "False"
    table_values(1, 6) = ""
    table_values(1, 7) = ""
    table_values(1, 8) = "AllowBlank"
    table_values(2, 1) = "別名"
    table_values(2, 2) = ""
    table_values(2, 3) = ""
    table_values(2, 4) = ""
    table_values(2, 5) = "False"
    table_values(2, 6) = "[元列]"
    table_values(2, 7) = ""
    table_values(2, 8) = "AllowBlank"

    Call pSetDetailTableValues(WsStub, table_values)
End Sub

Private Sub pSetDetailColumnTableWithOutputValueType(ByVal WsStub As WorksheetServiceTestDouble)
    Dim table_values() As Variant
    ReDim table_values(1 To 2, 1 To 8)
    table_values(1, 1) = "文字列列"
    table_values(1, 2) = "#text"
    table_values(1, 3) = "InnerText"
    table_values(1, 4) = ""
    table_values(1, 5) = "False"
    table_values(1, 6) = ""
    table_values(1, 7) = "String"
    table_values(1, 8) = "ErrorIfBlank"
    table_values(2, 1) = "自動列"
    table_values(2, 2) = "#auto"
    table_values(2, 3) = "InnerText"
    table_values(2, 4) = ""
    table_values(2, 5) = "False"
    table_values(2, 6) = ""
    table_values(2, 7) = "Auto"
    table_values(2, 8) = "AllowBlank"

    Call pSetDetailTableValues(WsStub, table_values)
End Sub

Private Sub pSetDetailColumnTableWithInvalidOutputValueType(ByVal WsStub As WorksheetServiceTestDouble)
    Dim table_values() As Variant
    ReDim table_values(1 To 1, 1 To 8)
    table_values(1, 1) = "不正型列"
    table_values(1, 2) = "#value"
    table_values(1, 3) = "InnerText"
    table_values(1, 4) = ""
    table_values(1, 5) = "False"
    table_values(1, 6) = ""
    table_values(1, 7) = "Number"
    table_values(1, 8) = "AllowBlank"

    Call pSetDetailTableValues(WsStub, table_values)
End Sub

Private Sub pSetDetailColumnTableWithLfOutputColumnName(ByVal WsStub As WorksheetServiceTestDouble)
    Dim table_values() As Variant
    ReDim table_values(1 To 1, 1 To 8)
    table_values(1, 1) = "件名" & vbLf & "詳細"
    table_values(1, 2) = "#subject"
    table_values(1, 3) = "InnerText"
    table_values(1, 4) = ""
    table_values(1, 5) = "False"
    table_values(1, 6) = ""
    table_values(1, 7) = ""
    table_values(1, 8) = "AllowBlank"

    Call pSetDetailTableValues(WsStub, table_values)
End Sub

Private Sub pSetDetailColumnTable(ByVal WsStub As WorksheetServiceTestDouble)
    Dim table_values() As Variant
    ReDim table_values(1 To 2, 1 To 8)
    table_values(1, 1) = "対象ID"
    table_values(1, 2) = "#target-id"
    table_values(1, 3) = "InnerText"
    table_values(1, 4) = ""
    table_values(1, 5) = "True"
    table_values(1, 6) = ""
    table_values(1, 7) = ""
    table_values(1, 8) = "ErrorIfBlank"
    table_values(2, 1) = "件名"
    table_values(2, 2) = "#title"
    table_values(2, 3) = "TextContent"
    table_values(2, 4) = ""
    table_values(2, 5) = "False"
    table_values(2, 6) = ""
    table_values(2, 7) = ""
    table_values(2, 8) = "AllowBlank"

    Call pSetDetailTableValues(WsStub, table_values)
End Sub

Private Sub pSetDetailTableValues(ByVal WsStub As WorksheetServiceTestDouble, ByRef TableValues() As Variant)
    Dim header_names() As String
    header_names = pDetailColumnDefinitionHeaders()
    Call pSetVirtualTable(WsStub, "DetailColumnDefinitions", header_names, TableValues, 31, 2)
End Sub

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

Private Sub pSetVirtualTable( _
        ByVal WsStub As WorksheetServiceTestDouble, _
        ByVal SectionName As String, _
        ByRef HeaderNames() As String, _
        ByRef TableValues() As Variant, _
        ByVal StartRow As Long, _
        ByVal StartColumn As Long)

    Dim header_idx As Long
    For header_idx = LBound(HeaderNames) To UBound(HeaderNames)
        Dim value_col_idx As Long
        value_col_idx = LBound(TableValues, 2) + header_idx - LBound(HeaderNames)

        Dim value_range As WorksheetRangeBounds
        Set value_range = New_RangeBounds( _
                Row:=StartRow, _
                Column:=StartColumn + header_idx - LBound(HeaderNames), _
                FinishRow:=StartRow + UBound(TableValues, 1) - LBound(TableValues, 1), _
                FinishColumn:=StartColumn + header_idx - LBound(HeaderNames), _
                Sheet:=G_WEB_SETTINGS_SHEET)

        Call pCurrentSettingsSheet.Store.SetReturn("GetItemRange", value_range, SectionName, HeaderNames(header_idx))

        Dim value_row_idx As Long
        For value_row_idx = LBound(TableValues, 1) To UBound(TableValues, 1)
            Call WsStub.Store.SetReturn( _
                    "ReadCell", _
                    CStr(TableValues(value_row_idx, value_col_idx)), _
                    value_range.GetCell(value_row_idx - LBound(TableValues, 1) + 1, 1), _
                    False)
        Next value_row_idx
    Next header_idx
End Sub
