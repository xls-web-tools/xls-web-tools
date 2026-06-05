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
    Call pSetSetting(ws_stub, "ListPageSelector", "#list-ready", 9)
    Call pSetSetting(ws_stub, "ListTransitionOperationName", "OpenList", 10)
    Call pSetSetting(ws_stub, "ListItemTargetIdSelector", "#list-item-target-id", 11)
    Call pSetSetting(ws_stub, "DetailTransitionOperationName", "OpenDetail", 12)
    Call pSetSetting(ws_stub, "TargetIdSelector", "#target-id", 13)
    Call pSetSetting(ws_stub, "ReturnToListOperationName", "ReturnToList", 14)
    Call pSetSetting(ws_stub, "ExistingRowMode", G_WEB_ROW_MODE_RETRY_ERROR, 15)
    Call pSetSetting(ws_stub, "TimeoutSeconds", "45", 16)

    Dim tool_settings As IToolSettings
    Set tool_settings = New ToolSettings

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

    Dim actual_list_selector As String
    actual_list_selector = tool_settings.ListPageSelector

    Dim actual_list_operation_name As String
    actual_list_operation_name = tool_settings.ListTransitionOperationName

    Dim actual_list_target_selector As String
    actual_list_target_selector = tool_settings.ListItemTargetIdSelector

    Dim actual_detail_operation_name As String
    actual_detail_operation_name = tool_settings.DetailTransitionOperationName

    Dim actual_target_selector As String
    actual_target_selector = tool_settings.TargetIdSelector

    Dim actual_return_operation_name As String
    actual_return_operation_name = tool_settings.ReturnToListOperationName

    Dim actual_existing_mode As String
    actual_existing_mode = tool_settings.ExistingRowMode

    Dim actual_timeout As Long
    actual_timeout = tool_settings.TimeoutSeconds

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Driver\msedgedriver.exe", actual_driver_path
    Assert.EqualsNumeric 9516, actual_port
    Assert.Equals "C:\Profile", actual_profile_path
    Assert.IsFalse actual_headless
    Assert.Equals "https://example.test/start", actual_start_url
    Assert.Equals "result", actual_output_sheet
    Assert.Equals "#top-ready", actual_authenticated_selector
    Assert.Equals "#list-ready", actual_list_selector
    Assert.Equals "OpenList", actual_list_operation_name
    Assert.Equals "#list-item-target-id", actual_list_target_selector
    Assert.Equals "OpenDetail", actual_detail_operation_name
    Assert.Equals "#target-id", actual_target_selector
    Assert.Equals "ReturnToList", actual_return_operation_name
    Assert.Equals G_WEB_ROW_MODE_RETRY_ERROR, actual_existing_mode
    Assert.EqualsNumeric 45, actual_timeout
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
    Call pSetSetting(ws_stub, "ListPageSelector", "#list-ready", 9)
    Call pSetSetting(ws_stub, "ListTransitionOperationName", "OpenList", 10)
    Call pSetMissingSetting(ws_stub, "ExistingRowMode")
    Call pSetMissingSetting(ws_stub, "TimeoutSeconds")

    Dim tool_settings As IToolSettings
    Set tool_settings = New ToolSettings

    ' --- Act ---
    Dim actual_driver_path As String
    actual_driver_path = tool_settings.WebDriverPath

    Dim actual_port As Long
    actual_port = tool_settings.WebDriverPort

    Dim actual_profile_path As String
    actual_profile_path = tool_settings.BrowserProfilePath

    Dim actual_headless As Boolean
    actual_headless = tool_settings.Headless

    Dim actual_existing_mode As String
    actual_existing_mode = tool_settings.ExistingRowMode

    Dim actual_timeout As Long
    actual_timeout = tool_settings.TimeoutSeconds

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Workbook\bin\msedgedriver.exe", actual_driver_path
    Assert.EqualsNumeric 9515, actual_port
    Assert.Equals "C:\Workbook\browser-profile", actual_profile_path
    Assert.IsTrue actual_headless
    Assert.Equals G_WEB_ROW_MODE_SKIP_EXISTING, actual_existing_mode
    Assert.EqualsNumeric 30, actual_timeout
End Sub

Public Sub Test_ToolSettings_Settingsシート_操作定義と列定義を読み取る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetTransitionTable(ws_stub)
    Call pSetDetailColumnTable(ws_stub)

    Dim tool_settings As IToolSettings
    Set tool_settings = New ToolSettings

    ' --- Act ---
    Dim operations As ObjectList
    Set operations = tool_settings.TransitionOperations

    Dim detail_columns As ObjectList
    Set detail_columns = tool_settings.DetailColumnDefinitions

    Dim first_operation As TransitionOperation
    Set first_operation = operations.Item(0)

    Dim first_column As DetailColumnDefinition
    Set first_column = detail_columns.Item(0)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, operations.Count
    Assert.Equals "OpenList", first_operation.OperationName
    Assert.Equals "css selector", first_operation.LocatorType
    Assert.Equals "#open-list", first_operation.LocatorValue
    Assert.Equals "ListReady", first_operation.WaitConditionName
    Assert.EqualsNumeric 2, detail_columns.Count
    Assert.Equals "対象ID", first_column.OutputColumnName
    Assert.Equals "#target-id", first_column.Selector
    Assert.Equals "InnerText", first_column.ExtractType
    Assert.IsTrue first_column.IsRequired
End Sub

Public Sub Test_ToolSettings_Settingsシート_必須設定不足はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseSettingsStubs("C:\Workbook")
    Call pSetMissingSetting(ws_stub, "StartUrl")

    Dim tool_settings As IToolSettings
    Set tool_settings = New ToolSettings

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

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = pParamSearchRange()

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = pParamUsedRange()

    Call ws_stub.Store.SetReturn("GetUsedRangeBounds", used_bounds, search_bounds, True, True, True, False)

    Set pUseSettingsStubs = ws_stub
End Function

Private Function pParamSearchRange() As WorksheetRangeBounds
    Set pParamSearchRange = New_RangeBounds( _
            Row:=2, _
            Column:=1, _
            FinishRow:=G_ROW_MAX, _
            FinishColumn:=1, _
            Sheet:=G_WEB_SETTINGS_SHEET)
End Function

Private Function pParamUsedRange() As WorksheetRangeBounds
    Set pParamUsedRange = New_RangeBounds( _
            Row:=2, _
            Column:=1, _
            FinishRow:=100, _
            FinishColumn:=1, _
            Sheet:=G_WEB_SETTINGS_SHEET)
End Function

Private Sub pSetSetting(ByVal WsStub As WorksheetServiceTestDouble, ByVal SettingName As String, ByVal SettingValue As String, ByVal RowIndex As Long)
    Dim found_list As ObjectList
    Set found_list = New ObjectList
    Call found_list.Add(New_RangeBounds(Row:=RowIndex, Column:=1, Sheet:=G_WEB_SETTINGS_SHEET))

    Call WsStub.Store.SetReturn("Find", found_list, SettingName, pParamUsedRange(), True, True, True, True)
    Call WsStub.Store.SetReturn("ReadCell", SettingValue, New_RangeBounds(Row:=RowIndex, Column:=2, Sheet:=G_WEB_SETTINGS_SHEET), False)
End Sub

Private Sub pSetMissingSetting(ByVal WsStub As WorksheetServiceTestDouble, ByVal SettingName As String)
    Dim found_list As ObjectList
    Set found_list = New ObjectList

    Call WsStub.Store.SetReturn("Find", found_list, SettingName, pParamUsedRange(), True, True, True, True)
End Sub

Private Sub pSetTransitionTable(ByVal WsStub As WorksheetServiceTestDouble)
    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=2, Column:=4, FinishRow:=G_ROW_MAX, FinishColumn:=8, Sheet:=G_WEB_SETTINGS_SHEET)

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = New_RangeBounds(Row:=2, Column:=4, FinishRow:=3, FinishColumn:=8, Sheet:=G_WEB_SETTINGS_SHEET)

    Dim table_values() As Variant
    ReDim table_values(1 To 2, 1 To 5)
    table_values(1, 1) = "OpenList"
    table_values(1, 2) = "css selector"
    table_values(1, 3) = "#open-list"
    table_values(1, 4) = ""
    table_values(1, 5) = "ListReady"
    table_values(2, 1) = "OpenDetail"
    table_values(2, 2) = "css selector"
    table_values(2, 3) = ".detail-link"
    table_values(2, 4) = ""
    table_values(2, 5) = "DetailReady"

    Call WsStub.Store.SetReturn("GetUsedRangeBounds", used_bounds, search_bounds, True, True, True, False)
    Call WsStub.Store.SetReturn("ReadRange", table_values, used_bounds)
End Sub

Private Sub pSetDetailColumnTable(ByVal WsStub As WorksheetServiceTestDouble)
    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=2, Column:=10, FinishRow:=G_ROW_MAX, FinishColumn:=15, Sheet:=G_WEB_SETTINGS_SHEET)

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = New_RangeBounds(Row:=2, Column:=10, FinishRow:=3, FinishColumn:=15, Sheet:=G_WEB_SETTINGS_SHEET)

    Dim table_values() As Variant
    ReDim table_values(1 To 2, 1 To 6)
    table_values(1, 1) = "対象ID"
    table_values(1, 2) = "#target-id"
    table_values(1, 3) = "InnerText"
    table_values(1, 4) = ""
    table_values(1, 5) = "True"
    table_values(1, 6) = "ErrorIfBlank"
    table_values(2, 1) = "件名"
    table_values(2, 2) = "#title"
    table_values(2, 3) = "TextContent"
    table_values(2, 4) = ""
    table_values(2, 5) = "False"
    table_values(2, 6) = "AllowBlank"

    Call WsStub.Store.SetReturn("GetUsedRangeBounds", used_bounds, search_bounds, True, True, True, False)
    Call WsStub.Store.SetReturn("ReadRange", table_values, used_bounds)
End Sub
