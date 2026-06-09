Attribute VB_Name = "Test_WebCollectionRunner"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Web collection runner のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_WebCollectionRunner_現在ページを対象ID主キーで巡回し既存OKはスキップする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Call fs_stub.Store.SetReturn("CreateTemporaryDirectory", "C:\Temp\xls-web-tools_tmp123.tmp", "xls-web-tools_")

    Dim crdownload_files() As String
    crdownload_files = EmptyStringArray()
    Call fs_stub.Store.SetReturn("GetFileList", crdownload_files, "C:\Temp\xls-web-tools_tmp123.tmp", "\.crdownload$", "", True)

    Dim completed_files(0 To 0) As String
    completed_files(0) = "C:\Temp\xls-web-tools_tmp123.tmp\report.pdf"
    Call fs_stub.Store.SetReturn("GetFileList", completed_files, "C:\Temp\xls-web-tools_tmp123.tmp", "", "\.crdownload$", True)
    Call fs_stub.Store.SetReturn("CreateDirectory", True, "D:\Root\T-002", False, True)
    Set FsSrv = fs_stub

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.OutputSheetName = "output"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemSelector = "#list tbody tr"
    tool_settings.ListItemTargetIdSelector = "#list tbody tr:nth-child({{rowNumber}}) td.id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"
    tool_settings.ExistingRowMode = G_WEB_ROW_MODE_SKIP_EXISTING
    tool_settings.TimeoutSeconds = 1
    tool_settings.DownloadEnabled = True
    tool_settings.DownloadRootPath = "D:\Root"
    tool_settings.DownloadLinkSelector = "#download"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "", "", Script:="openDetail({{index}})", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Call pPrepareExistingOutputRows(ws_stub)

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""],""prefs"":{""download.default_directory"":""C:\\Temp\\xls-web-tools_tmp123.tmp"",""download.prompt_for_download"":false,""download.directory_upgrade"":true}}}}}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", pCssFindBody("#top-ready"))
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#open-list"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", pCssFindBody("#list-ready"))
    Call client_double.Store.SetReturn("Execute", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""row-1""},{""element-6066-11e4-a52e-4f735466cecf"":""row-2""}]}", "POST", "/session/abc/elements", pCssFindBody("#list tbody tr"))
    Call pSetTextElement(client_double, "#list tbody tr:nth-child(1) td.id", "list-target-1", "T-001")
    Call pSetTextElement(client_double, "#list tbody tr:nth-child(2) td.id", "list-target-2", "T-002")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/execute/sync", "{""script"":""openDetail(1)"",""args"":[]}")
    Call pSetTextElement(client_double, "#target-id", "target-element", "T-002")
    Call client_double.Store.SetReturn("Execute.AnyRequestBody", "{""value"": [""案件B""]}", "POST", "/session/abc/execute/sync")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", "{""id"":null}")
    Call client_double.Store.SetReturn("Execute", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""download-1""}]}", "POST", "/session/abc/elements", pCssFindBody("#download"))
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""download-1""}}", "POST", "/session/abc/element", pCssFindBody("#download"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/download-1/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""return-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#return-list"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/return-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Set ProgStat = Nothing

    Dim runner As WebCollectionRunner
    Set runner = New_WebCollectionRunner(process, session_client, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = runner.Run()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.EqualsNumeric 1, runner.SucceededCount
    Assert.EqualsNumeric 1, runner.SkippedCount
    Assert.EqualsNumeric 0, runner.ErrorCount
    Assert.EqualsNumeric 1, runner.PageCount
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", "{""script"":""openDetail(1)"",""args"":[]}")
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", "{""script"":""openDetail(0)"",""args"":[]}")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/return-list-element/click", "{}")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
    Call pAssertWrittenCell(Assert, ws_stub, 6, 2, "T-002")
    Call pAssertWrittenCell(Assert, ws_stub, 6, 3, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 6, 4, "")
    Call pAssertWrittenCell(Assert, ws_stub, 6, 5, G_WEB_DOWNLOAD_STATUS_DOWNLOADED)
    Call pAssertWrittenCell(Assert, ws_stub, 6, 6, "案件B")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("MoveFile", "C:\Temp\xls-web-tools_tmp123.tmp\report.pdf", "D:\Root\T-002\002_report.pdf", False)
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=5, Column:=3, Sheet:="output"))
    Assert.IsNothing ProgStat
End Sub

Public Sub Test_WebCollectionRunner_初回ページ0件なら収集進捗を開始しない(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.OutputSheetName = "output"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemSelector = "#list tbody tr"
    tool_settings.ListItemTargetIdSelector = "#list tbody tr:nth-child({{rowNumber}}) td.id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"
    tool_settings.ExistingRowMode = G_WEB_ROW_MODE_SKIP_EXISTING
    tool_settings.TimeoutSeconds = 1

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "", "", Script:="openDetail({{index}})", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", pCssFindBody("#top-ready"))
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#open-list"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", pCssFindBody("#list-ready"))
    Call client_double.Store.SetReturn("Execute", "{""value"":[]}", "POST", "/session/abc/elements", pCssFindBody("#list tbody tr"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Set ProgStat = New ProgressStatus
    Application.StatusBar = False

    Dim runner As WebCollectionRunner
    Set runner = New_WebCollectionRunner(process, session_client, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = runner.Run()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.EqualsNumeric 1, runner.PageCount
    Assert.Equals "", ProgStat.TaskName
    Assert.EqualsNumeric 100, ProgStat.TotalValue
    Assert.EqualsNumeric 0, ProgStat.ProcessedValue
    Assert.IsFalse ProgStat.IsComplete
    Assert.Equals False, CBool(Application.StatusBar)
End Sub

Public Sub Test_WebCollectionRunner_既存ERROR行が条件不一致なら既存行を更新せず出力対象外に数える(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.OutputSheetName = "output"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemSelector = "#list tbody tr"
    tool_settings.ListItemTargetIdSelector = "#list tbody tr:nth-child({{rowNumber}}) td.id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"
    tool_settings.ExistingRowMode = G_WEB_ROW_MODE_SKIP_EXISTING
    tool_settings.OutputConditionExpression = "[判定] == ""対象"""
    tool_settings.TimeoutSeconds = 1

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "", "", Script:="openDetail({{index}})", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("判定", "#decision"))
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject", IsRequired:=True, BlankMode:="ErrorIfBlank"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim output_target_search_bounds As WorksheetRangeBounds
    Set output_target_search_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=G_ROW_MAX, FinishColumn:=2, Sheet:="output")

    Dim output_found_rows As ObjectList
    Set output_found_rows = New_ObjectList("WorksheetRangeBounds")
    Call output_found_rows.Add(New_RangeBounds(Row:=6, Column:=2, Sheet:="output"))
    Call ws_stub.Store.SetReturn("Find", output_found_rows, "T-001", output_target_search_bounds, True, True, True, True)
    Call ws_stub.Store.SetReturn("ReadCell", G_WEB_STATUS_ERROR, New_RangeBounds(Row:=6, Column:=3, Sheet:="output"), False)

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", pCssFindBody("#top-ready"))
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#open-list"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", pCssFindBody("#list-ready"))
    Call client_double.Store.SetReturn("Execute", "{""value"": [{""element-6066-11e4-a52e-4f735466cecf"":""row-1""}]}", "POST", "/session/abc/elements", pCssFindBody("#list tbody tr"))
    Call pSetTextElement(client_double, "#list tbody tr:nth-child(1) td.id", "list-target-1", "T-001")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/execute/sync", "{""script"":""openDetail(0)"",""args"":[]}")
    Call pSetTextElement(client_double, "#target-id", "target-element", "T-001")
    Call client_double.Store.SetReturn("Execute.AnyRequestBody", "{""value"": [""対象外""]}", "POST", "/session/abc/execute/sync")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""return-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#return-list"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/return-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Set ProgStat = Nothing

    Dim runner As WebCollectionRunner
    Set runner = New_WebCollectionRunner(process, session_client, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = runner.Run()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.EqualsNumeric 0, runner.SucceededCount
    Assert.EqualsNumeric 0, runner.SkippedCount
    Assert.EqualsNumeric 1, runner.OutputExcludedCount
    Assert.EqualsNumeric 0, runner.ErrorCount
    Assert.EqualsNumeric 1, runner.PageCount
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/return-list-element/click", "{}")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=2, Column:=2, Sheet:="output"))
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=6, Column:=3, Sheet:="output"))
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Public Sub Test_WebCollectionRunner_次ページがあれば全ページを巡回する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.OutputSheetName = "output"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemSelector = "#list tbody tr"
    tool_settings.ListItemTargetIdSelector = "#list tbody tr:nth-child({{rowNumber}}) td.id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"
    tool_settings.NextPageOperationName = "NextPage"
    tool_settings.NextPageAvailableSelector = "#next-page"
    tool_settings.ExistingRowMode = G_WEB_ROW_MODE_SKIP_EXISTING
    tool_settings.TimeoutSeconds = 1

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "", "", Script:="openDetail({{index}})", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("NextPage", "css selector", "#next-link", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Call pPrepareExistingOkOutputRows(ws_stub)

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", pCssFindBody("#top-ready"))
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#open-list"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", pCssFindBody("#list-ready"))
    Call client_double.Store.SetReturn("Execute.Sequence", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""row-page-1""}]}", "POST", "/session/abc/elements", pCssFindBody("#list tbody tr"), CLng(0))
    Call client_double.Store.SetReturn("Execute.Sequence", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""row-page-2""}]}", "POST", "/session/abc/elements", pCssFindBody("#list tbody tr"), CLng(1))
    Call client_double.Store.SetReturn("Execute.Sequence", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""next-page""}]}", "POST", "/session/abc/elements", pCssFindBody("#next-page"), CLng(0))
    Call client_double.Store.SetReturn("Execute.Sequence", "{""value"":[]}", "POST", "/session/abc/elements", pCssFindBody("#next-page"), CLng(1))
    Call pSetTextElementSequence(client_double, "#list tbody tr:nth-child(1) td.id", "list-target", Array("T-001", "T-001", "T-002", "T-002", "T-002"))
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""next-link-element""}}", "POST", "/session/abc/element", pCssFindBody("#next-link"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/next-link-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Set ProgStat = New ProgressStatus

    Dim runner As WebCollectionRunner
    Set runner = New_WebCollectionRunner(process, session_client, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = runner.Run()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.EqualsNumeric 0, runner.SucceededCount
    Assert.EqualsNumeric 2, runner.SkippedCount
    Assert.EqualsNumeric 0, runner.ErrorCount
    Assert.EqualsNumeric 2, runner.PageCount
    Assert.Equals "取得中", ProgStat.TaskName
    Assert.EqualsNumeric 2, ProgStat.TotalValue
    Assert.EqualsNumeric 2, ProgStat.ProcessedValue
    Assert.IsTrue ProgStat.IsComplete
    Assert.Equals False, CBool(Application.StatusBar)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/next-link-element/click", "{}")
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", "{""script"":""openDetail(0)"",""args"":[]}")
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", "{""script"":""openDetail(1)"",""args"":[]}")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Public Sub Test_WebCollectionRunner_次ページ操作後に先頭対象IDが変わらなければ中断する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.OutputSheetName = "output"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemSelector = "#list tbody tr"
    tool_settings.ListItemTargetIdSelector = "#list tbody tr:nth-child({{rowNumber}}) td.id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"
    tool_settings.NextPageOperationName = "NextPage"
    tool_settings.NextPageAvailableSelector = "#next-page"
    tool_settings.ExistingRowMode = G_WEB_ROW_MODE_SKIP_EXISTING
    tool_settings.TimeoutSeconds = 0

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "", "", Script:="openDetail({{index}})", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("NextPage", "css selector", "#next-link", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Call pPrepareSingleExistingOkOutputRow(ws_stub, "T-001", 5)

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", pCssFindBody("#top-ready"))
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#open-list"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", pCssFindBody("#list-ready"))
    Call client_double.Store.SetReturn("Execute.Sequence", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""row-page-1""}]}", "POST", "/session/abc/elements", pCssFindBody("#list tbody tr"), CLng(0))
    Call client_double.Store.SetReturn("Execute.Sequence", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""next-page""}]}", "POST", "/session/abc/elements", pCssFindBody("#next-page"), CLng(0))
    Call pSetTextElementSequence(client_double, "#list tbody tr:nth-child(1) td.id", "list-target", Array("T-001", "T-001", "T-001"))
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""next-link-element""}}", "POST", "/session/abc/element", pCssFindBody("#next-link"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/next-link-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Set ProgStat = New ProgressStatus

    Dim runner As WebCollectionRunner
    Set runner = New_WebCollectionRunner(process, session_client, tool_settings)

    ' --- Act ---
    Call runner.Run

    ' --- Assert ---
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue 0 < InStr(1, Err.Description, "ページが進みません", vbTextCompare)
    Assert.EqualsNumeric 1, runner.SkippedCount
    Assert.EqualsNumeric 1, runner.PageCount
    Assert.EqualsNumeric 0, ProgStat.ProcessedValue
    Assert.IsFalse ProgStat.IsComplete
    Assert.Equals False, CBool(Application.StatusBar)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/next-link-element/click", "{}")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Private Sub pPrepareExistingOutputRows(ByVal WsStub As WorksheetServiceTestDouble)
    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=G_ROW_MAX, FinishColumn:=2, Sheet:="output")

    Dim ok_rows As ObjectList
    Set ok_rows = New_ObjectList("WorksheetRangeBounds")
    Call ok_rows.Add(New_RangeBounds(Row:=5, Column:=2, Sheet:="output"))
    Call WsStub.Store.SetReturn("Find", ok_rows, "T-001", search_bounds, True, True, True, True)
    Call WsStub.Store.SetReturn("ReadCell", G_WEB_STATUS_OK, New_RangeBounds(Row:=5, Column:=3, Sheet:="output"), False)

    Dim error_rows As ObjectList
    Set error_rows = New_ObjectList("WorksheetRangeBounds")
    Call error_rows.Add(New_RangeBounds(Row:=6, Column:=2, Sheet:="output"))
    Call WsStub.Store.SetReturn("Find", error_rows, "T-002", search_bounds, True, True, True, True)
    Call WsStub.Store.SetReturn("ReadCell", G_WEB_STATUS_ERROR, New_RangeBounds(Row:=6, Column:=3, Sheet:="output"), False)
End Sub

Private Sub pPrepareExistingOkOutputRows(ByVal WsStub As WorksheetServiceTestDouble)
    Call pPrepareSingleExistingOkOutputRow(WsStub, "T-001", 5)
    Call pPrepareSingleExistingOkOutputRow(WsStub, "T-002", 6)
End Sub

Private Sub pPrepareSingleExistingOkOutputRow( _
        ByVal WsStub As WorksheetServiceTestDouble, _
        ByVal TargetId As String, _
        ByVal OutputRow As Long)

    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=G_ROW_MAX, FinishColumn:=2, Sheet:="output")

    Dim found_rows As ObjectList
    Set found_rows = New_ObjectList("WorksheetRangeBounds")
    Call found_rows.Add(New_RangeBounds(Row:=OutputRow, Column:=2, Sheet:="output"))
    Call WsStub.Store.SetReturn("Find", found_rows, TargetId, search_bounds, True, True, True, True)
    Call WsStub.Store.SetReturn("ReadCell", G_WEB_STATUS_OK, New_RangeBounds(Row:=OutputRow, Column:=3, Sheet:="output"), False)
End Sub

Private Function pCssFindBody(ByVal Selector As String) As String
    pCssFindBody = "{""using"":""css selector"",""value"":""" & Selector & """}"
End Function

Private Sub pSetTextElement( _
        ByVal ClientDouble As WebDriverClientTestDouble, _
        ByVal Selector As String, _
        ByVal ElementId As String, _
        ByVal ElementText As String)

    Call ClientDouble.Store.SetReturn( _
            "Execute", _
            "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""" & ElementId & """}}", _
            "POST", _
            "/session/abc/element", _
            pCssFindBody(Selector))
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":""" & ElementText & """}", "GET", "/session/abc/element/" & ElementId & "/text", "")
End Sub

Private Sub pSetTextElementSequence( _
        ByVal ClientDouble As WebDriverClientTestDouble, _
        ByVal Selector As String, _
        ByVal ElementId As String, _
        ByVal ElementTexts As Variant)

    Call ClientDouble.Store.SetReturn( _
            "Execute", _
            "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""" & ElementId & """}}", _
            "POST", _
            "/session/abc/element", _
            pCssFindBody(Selector))

    Dim text_idx As Long
    For text_idx = LBound(ElementTexts) To UBound(ElementTexts)
        Call ClientDouble.Store.SetReturn("Execute.Sequence", "{""value"":""" & CStr(ElementTexts(text_idx)) & """}", "GET", "/session/abc/element/" & ElementId & "/text", "", CLng(text_idx))
    Next text_idx
End Sub

Private Sub pAssertWrittenCell( _
        ByVal Assert As UnitTestAssert, _
        ByVal WsStub As WorksheetServiceTestDouble, _
        ByVal RowIndex As Long, _
        ByVal ColumnIndex As Long, _
        ByVal ExpectedValue As String)

    Dim target_bounds As WorksheetRangeBounds
    Set target_bounds = New_RangeBounds(Row:=RowIndex, Column:=ColumnIndex, Sheet:="output")

    Assert.EqualsNumeric 1, WsStub.Store.GetCallCount("WriteCell", target_bounds), CaseName:="Row " & CStr(RowIndex) & ", Column " & CStr(ColumnIndex) & " call count"

    Dim call_record As TestDoubleCallRecord
    Set call_record = WsStub.Store.GetLatestCall("WriteCell", target_bounds)

    Assert.Equals ExpectedValue, CStr(call_record.GetArgument(1)), CaseName:="Row " & CStr(RowIndex) & ", Column " & CStr(ColumnIndex)
End Sub
