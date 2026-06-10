Attribute VB_Name = "Test_WebNavDiagnosticRunner"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Web navigation diagnostic runner のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_WebNavDiagnosticRunner_StartUrlから一覧画面到達まで診断する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim auth_find_body As String
    auth_find_body = "{""using"":""css selector"",""value"":""#top-ready""}"

    Dim open_list_find_body As String
    open_list_find_body = "{""using"":""css selector"",""value"":""#open-list""}"

    Dim list_find_body As String
    list_find_body = "{""using"":""css selector"",""value"":""#list-ready""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", auth_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", open_list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = runner.Run()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.EqualsNumeric 1, process.Store.GetCallCount("Start")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.IsFalse process.IsRunning
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", auth_find_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", open_list_find_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/open-list-element/click", "{}")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", list_find_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Public Sub Test_WebNavDiagnosticRunner_先頭一覧項目から詳細ページへ入り対象IDを抽出する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemTargetIdSelector = "#list-item-target-id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "css selector", ".first-detail-link", WaitConditionName:="DetailReady"))
    Set tool_settings.TransitionOperations = operations

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim auth_find_body As String
    auth_find_body = "{""using"":""css selector"",""value"":""#top-ready""}"

    Dim open_list_find_body As String
    open_list_find_body = "{""using"":""css selector"",""value"":""#open-list""}"

    Dim list_find_body As String
    list_find_body = "{""using"":""css selector"",""value"":""#list-ready""}"

    Dim list_target_find_body As String
    list_target_find_body = "{""using"":""css selector"",""value"":""#list-item-target-id""}"

    Dim open_detail_find_body As String
    open_detail_find_body = "{""using"":""css selector"",""value"":"".first-detail-link""}"

    Dim target_find_body As String
    target_find_body = "{""using"":""css selector"",""value"":""#target-id""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", auth_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", open_list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-target-element""}}", "POST", "/session/abc/element", list_target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-001""}", "GET", "/session/abc/element/list-target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""detail-link-element""}}", "POST", "/session/abc/element", open_detail_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/detail-link-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""target-element""}}", "POST", "/session/abc/element", target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-001""}", "GET", "/session/abc/element/target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = runner.Run()

    Dim actual_target_id As String
    actual_target_id = runner.TargetId

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.Equals "T-001", actual_target_id
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "GET", "/session/abc/element/list-target-element/text", "")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/detail-link-element/click", "{}")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "GET", "/session/abc/element/target-element/text", "")
End Sub

Public Sub Test_WebNavDiagnosticRunner_詳細列定義に基づく診断出力行を書く(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub
    Call pPrepareEmptyOutput(ws_stub)

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Call fs_stub.Store.SetReturn("CreateTemporaryDirectory", "C:\Temp\xls-web-tools_tmp123.tmp", "xls-web-tools_")
    Set FsSrv = fs_stub

    Dim output_target_search_bounds As WorksheetRangeBounds
    Set output_target_search_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=1, Sheet:="output")

    Dim output_found_rows As ObjectList
    Set output_found_rows = New_ObjectList("WorksheetRangeBounds")
    Call ws_stub.Store.SetReturn("Find", output_found_rows, "T-001", output_target_search_bounds, True, True, True, True)
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.OutputSheetName = "output"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemTargetIdSelector = "#list-item-target-id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.DownloadEnabled = True
    tool_settings.DownloadRootPath = "D:\Root"
    tool_settings.DownloadLinkSelector = "#download"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "css selector", ".first-detail-link", WaitConditionName:="DetailReady"))
    Set tool_settings.TransitionOperations = operations

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject"))
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""],""prefs"":{""download.default_directory"":""C:\\Temp\\xls-web-tools_tmp123.tmp"",""download.prompt_for_download"":false,""download.directory_upgrade"":true}}}}}"

    Dim auth_find_body As String
    auth_find_body = "{""using"":""css selector"",""value"":""#top-ready""}"

    Dim open_list_find_body As String
    open_list_find_body = "{""using"":""css selector"",""value"":""#open-list""}"

    Dim list_find_body As String
    list_find_body = "{""using"":""css selector"",""value"":""#list-ready""}"

    Dim list_target_find_body As String
    list_target_find_body = "{""using"":""css selector"",""value"":""#list-item-target-id""}"

    Dim open_detail_find_body As String
    open_detail_find_body = "{""using"":""css selector"",""value"":"".first-detail-link""}"

    Dim target_find_body As String
    target_find_body = "{""using"":""css selector"",""value"":""#target-id""}"

    Dim subject_find_body As String
    subject_find_body = "{""using"":""css selector"",""value"":""#subject""}"

    Dim requester_find_body As String
    requester_find_body = "{""using"":""css selector"",""value"":""#requester""}"

    Dim download_find_body As String
    download_find_body = "{""using"":""css selector"",""value"":""#download""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", auth_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", open_list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-target-element""}}", "POST", "/session/abc/element", list_target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-001""}", "GET", "/session/abc/element/list-target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""detail-link-element""}}", "POST", "/session/abc/element", open_detail_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/detail-link-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""target-element""}}", "POST", "/session/abc/element", target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-001""}", "GET", "/session/abc/element/target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""subject-element""}}", "POST", "/session/abc/element", subject_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""案件A""}", "GET", "/session/abc/element/subject-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""requester-element""}}", "POST", "/session/abc/element", requester_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""山田太郎""}", "GET", "/session/abc/element/requester-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", "{""id"":null}")
    Call client_double.Store.SetReturn("Execute", "{""value"":[]}", "POST", "/session/abc/elements", download_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Call runner.Run

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 2, 1, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 2, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 2, 3, "")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 4, G_WEB_DOWNLOAD_STATUS_NO_FILE)
    Call pAssertWrittenCell(Assert, ws_stub, 2, 5, "案件A")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 6, "山田太郎")
End Sub

Public Sub Test_WebNavDiagnosticRunner_詳細列定義なしでも固定管理列だけの診断行を書く(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim output_target_search_bounds As WorksheetRangeBounds
    Set output_target_search_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=1, Sheet:="output")

    Dim output_found_rows As ObjectList
    Set output_found_rows = New_ObjectList("WorksheetRangeBounds")
    Call ws_stub.Store.SetReturn("Find", output_found_rows, "T-001", output_target_search_bounds, True, True, True, True)

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.OutputSheetName = "output"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemTargetIdSelector = "#list-item-target-id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "css selector", ".first-detail-link", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

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
    Call pSetTextElement(client_double, "#list-item-target-id", "list-target-element", "T-001")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""detail-link-element""}}", "POST", "/session/abc/element", pCssFindBody(".first-detail-link"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/detail-link-element/click", "{}")
    Call pSetTextElement(client_double, "#target-id", "target-element", "T-001")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""return-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#return-list"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/return-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Call runner.Run

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 1, 1, G_WEB_OUTPUT_COL_TARGET_ID)
    Call pAssertWrittenCell(Assert, ws_stub, 1, 2, G_WEB_OUTPUT_COL_STATUS)
    Call pAssertWrittenCell(Assert, ws_stub, 1, 3, G_WEB_OUTPUT_COL_ERROR)
    Call pAssertWrittenCell(Assert, ws_stub, 1, 4, G_WEB_OUTPUT_COL_DOWNLOAD_STATUS)
    Call pAssertWrittenCell(Assert, ws_stub, 2, 1, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 2, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 2, 3, "")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 4, "")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=2, Column:=5, Sheet:="output"))
End Sub

Public Sub Test_WebNavDiagnosticRunner_詳細列定義なしでDownloadRequiredならNO_FILEをERROR行にする(ByVal Assert As UnitTestAssert)
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
    Set FsSrv = fs_stub

    Dim output_target_search_bounds As WorksheetRangeBounds
    Set output_target_search_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=1, Sheet:="output")

    Dim output_found_rows As ObjectList
    Set output_found_rows = New_ObjectList("WorksheetRangeBounds")
    Call ws_stub.Store.SetReturn("Find", output_found_rows, "T-001", output_target_search_bounds, True, True, True, True)

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.OutputSheetName = "output"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemTargetIdSelector = "#list-item-target-id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"
    tool_settings.DownloadEnabled = True
    tool_settings.DownloadRequired = True
    tool_settings.DownloadRootPath = "D:\Root"
    tool_settings.DownloadLinkSelector = "#download"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "css selector", ".first-detail-link", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

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
    Call pSetTextElement(client_double, "#list-item-target-id", "list-target-element", "T-001")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""detail-link-element""}}", "POST", "/session/abc/element", pCssFindBody(".first-detail-link"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/detail-link-element/click", "{}")
    Call pSetTextElement(client_double, "#target-id", "target-element", "T-001")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", "{""id"":null}")
    Call client_double.Store.SetReturn("Execute", "{""value"":[]}", "POST", "/session/abc/elements", pCssFindBody("#download"))
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""return-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#return-list"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/return-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Call runner.Run

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 2, 1, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 2, G_WEB_STATUS_ERROR)
    Call pAssertWrittenCellContains(Assert, ws_stub, 2, 3, "ダウンロードリンクがありません")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 4, G_WEB_DOWNLOAD_STATUS_NO_FILE)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/elements", pCssFindBody("#download"))
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=2, Column:=5, Sheet:="output"))
End Sub

Public Sub Test_WebNavDiagnosticRunner_詳細列定義なしでもダウンロード済みファイルを保存する(ByVal Assert As UnitTestAssert)
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
    Call fs_stub.Store.SetReturn("GetFileList", crdownload_files, "C:\Temp\xls-web-tools_tmp123.tmp", "\.(crdownload|tmp)$", "", True)

    Dim completed_files(0 To 0) As String
    completed_files(0) = "C:\Temp\xls-web-tools_tmp123.tmp\report.pdf"
    Call fs_stub.Store.SetReturn("GetFileList", completed_files, "C:\Temp\xls-web-tools_tmp123.tmp", "", "\.(crdownload|tmp)$", True)
    Call fs_stub.Store.SetReturn("CreateDirectory", True, "D:\Root\T-001", False, True)
    Set FsSrv = fs_stub

    Dim output_target_search_bounds As WorksheetRangeBounds
    Set output_target_search_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=1, Sheet:="output")

    Dim output_found_rows As ObjectList
    Set output_found_rows = New_ObjectList("WorksheetRangeBounds")
    Call ws_stub.Store.SetReturn("Find", output_found_rows, "T-001", output_target_search_bounds, True, True, True, True)

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.OutputSheetName = "output"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemTargetIdSelector = "#list-item-target-id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"
    tool_settings.DownloadEnabled = True
    tool_settings.DownloadRootPath = "D:\Root"
    tool_settings.DownloadLinkSelector = "#download"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "css selector", ".first-detail-link", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

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
    Call pSetTextElement(client_double, "#list-item-target-id", "list-target-element", "T-001")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""detail-link-element""}}", "POST", "/session/abc/element", pCssFindBody(".first-detail-link"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/detail-link-element/click", "{}")
    Call pSetTextElement(client_double, "#target-id", "target-element", "T-001")
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

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Call runner.Run

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 2, 1, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 2, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 2, 3, "")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 4, G_WEB_DOWNLOAD_STATUS_DOWNLOADED)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/elements", pCssFindBody("#download"))
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("MoveFile", "C:\Temp\xls-web-tools_tmp123.tmp\report.pdf", "D:\Root\T-001\001_report.pdf", False)
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=2, Column:=5, Sheet:="output"))
End Sub

Public Sub Test_WebNavDiagnosticRunner_条件不一致なら診断出力行を書かない(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub
    Call pPrepareEmptyOutput(ws_stub)

    Dim output_target_search_bounds As WorksheetRangeBounds
    Set output_target_search_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=1, Sheet:="output")

    Dim output_found_rows As ObjectList
    Set output_found_rows = New_ObjectList("WorksheetRangeBounds")
    Call ws_stub.Store.SetReturn("Find", output_found_rows, "T-001", output_target_search_bounds, True, True, True, True)

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.OutputSheetName = "output"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemTargetIdSelector = "#list-item-target-id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"
    tool_settings.OutputConditionExpression = "[判定] == ""対象"""

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "css selector", ".first-detail-link", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("判定", "#decision"))
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject", IsRequired:=True, BlankMode:="ErrorIfBlank"))
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
    Call pSetTextElement(client_double, "#list-item-target-id", "list-target-element", "T-001")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""detail-link-element""}}", "POST", "/session/abc/element", pCssFindBody(".first-detail-link"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/detail-link-element/click", "{}")
    Call pSetTextElement(client_double, "#target-id", "target-element", "T-001")
    Call client_double.Store.SetReturn("Execute.AnyRequestBody", "{""value"": [""対象外""]}", "POST", "/session/abc/execute/sync")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""return-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#return-list"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/return-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = runner.Run()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.IsTrue runner.IsOutputExcluded
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=2, Column:=1, Sheet:="output"))
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/return-list-element/click", "{}")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub
Public Sub Test_WebNavDiagnosticRunner_必須詳細列が見つからない場合はERROR行を書く(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub
    Call pPrepareEmptyOutput(ws_stub)

    Dim output_target_search_bounds As WorksheetRangeBounds
    Set output_target_search_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=1, Sheet:="output")

    Dim output_found_rows As ObjectList
    Set output_found_rows = New_ObjectList("WorksheetRangeBounds")
    Call ws_stub.Store.SetReturn("Find", output_found_rows, "T-001", output_target_search_bounds, True, True, True, True)
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.OutputSheetName = "output"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemTargetIdSelector = "#list-item-target-id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "css selector", ".first-detail-link", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject", IsRequired:=True))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim auth_find_body As String
    auth_find_body = "{""using"":""css selector"",""value"":""#top-ready""}"

    Dim open_list_find_body As String
    open_list_find_body = "{""using"":""css selector"",""value"":""#open-list""}"

    Dim list_find_body As String
    list_find_body = "{""using"":""css selector"",""value"":""#list-ready""}"

    Dim list_target_find_body As String
    list_target_find_body = "{""using"":""css selector"",""value"":""#list-item-target-id""}"

    Dim open_detail_find_body As String
    open_detail_find_body = "{""using"":""css selector"",""value"":"".first-detail-link""}"

    Dim target_find_body As String
    target_find_body = "{""using"":""css selector"",""value"":""#target-id""}"

    Dim subject_find_body As String
    subject_find_body = "{""using"":""css selector"",""value"":""#subject""}"

    Dim return_list_find_body As String
    return_list_find_body = "{""using"":""css selector"",""value"":""#return-list""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", auth_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", open_list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-target-element""}}", "POST", "/session/abc/element", list_target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-001""}", "GET", "/session/abc/element/list-target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""detail-link-element""}}", "POST", "/session/abc/element", open_detail_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/detail-link-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""target-element""}}", "POST", "/session/abc/element", target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-001""}", "GET", "/session/abc/element/target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""error"":""no such element"",""message"":""missing subject""}}", "POST", "/session/abc/element", subject_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""return-list-element""}}", "POST", "/session/abc/element", return_list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/return-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Call runner.Run

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 2, 1, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 2, G_WEB_STATUS_ERROR)
    Call pAssertWrittenCellContains(Assert, ws_stub, 2, 3, "no such element")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 5, "")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/return-list-element/click", "{}")
    Assert.EqualsNumeric 2, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", list_find_body)
End Sub

Public Sub Test_WebNavDiagnosticRunner_一覧上と詳細ページの対象ID不一致はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemTargetIdSelector = "#list-item-target-id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "css selector", ".first-detail-link", WaitConditionName:="DetailReady"))
    Set tool_settings.TransitionOperations = operations

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim auth_find_body As String
    auth_find_body = "{""using"":""css selector"",""value"":""#top-ready""}"

    Dim open_list_find_body As String
    open_list_find_body = "{""using"":""css selector"",""value"":""#open-list""}"

    Dim list_find_body As String
    list_find_body = "{""using"":""css selector"",""value"":""#list-ready""}"

    Dim list_target_find_body As String
    list_target_find_body = "{""using"":""css selector"",""value"":""#list-item-target-id""}"

    Dim open_detail_find_body As String
    open_detail_find_body = "{""using"":""css selector"",""value"":"".first-detail-link""}"

    Dim target_find_body As String
    target_find_body = "{""using"":""css selector"",""value"":""#target-id""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", auth_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", open_list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-target-element""}}", "POST", "/session/abc/element", list_target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-001""}", "GET", "/session/abc/element/list-target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""detail-link-element""}}", "POST", "/session/abc/element", open_detail_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/detail-link-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""target-element""}}", "POST", "/session/abc/element", target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-999""}", "GET", "/session/abc/element/target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Call runner.Run

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "一覧上の対象IDと詳細ページ上の対象IDが一致しません", vbTextCompare)
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Public Sub Test_WebNavDiagnosticRunner_可視ブラウザ診断中のエラーではブラウザを残す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = False
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.TimeoutSeconds = 1

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile""]}}}}"

    Dim auth_find_body As String
    auth_find_body = "{""using"":""css selector"",""value"":""#top-ready""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetError("Execute", vbObjectError + 1, "Class WebDriverClient", "WebDriver HTTP error 404。Not Found", "POST", "/session/abc/element", auth_find_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Call runner.Run

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "ブラウザを残しました", vbTextCompare)
    Assert.EqualsNumeric 1, process.Store.GetCallCount("Start")
    Assert.EqualsNumeric 0, process.Store.GetCallCount("StopProcess")
    Assert.IsTrue process.IsRunning
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Public Sub Test_WebNavDiagnosticRunner_詳細ページ後に一覧復帰リンクで戻る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemTargetIdSelector = "#list-item-target-id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "css selector", ".first-detail-link", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim auth_find_body As String
    auth_find_body = "{""using"":""css selector"",""value"":""#top-ready""}"

    Dim open_list_find_body As String
    open_list_find_body = "{""using"":""css selector"",""value"":""#open-list""}"

    Dim list_find_body As String
    list_find_body = "{""using"":""css selector"",""value"":""#list-ready""}"

    Dim list_target_find_body As String
    list_target_find_body = "{""using"":""css selector"",""value"":""#list-item-target-id""}"

    Dim open_detail_find_body As String
    open_detail_find_body = "{""using"":""css selector"",""value"":"".first-detail-link""}"

    Dim target_find_body As String
    target_find_body = "{""using"":""css selector"",""value"":""#target-id""}"

    Dim return_list_find_body As String
    return_list_find_body = "{""using"":""css selector"",""value"":""#return-list""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", auth_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", open_list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-target-element""}}", "POST", "/session/abc/element", list_target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-001""}", "GET", "/session/abc/element/list-target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""detail-link-element""}}", "POST", "/session/abc/element", open_detail_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/detail-link-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""target-element""}}", "POST", "/session/abc/element", target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-001""}", "GET", "/session/abc/element/target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""return-list-element""}}", "POST", "/session/abc/element", return_list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/return-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Call runner.Run

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/return-list-element/click", "{}")
    Assert.EqualsNumeric 2, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", list_find_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub
Public Sub Test_WebNavDiagnosticRunner_一覧復帰失敗は復帰不能エラーにする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"
    tool_settings.ListItemTargetIdSelector = "#list-item-target-id"
    tool_settings.DetailTransitionOperationName = "OpenDetail"
    tool_settings.TargetIdSelector = "#target-id"
    tool_settings.ReturnToListOperationName = "ReturnToList"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Call operations.Add(New_TransitionOperation("OpenDetail", "css selector", ".first-detail-link", WaitConditionName:="DetailReady"))
    Call operations.Add(New_TransitionOperation("ReturnToList", "css selector", "#return-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim auth_find_body As String
    auth_find_body = "{""using"":""css selector"",""value"":""#top-ready""}"

    Dim open_list_find_body As String
    open_list_find_body = "{""using"":""css selector"",""value"":""#open-list""}"

    Dim list_find_body As String
    list_find_body = "{""using"":""css selector"",""value"":""#list-ready""}"

    Dim list_target_find_body As String
    list_target_find_body = "{""using"":""css selector"",""value"":""#list-item-target-id""}"

    Dim open_detail_find_body As String
    open_detail_find_body = "{""using"":""css selector"",""value"":"".first-detail-link""}"

    Dim target_find_body As String
    target_find_body = "{""using"":""css selector"",""value"":""#target-id""}"

    Dim return_list_find_body As String
    return_list_find_body = "{""using"":""css selector"",""value"":""#return-list""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", auth_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", open_list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-target-element""}}", "POST", "/session/abc/element", list_target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-001""}", "GET", "/session/abc/element/list-target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""detail-link-element""}}", "POST", "/session/abc/element", open_detail_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/detail-link-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""target-element""}}", "POST", "/session/abc/element", target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""T-001""}", "GET", "/session/abc/element/target-element/text", "")
    Call client_double.Store.SetReturn("Execute", "{""value"":{""error"":""no such element"",""message"":""missing return link""}}", "POST", "/session/abc/element", return_list_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim runner As WebNavDiagnosticRunner
    Set runner = New_WebNavDiagnosticRunner(process, session_client, tool_settings)

    ' --- Act ---
    Call runner.Run

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "一覧画面へ復帰できません", vbTextCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "missing return link", vbTextCompare)
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub
Private Sub pPrepareEmptyOutput(ByVal WsStub As WorksheetServiceTestDouble)
    Dim header_search_bounds As WorksheetRangeBounds
    Set header_search_bounds = New_RangeBounds(Row:=1, Column:=5, FinishRow:=1, FinishColumn:=G_COL_MAX, Sheet:="output")

    Dim header_used_bounds As WorksheetRangeBounds
    Set header_used_bounds = New_RangeBounds(Row:=1, Column:=5, FinishRow:=1, FinishColumn:=6, Sheet:="output")

    Dim header_values(1 To 1, 1 To 2) As Variant
    header_values(1, 1) = "件名"
    header_values(1, 2) = "申請者"

    Call WsStub.Store.SetReturn("GetUsedRangeBounds", header_used_bounds, header_search_bounds, True, True, True, False)
    Call WsStub.Store.SetReturn("ReadRange", header_values, header_used_bounds)
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

Private Function pGetWrittenCellValue( _
        ByVal Assert As UnitTestAssert, _
        ByVal WsStub As WorksheetServiceTestDouble, _
        ByVal RowIndex As Long, _
        ByVal ColumnIndex As Long) As String

    Dim target_bounds As WorksheetRangeBounds
    Set target_bounds = New_RangeBounds(Row:=RowIndex, Column:=ColumnIndex, Sheet:="output")

    Assert.EqualsNumeric 1, WsStub.Store.GetCallCount("WriteCell", target_bounds), CaseName:="Row " & CStr(RowIndex) & ", Column " & CStr(ColumnIndex) & " call count"

    Dim call_record As TestDoubleCallRecord
    Set call_record = WsStub.Store.GetLatestCall("WriteCell", target_bounds)

    pGetWrittenCellValue = CStr(call_record.GetArgument(1))
End Function

Private Sub pAssertWrittenCellContains( _
        ByVal Assert As UnitTestAssert, _
        ByVal WsStub As WorksheetServiceTestDouble, _
        ByVal RowIndex As Long, _
        ByVal ColumnIndex As Long, _
        ByVal ExpectedPart As String)

    Dim actual_value As String
    actual_value = pGetWrittenCellValue(Assert, WsStub, RowIndex, ColumnIndex)

    Assert.IsTrue 0 < InStr(1, actual_value, ExpectedPart, vbTextCompare), CaseName:="Row " & CStr(RowIndex) & ", Column " & CStr(ColumnIndex) & " contains"
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
