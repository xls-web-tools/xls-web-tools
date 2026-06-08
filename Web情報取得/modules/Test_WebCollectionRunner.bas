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
    tool_settings.ExistingRowMode = G_WEB_ROW_MODE_RETRY_ERROR
    tool_settings.TimeoutSeconds = 1

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
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

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
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""return-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#return-list"))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/return-list-element/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

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
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", "{""script"":""openDetail(1)"",""args"":[]}")
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", "{""script"":""openDetail(0)"",""args"":[]}")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/return-list-element/click", "{}")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
    Call pAssertWrittenCell(Assert, ws_stub, 6, 2, "T-002")
    Call pAssertWrittenCell(Assert, ws_stub, 6, 3, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 6, 4, "")
    Call pAssertWrittenCell(Assert, ws_stub, 6, 5, "案件B")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=5, Column:=3, Sheet:="output"))
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