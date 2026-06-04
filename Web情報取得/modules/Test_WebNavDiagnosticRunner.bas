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
