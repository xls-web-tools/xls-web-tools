Attribute VB_Name = "Test_WebCollectionRunLifecycle"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! WebCollectionRunLifecycle のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_WebCollectionRunLifecycle_一覧画面到達後にBodyを実行してSessionを正常終了する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = pCreateToolSettings()

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call pPrepareSuccessfulPreparationClient(client_double)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, tool_settings)

    Dim body As WebCollectionRunBodyTestDouble
    Set body = New WebCollectionRunBodyTestDouble

    Set ProgStat = New ProgressStatus

    Dim run_lifecycle As WebCollectionRunLifecycle
    Set run_lifecycle = New_WebCollectionRunLifecycle(lifecycle, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = run_lifecycle.Run(body)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.EqualsNumeric 1, body.Store.GetCallCount("RunAfterPreparation")
    Assert.EqualsNumeric 0, body.Store.GetCallCount("ClearProgressAfterError")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("Start")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.IsFalse process.IsRunning
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
    Assert.Equals "取得準備中", ProgStat.TaskName
    Assert.EqualsNumeric 5, ProgStat.ProcessedValue
End Sub

Public Sub Test_WebCollectionRunLifecycle_Body失敗時はBody進捗とSessionをCleanupして元エラーを再送出する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = pCreateToolSettings()

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call pPrepareSuccessfulPreparationClient(client_double)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, tool_settings)

    Dim body As WebCollectionRunBodyTestDouble
    Set body = New WebCollectionRunBodyTestDouble
    Call body.Store.SetError("RunAfterPreparation", vbObjectError + 333, "Class Body", "body failed")

    Set ProgStat = New ProgressStatus

    Dim run_lifecycle As WebCollectionRunLifecycle
    Set run_lifecycle = New_WebCollectionRunLifecycle(lifecycle, tool_settings)

    ' --- Act ---
    Call run_lifecycle.Run(body)

    ' --- Assert ---
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric vbObjectError + 333, Err.Number
    Assert.Equals "Class Body", Err.Source
    Assert.Equals "body failed", Err.Description
    Err.Clear
    Assert.EqualsNumeric 1, body.Store.GetCallCount("ClearProgressAfterError")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.IsFalse process.IsRunning
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Public Sub Test_WebCollectionRunLifecycle_BrowserProfilePath作成中止ならBodyを実行せず中断する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Missing\Profile"
    tool_settings.StartUrl = "https://example.test/start"
    tool_settings.AuthenticatedStartSelector = "#top-ready"
    tool_settings.ListPageSelector = "#list-ready"
    tool_settings.ListTransitionOperationName = "OpenList"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Set FsSrv = fs_stub
    Call fs_stub.Store.SetReturn("IsDirectory", False, "C:\Missing\Profile")

    Dim prompt As BrowserProfilePromptTestDouble
    Set prompt = New BrowserProfilePromptTestDouble
    Call prompt.Store.SetReturn("ConfirmCreateDirectory", False, "C:\Missing\Profile")

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, tool_settings)
    Set lifecycle.BrowserProfilePrompt = prompt

    Dim body As WebCollectionRunBodyTestDouble
    Set body = New WebCollectionRunBodyTestDouble

    Dim run_lifecycle As WebCollectionRunLifecycle
    Set run_lifecycle = New_WebCollectionRunLifecycle(lifecycle, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = run_lifecycle.Run(body)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", actual_session_id
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("IsDirectory", "C:\Missing\Profile")
    Assert.EqualsNumeric 1, prompt.Store.GetCallCount("ConfirmCreateDirectory", "C:\Missing\Profile")
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("CreateDirectory", "C:\Missing\Profile", False, True)
    Assert.EqualsNumeric 0, process.Store.GetCallCount("Start")
    Assert.EqualsNumeric 0, client_double.Store.GetCallCountAll("Execute")
    Assert.EqualsNumeric 0, body.Store.GetCallCount("RunAfterPreparation")
    Assert.EqualsNumeric 0, body.Store.GetCallCount("ClearProgressAfterError")
End Sub
Private Function pCreateToolSettings() As ToolSettingsTestDouble
    Dim result_value As ToolSettingsTestDouble
    Set result_value = New ToolSettingsTestDouble
    result_value.Headless = True
    result_value.BrowserProfilePath = "C:\Profile"
    Call pUseProfileDirectory("C:\Profile", True)
    result_value.StartUrl = "https://example.test/start"
    result_value.AuthenticatedStartSelector = "#top-ready"
    result_value.ListPageSelector = "#list-ready"
    result_value.ListTransitionOperationName = "OpenList"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Set result_value.TransitionOperations = operations

    Set pCreateToolSettings = result_value
End Function

Private Sub pUseProfileDirectory(ByVal DirectoryPath As String, ByVal Exists As Boolean)
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Set FsSrv = fs_stub
    Call fs_stub.Store.SetReturn("IsDirectory", Exists, DirectoryPath)
End Sub

Private Sub pPrepareSuccessfulPreparationClient(ByVal ClientDouble As WebDriverClientTestDouble)
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", pCreateSessionBody())
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""auth-element""}}", "POST", "/session/abc/element", pCssFindBody("#top-ready"))
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""open-list-element""}}", "POST", "/session/abc/element", pCssFindBody("#open-list"))
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/open-list-element/click", "{}")
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""list-element""}}", "POST", "/session/abc/element", pCssFindBody("#list-ready"))
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")
End Sub

Private Function pCreateSessionBody() As String
    pCreateSessionBody = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"
End Function

Private Function pCssFindBody(ByVal Selector As String) As String
    pCssFindBody = "{""using"":""css selector"",""value"":""" & Selector & """}"
End Function
