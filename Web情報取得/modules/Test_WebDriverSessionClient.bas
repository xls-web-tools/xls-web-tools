Attribute VB_Name = "Test_WebDriverSessionClient"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! WebDriver session client のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_WebDriverSessionClient_HeadlessTrueで不可視ブラウザCapabilitiesを作る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim expected_body As String
    expected_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", expected_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = session_client.CreateSession()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session", expected_body)
End Sub

Public Sub Test_WebDriverSessionClient_QuitSessionでWebDriverSessionを終了する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim session_id As String
    session_id = session_client.CreateSession()
    Err.Clear

    ' --- Act ---
    CallByName session_client, "QuitSession", VbMethod

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", session_id
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Public Sub Test_WebDriverSessionClient_NavigateToStartUrlSendsUrlCommand(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.StartUrl = "https://example.test/start"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.NavigateToStartUrl

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/url", "{""url"":""https://example.test/start""}")
End Sub

Public Sub Test_WebDriverSessionClient_WaitForSelectorでSelector出現を判定する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim find_body As String
    find_body = "{""using"":""css selector"",""value"":""#top-ready""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""element-1""}}", "POST", "/session/abc/element", find_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Dim actual_element_id As String
    actual_element_id = session_client.WaitForSelector("#top-ready")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "element-1", actual_element_id
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", find_body)
End Sub

Public Sub Test_WebDriverSessionClient_WaitForSelectorでFrame内Selector出現を判定する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim switch_top_body As String
    switch_top_body = "{""id"":null}"

    Dim frame_find_body As String
    frame_find_body = "{""using"":""css selector"",""value"":""iframe[name='right']""}"

    Dim switch_frame_body As String
    switch_frame_body = "{""id"":{""element-6066-11e4-a52e-4f735466cecf"":""frame-1""}}"

    Dim target_find_body As String
    target_find_body = "{""using"":""css selector"",""value"":""#list-ready""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", switch_top_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""frame-1""}}", "POST", "/session/abc/element", frame_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", switch_frame_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""target-1""}}", "POST", "/session/abc/element", target_find_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Dim actual_element_id As String
    actual_element_id = session_client.WaitForSelector("iframe[name='right'] >> #list-ready")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "target-1", actual_element_id
    Assert.EqualsNumeric 2, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/frame", switch_top_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", frame_find_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/frame", switch_frame_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", target_find_body)
End Sub
Public Sub Test_WebDriverSessionClient_RunTransitionOperationでClickを優先する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim find_body As String
    find_body = "{""using"":""css selector"",""value"":""#open-list""}"

    Dim script_body As String
    script_body = "{""script"":""openList()"",""args"":[]}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""element-1""}}", "POST", "/session/abc/element", find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/element-1/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/execute/sync", script_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim operation As TransitionOperation
    Set operation = New_TransitionOperation("OpenList", "css selector", "#open-list", Script:="openList()", WaitConditionName:="ListReady")

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.RunTransitionOperation(operation)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", find_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/element-1/click", "{}")
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", script_body)
End Sub

Public Sub Test_WebDriverSessionClient_RunTransitionOperationでFrame内要素をClickする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim switch_top_body As String
    switch_top_body = "{""id"":null}"

    Dim frame_find_body As String
    frame_find_body = "{""using"":""css selector"",""value"":""iframe[name='right']""}"

    Dim switch_frame_body As String
    switch_frame_body = "{""id"":{""element-6066-11e4-a52e-4f735466cecf"":""frame-1""}}"

    Dim target_find_body As String
    target_find_body = "{""using"":""css selector"",""value"":""#list tbody tr:first-child td:nth-child(2)""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", switch_top_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""frame-1""}}", "POST", "/session/abc/element", frame_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", switch_frame_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""target-1""}}", "POST", "/session/abc/element", target_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/target-1/click", "{}")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim operation As TransitionOperation
    Set operation = New_TransitionOperation("OpenList", "css selector", "iframe[name='right'] >> #list tbody tr:first-child td:nth-child(2)", WaitConditionName:="ListReady")

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.RunTransitionOperation(operation)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/frame", switch_top_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", frame_find_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/frame", switch_frame_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", target_find_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/target-1/click", "{}")
End Sub
Public Sub Test_WebDriverSessionClient_RunTransitionOperationでScriptOnlyを実行する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim script_body As String
    script_body = "{""script"":""openList()"",""args"":[]}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/execute/sync", script_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim operation As TransitionOperation
    Set operation = New_TransitionOperation("OpenList", "", "", Script:="openList()", WaitConditionName:="ListReady")

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.RunTransitionOperation(operation)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", script_body)
End Sub

Public Sub Test_WebDriverSessionClient_HeadlessFalseで可視ブラウザCapabilitiesを作る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = False
    tool_settings.BrowserProfilePath = "D:\DedicatedProfile"

    Dim expected_body As String
    expected_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=D:\\DedicatedProfile""]}}}}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""visible""}}", "POST", "/session", expected_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = session_client.CreateSession()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "visible", actual_session_id
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session", expected_body)
End Sub

Public Sub Test_WebDriverSessionClient_WebDriverErrorResponseを原因付きで再送出する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""error"":""session not created"",""message"":""profile locked""}}", "POST", "/session", create_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = session_client.CreateSession()

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "session not created", vbTextCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "profile locked", vbTextCompare)
End Sub
