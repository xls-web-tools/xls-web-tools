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

Public Sub Test_WebDriverSessionClient_DownloadEnabledTrueで一時ダウンロード領域をCapabilitiesに入れる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.DownloadEnabled = True

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Call fs_stub.Store.SetReturn("CreateTemporaryDirectory", "C:\Temp\xls-web-tools_tmp123.tmp", "xls-web-tools_")
    Set FsSrv = fs_stub

    Dim expected_body As String
    expected_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""],""prefs"":{""download.default_directory"":""C:\\Temp\\xls-web-tools_tmp123.tmp"",""download.prompt_for_download"":false,""download.directory_upgrade"":true}}}}}"

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
    Assert.Equals "C:\Temp\xls-web-tools_tmp123.tmp", session_client.DownloadDirectoryPath
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("CreateTemporaryDirectory", "xls-web-tools_")
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

Public Sub Test_WebDriverSessionClient_RunTransitionOperationでLocatorInnerText一致要素をClickする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim elements_body As String
    elements_body = "{""using"":""css selector"",""value"":"".menu-button""}"

    Dim element_1_text_body As String
    element_1_text_body = pInnerTextScriptBody("element-1")

    Dim element_2_text_body As String
    element_2_text_body = pInnerTextScriptBody("element-2")

    Dim fallback_script_body As String
    fallback_script_body = "{""script"":""fallback()"",""args"":[]}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""element-1""},{""element-6066-11e4-a52e-4f735466cecf"":""element-2""}]}", "POST", "/session/abc/elements", elements_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""対象外""}", "POST", "/session/abc/execute/sync", element_1_text_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""一覧\r\nを   開く""}", "POST", "/session/abc/execute/sync", element_2_text_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/element-2/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/execute/sync", fallback_script_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim operation As TransitionOperation
    Set operation = New_TransitionOperation("OpenList", "css selector", ".menu-button", LocatorInnerText:="一覧 を 開く", Script:="fallback()", WaitConditionName:="ListReady")

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.RunTransitionOperation(operation)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/elements", elements_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", element_1_text_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", element_2_text_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/element-2/click", "{}")
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", fallback_script_body)
End Sub

Public Sub Test_WebDriverSessionClient_RunTransitionOperationでLocatorInnerText一致0件ならScriptFallbackしない(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim elements_body As String
    elements_body = "{""using"":""css selector"",""value"":"".menu-button""}"

    Dim element_1_text_body As String
    element_1_text_body = pInnerTextScriptBody("element-1")

    Dim element_2_text_body As String
    element_2_text_body = pInnerTextScriptBody("element-2")

    Dim fallback_script_body As String
    fallback_script_body = "{""script"":""fallback()"",""args"":[]}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""element-1""},{""element-6066-11e4-a52e-4f735466cecf"":""element-2""}]}", "POST", "/session/abc/elements", elements_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""対象外""}", "POST", "/session/abc/execute/sync", element_1_text_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""別項目""}", "POST", "/session/abc/execute/sync", element_2_text_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/execute/sync", fallback_script_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim operation As TransitionOperation
    Set operation = New_TransitionOperation("OpenList", "css selector", ".menu-button", LocatorInnerText:="一覧を開く", Script:="fallback()", WaitConditionName:="ListReady")

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.RunTransitionOperation(operation)

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "OperationName=OpenList", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "LocatorValue=.menu-button", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "LocatorInnerText=一覧を開く", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "MatchCount=0", vbBinaryCompare)
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", fallback_script_body)
End Sub

Public Sub Test_WebDriverSessionClient_RunTransitionOperationでLocatorInnerText一致複数件ならScriptFallbackしない(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim elements_body As String
    elements_body = "{""using"":""css selector"",""value"":"".menu-button""}"

    Dim element_1_text_body As String
    element_1_text_body = pInnerTextScriptBody("element-1")

    Dim element_2_text_body As String
    element_2_text_body = pInnerTextScriptBody("element-2")

    Dim fallback_script_body As String
    fallback_script_body = "{""script"":""fallback()"",""args"":[]}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""element-1""},{""element-6066-11e4-a52e-4f735466cecf"":""element-2""}]}", "POST", "/session/abc/elements", elements_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""一覧 を 開く""}", "POST", "/session/abc/execute/sync", element_1_text_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""一覧\nを    開く""}", "POST", "/session/abc/execute/sync", element_2_text_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/execute/sync", fallback_script_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim operation As TransitionOperation
    Set operation = New_TransitionOperation("OpenList", "css selector", ".menu-button", LocatorInnerText:="一覧 を 開く", Script:="fallback()", WaitConditionName:="ListReady")

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.RunTransitionOperation(operation)

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "MatchCount=2", vbBinaryCompare)
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", fallback_script_body)
End Sub

Public Sub Test_WebDriverSessionClient_RunTransitionOperationでLocatorInnerTextは文字種を区別する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim elements_body As String
    elements_body = "{""using"":""css selector"",""value"":"".menu-button""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""element-1""},{""element-6066-11e4-a52e-4f735466cecf"":""element-2""},{""element-6066-11e4-a52e-4f735466cecf"":""element-3""}]}", "POST", "/session/abc/elements", elements_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""ABC ア あ""}", "POST", "/session/abc/execute/sync", pInnerTextScriptBody("element-1"))
    Call client_double.Store.SetReturn("Execute", "{""value"":""abc ｱ あ""}", "POST", "/session/abc/execute/sync", pInnerTextScriptBody("element-2"))
    Call client_double.Store.SetReturn("Execute", "{""value"":""abc ア ア""}", "POST", "/session/abc/execute/sync", pInnerTextScriptBody("element-3"))

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim operation As TransitionOperation
    Set operation = New_TransitionOperation("OpenList", "css selector", ".menu-button", LocatorInnerText:="abc ア あ", WaitConditionName:="ListReady")

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.RunTransitionOperation(operation)

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "MatchCount=0", vbBinaryCompare)
End Sub

Public Sub Test_WebDriverSessionClient_RunTransitionOperationでLocatorInnerText対象がStaleなら再探索する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim elements_body As String
    elements_body = "{""using"":""css selector"",""value"":"".menu-button""}"

    Dim old_text_body As String
    old_text_body = pInnerTextScriptBody("element-old")

    Dim new_text_body As String
    new_text_body = pInnerTextScriptBody("element-new")

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute.Sequence", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""element-old""}]}", "POST", "/session/abc/elements", elements_body, CLng(0))
    Call client_double.Store.SetReturn("Execute.Sequence", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""element-new""}]}", "POST", "/session/abc/elements", elements_body, CLng(1))
    Call client_double.Store.SetReturn("Execute", "{""value"":""一覧を開く""}", "POST", "/session/abc/execute/sync", old_text_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""一覧を開く""}", "POST", "/session/abc/execute/sync", new_text_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""error"":""stale element reference"",""message"":""stale element reference""}}", "POST", "/session/abc/element/element-old/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/element-new/click", "{}")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim operation As TransitionOperation
    Set operation = New_TransitionOperation("OpenList", "css selector", ".menu-button", LocatorInnerText:="一覧を開く", WaitConditionName:="ListReady")

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.RunTransitionOperation(operation)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/elements", elements_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", old_text_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", new_text_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/element-old/click", "{}")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/element-new/click", "{}")
End Sub

Public Sub Test_WebDriverSessionClient_RunTransitionOperationでFrame内LocatorInnerText一致要素をClickする(ByVal Assert As UnitTestAssert)
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

    Dim elements_body As String
    elements_body = "{""using"":""css selector"",""value"":"".menu-button""}"

    Dim element_1_text_body As String
    element_1_text_body = pInnerTextScriptBody("target-1")

    Dim element_2_text_body As String
    element_2_text_body = pInnerTextScriptBody("target-2")

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", switch_top_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""frame-1""}}", "POST", "/session/abc/element", frame_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", switch_frame_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""target-1""},{""element-6066-11e4-a52e-4f735466cecf"":""target-2""}]}", "POST", "/session/abc/elements", elements_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""対象外""}", "POST", "/session/abc/execute/sync", element_1_text_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""詳細を開く""}", "POST", "/session/abc/execute/sync", element_2_text_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/target-2/click", "{}")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim operation As TransitionOperation
    Set operation = New_TransitionOperation("OpenDetail", "css selector", "iframe[name='right'] >> .menu-button", LocatorInnerText:="詳細を開く", WaitConditionName:="DetailReady")

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.RunTransitionOperation(operation)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/frame", switch_top_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", frame_find_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/frame", switch_frame_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/elements", elements_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", element_1_text_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", element_2_text_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/target-2/click", "{}")
End Sub

Public Sub Test_WebDriverSessionClient_RunTransitionOperationでLocatorInnerText確定後Click失敗ならScriptFallbackする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim elements_body As String
    elements_body = "{""using"":""css selector"",""value"":"".menu-button""}"

    Dim element_text_body As String
    element_text_body = pInnerTextScriptBody("element-1")

    Dim fallback_script_body As String
    fallback_script_body = "{""script"":""fallback()"",""args"":[]}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""element-1""}]}", "POST", "/session/abc/elements", elements_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""一覧を開く""}", "POST", "/session/abc/execute/sync", element_text_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""error"":""element click intercepted"",""message"":""blocked""}}", "POST", "/session/abc/element/element-1/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/execute/sync", fallback_script_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim operation As TransitionOperation
    Set operation = New_TransitionOperation("OpenList", "css selector", ".menu-button", LocatorInnerText:="一覧を開く", Script:="fallback()", WaitConditionName:="ListReady")

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.RunTransitionOperation(operation)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/elements", elements_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/element-1/click", "{}")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", fallback_script_body)
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
Public Sub Test_WebDriverSessionClient_RunTransitionOperationで通常Click不能ならScriptClickする(ByVal Assert As UnitTestAssert)
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

    Dim script_click_body As String
    script_click_body = "{""script"":""arguments[0].click();"",""args"":[{""element-6066-11e4-a52e-4f735466cecf"":""element-1""}]}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""element-1""}}", "POST", "/session/abc/element", find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""error"":""element not interactable"",""message"":""element not interactable""}}", "POST", "/session/abc/element/element-1/click", "{}")
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/execute/sync", script_click_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim operation As TransitionOperation
    Set operation = New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady")

    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Call session_client.RunTransitionOperation(operation)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", find_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/element-1/click", "{}")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", script_click_body)
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

Public Sub Test_WebDriverSessionClient_ReadDetailColumnValuesで抽出種別ごとの値を一括抽出する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject", ExtractType:="TextContent"))
    Call detail_defs.Add(New_DetailColumnDefinition("リンク", "iframe[name='right'] >> #link", ExtractType:="Attribute", AttributeName:="href"))
    Call detail_defs.Add(New_DetailColumnDefinition("担当者", "#owner", ExtractType:="InnerText"))

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute.AnyRequestBody", "{""value"": [""案件A"", ""https://example.test/detail"", ""山田太郎""]}", "POST", "/session/abc/execute/sync")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)
    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Dim actual_values As ArrayObject
    Set actual_values = session_client.ReadDetailColumnValues(detail_defs)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "案件A", CStr(actual_values.Item(0))
    Assert.Equals "https://example.test/detail", CStr(actual_values.Item(1))
    Assert.Equals "山田太郎", CStr(actual_values.Item(2))
    Assert.EqualsNumeric 1, pCountExecuteCalls(client_double, "POST", "/session/abc/execute/sync")
End Sub

Public Sub Test_WebDriverSessionClient_CountElementsでFrame内の一致件数を返す(ByVal Assert As UnitTestAssert)
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

    Dim list_item_find_body As String
    list_item_find_body = "{""using"":""css selector"",""value"":""#list tbody tr""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", switch_top_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""frame-1""}}", "POST", "/session/abc/element", frame_find_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", switch_frame_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""row-1""},{""element-6066-11e4-a52e-4f735466cecf"":""row-2""}]}", "POST", "/session/abc/elements", list_item_find_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)
    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Dim actual_count As Long
    actual_count = session_client.CountElements("iframe[name='right'] >> #list tbody tr")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_count
    Assert.EqualsNumeric 2, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/frame", switch_top_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element", frame_find_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/frame", switch_frame_body)
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/elements", list_item_find_body)
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

Private Function pInnerTextScriptBody(ByVal ElementId As String) As String
    pInnerTextScriptBody = "{""script"":""var value = arguments[0].innerText; return value === null || value === undefined ? '' : String(value);"",""args"":[{""element-6066-11e4-a52e-4f735466cecf"":""" & ElementId & """}]}"
End Function

Private Function pCountExecuteCalls( _
        ByVal ClientDouble As WebDriverClientTestDouble, _
        ByVal HttpMethod As String, _
        ByVal EndpointPath As String) As Long

    Dim result_count As Long
    result_count = 0

    Dim calls As ObjectList
    Set calls = ClientDouble.Store.GetCallsAll("Execute")

    Dim call_idx As Long
    For call_idx = 0 To calls.Count - 1
        Dim call_record As TestDoubleCallRecord
        Set call_record = calls.Item(call_idx)
        If CStr(call_record.GetArgument(0)) = HttpMethod And CStr(call_record.GetArgument(1)) = EndpointPath Then
            result_count = result_count + 1
        End If
    Next call_idx

    pCountExecuteCalls = result_count
End Function
