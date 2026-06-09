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

Public Sub Test_WebDriverSessionClient_DownloadLinkedFile_リンク0件ならNO_FILEを返す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.DownloadEnabled = True
    tool_settings.DownloadLinkSelector = "#download"

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Call fs_stub.Store.SetReturn("CreateTemporaryDirectory", "C:\Temp\xls-web-tools_tmp123.tmp", "xls-web-tools_")
    Set FsSrv = fs_stub

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""],""prefs"":{""download.default_directory"":""C:\\Temp\\xls-web-tools_tmp123.tmp"",""download.prompt_for_download"":false,""download.directory_upgrade"":true}}}}}"

    Dim find_elements_body As String
    find_elements_body = "{""using"":""css selector"",""value"":""#download""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", "{""id"":null}")
    Call client_double.Store.SetReturn("Execute", "{""value"":[]}", "POST", "/session/abc/elements", find_elements_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)
    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Dim actual_file_path As String
    Dim actual_status As String
    actual_status = session_client.DownloadLinkedFile(actual_file_path)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_NO_FILE, actual_status
    Assert.Equals "", actual_file_path
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/elements", find_elements_body)
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCountAll("GetFileList")
End Sub
Public Sub Test_WebDriverSessionClient_DownloadLinkedFile_リンク1件ならクリックして完了ファイルを返す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.DownloadEnabled = True
    tool_settings.DownloadLinkSelector = "#download"

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Call fs_stub.Store.SetReturn("CreateTemporaryDirectory", "C:\Temp\xls-web-tools_tmp123.tmp", "xls-web-tools_")

    Dim crdownload_files() As String
    crdownload_files = EmptyStringArray()
    Call fs_stub.Store.SetReturn("GetFileList", crdownload_files, "C:\Temp\xls-web-tools_tmp123.tmp", "\.crdownload$", "", True)

    Dim completed_files(0 To 0) As String
    completed_files(0) = "C:\Temp\xls-web-tools_tmp123.tmp\report.pdf"
    Call fs_stub.Store.SetReturn("GetFileList", completed_files, "C:\Temp\xls-web-tools_tmp123.tmp", "", "\.crdownload$", True)
    Set FsSrv = fs_stub

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""],""prefs"":{""download.default_directory"":""C:\\Temp\\xls-web-tools_tmp123.tmp"",""download.prompt_for_download"":false,""download.directory_upgrade"":true}}}}}"

    Dim selector_body As String
    selector_body = "{""using"":""css selector"",""value"":""#download""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", "{""id"":null}")
    Call client_double.Store.SetReturn("Execute", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""download-1""}]}", "POST", "/session/abc/elements", selector_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""download-1""}}", "POST", "/session/abc/element", selector_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/element/download-1/click", "{}")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)
    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Dim actual_file_path As String
    Dim actual_status As String
    actual_status = session_client.DownloadLinkedFile(actual_file_path)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_DOWNLOADED, actual_status
    Assert.Equals "C:\Temp\xls-web-tools_tmp123.tmp\report.pdf", actual_file_path
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/download-1/click", "{}")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetFileList", "C:\Temp\xls-web-tools_tmp123.tmp", "\.crdownload$", "", True)
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetFileList", "C:\Temp\xls-web-tools_tmp123.tmp", "", "\.crdownload$", True)
End Sub

Public Sub Test_WebDriverSessionClient_DownloadLinkedFile_A要素HrefならUrl遷移でダウンロードする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    tool_settings.DownloadEnabled = True
    tool_settings.DownloadLinkSelector = "#download"

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Call fs_stub.Store.SetReturn("CreateTemporaryDirectory", "C:\Temp\xls-web-tools_tmp123.tmp", "xls-web-tools_")

    Dim crdownload_files() As String
    crdownload_files = EmptyStringArray()
    Call fs_stub.Store.SetReturn("GetFileList", crdownload_files, "C:\Temp\xls-web-tools_tmp123.tmp", "\.crdownload$", "", True)

    Dim completed_files(0 To 0) As String
    completed_files(0) = "C:\Temp\xls-web-tools_tmp123.tmp\report.pdf"
    Call fs_stub.Store.SetReturn("GetFileList", completed_files, "C:\Temp\xls-web-tools_tmp123.tmp", "", "\.crdownload$", True)
    Set FsSrv = fs_stub

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""],""prefs"":{""download.default_directory"":""C:\\Temp\\xls-web-tools_tmp123.tmp"",""download.prompt_for_download"":false,""download.directory_upgrade"":true}}}}}"

    Dim selector_body As String
    selector_body = "{""using"":""css selector"",""value"":""#download""}"

    Dim href_script_body As String
    href_script_body = "{""script"":""return arguments[0].href || arguments[0].getAttribute('href') || '';"",""args"": [{""element-6066-11e4-a52e-4f735466cecf"":""download-1""}]}"

    Dim navigate_body As String
    navigate_body = "{""url"":""https://example.test/files/report.pdf""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/frame", "{""id"":null}")
    Call client_double.Store.SetReturn("Execute", "{""value"":[{""element-6066-11e4-a52e-4f735466cecf"":""download-1""}]}", "POST", "/session/abc/elements", selector_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""download-1""}}", "POST", "/session/abc/element", selector_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":""https://example.test/files/report.pdf""}", "POST", "/session/abc/execute/sync", href_script_body)
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/abc/url", navigate_body)

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)
    Call session_client.CreateSession
    Err.Clear

    ' --- Act ---
    Dim actual_file_path As String
    Dim actual_status As String
    actual_status = session_client.DownloadLinkedFile(actual_file_path)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_DOWNLOADED, actual_status
    Assert.Equals "C:\Temp\xls-web-tools_tmp123.tmp\report.pdf", actual_file_path
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/execute/sync", href_script_body)
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "GET", "/session/abc/element/download-1/attribute/href", "")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/url", navigate_body)
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "POST", "/session/abc/element/download-1/click", "{}")
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
