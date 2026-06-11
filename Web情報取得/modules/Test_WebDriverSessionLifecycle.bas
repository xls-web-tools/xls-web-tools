Attribute VB_Name = "Test_WebDriverSessionLifecycle"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! WebDriver session lifecycle のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_WebDriverSessionLifecycle_StartSessionからFinishSessionまで通す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    Call pUseProfileDirectory("C:\Profile", True)

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", pCreateSessionBody(True))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = lifecycle.StartSession()
    Call lifecycle.FinishSession

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.Equals "", lifecycle.SessionId
    Assert.EqualsNumeric 1, process.Store.GetCallCount("Start")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.IsFalse process.IsRunning
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session", pCreateSessionBody(True))
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Public Sub Test_WebDriverSessionLifecycle_StartSession前にBrowserProfilePathの存在を確認する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = pUseProfileDirectory("C:\Profile", True)

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", pCreateSessionBody(True))

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, tool_settings)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = lifecycle.StartSession()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("IsDirectory", "C:\Profile")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("Start")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session", pCreateSessionBody(True))
End Sub
Public Sub Test_WebDriverSessionLifecycle_BrowserProfilePath不存在で作成続行なら再帰作成してSessionを開始する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Missing\Profile"

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = pUseProfileDirectory("C:\Missing\Profile", False)
    Call fs_stub.Store.SetReturn("CreateDirectory", True, "C:\Missing\Profile", False, True)

    Dim prompt As BrowserProfilePromptTestDouble
    Set prompt = New BrowserProfilePromptTestDouble
    Call prompt.Store.SetReturn("ConfirmCreateDirectory", True, "C:\Missing\Profile")

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", pCreateSessionBodyWithProfile("C:\Missing\Profile", True))

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, tool_settings)
    Set lifecycle.BrowserProfilePrompt = prompt

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = lifecycle.StartSession()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("IsDirectory", "C:\Missing\Profile")
    Assert.EqualsNumeric 1, prompt.Store.GetCallCount("ConfirmCreateDirectory", "C:\Missing\Profile")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("CreateDirectory", "C:\Missing\Profile", False, True)
    Assert.EqualsNumeric 1, process.Store.GetCallCount("Start")
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session", pCreateSessionBodyWithProfile("C:\Missing\Profile", True))
End Sub
Public Sub Test_WebDriverSessionLifecycle_BrowserProfilePath不存在で作成中止ならSessionを開始しない(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Missing\Profile"

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = pUseProfileDirectory("C:\Missing\Profile", False)

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

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = lifecycle.StartSession()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", actual_session_id
    Assert.Equals "", lifecycle.SessionId
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("IsDirectory", "C:\Missing\Profile")
    Assert.EqualsNumeric 1, prompt.Store.GetCallCount("ConfirmCreateDirectory", "C:\Missing\Profile")
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("CreateDirectory", "C:\Missing\Profile", False, True)
    Assert.EqualsNumeric 0, process.Store.GetCallCount("Start")
    Assert.EqualsNumeric 0, client_double.Store.GetCallCountAll("Execute")
End Sub
Public Sub Test_WebDriverSessionLifecycle_CleanupAfterErrorはHeadlessならSessionとProcessを閉じる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    Call pUseProfileDirectory("C:\Profile", True)

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", pCreateSessionBody(True))
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, tool_settings, KeepVisibleBrowserOnError:=True)
    Call lifecycle.StartSession
    Err.Clear

    ' --- Act ---
    Dim err_desc As String
    err_desc = "失敗しました。"
    Call lifecycle.CleanupAfterError(err_desc)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "失敗しました。", err_desc
    Assert.Equals "", lifecycle.SessionId
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.IsFalse process.IsRunning
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Public Sub Test_WebDriverSessionLifecycle_可視ブラウザ診断中のエラーではSessionとProcessを残す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = False
    tool_settings.BrowserProfilePath = "C:\Profile"
    Call pUseProfileDirectory("C:\Profile", True)

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", pCreateSessionBody(False))

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, tool_settings, KeepVisibleBrowserOnError:=True)
    Call lifecycle.StartSession
    Err.Clear

    ' --- Act ---
    Dim err_desc As String
    err_desc = "失敗しました。"
    Call lifecycle.CleanupAfterError(err_desc)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue 0 < InStr(1, err_desc, "ブラウザを残しました", vbTextCompare)
    Assert.Equals "abc", lifecycle.SessionId
    Assert.EqualsNumeric 1, process.Store.GetCallCount("Start")
    Assert.EqualsNumeric 0, process.Store.GetCallCount("StopProcess")
    Assert.IsTrue process.IsRunning
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Public Sub Test_WebDriverSessionLifecycle_StartSession失敗後のCleanupはProcessを閉じる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"
    Call pUseProfileDirectory("C:\Profile", True)

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetError("Execute", vbObjectError + 1, "Class WebDriverClient", "session 作成失敗", "POST", "/session", pCreateSessionBody(True))

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(client_double, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, tool_settings)

    ' --- Act ---
    Call lifecycle.StartSession

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Err.Clear

    Dim err_desc As String
    err_desc = "失敗しました。"
    Call lifecycle.CleanupAfterError(err_desc)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", lifecycle.SessionId
    Assert.EqualsNumeric 1, process.Store.GetCallCount("Start")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.IsFalse process.IsRunning
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub

Private Function pUseProfileDirectory(ByVal DirectoryPath As String, ByVal Exists As Boolean) As FileSystemServiceTestDouble
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Set FsSrv = fs_stub
    Call fs_stub.Store.SetReturn("IsDirectory", Exists, DirectoryPath)

    Set pUseProfileDirectory = fs_stub
End Function
Private Function pCreateSessionBodyWithProfile(ByVal BrowserProfilePath As String, ByVal Headless As Boolean) As String
    Dim profile_path As String
    profile_path = Replace(BrowserProfilePath, "\", "\\")

    Dim result_value As String
    result_value = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=" & profile_path & """"
    If Headless Then result_value = result_value & ",""--headless=new"""
    result_value = result_value & "]}}}}"

    pCreateSessionBodyWithProfile = result_value
End Function
Private Function pCreateSessionBody(ByVal Headless As Boolean) As String
    Dim result_value As String
    If Headless Then
        result_value = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"
    Else
        result_value = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile""]}}}}"
    End If

    pCreateSessionBody = result_value
End Function
