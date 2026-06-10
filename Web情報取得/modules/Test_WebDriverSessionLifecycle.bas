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

Public Sub Test_WebDriverSessionLifecycle_CleanupAfterErrorはHeadlessならSessionとProcessを閉じる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

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

Private Function pCreateSessionBody(ByVal Headless As Boolean) As String
    Dim result_value As String
    If Headless Then
        result_value = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"
    Else
        result_value = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile""]}}}}"
    End If

    pCreateSessionBody = result_value
End Function
