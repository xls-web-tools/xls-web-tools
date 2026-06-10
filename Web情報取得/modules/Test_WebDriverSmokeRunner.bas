Attribute VB_Name = "Test_WebDriverSmokeRunner"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! WebDriver smoke path のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_WebDriverSmokeRunner_Process起動からSessionQuitまで通す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.Headless = True
    tool_settings.BrowserProfilePath = "C:\Profile"

    Dim create_body As String
    create_body = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]}}}}"

    Dim web_driver_client As WebDriverClientTestDouble
    Set web_driver_client = New WebDriverClientTestDouble
    Call web_driver_client.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", create_body)
    Call web_driver_client.Store.SetReturn("Execute", "{""value"":null}", "DELETE", "/session/abc", "")

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(web_driver_client, tool_settings)

    Dim process As WebDriverProcessTestDouble
    Set process = New WebDriverProcessTestDouble

    Dim lifecycle As WebDriverSessionLifecycle
    Set lifecycle = New_WebDriverSessionLifecycle(process, session_client, tool_settings)

    Dim runner As WebDriverSmokeRunner
    Set runner = New_WebDriverSmokeRunner(lifecycle)

    ' --- Act ---
    Dim actual_session_id As String
    actual_session_id = runner.Run()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", actual_session_id
    Assert.EqualsNumeric 1, process.Store.GetCallCount("Start")
    Assert.EqualsNumeric 1, process.Store.GetCallCount("StopProcess")
    Assert.IsFalse process.IsRunning
    Assert.EqualsNumeric 1, web_driver_client.Store.GetCallCount("Execute", "POST", "/session", create_body)
    Assert.EqualsNumeric 1, web_driver_client.Store.GetCallCount("Execute", "DELETE", "/session/abc", "")
End Sub
