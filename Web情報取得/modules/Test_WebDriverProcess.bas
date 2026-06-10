Attribute VB_Name = "Test_WebDriverProcess"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! WebDriver process 管理のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_WebDriverProcess_WebDriver実行ファイルがない場合は配置不足エラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.WebDriverPath = "C:\Missing\msedgedriver.exe"
    tool_settings.WebDriverPort = 9515

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Set FsSrv = fs_stub

    Dim port_probe As WebDriverPortProbeTestDouble
    Set port_probe = New WebDriverPortProbeTestDouble

    Dim process As WebDriverProcess
    Set process = New_WebDriverProcess(port_probe)

    ' --- Act ---
    Call process.Start(tool_settings)

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "配置", vbTextCompare)
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("IsFile", "C:\Missing\msedgedriver.exe")
    Assert.EqualsNumeric 0, port_probe.Store.GetCallCount("IsPortInUse", CLng(9515))
    Assert.IsFalse process.IsRunning
End Sub

Public Sub Test_WebDriverProcess_同じPortが使用中の場合は他プロセスを終了せず中断する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.WebDriverPath = "C:\Driver\msedgedriver.exe"
    tool_settings.WebDriverPort = 9515

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Set FsSrv = fs_stub
    Call fs_stub.Store.SetReturn("IsFile", True, "C:\Driver\msedgedriver.exe")

    Dim port_probe As WebDriverPortProbeTestDouble
    Set port_probe = New WebDriverPortProbeTestDouble
    Call port_probe.Store.SetReturn("IsPortInUse", True, CLng(9515))

    Dim process As WebDriverProcess
    Set process = New_WebDriverProcess(port_probe)

    ' --- Act ---
    Call process.Start(tool_settings)

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "使用中", vbTextCompare)
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("IsFile", "C:\Driver\msedgedriver.exe")
    Assert.EqualsNumeric 1, port_probe.Store.GetCallCount("IsPortInUse", CLng(9515))
    Assert.IsFalse process.IsRunning
End Sub
