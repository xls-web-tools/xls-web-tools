Attribute VB_Name = "Test_DownloadedFileSaver"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! ダウンロード済みファイル保存のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_DownloadedFileSaver_通常ファイルを対象IDフォルダーへ末尾3桁接頭辞付きで保存する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.DownloadRootPath = "D:\Root"

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Set FsSrv = fs_stub
    Call fs_stub.Store.SetReturn("CreateDirectory", True, "D:\Root\ABC123", False, True)

    Dim saver As DownloadedFileSaver
    Set saver = New_DownloadedFileSaver(tool_settings)

    ' --- Act ---
    Dim actual_status As String
    actual_status = saver.SaveDownloadedFile("C:\Temp\report.pdf", "ABC123")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_DOWNLOADED, actual_status
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("CreateDirectory", "D:\Root\ABC123", False, True)
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("MoveFile", "C:\Temp\report.pdf", "D:\Root\ABC123\123_report.pdf", False)
End Sub
Public Sub Test_DownloadedFileSaver_出力先ファイルが既にある場合は移動せずSKIPPEDを返す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.DownloadRootPath = "D:\Root"

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Set FsSrv = fs_stub
    Call fs_stub.Store.SetReturn("PathExists", True, "D:\Root\ABC123\123_report.pdf")

    Dim saver As DownloadedFileSaver
    Set saver = New_DownloadedFileSaver(tool_settings)

    ' --- Act ---
    Dim actual_status As String
    actual_status = saver.SaveDownloadedFile("C:\Temp\report.pdf", "ABC123")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_SKIPPED, actual_status
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCountAll("CreateDirectory")
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCountAll("MoveFile")
End Sub
Public Sub Test_DownloadedFileSaver_ZIP直下が1フォルダーなら中身を対象IDフォルダー直下へ出力する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.DownloadRootPath = "D:\Root"

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Set FsSrv = fs_stub
    Call fs_stub.Store.SetReturn("CreateTemporaryDirectory", "C:\Extract", "xls-web-tools_zip_")
    Call fs_stub.Store.SetReturn("CreateDirectory", True, "D:\Root\ABC123", False, True)

    Dim root_files() As String
    root_files = EmptyStringArray()
    Call fs_stub.Store.SetReturn("GetFileList", root_files, "C:\Extract", "", "", True)

    Dim root_dirs(0 To 0) As String
    root_dirs(0) = "C:\Extract\payload"
    Call fs_stub.Store.SetReturn("GetDirectoryList", root_dirs, "C:\Extract", "", "", True)

    Dim payload_files(0 To 0) As String
    payload_files(0) = "C:\Extract\payload\report.csv"
    Call fs_stub.Store.SetReturn("GetFileList", payload_files, "C:\Extract\payload", "", "", True)

    Dim payload_dirs(0 To 0) As String
    payload_dirs(0) = "C:\Extract\payload\docs"
    Call fs_stub.Store.SetReturn("GetDirectoryList", payload_dirs, "C:\Extract\payload", "", "", True)

    Dim zip_stub As ZipExtractorTestDouble
    Set zip_stub = New ZipExtractorTestDouble

    Dim saver As DownloadedFileSaver
    Set saver = New_DownloadedFileSaver(tool_settings, zip_stub)

    ' --- Act ---
    Dim actual_status As String
    actual_status = saver.SaveDownloadedFile("C:\Temp\bundle.zip", "ABC123")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_DOWNLOADED, actual_status
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("CreateTemporaryDirectory", "xls-web-tools_zip_")
    Assert.EqualsNumeric 1, zip_stub.Store.GetCallCount("ExtractZip", "C:\Temp\bundle.zip", "C:\Extract")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("MoveFile", "C:\Extract\payload\report.csv", "D:\Root\ABC123\123_report.csv", False)
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("MoveDirectory", "C:\Extract\payload\docs", "D:\Root\ABC123\docs", False)
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("MoveFile", "C:\Temp\bundle.zip", "D:\Root\ABC123\123_bundle.zip", False)
End Sub
Public Sub Test_DownloadedFileSaver_対象IDがWindowsフォルダー名として不正ならERRORを返す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.DownloadRootPath = "D:\Root"

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Set FsSrv = fs_stub

    Dim saver As DownloadedFileSaver
    Set saver = New_DownloadedFileSaver(tool_settings)

    ' --- Act ---
    Dim actual_status As String
    actual_status = saver.SaveDownloadedFile("C:\Temp\report.pdf", "ABC:123")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_ERROR, actual_status
    Assert.IsTrue 0 < InStr(1, saver.LastErrorDescription, "対象ID", vbTextCompare)
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCountAll("MoveFile")
End Sub
