Attribute VB_Name = "Test_DetailFileDownloader"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 詳細ページのダウンロード処理のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Private Const C_DOWNLOAD_DIR As String = "C:\Temp\xls-web-tools_tmp123.tmp"
Private Const C_TARGET_ID As String = "ABC123"
Private Const C_DOWNLOAD_SELECTOR As String = "#download"
Private Const C_SESSION_ID As String = "abc"
Private Const C_ELEMENT_ID As String = "download-1"
Private Const C_DOWNLOADED_FILE As String = "C:\Temp\xls-web-tools_tmp123.tmp\report.pdf"

Public Sub Test_DetailFileDownloader_リンク0件ならNO_FILEを返す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = pNewSettings()

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = pNewFileSystem()

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call pRegisterCreateSession(client_double)
    Call pRegisterNoDownloadElement(client_double)

    Dim downloader As DetailFileDownloader
    Set downloader = pCreateDownloader(client_double, tool_settings, fs_stub)
    Err.Clear

    ' --- Act ---
    Dim actual_result As DetailFileDownloadResult
    Set actual_result = downloader.DownloadForDetail(C_TARGET_ID)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_NO_FILE, actual_result.DownloadStatus
    Assert.Equals "ダウンロードリンクがありません。", actual_result.ErrorDescription
    Assert.IsFalse actual_result.RequiresDetailError
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/" & C_SESSION_ID & "/elements", pCssFindBody(C_DOWNLOAD_SELECTOR))
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCountAll("GetFileList")
End Sub

Public Sub Test_DetailFileDownloader_リンク1件ならクリックして保存する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = pNewSettings()

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = pNewFileSystem()
    Call pRegisterEmptyStaging(fs_stub)
    Call pRegisterDownloadedPdf(fs_stub, 1024#)
    Call fs_stub.Store.SetReturn("CreateDirectory", True, "D:\Root\ABC123", False, True)

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call pRegisterCreateSession(client_double)
    Call pRegisterDownloadElement(client_double)
    Call pRegisterClick(client_double)

    Dim downloader As DetailFileDownloader
    Set downloader = pCreateDownloader(client_double, tool_settings, fs_stub)
    Err.Clear

    ' --- Act ---
    Dim actual_result As DetailFileDownloadResult
    Set actual_result = downloader.DownloadForDetail(C_TARGET_ID)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_DOWNLOADED, actual_result.DownloadStatus
    Assert.Equals "", actual_result.ErrorDescription
    Assert.IsFalse actual_result.RequiresDetailError
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/" & C_SESSION_ID & "/element/" & C_ELEMENT_ID & "/click", "{}")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("MoveFile", C_DOWNLOADED_FILE, "D:\Root\ABC123\123_report.pdf", False)
End Sub

Public Sub Test_DetailFileDownloader_tmpファイルは完了扱いしない(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = pNewSettings(TimeoutSeconds:=0)

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = pNewFileSystem()
    Call pRegisterEmptyStaging(fs_stub)

    Dim incomplete_files(0 To 0) As String
    incomplete_files(0) = "C:\Temp\xls-web-tools_tmp123.tmp\download.tmp"
    Call fs_stub.Store.SetReturn("GetFileList", incomplete_files, C_DOWNLOAD_DIR, "\.(crdownload|tmp)$", "", True)

    Dim completed_files() As String
    completed_files = EmptyStringArray()
    Call fs_stub.Store.SetReturn("GetFileList", completed_files, C_DOWNLOAD_DIR, "", "\.(crdownload|tmp)$", True)

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call pRegisterCreateSession(client_double)
    Call pRegisterDownloadElement(client_double)
    Call pRegisterClick(client_double)

    Dim downloader As DetailFileDownloader
    Set downloader = pCreateDownloader(client_double, tool_settings, fs_stub)
    Err.Clear

    ' --- Act ---
    Dim actual_result As DetailFileDownloadResult
    Set actual_result = downloader.DownloadForDetail(C_TARGET_ID)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_ERROR, actual_result.DownloadStatus
    Assert.IsTrue 0 < InStr(1, actual_result.ErrorDescription, "タイムアウト", vbTextCompare)
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCountAll("GetFileSize")
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCountAll("MoveFile")
End Sub

Public Sub Test_DetailFileDownloader_ZeroByteFileIsNotCompleted(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = pNewSettings(TimeoutSeconds:=0)

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = pNewFileSystem()
    Call pRegisterEmptyStaging(fs_stub)
    Call pRegisterDownloadedPdf(fs_stub, 0#)

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call pRegisterCreateSession(client_double)
    Call pRegisterDownloadElement(client_double)
    Call pRegisterClick(client_double)

    Dim downloader As DetailFileDownloader
    Set downloader = pCreateDownloader(client_double, tool_settings, fs_stub)
    Err.Clear

    ' --- Act ---
    Dim actual_result As DetailFileDownloadResult
    Set actual_result = downloader.DownloadForDetail(C_TARGET_ID)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_ERROR, actual_result.DownloadStatus
    Assert.IsTrue 0 < InStr(1, actual_result.ErrorDescription, "タイムアウト", vbTextCompare)
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetFileSize", C_DOWNLOADED_FILE)
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCountAll("MoveFile")
End Sub

Public Sub Test_DetailFileDownloader_開始前にStagingを空にする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = pNewSettings()

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = pNewFileSystem()

    Dim staging_files(0 To 1) As String
    staging_files(0) = "C:\Temp\xls-web-tools_tmp123.tmp\stale.pdf"
    staging_files(1) = "C:\Temp\xls-web-tools_tmp123.tmp\stale.crdownload"
    Call fs_stub.Store.SetReturn("GetFileList", staging_files, C_DOWNLOAD_DIR, "", "", True)
    Call fs_stub.Store.SetReturn("RemoveFile", True, "C:\Temp\xls-web-tools_tmp123.tmp\stale.pdf", False)
    Call fs_stub.Store.SetReturn("RemoveFile", True, "C:\Temp\xls-web-tools_tmp123.tmp\stale.crdownload", False)
    Call pRegisterDownloadedPdf(fs_stub, 1024#)
    Call fs_stub.Store.SetReturn("CreateDirectory", True, "D:\Root\ABC123", False, True)

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call pRegisterCreateSession(client_double)
    Call pRegisterDownloadElement(client_double)
    Call pRegisterClick(client_double)

    Dim downloader As DetailFileDownloader
    Set downloader = pCreateDownloader(client_double, tool_settings, fs_stub)
    Err.Clear

    ' --- Act ---
    Dim actual_result As DetailFileDownloadResult
    Set actual_result = downloader.DownloadForDetail(C_TARGET_ID)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_DOWNLOADED, actual_result.DownloadStatus
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("RemoveFile", "C:\Temp\xls-web-tools_tmp123.tmp\stale.pdf", False)
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("RemoveFile", "C:\Temp\xls-web-tools_tmp123.tmp\stale.crdownload", False)
End Sub

Public Sub Test_DetailFileDownloader_A要素HrefならUrl遷移で保存する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = pNewSettings()

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = pNewFileSystem()
    Call pRegisterEmptyStaging(fs_stub)
    Call pRegisterDownloadedPdf(fs_stub, 1024#)
    Call fs_stub.Store.SetReturn("CreateDirectory", True, "D:\Root\ABC123", False, True)

    Dim navigate_body As String
    navigate_body = "{""url"":""https://example.test/files/report.pdf""}"

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call pRegisterCreateSession(client_double)
    Call pRegisterDownloadElement(client_double)
    Call client_double.Store.SetReturn("Execute", "{""value"":""https://example.test/files/report.pdf""}", "POST", "/session/" & C_SESSION_ID & "/execute/sync", pHrefScriptBody())
    Call client_double.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/" & C_SESSION_ID & "/url", navigate_body)

    Dim downloader As DetailFileDownloader
    Set downloader = pCreateDownloader(client_double, tool_settings, fs_stub)
    Err.Clear

    ' --- Act ---
    Dim actual_result As DetailFileDownloadResult
    Set actual_result = downloader.DownloadForDetail(C_TARGET_ID)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_DOWNLOADED, actual_result.DownloadStatus
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/" & C_SESSION_ID & "/execute/sync", pHrefScriptBody())
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session/" & C_SESSION_ID & "/url", navigate_body)
    Assert.EqualsNumeric 0, client_double.Store.GetCallCount("Execute", "POST", "/session/" & C_SESSION_ID & "/element/" & C_ELEMENT_ID & "/click", "{}")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("MoveFile", C_DOWNLOADED_FILE, "D:\Root\ABC123\123_report.pdf", False)
End Sub

Public Sub Test_DetailFileDownloader_DownloadRequiredならNO_FILEを詳細エラー扱いにする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = pNewSettings(DownloadRequired:=True)

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = pNewFileSystem()

    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call pRegisterCreateSession(client_double)
    Call pRegisterNoDownloadElement(client_double)

    Dim downloader As DetailFileDownloader
    Set downloader = pCreateDownloader(client_double, tool_settings, fs_stub)
    Err.Clear

    ' --- Act ---
    Dim actual_result As DetailFileDownloadResult
    Set actual_result = downloader.DownloadForDetail(C_TARGET_ID)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals G_WEB_DOWNLOAD_STATUS_NO_FILE, actual_result.DownloadStatus
    Assert.IsTrue actual_result.RequiresDetailError
End Sub

Private Function pNewSettings( _
        Optional ByVal DownloadRequired As Boolean = False, _
        Optional ByVal TimeoutSeconds As Long = 30) As ToolSettingsTestDouble

    Dim result_value As ToolSettingsTestDouble
    Set result_value = New ToolSettingsTestDouble
    result_value.Headless = True
    result_value.BrowserProfilePath = "C:\Profile"
    result_value.DownloadEnabled = True
    result_value.DownloadRequired = DownloadRequired
    result_value.DownloadRootPath = "D:\Root"
    result_value.DownloadLinkSelector = C_DOWNLOAD_SELECTOR
    result_value.TimeoutSeconds = TimeoutSeconds

    Set pNewSettings = result_value
End Function

Private Function pNewFileSystem() As FileSystemServiceTestDouble
    Dim result_value As FileSystemServiceTestDouble
    Set result_value = New FileSystemServiceTestDouble
    Call result_value.Store.SetReturn("CreateTemporaryDirectory", C_DOWNLOAD_DIR, "xls-web-tools_")

    Set pNewFileSystem = result_value
End Function

Private Function pCreateDownloader( _
        ByVal ClientDouble As WebDriverClientTestDouble, _
        ByVal Settings As ToolSettingsTestDouble, _
        ByVal FileSystem As FileSystemServiceTestDouble) As DetailFileDownloader

    Dim session_client As WebDriverSessionClient
    Set session_client = New_WebDriverSessionClient(ClientDouble, Settings, FileSystem)
    Call session_client.CreateSession

    Set pCreateDownloader = New_DetailFileDownloader(session_client, Settings, FileSystem)
End Function

Private Sub pRegisterCreateSession(ByVal ClientDouble As WebDriverClientTestDouble)
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", pCreateSessionBody())
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/" & C_SESSION_ID & "/frame", "{""id"":null}")
End Sub

Private Sub pRegisterNoDownloadElement(ByVal ClientDouble As WebDriverClientTestDouble)
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":[]}", "POST", "/session/" & C_SESSION_ID & "/elements", pCssFindBody(C_DOWNLOAD_SELECTOR))
End Sub

Private Sub pRegisterDownloadElement(ByVal ClientDouble As WebDriverClientTestDouble)
    Call ClientDouble.Store.SetReturn("Execute", "{""value"": [{""element-6066-11e4-a52e-4f735466cecf"":""download-1""}]}", "POST", "/session/" & C_SESSION_ID & "/elements", pCssFindBody(C_DOWNLOAD_SELECTOR))
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":{""element-6066-11e4-a52e-4f735466cecf"":""download-1""}}", "POST", "/session/" & C_SESSION_ID & "/element", pCssFindBody(C_DOWNLOAD_SELECTOR))
End Sub

Private Sub pRegisterClick(ByVal ClientDouble As WebDriverClientTestDouble)
    Call ClientDouble.Store.SetReturn("Execute", "{""value"":null}", "POST", "/session/" & C_SESSION_ID & "/element/" & C_ELEMENT_ID & "/click", "{}")
End Sub

Private Sub pRegisterEmptyStaging(ByVal FileSystem As FileSystemServiceTestDouble)
    Dim staging_files() As String
    staging_files = EmptyStringArray()
    Call FileSystem.Store.SetReturn("GetFileList", staging_files, C_DOWNLOAD_DIR, "", "", True)
End Sub

Private Sub pRegisterDownloadedPdf(ByVal FileSystem As FileSystemServiceTestDouble, ByVal FileSize As Double)
    Dim incomplete_files() As String
    incomplete_files = EmptyStringArray()
    Call FileSystem.Store.SetReturn("GetFileList", incomplete_files, C_DOWNLOAD_DIR, "\.(crdownload|tmp)$", "", True)

    Dim completed_files(0 To 0) As String
    completed_files(0) = C_DOWNLOADED_FILE
    Call FileSystem.Store.SetReturn("GetFileList", completed_files, C_DOWNLOAD_DIR, "", "\.(crdownload|tmp)$", True)
    Call FileSystem.Store.SetReturn("IsFile", True, C_DOWNLOADED_FILE)
    Call FileSystem.Store.SetReturn("GetFileSize", FileSize, C_DOWNLOADED_FILE)
End Sub

Private Function pCreateSessionBody() As String
    pCreateSessionBody = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge""," _
            & """ms:edgeOptions"":{""args"":[""--user-data-dir=C:\\Profile"",""--headless=new""]," _
            & """prefs"":{""download.default_directory"":""C:\\Temp\\xls-web-tools_tmp123.tmp""," _
            & """download.prompt_for_download"":false,""download.directory_upgrade"":true}}}}}"
End Function

Private Function pCssFindBody(ByVal Selector As String) As String
    pCssFindBody = "{""using"":""css selector"",""value"":""" & Selector & """}"
End Function

Private Function pHrefScriptBody() As String
    pHrefScriptBody = "{""script"":""return arguments[0].href || arguments[0].getAttribute('href') || '';""," _
            & """args"": [{""element-6066-11e4-a52e-4f735466cecf"":""" & C_ELEMENT_ID & """}]}"
End Function
