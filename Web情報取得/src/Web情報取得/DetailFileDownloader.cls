VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailFileDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 詳細ページ上のダウンロード対象ファイルを取得して保存します。
'!
' #############################################################################

Private Const C_INCOMPLETE_DOWNLOAD_FILE_REGEX As String = "\.(crdownload|tmp)$"

Private pSessionClient As WebDriverSessionClient
Private pSettings As IToolSettings
Private pFileSystem As IFileSystemService
Private pFileSaver As DownloadedFileSaver
Private pIsInitialized As Boolean

'* 詳細ページのダウンロード処理を初期化します。
'*
'* @param SessionClient WebDriver session client。
'* @param Settings Web情報取得の設定。
'* @param FileSaver ダウンロード済みファイル保存処理。
'*
'* @details
'* browser 上のリンク操作と最終保存処理を外部から受け取り、ファイル操作は FsSrv を使用します。
Public Sub Initialize( _
        ByVal SessionClient As WebDriverSessionClient, _
        ByVal Settings As IToolSettings, _
        ByVal FileSaver As DownloadedFileSaver)

    If SessionClient Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailFileDownloader.Initialize", Description:="SessionClient は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailFileDownloader.Initialize", Description:="Settings は必須です。"
    End If
    If FileSaver Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailFileDownloader.Initialize", Description:="FileSaver は必須です。"
    End If

    Set pSessionClient = SessionClient
    Set pSettings = Settings
    Set pFileSystem = FsSrv
    Set pFileSaver = FileSaver
    pIsInitialized = True
End Sub

'* 詳細ページ上のダウンロード対象ファイルを取得して保存します。
'*
'* @param TargetId 保存先ファイル名の決定に使う対象ID。
'* @return 詳細ページのダウンロード処理結果。
'*
'* @details
'* ダウンロードリンク数を判定し、1 件だけ存在する場合は browser-resolved href またはクリックで staging 領域へ取得して保存処理へ渡します。
Public Function DownloadForDetail(ByVal TargetId As String) As DetailFileDownloadResult
    Call pCheckInit

    If Not pSettings.DownloadEnabled Then
        Set DownloadForDetail = New_DetailFileDownloadResult("", "", False)
        Exit Function
    End If

    Dim error_description As String
    error_description = ""

    Dim result_status As String
    result_status = pDownloadCurrentDetailFile(TargetId, error_description)

    Set DownloadForDetail = New_DetailFileDownloadResult( _
            result_status, _
            error_description, _
            pDownloadStatusRequiresDetailError(result_status))
End Function

Private Function pDownloadCurrentDetailFile( _
        ByVal TargetId As String, _
        ByRef ErrorDescription As String) As String

    ErrorDescription = ""
    On Error GoTo ON_ERROR

    Dim downloaded_file_path As String
    Dim result_status As String
    result_status = pDownloadToStaging(downloaded_file_path, ErrorDescription)

    If result_status = G_WEB_DOWNLOAD_STATUS_DOWNLOADED Then
        result_status = pFileSaver.SaveDownloadedFile(downloaded_file_path, TargetId)
        If result_status = G_WEB_DOWNLOAD_STATUS_ERROR Then ErrorDescription = pFileSaver.LastErrorDescription
    End If

    On Error GoTo 0
    pDownloadCurrentDetailFile = result_status
    Exit Function

ON_ERROR:
    ErrorDescription = Err.Description
    Err.Clear
    On Error GoTo 0

    pDownloadCurrentDetailFile = G_WEB_DOWNLOAD_STATUS_ERROR
End Function

Private Function pDownloadStatusRequiresDetailError(ByVal DownloadStatus As String) As Boolean
    pDownloadStatusRequiresDetailError = (pSettings.DownloadRequired _
            And (DownloadStatus = G_WEB_DOWNLOAD_STATUS_ERROR Or DownloadStatus = G_WEB_DOWNLOAD_STATUS_NO_FILE))
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailFileDownloader.pCheckInit", Description:="初期化されていません。"
    End If
End Sub

Private Function pDownloadToStaging( _
        ByRef DownloadedFilePath As String, _
        ByRef ErrorDescription As String) As String

    DownloadedFilePath = ""
    ErrorDescription = ""

    On Error GoTo ON_ERROR

    Dim link_count As Long
    link_count = pSessionClient.CountElements(pSettings.DownloadLinkSelector)

    Dim result_status As String
    Select Case link_count
     Case 0
        ErrorDescription = "ダウンロードリンクがありません。"
        result_status = G_WEB_DOWNLOAD_STATUS_NO_FILE
     Case 1
        Call pClearDownloadDirectory
        If Not pSessionClient.TryNavigateDownloadLink("css selector", pSettings.DownloadLinkSelector) Then
            Call pSessionClient.ClickElement("css selector", pSettings.DownloadLinkSelector)
        End If
        result_status = pWaitForDownloadedFile(DownloadedFilePath, ErrorDescription)
     Case Else
        ErrorDescription = "ダウンロードリンクが複数あります。(" & CStr(link_count) & ")"
        result_status = G_WEB_DOWNLOAD_STATUS_ERROR
    End Select

    On Error GoTo 0
    pDownloadToStaging = result_status
    Exit Function

ON_ERROR:
    ErrorDescription = Err.Description
    Err.Clear
    On Error GoTo 0

    pDownloadToStaging = G_WEB_DOWNLOAD_STATUS_ERROR
End Function

Private Function pWaitForDownloadedFile( _
        ByRef DownloadedFilePath As String, _
        ByRef ErrorDescription As String) As String

    Call pRequireDownloadDirectory

    Dim deadline_time As Date
    deadline_time = DateAdd("s", pSettings.TimeoutSeconds, Now)

    Do
        Dim incomplete_files() As String
        incomplete_files = pFileSystem.GetFileList(pSessionClient.DownloadDirectoryPath, CatchRegExp:=C_INCOMPLETE_DOWNLOAD_FILE_REGEX)

        Dim completed_files() As String
        completed_files = pFileSystem.GetFileList(pSessionClient.DownloadDirectoryPath, IgnoreRegExp:=C_INCOMPLETE_DOWNLOAD_FILE_REGEX)

        If pStringArrayCount(incomplete_files) = 0 Then
            Select Case pStringArrayCount(completed_files)
             Case 1
                Dim candidate_file_path As String
                candidate_file_path = completed_files(LBound(completed_files))
                If pIsDownloadFileReady(candidate_file_path) Then
                    DownloadedFilePath = candidate_file_path
                    pWaitForDownloadedFile = G_WEB_DOWNLOAD_STATUS_DOWNLOADED
                    Exit Function
                End If
             Case Is > 1
                ErrorDescription = "ダウンロード完了ファイルが複数あります。(" & CStr(pStringArrayCount(completed_files)) & ")"
                pWaitForDownloadedFile = G_WEB_DOWNLOAD_STATUS_ERROR
                Exit Function
            End Select
        End If

        DoEvents
    Loop While Now < deadline_time

    ErrorDescription = "ダウンロード完了を待機中にタイムアウトしました。"
    pWaitForDownloadedFile = G_WEB_DOWNLOAD_STATUS_ERROR
End Function

Private Function pIsDownloadFileReady(ByVal FilePath As String) As Boolean
    On Error GoTo NOT_READY

    Dim result_value As Boolean
    result_value = False

    If pFileSystem.IsFile(FilePath) Then
        If 0# < pFileSystem.GetFileSize(FilePath) Then
            result_value = pFileSystem.IsFile(FilePath)
        End If
    End If

    On Error GoTo 0
    pIsDownloadFileReady = result_value
    Exit Function

NOT_READY:
    Err.Clear
    On Error GoTo 0
    pIsDownloadFileReady = False
End Function

Private Sub pClearDownloadDirectory()
    Call pRequireDownloadDirectory

    Dim staging_files() As String
    staging_files = pFileSystem.GetFileList(pSessionClient.DownloadDirectoryPath)

    If pStringArrayCount(staging_files) = 0 Then Exit Sub

    Dim file_index As Long
    For file_index = LBound(staging_files) To UBound(staging_files)
        Call pFileSystem.RemoveFile(staging_files(file_index))
    Next file_index
End Sub

Private Sub pRequireDownloadDirectory()
    If pSessionClient.DownloadDirectoryPath = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailFileDownloader.pRequireDownloadDirectory", Description:="一時ダウンロード領域がありません。"
    End If
End Sub

Private Function pStringArrayCount(ByRef Values() As String) As Long
    If IsEmptyArray(Values) Then
        pStringArrayCount = 0
    Else
        pStringArrayCount = UBound(Values) - LBound(Values) + 1
    End If
End Function
