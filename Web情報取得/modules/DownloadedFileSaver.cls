VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DownloadedFileSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! ダウンロード済みファイルを対象ID別の出力先へ保存します。
'!
' #############################################################################

Private Const C_ZIP_EXTRACT_DIRECTORY_PREFIX As String = "xls-web-tools_zip_"

Private pSettings As IToolSettings
Private pFileSystem As IFileSystemService
Private pZipExtractor As IZipExtractor
Private pLastErrorDescription As String
Private pIsInitialized As Boolean

'* 直近の保存エラー内容。
Public Property Get LastErrorDescription() As String
    Call pCheckInit
    LastErrorDescription = pLastErrorDescription
End Property

'* DownloadedFileSaver を初期化します。
'*
'* @param Settings Web情報取得の設定。
'* @param ZipExtractor ZIP 展開処理。
'*
'* @details
'* ダウンロード保存先、ZIP flatten、ファイル衝突判定を設定に従って実行します。ファイル操作は FsSrv、ZIP 展開は ZipExtractor を使用します。
Public Sub Initialize( _
        ByVal Settings As IToolSettings, _
        ByVal ZipExtractor As IZipExtractor)

    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DownloadedFileSaver", Description:="Settings は必須です。"
    End If
    If ZipExtractor Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DownloadedFileSaver", Description:="ZipExtractor は必須です。"
    End If

    Set pSettings = Settings
    Set pFileSystem = FsSrv
    Set pZipExtractor = ZipExtractor
    pLastErrorDescription = ""
    pIsInitialized = True
End Sub

'* ダウンロード済みファイルを保存し、ダウンロード状態を返します。
'*
'* @param DownloadedFilePath staging 領域に取得済みのファイルパス。
'* @param TargetId 保存先ファイル名の決定に使う対象ID。
'* @return ダウンロード状態。
'*
'* @details
'* DownloadRootPath 配下へファイルを保存します。ZIP flatten 設定時は展開結果を保存し、保存できない理由は LastErrorDescription に保持します。
Public Function SaveDownloadedFile( _
        ByVal DownloadedFilePath As String, _
        ByVal TargetId As String) As String

    Call pCheckInit

    pLastErrorDescription = ""
    On Error GoTo ON_ERROR

    Dim target_id_text As String
    target_id_text = Trim$(TargetId)
    Call pValidateTargetId(target_id_text)

    Dim target_dir_path As String
    target_dir_path = JoinPath(pSettings.DownloadRootPath, target_id_text)

    Dim result_status As String
    If pIsZipFile(DownloadedFilePath) Then
        result_status = pSaveZipFile(DownloadedFilePath, target_id_text, target_dir_path)
    Else
        result_status = pSaveRegularFile(DownloadedFilePath, target_id_text, target_dir_path)
    End If

    On Error GoTo 0
    SaveDownloadedFile = result_status
    Exit Function

ON_ERROR:
    pLastErrorDescription = Err.Description
    Err.Clear
    On Error GoTo 0

    SaveDownloadedFile = G_WEB_DOWNLOAD_STATUS_ERROR
End Function

Private Function pSaveRegularFile( _
        ByVal DownloadedFilePath As String, _
        ByVal TargetId As String, _
        ByVal TargetDirectoryPath As String) As String

    Dim destination_path As String
    destination_path = JoinPath(TargetDirectoryPath, pBuildPrefixedFileName(DownloadedFilePath, TargetId))

    Dim result_status As String
    If pFileSystem.PathExists(destination_path) Then
        result_status = G_WEB_DOWNLOAD_STATUS_SKIPPED
    Else
        Call pEnsureTargetDirectory(TargetDirectoryPath)
        Call pFileSystem.MoveFile(DownloadedFilePath, destination_path)
        result_status = G_WEB_DOWNLOAD_STATUS_DOWNLOADED
    End If

    pSaveRegularFile = result_status
End Function

Private Function pSaveZipFile( _
        ByVal ZipFilePath As String, _
        ByVal TargetId As String, _
        ByVal TargetDirectoryPath As String) As String

    Dim extract_dir_path As String
    extract_dir_path = pFileSystem.CreateTemporaryDirectory(C_ZIP_EXTRACT_DIRECTORY_PREFIX)

    On Error GoTo ON_ERROR

    Call pZipExtractor.ExtractZip(ZipFilePath, extract_dir_path)

    Dim source_root_path As String
    source_root_path = pResolveZipSourceRoot(extract_dir_path)

    Dim output_files() As String
    output_files = pFileSystem.GetFileList(source_root_path)

    Dim output_dirs() As String
    output_dirs = pFileSystem.GetDirectoryList(source_root_path)

    Dim result_status As String
    If pStringArrayCount(output_files) = 0 And pStringArrayCount(output_dirs) = 0 Then
        pLastErrorDescription = "ZIP 展開後の出力対象がありません。"
        result_status = G_WEB_DOWNLOAD_STATUS_ERROR
    ElseIf pZipOutputHasCollision(output_files, output_dirs, TargetId, TargetDirectoryPath) Then
        result_status = G_WEB_DOWNLOAD_STATUS_SKIPPED
    Else
        Call pEnsureTargetDirectory(TargetDirectoryPath)
        Call pMoveZipOutputFiles(output_files, TargetId, TargetDirectoryPath)
        Call pMoveZipOutputDirectories(output_dirs, TargetDirectoryPath)
        result_status = G_WEB_DOWNLOAD_STATUS_DOWNLOADED
    End If

    On Error GoTo 0
    Call pTryRemoveDirectory(extract_dir_path)

    pSaveZipFile = result_status
    Exit Function

ON_ERROR:
    pLastErrorDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    Call pTryRemoveDirectory(extract_dir_path)

    pSaveZipFile = G_WEB_DOWNLOAD_STATUS_ERROR
End Function

Private Sub pValidateTargetId(ByVal TargetId As String)
    If TargetId = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DownloadedFileSaver", Description:="対象IDが空です。"
    End If

    If pContainsInvalidFolderNameCharacter(TargetId) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DownloadedFileSaver", Description:="対象IDを Windows フォルダー名として使えません。(" & TargetId & ")"
    End If

    If Right$(TargetId, 1) = "." Or Right$(TargetId, 1) = " " Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DownloadedFileSaver", Description:="対象IDの末尾が Windows フォルダー名として使えません。(" & TargetId & ")"
    End If

    If pIsReservedFolderName(TargetId) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DownloadedFileSaver", Description:="対象IDが Windows の予約名です。(" & TargetId & ")"
    End If
End Sub

Private Function pIsZipFile(ByVal FilePath As String) As Boolean
    pIsZipFile = (StrComp(GetLeafFromPath(FilePath, BaseName:=False, Extension:=True), ".zip", vbTextCompare) = 0)
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DownloadedFileSaver", Description:="初期化されていません。"
    End If
End Sub

Private Function pResolveZipSourceRoot(ByVal ExtractDirectoryPath As String) As String
    Dim root_files() As String
    root_files = pFileSystem.GetFileList(ExtractDirectoryPath)

    Dim root_dirs() As String
    root_dirs = pFileSystem.GetDirectoryList(ExtractDirectoryPath)

    Dim result_path As String
    If pStringArrayCount(root_files) = 0 And pStringArrayCount(root_dirs) = 1 Then
        result_path = root_dirs(LBound(root_dirs))
    Else
        result_path = ExtractDirectoryPath
    End If

    pResolveZipSourceRoot = result_path
End Function

Private Function pZipOutputHasCollision( _
        ByRef OutputFiles() As String, _
        ByRef OutputDirectories() As String, _
        ByVal TargetId As String, _
        ByVal TargetDirectoryPath As String) As Boolean

    Dim file_idx As Long
    If Not IsEmptyArray(OutputFiles) Then
        For file_idx = LBound(OutputFiles) To UBound(OutputFiles)
            If pFileSystem.PathExists(JoinPath(TargetDirectoryPath, pBuildPrefixedFileName(OutputFiles(file_idx), TargetId))) Then
                pZipOutputHasCollision = True
                Exit Function
            End If
        Next file_idx
    End If

    Dim dir_idx As Long
    If Not IsEmptyArray(OutputDirectories) Then
        For dir_idx = LBound(OutputDirectories) To UBound(OutputDirectories)
            If pFileSystem.PathExists(JoinPath(TargetDirectoryPath, GetLeafFromPath(OutputDirectories(dir_idx), IgnoreEndSep:=True))) Then
                pZipOutputHasCollision = True
                Exit Function
            End If
        Next dir_idx
    End If

    pZipOutputHasCollision = False
End Function

Private Sub pMoveZipOutputFiles( _
        ByRef OutputFiles() As String, _
        ByVal TargetId As String, _
        ByVal TargetDirectoryPath As String)

    If IsEmptyArray(OutputFiles) Then Exit Sub

    Dim file_idx As Long
    For file_idx = LBound(OutputFiles) To UBound(OutputFiles)
        Call pFileSystem.MoveFile( _
                OutputFiles(file_idx), _
                JoinPath(TargetDirectoryPath, pBuildPrefixedFileName(OutputFiles(file_idx), TargetId)))
    Next file_idx
End Sub

Private Sub pMoveZipOutputDirectories( _
        ByRef OutputDirectories() As String, _
        ByVal TargetDirectoryPath As String)

    If IsEmptyArray(OutputDirectories) Then Exit Sub

    Dim dir_idx As Long
    For dir_idx = LBound(OutputDirectories) To UBound(OutputDirectories)
        Call pFileSystem.MoveDirectory( _
                OutputDirectories(dir_idx), _
                JoinPath(TargetDirectoryPath, GetLeafFromPath(OutputDirectories(dir_idx), IgnoreEndSep:=True)))
    Next dir_idx
End Sub

Private Sub pEnsureTargetDirectory(ByVal TargetDirectoryPath As String)
    If pFileSystem.PathExists(TargetDirectoryPath) Then Exit Sub

    Call pFileSystem.CreateDirectory(TargetDirectoryPath, Recursive:=True)
End Sub

Private Function pBuildPrefixedFileName( _
        ByVal FilePath As String, _
        ByVal TargetId As String) As String

    pBuildPrefixedFileName = pBuildTargetIdPrefix(TargetId) & "_" & GetLeafFromPath(FilePath)
End Function

Private Function pBuildTargetIdPrefix(ByVal TargetId As String) As String
    If Len(TargetId) <= 3 Then
        pBuildTargetIdPrefix = TargetId
    Else
        pBuildTargetIdPrefix = Right$(TargetId, 3)
    End If
End Function

Private Function pContainsInvalidFolderNameCharacter(ByVal FolderName As String) As Boolean
    Dim invalid_chars As Variant
    invalid_chars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")

    Dim char_idx As Long
    For char_idx = LBound(invalid_chars) To UBound(invalid_chars)
        If 0 < InStr(1, FolderName, CStr(invalid_chars(char_idx)), vbBinaryCompare) Then
            pContainsInvalidFolderNameCharacter = True
            Exit Function
        End If
    Next char_idx

    Dim text_idx As Long
    For text_idx = 1 To Len(FolderName)
        If AscW(Mid$(FolderName, text_idx, 1)) < 32 Then
            pContainsInvalidFolderNameCharacter = True
            Exit Function
        End If
    Next text_idx

    pContainsInvalidFolderNameCharacter = False
End Function

Private Function pIsReservedFolderName(ByVal FolderName As String) As Boolean
    Dim base_name As String
    base_name = UCase$(GetLeafFromPath(FolderName, BaseName:=True, Extension:=False))

    Select Case base_name
     Case "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
        pIsReservedFolderName = True
     Case Else
        pIsReservedFolderName = False
    End Select
End Function

Private Function pStringArrayCount(ByRef Values() As String) As Long
    If IsEmptyArray(Values) Then
        pStringArrayCount = 0
    Else
        pStringArrayCount = UBound(Values) - LBound(Values) + 1
    End If
End Function

Private Sub pTryRemoveDirectory(ByVal DirectoryPath As String)
    If Trim$(DirectoryPath) = "" Then Exit Sub

    On Error Resume Next
    Call pFileSystem.RemoveDirectory(DirectoryPath, Force:=True)
    On Error GoTo 0
End Sub
