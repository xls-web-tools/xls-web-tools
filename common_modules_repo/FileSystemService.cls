VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFileSystemService

' #############################################################################
'!
'! @brief
'! ファイル システムへのアクセスを管理するサービスの実装です。
'!
' #############################################################################

'Private pFSO As Object
Private pFSO As Scripting.FileSystemObject

Private Const C_ERR_FILE_NOT_FOUND As Long = 53
Private Const C_ERR_PATH_NOT_FOUND As Long = 76

Private Sub Class_Initialize()
    Set pFSO = CreateObject("Scripting.FileSystemObject")
    'Set pFSO = New Scripting.FileSystemObject
End Sub

'* 指定されたパスを絶対パスにします。
'*
'* @param PathLikeString パス文字列。
'* @return 絶対パス。
'*
'* @details
'* 指定されたパスを絶対パスにします。もともと絶対パスの場合も正規化した絶対パスを返します。
'* 相対パスの場合は、WorkbookService.GetThisWorkbookDirectoryPath の結果を基準にします。
Public Function GetAbsolutePath(ByVal PathLikeString As String) As String
    If IsAbsolutePath(PathLikeString) Then
        GetAbsolutePath = GetAbsolutePathFromParent(PathLikeString, PathLikeString)
    Else
        Call pRequireWorkbookService("GetAbsolutePath")
        GetAbsolutePath = GetAbsolutePathFromParent(WbSrv.GetThisWorkbookDirectoryPath(), PathLikeString)
    End If
End Function
Private Function IFileSystemService_GetAbsolutePath(ByVal PathLikeString As String) As String
    IFileSystemService_GetAbsolutePath = GetAbsolutePath(PathLikeString)
End Function

'* OS ユーザー一時フォルダーのパスを取得します。
'*
'* @return OS ユーザー一時フォルダーの絶対パス。
'*
'* @details
'* OS が示すユーザー一時フォルダーの絶対 Windows パスを返します。末尾のパス区切り文字は付けません。
Public Function GetTemporaryDirectoryPath() As String
    Dim result_value As String
    result_value = pFSO.GetSpecialFolder(TemporaryFolder).Path

    Do While Len(result_value) > Len(GetPathRoot(result_value)) And Right$(result_value, 1) = G_FS_PATH_SEP
        result_value = Left$(result_value, Len(result_value) - 1)
    Loop

    GetTemporaryDirectoryPath = result_value
End Function
Private Function IFileSystemService_GetTemporaryDirectoryPath() As String
    IFileSystemService_GetTemporaryDirectoryPath = GetTemporaryDirectoryPath()
End Function

'* OS ユーザー一時フォルダー直下に一時ディレクトリを作成します。
'*
'* @param NamePrefix [省略可] 一時ディレクトリ名の前に付ける文字列。
'* @return 作成した一時ディレクトリの絶対パス。
'*
'* @details
'* OS ユーザー一時フォルダー直下に衝突しにくい名前のディレクトリを作成し、作成した絶対 Windows パスを返します。
'* NamePrefix は変換しません。無効なフォルダー名になる場合は作成エラーを送出します。
Public Function CreateTemporaryDirectory(Optional ByVal NamePrefix As String = "") As String
    Const C_MAX_ATTEMPT As Long = 100

    Dim parent_path As String
    parent_path = GetTemporaryDirectoryPath()

    Dim attempt_idx As Long
    For attempt_idx = 1 To C_MAX_ATTEMPT
        Dim dir_path As String
        dir_path = JoinPath(parent_path, NamePrefix & pFSO.GetTempName())

        If Not PathExists(dir_path) Then
            On Error Resume Next
            Call pFSO.CreateFolder(dir_path)

            Dim err_num As Long
            Dim err_source As String
            Dim err_desc As String
            Dim err_help_file As String
            Dim err_help_context As Long
            err_num = Err.Number
            err_source = Err.Source
            err_desc = Err.Description
            err_help_file = Err.HelpFile
            err_help_context = Err.HelpContext
            Err.Clear
            On Error GoTo 0

            If err_num = 0 Then
                CreateTemporaryDirectory = dir_path
                Exit Function
            End If

            If Not PathExists(dir_path) Then
                Err.Raise Number:=err_num, Source:=err_source, Description:=err_desc, HelpFile:=err_help_file, HelpContext:=err_help_context
            End If
        End If
    Next attempt_idx

    Err.Raise Number:=vbObjectError + 1, Source:="Class FileSystemService", Description:="一時ディレクトリを作成できませんでした。"
End Function
Private Function IFileSystemService_CreateTemporaryDirectory(Optional ByVal NamePrefix As String = "") As String
    IFileSystemService_CreateTemporaryDirectory = CreateTemporaryDirectory(NamePrefix:=NamePrefix)
End Function

Private Sub pRequireWorkbookService(ByVal MethodName As String)
    If WbSrv Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class FileSystemService" & "." & MethodName, Description:="InitializeCommonService が呼ばれていません。"
    End If
End Sub

' 指定されたディレクトリにあるファイルの絶対パス一覧を得ます。
'
' @param DirectoryPath ディレクトリのパス。
' @param CatchRegExp [省略可] リストに含めるファイル名の正規表現。
' @param IgnoreRegExp [省略可] リストに含めないファイル名の正規表現。
' @param IgnoreCase [省略可] 正規表現の大小文字を無視するか。
' @return 絶対パスの配列。
'
' @details
' 指定されたディレクトリにあるファイルの絶対パス一覧を配列で得ます。該当なしの場合は 0 ベースの空配列を返します。
' 戻り順は FileSystemObject / OS の列挙順のままで、ソートしません。
' 列挙中に個別のファイルが削除またはリネームされ、名前またはパスの取得時にファイルまたはパスが見つからないエラーとなった場合は、その項目を除外して続行します。
Public Function GetFileList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)

    Dim file_list As Object 'Scripting.Files
    Set file_list = pFSO.GetFolder(dir_path).Files


    GetFileList = pFilterFiles(file_list, CatchRegExp, IgnoreRegExp, IgnoreCase)
End Function
Private Function IFileSystemService_GetFileList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    IFileSystemService_GetFileList = GetFileList(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp, IgnoreCase:=IgnoreCase)
End Function

' 指定されたディレクトリにあるサブ ディレクトリの絶対パス一覧を得ます。
'
' @param DirectoryPath ディレクトリのパス。
' @param CatchRegExp [省略可] リストに含めるサブ ディレクトリ名の正規表現。
' @param IgnoreRegExp [省略可] リストに含めないサブ ディレクトリ名の正規表現。
' @param IgnoreCase [省略可] 正規表現の大小文字を無視するか。
' @return 絶対パスの配列。
'
' @details
' 指定されたディレクトリにあるサブ ディレクトリの絶対パス一覧を配列で得ます。該当なしの場合は 0 ベースの空配列を返します。
' 戻り順は FileSystemObject / OS の列挙順のままで、ソートしません。
' 列挙中に個別のサブ ディレクトリが削除またはリネームされ、名前またはパスの取得時にファイルまたはパスが見つからないエラーとなった場合は、その項目を除外して続行します。
Public Function GetDirectoryList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)

    Dim folder_list As Object 'Scripting.Folders
    Set folder_list = pFSO.GetFolder(dir_path).SubFolders

    'GetDirectoryList = pFilterFolders(folder_list, CatchRegExp, IgnoreRegExp)
    GetDirectoryList = pFilterFiles(folder_list, CatchRegExp, IgnoreRegExp, IgnoreCase)
End Function
Private Function IFileSystemService_GetDirectoryList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    IFileSystemService_GetDirectoryList = GetDirectoryList(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp, IgnoreCase:=IgnoreCase)
End Function

Private Sub pCreateRegExpObject(ByRef CatchRegExpObject As Object, ByRef IgnoreRegExpObject As Object, ByVal CatchRegExp As String, ByVal IgnoreRegExp As String, ByVal IgnoreCase As Boolean)
    If CatchRegExp <> "" Then
        Set CatchRegExpObject = CreateObject("VBScript.RegExp")
        CatchRegExpObject.Pattern = CatchRegExp
        CatchRegExpObject.IgnoreCase = IgnoreCase
    End If

    If IgnoreRegExp <> "" Then
        Set IgnoreRegExpObject = CreateObject("VBScript.RegExp")
        IgnoreRegExpObject.Pattern = IgnoreRegExp
        IgnoreRegExpObject.IgnoreCase = IgnoreCase
    End If
End Sub

Private Function pIsMatchRegExp(ByVal Expression As String, ByVal CatchRegExpObject As Object, ByVal IgnoreRegExpObject As Object) As Boolean
    If pIsMatchRegExpCore(Expression, CatchRegExpObject, True) Then
        If pIsMatchRegExpCore(Expression, IgnoreRegExpObject, False) Then
            pIsMatchRegExp = False
        Else
            pIsMatchRegExp = True
        End If
    Else
        pIsMatchRegExp = False
    End If
End Function

Private Function pIsMatchRegExpCore(ByVal Expression As String, ByVal RegExpObject As Object, ByVal ResultWhenNothing As Boolean) As Boolean
    If RegExpObject Is Nothing Then
        pIsMatchRegExpCore = ResultWhenNothing
    Else
        If RegExpObject.Test(Expression) Then
            pIsMatchRegExpCore = True
        Else
            pIsMatchRegExpCore = False
        End If
    End If
End Function

Private Function pFilterFiles(ByVal FileList As Object, ByVal CatchRegExp As String, ByVal IgnoreRegExp As String, ByVal IgnoreCase As Boolean) As String()
    Dim catch_re As Object, ignore_re As Object 'RegExp
    Call pCreateRegExpObject(catch_re, ignore_re, CatchRegExp, IgnoreRegExp, IgnoreCase)

    Dim result_arr() As String
    Dim arr_item As Object 'Scripting.File / Scripting.Folder
    For Each arr_item In FileList
        Dim item_name As String
        Dim item_path As String
        If pTryReadFileListItem(item_name, item_path, arr_item) Then
            If pIsMatchRegExp(item_name, catch_re, ignore_re) Then
                Call pAddItemToStringArray(result_arr, item_path)
            End If
        End If
    Next arr_item

    If IsEmptyArray(result_arr) Then
        pFilterFiles = EmptyStringArray()
    Else
        pFilterFiles = result_arr
    End If
End Function

Private Function pTryReadFileListItem(ByRef ItemName As String, ByRef ItemPath As String, ByVal FileListItem As Object) As Boolean
    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    Dim err_help_file As String
    Dim err_help_context As Long

    Err.Clear
    On Error Resume Next
    ItemName = FileListItem.Name
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description
    err_help_file = Err.HelpFile
    err_help_context = Err.HelpContext
    Err.Clear
    On Error GoTo 0

    If err_num <> 0 Then
        If pIsMissingEnumeratedItemError(err_num) Then
            pTryReadFileListItem = False
            Exit Function
        End If
        Err.Raise Number:=err_num, Source:=err_source, Description:=err_desc, HelpFile:=err_help_file, HelpContext:=err_help_context
    End If

    Err.Clear
    On Error Resume Next
    ItemPath = FileListItem.Path
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description
    err_help_file = Err.HelpFile
    err_help_context = Err.HelpContext
    Err.Clear
    On Error GoTo 0

    If err_num <> 0 Then
        If pIsMissingEnumeratedItemError(err_num) Then
            pTryReadFileListItem = False
            Exit Function
        End If
        Err.Raise Number:=err_num, Source:=err_source, Description:=err_desc, HelpFile:=err_help_file, HelpContext:=err_help_context
    End If

    pTryReadFileListItem = True
End Function

Private Function pIsMissingEnumeratedItemError(ByVal ErrorNumber As Long) As Boolean
    Select Case ErrorNumber
        Case C_ERR_FILE_NOT_FOUND, C_ERR_PATH_NOT_FOUND
            pIsMissingEnumeratedItemError = True
        Case Else
            pIsMissingEnumeratedItemError = False
    End Select
End Function

Private Sub pAddItemToStringArray(ByRef TargetArray() As String, ByVal ItemToBeAdded As String)
    Dim new_ubound As Long

    If (Not TargetArray) = -1 Then
        ReDim Preserve TargetArray(0 To 0) As String
    Else
        new_ubound = UBound(TargetArray) + 1
        ReDim Preserve TargetArray(LBound(TargetArray) To new_ubound) As String
    End If
    TargetArray(new_ubound) = ItemToBeAdded
End Sub

'Private Function pFilterFolders(FolderList As Object, CatchRegExp As String, IgnoreRegExp As String) As String()
'    Dim catch_re As Object, ignore_re As Object 'RegExp
'    Call pCreateRegExpObject(catch_re, ignore_re, CatchRegExp, IgnoreRegExp, IgnoreCase)
'
'    Dim result_arr() As String
'
'    Dim arr_item As Object 'Scripting.Folder
'    For Each arr_item In FolderList
'        If pIsMatchRegExp(arr_item.Name, catch_re, ignore_re) Then
'
'            Call pAddItemToStringArray(result_arr, arr_item.Path)
'        End If
'    Next arr_item
'
'    pFilterFolders = result_arr
'End Function

' 指定されたパスに実体が存在するかを確認します。
'
' @param PathLikeString パス文字列。
' @return 実体が存在するか。
'
' @details
' 指定されたパスに実体が存在するかを確認します。ファイルかディレクトリがある場合に True となります。
Public Function PathExists(ByVal PathLikeString As String) As Boolean
    Dim fs_path As String
    fs_path = GetAbsolutePath(PathLikeString)

    If pFSO.FileExists(fs_path) Or pFSO.FolderExists(fs_path) Then
        PathExists = True
    Else
        PathExists = False
    End If
End Function
Private Function IFileSystemService_PathExists(ByVal PathLikeString As String) As Boolean
    IFileSystemService_PathExists = PathExists(PathLikeString)
End Function

' 指定されたパスにファイルが存在するかを確認します。
'
' @param FilePath ファイルのパス。
' @return 実体が存在するか。
'
' @details
' 指定されたパスにファイルが存在するかを確認します。
Public Function IsFile(ByVal FilePath As String) As Boolean
    Dim file_path As String
    file_path = GetAbsolutePath(FilePath)

    If pFSO.FileExists(file_path) Then
        IsFile = True
    Else
        IsFile = False
    End If
End Function
Private Function IFileSystemService_IsFile(ByVal FilePath As String) As Boolean
    IFileSystemService_IsFile = IsFile(FilePath)
End Function

' 指定されたパスにディレクトリが存在するかを確認します。
'
' @param DirectoryPath ディレクトリのパス。
' @return 実体が存在するか。
'
' @details
' 指定されたパスにディレクトリが存在するかを確認します。
Public Function IsDirectory(ByVal DirectoryPath As String) As Boolean
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)

    If pFSO.FolderExists(dir_path) Then
        IsDirectory = True
    Else
        IsDirectory = False
    End If
End Function
Private Function IFileSystemService_IsDirectory(ByVal DirectoryPath As String) As Boolean
    IFileSystemService_IsDirectory = IsDirectory(DirectoryPath)
End Function

' 最終更新日時を得ます
'
' @param PathLikeString パス文字列。
' @return 最終更新日時。
'
' @details
' 最終更新日時を得ます。
Public Function GetLastModified(ByVal PathLikeString As String) As Date
    Dim path_str As String
    path_str = GetAbsolutePath(PathLikeString)
    Dim path_obj As Object
    If IsFile(path_str) Then
        Set path_obj = pFSO.GetFile(path_str)
    ElseIf IsDirectory(path_str) Then
        Set path_obj = pFSO.GetFolder(path_str)
    Else
        Err.Raise vbObjectError + 1, "Class FileSystemService", "更新日時取得対象のファイルまたはディレクトリがありません。(" & path_str & ")"
    End If
    GetLastModified = path_obj.DateLastModified
End Function
Private Function IFileSystemService_GetLastModified(ByVal PathLikeString As String) As Date
    IFileSystemService_GetLastModified = GetLastModified(PathLikeString)
End Function

' ファイルサイズをバイト単位で得ます。
'
' @param FilePath ファイルのパス。
' @return ファイルサイズ。単位はバイトです。
'
' @details
' 指定されたファイルのサイズをバイト単位で得ます。
Public Function GetFileSize(ByVal FilePath As String) As Double
    Dim file_path As String
    file_path = GetAbsolutePath(FilePath)

    If pFSO.FileExists(file_path) Then
        GetFileSize = pFSO.GetFile(file_path).Size
    ElseIf pFSO.FolderExists(file_path) Then
        Err.Raise vbObjectError + 1, "Class FileSystemService.GetFileSize", "ファイルサイズ取得対象がファイルではありません。(" & file_path & ")"
    Else
        Err.Raise vbObjectError + 1, "Class FileSystemService.GetFileSize", "ファイルサイズ取得対象のファイルがありません。(" & file_path & ")"
    End If
End Function
Private Function IFileSystemService_GetFileSize(ByVal FilePath As String) As Double
    IFileSystemService_GetFileSize = GetFileSize(FilePath)
End Function

' ディレクトリを作成します。
'
' @param DirectoryPath ディレクトリのパス。
' @param Force [省略可] 作成先のパスにすでになにかが存在する場合に、上書き (削除) するか。
' @param Recursive [省略可] 親ディレクトリが存在しない場合に、再帰的に作成するか。
' @return 実際に作成したか。
'
' @details
' 指定されたパスにディレクトリを作成します。すでにディレクトリが存在していた場合は何もせず、戻り値が False になります。
' 実際にディレクトリを作成した場合は、戻り値が True になります。
Public Function CreateDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False, Optional ByVal Recursive As Boolean = False) As Boolean
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)

    If IsDirectory(dir_path) Then
        Debug.Print dir_path & " を作成しようとしましたが、すでに存在します。"
        CreateDirectory = False
        Exit Function
    End If

    Dim parent_dir As String
    parent_dir = GetParentPath(dir_path, IgnoreEndSep:=True)

    If IsDirectory(parent_dir) Then
        ' Force なら既存ファイルを消す
        If Force And IsFile(dir_path) Then Call RemoveFile(dir_path, Force:=True)
        ' 親ディレクトリが存在するので、普通に作成
        Call pFSO.CreateFolder(dir_path)
        CreateDirectory = True
    Else
        ' 親ディレクトリが存在しない場合
        If Not Recursive Then Err.Raise vbObjectError + 1, "Class FileSystemService", "親ディレクトリがありません。(" & parent_dir & ")"

        Dim dir_arr() As String

        ' ディレクトリ名を配列に格納しながら、親をたどっていく
        Do
            Call pAddItemToStringArray(dir_arr, GetLeafFromPath(dir_path, IgnoreEndSep:=True))
            dir_path = parent_dir
            parent_dir = GetParentPath(dir_path, IgnoreEndSep:=True)
        Loop While dir_path <> "" And Not IsDirectory(parent_dir)

        If dir_path = "" Then
            ' ルートが見つからなかった場合？
            Err.Raise vbObjectError + 1, "Class FileSystemService", "ディレクトリを作成できませんでした。(" & dir_path & ")"
        Else
            ' Force なら既存ファイルを消す
            If Force And IsFile(dir_path) Then Call RemoveFile(dir_path, Force:=True)
            ' ディレクトリを順に作成していく
            Dim item_idx As Long
            For item_idx = UBound(dir_arr) To LBound(dir_arr) Step -1
                Call pFSO.CreateFolder(dir_path)
                dir_path = JoinPath(dir_path, dir_arr(item_idx))
            Next item_idx
            Call pFSO.CreateFolder(dir_path)
            CreateDirectory = True
        End If
    End If
End Function
Private Function IFileSystemService_CreateDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False, Optional ByVal Recursive As Boolean = False) As Boolean
    IFileSystemService_CreateDirectory = CreateDirectory(DirectoryPath, Force:=Force, Recursive:=Recursive)
End Function

' ディレクトリを移動します。
'
' @param SourcePath 移動元ディレクトリのパス。
' @param DestinationPath 移動先ディレクトリのパス。
' @param Force [省略可] 移動先のパスにすでになにかが存在する場合に、上書きするか。
'
' @details
' 指定されたパスにディレクトリを作成します。DestinationPath がパス区切り文字で終わる場合、そのディレクトリのサブ ディレクトリとして、現在のディレクトリ名のまま移動します。
' Force:=True の場合は、CopyDirectory と同じ上書き規則で移動先へコピーしてから、移動元を削除します。既存の移動先ディレクトリ配下にある余剰ファイルは削除しません。
Public Sub MoveDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Dim src_path As String, dst_path As String
    src_path = GetAbsolutePath(SourcePath)
    dst_path = GetAbsolutePath(DestinationPath)

    If Not IsDirectory(src_path) Then Err.Raise vbObjectError + 1, "Class FileSystemService", "移動対象のディレクトリがありません。(" & src_path & ")"

    Dim ends_with_sep As Boolean
    ends_with_sep = pEndsWithPathSeparator(DestinationPath)

    Dim real_dst_path As String
    real_dst_path = pResolveOperationDestinationPath(src_path, dst_path, ends_with_sep)
    Call pValidateOperationDestination(src_path, real_dst_path, SourceIsDirectory:=True)

    If Force Then
        Call pFSO.CopyFolder(src_path, real_dst_path, OverWriteFiles:=True)
        Call RemoveDirectory(src_path, Force:=True)
    Else
        Call pFSO.MoveFolder(src_path, real_dst_path)
    End If
End Sub
Private Sub IFileSystemService_MoveDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call MoveDirectory(SourcePath, DestinationPath, Force:=Force)
End Sub

' ディレクトリをコピーします。
'
' @param SourcePath コピー元ディレクトリのパス。
' @param DestinationPath コピー先ディレクトリのパス。
' @param Force [省略可] コピー先のパスにすでになにかが存在する場合に、上書きするか。
'
' @details
' 指定されたパスにディレクトリを作成します。DestinationPath がパス区切り文字で終わる場合、そのディレクトリのサブ ディレクトリとして、現在のディレクトリ名のままコピーします。
' Force:=True の場合は、既存ファイルを上書きします。既存のコピー先ディレクトリ配下にある余剰ファイルは削除しません。
Public Sub CopyDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Dim src_path As String, dst_path As String
    src_path = GetAbsolutePath(SourcePath)
    dst_path = GetAbsolutePath(DestinationPath)

    If Not IsDirectory(src_path) Then Err.Raise vbObjectError + 1, "Class FileSystemService", "コピー対象のディレクトリがありません。(" & src_path & ")"

    Dim ends_with_sep As Boolean
    ends_with_sep = pEndsWithPathSeparator(DestinationPath)

    Dim real_dst_path As String
    real_dst_path = pResolveOperationDestinationPath(src_path, dst_path, ends_with_sep)
    Call pValidateOperationDestination(src_path, real_dst_path, SourceIsDirectory:=True)

    Call pFSO.CopyFolder(src_path, real_dst_path, OverWriteFiles:=Force)
End Sub
Private Sub IFileSystemService_CopyDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call CopyDirectory(SourcePath, DestinationPath, Force:=Force)
End Sub

' ディレクトリを削除します。
'
' @param DirectoryPath ディレクトリのパス。
' @param Force [省略可] 読み取り専用属性でも削除するか。
' @return 実際に削除したか。
'
' @details
' 指定されたパスにあるディレクトリを削除します。すでにディレクトリが存在していなかった場合は何もせず、戻り値が False になります。
' 実際にディレクトリを削除した場合は、戻り値が True になります。
Public Function RemoveDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False) As Boolean
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)

    If Not IsDirectory(dir_path) Then
        Debug.Print dir_path & "を消そうとしましたが､存在しませんでした｡"
        RemoveDirectory = False
        Exit Function
    End If

    Call pFSO.DeleteFolder(dir_path, Force:=Force) 'Force でなくとも空でないディレクトリも削除。これは、読み取り専用属性を無視するか否か
    RemoveDirectory = True
End Function
Private Function IFileSystemService_RemoveDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False) As Boolean
    IFileSystemService_RemoveDirectory = RemoveDirectory(DirectoryPath, Force:=Force)
End Function

' ファイルを移動します。
'
' @param SourcePath 移動元ファイルのパス。
' @param DestinationPath 移動先ファイルのパス。
' @param Force [省略可] 移動先のパスにすでになにかが存在する場合に、上書きするか。
'
' @details
' 指定されたパスにファイルを作成します。DestinationPath がパス区切り文字で終わる場合、そのディレクトリ配下のファイルとして、現在のファイル名のまま移動します。
' Force:=True の場合は、CopyFile と同じ上書き規則で移動先へコピーしてから、移動元を削除します。
Public Sub MoveFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Dim src_path As String, dst_path As String
    src_path = GetAbsolutePath(SourcePath)
    dst_path = GetAbsolutePath(DestinationPath)

    If Not IsFile(src_path) Then Err.Raise vbObjectError + 1, "Class FileSystemService", "移動対象のファイルがありません。(" & src_path & ")"

    Dim ends_with_sep As Boolean
    ends_with_sep = pEndsWithPathSeparator(DestinationPath)

    Dim real_dst_path As String
    real_dst_path = pResolveOperationDestinationPath(src_path, dst_path, ends_with_sep)
    Call pValidateOperationDestination(src_path, real_dst_path, SourceIsDirectory:=False)

    If Force Then
        Call pFSO.CopyFile(src_path, real_dst_path, OverWriteFiles:=True)
        Call RemoveFile(src_path, Force:=True)
    Else
        Call pFSO.MoveFile(src_path, real_dst_path)
    End If
End Sub
Private Sub IFileSystemService_MoveFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call MoveFile(SourcePath, DestinationPath, Force:=Force)
End Sub

' ファイルをコピーします。
'
' @param SourcePath コピー元ファイルのパス。
' @param DestinationPath コピー先ファイルのパス。
' @param Force [省略可] コピー先のパスにすでになにかが存在する場合に、上書きするか。
'
' @details
' 指定されたパスにファイルを作成します。DestinationPath がパス区切り文字で終わる場合、そのディレクトリ配下のファイルとして、現在のファイル名のままコピーします。
' Force:=True の場合は、既存ファイルを上書きします。
Public Sub CopyFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Dim src_path As String, dst_path As String
    src_path = GetAbsolutePath(SourcePath)
    dst_path = GetAbsolutePath(DestinationPath)

    If Not IsFile(src_path) Then Err.Raise vbObjectError + 1, "Class FileSystemService", "コピー対象のファイルがありません。(" & src_path & ")"

    Dim ends_with_sep As Boolean
    ends_with_sep = pEndsWithPathSeparator(DestinationPath)

    Dim real_dst_path As String
    real_dst_path = pResolveOperationDestinationPath(src_path, dst_path, ends_with_sep)
    Call pValidateOperationDestination(src_path, real_dst_path, SourceIsDirectory:=False)

    Call pFSO.CopyFile(src_path, real_dst_path, OverWriteFiles:=Force)
End Sub
Private Sub IFileSystemService_CopyFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call CopyFile(SourcePath, DestinationPath, Force:=Force)
End Sub

' ファイルを削除します。
'
' @param FilePath ファイルのパス。
' @param Force [省略可] 読み取り専用属性でも削除するか。
' @return 実際に削除したか。
'
' @details
' 指定されたパスにあるファイルを削除します。すでにファイルが存在していなかった場合は何もせず、戻り値が False になります。
' 実際にファイルを削除した場合は、戻り値が True になります。
Public Function RemoveFile(ByVal FilePath As String, Optional ByVal Force As Boolean = False) As Boolean
    Dim file_path As String
    file_path = GetAbsolutePath(FilePath)

    If Not IsFile(file_path) Then
        Debug.Print file_path & "を消そうとしましたが､存在しませんでした｡"
        RemoveFile = False
        Exit Function
    End If

    Call pFSO.DeleteFile(file_path, Force:=Force) 'ロックは関係ない、読み取り専用属性を無視するか否か
    RemoveFile = True
End Function
Private Function IFileSystemService_RemoveFile(ByVal FilePath As String, Optional ByVal Force As Boolean = False) As Boolean
    IFileSystemService_RemoveFile = RemoveFile(FilePath, Force:=Force)
End Function

' 指定されたフォルダー直下にある、指定された名前のファイルのうち、最新のものを取得します。
'
' @param DirectoryPath 検索対象のフォルダーのパス
' @param CatchRegExp [省略可] チェック対象とするファイル名の正規表現。
' @param IgnoreRegExp [省略可] チェック対象としないファイル名の正規表現。
' @param IgnoreCase [省略可] 正規表現の大小文字を無視するか。
' @return 最新のファイルのフルパス。該当するファイルがない場合は空文字列
'
' @details
' 指定されたフォルダー内のファイルを検索し、指定されたパターンに一致するファイルの中で
' 最も新しいもののフルパスを返します。走査順は FileSystemObject / OS の列挙順のままで、同一時刻の順序は保証しません。
Public Function GetNewestFile(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String
    Dim catch_re As Object, ignore_re As Object 'RegExp
    Call pCreateRegExpObject(catch_re, ignore_re, CatchRegExp, IgnoreRegExp, IgnoreCase)

    Dim dir_obj As Object 'Folder
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)
    Set dir_obj = pFSO.GetFolder(dir_path)

    Dim newest_file As Object 'File
    Dim file_item As Object 'File
    For Each file_item In dir_obj.Files
        If pIsMatchRegExp(file_item.Name, catch_re, ignore_re) Then
            If newest_file Is Nothing Then
                Set newest_file = file_item
            Else
                If newest_file.DateLastModified < file_item.DateLastModified Then
                    Set newest_file = file_item
                End If
            End If
        End If
    Next file_item

    If newest_file Is Nothing Then
        GetNewestFile = ""
    Else
        GetNewestFile = newest_file.Path
    End If
End Function
Private Function IFileSystemService_GetNewestFile(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String
    IFileSystemService_GetNewestFile = GetNewestFile(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp, IgnoreCase:=IgnoreCase)
End Function

' バックアップファイルを作成します。
'
' @param SourceFilePath [省略可] バックアップ対象のファイルパス (既定値は現在のワークブックのパス)
' @param DestinationDirectoryPath [省略可] バックアップ先のディレクトリパス (既定値は対象ファイルのディレクトリ)
' @param BackupSuffix [省略可] バックアップファイル名に追加するサフィックス (既定値はタイムスタンプ)
' @param Force [省略可] 既存のバックアップファイルを強制的に上書きするかどうか
' @param Move [省略可] バックアップ時に元のファイルを移動するかどうか (True の場合、移動)
' @param FailIfSourceNotFound 対象ファイルが存在しない場合にエラーを発生させるかどうか (True の場合、エラー)
' @return 作成されたバックアップファイルのフルパス
'
' @details
' 指定されたファイルのバックアップを作成します。バックアップ先のファイル名には元のファイル名に
' サフィックスが追加されます。既存のファイルがある場合は `Force` の設定に基づいて処理されます。
' `Force:=True` は既存のバックアップ ファイルを上書きします。
' `Force:=False` かつ `BackupSuffix:=""` の既定サフィックスが衝突した場合は、既存バックアップのパスを返し、内容比較や `Move:=True` の移動は行いません。
' `Force:=False` かつ `BackupSuffix` を明示した場合の衝突はエラーです。
Public Function CreateBackupFile( _
        Optional ByVal SourceFilePath As String = "", _
        Optional ByVal DestinationDirectoryPath As String = "", _
        Optional ByVal BackupSuffix As String = "", _
        Optional ByVal Force As Boolean = False, _
        Optional ByVal Move As Boolean = False, _
        Optional ByVal FailIfSourceNotFound As Boolean = True) As String

    Dim src_file_path As String
    Dim src_file_name As String
    Dim src_file_obj As File
    Dim file_base As String
    Dim file_ext As String
    Dim dst_file_path As String
    Dim dst_dir_path As String
    Dim suffix_str As String

    ' バックアップ対象ファイルが指定されなかったら、このワークブックが指定されたものとする。
    If SourceFilePath = "" Then
        src_file_path = ThisWorkbook.FullName
    Else
        src_file_path = GetAbsolutePath(SourceFilePath)
    End If

    ' バックアップ対象ファイルがない場合はエラー
    If Not IsFile(src_file_path) Then
        If FailIfSourceNotFound Then
            Err.Raise Number:=vbObjectError + 1, Source:="Sub CreateBackupFile", Description:="バックアップ対象のファイルがありません。(" & SourceFilePath & ")"
        Else
            Debug.Print "バックアップ対象のファイルがありませんでした。(" & src_file_path & ")"
            Exit Function
        End If
    End If

    ' バックアップ先ディレクトリが指定されなかった場合は、対象ファイルと同じディレクトリを出力先とする。
    If DestinationDirectoryPath = "" Then
        dst_dir_path = GetParentPath(src_file_path, IgnoreEndSep:=True)
    Else
        If IsAbsolutePath(DestinationDirectoryPath) Then
            dst_dir_path = DestinationDirectoryPath
        Else
            dst_dir_path = JoinPath(GetParentPath(src_file_path), DestinationDirectoryPath)
        End If
    End If

    ' バックアップ先ディレクトリがない場合は作成する。
    If IsFile(dst_dir_path) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Sub CreateBackupFile", Description:="バックアップ出力先がディレクトリではありません。(" & DestinationDirectoryPath & ")"
    ElseIf Not IsDirectory(dst_dir_path) Then
        Call CreateDirectory(dst_dir_path, Recursive:=True)
    End If

    ' サフィックスが指定されない場合は、最終更新日時を使用する
    If BackupSuffix = "" Then
        suffix_str = Format(GetLastModified(src_file_path), "_yyyy-MM-dd_HH-mm-ss")
    Else
        suffix_str = BackupSuffix
    End If

    ' 元のパスからファイル名を得る。
    file_base = GetLeafFromPath(src_file_path, BaseName:=True, Extension:=False)
    file_ext = GetLeafFromPath(src_file_path, BaseName:=False, Extension:=True)

    ' バックアップ先ファイルパスを得る。
    dst_file_path = JoinPath(dst_dir_path, file_base & suffix_str & file_ext)

    If IsFile(dst_file_path) Then
        If Not Force Then
            ' 強制フラグが立っていない場合、バックアップ対象ファイルがある場合はエラー
            If BackupSuffix = "" Then
                ' 時刻に基づいたサフィックスで重複する場合は、バックアップ済みなので処理終了
                Debug.Print "すでにバックアップ済みでした。(" & src_file_path & ")"
                CreateBackupFile = dst_file_path
                Exit Function
            Else
                Err.Raise Number:=vbObjectError + 1, Source:="Sub CreateBackupFile", Description:="バックアップ先のファイルがすでにあります。(" & dst_file_path & ")"
            End If
        End If
    ElseIf IsDirectory(dst_file_path) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Sub CreateBackupFile", Description:="バックアップ先のファイルと同名のディレクトリがあります。(" & dst_file_path & ")"
    End If

    ' ファイルをバックアップ
    If Move Then
        Call MoveFile(src_file_path, dst_file_path, Force:=Force)
    Else
        Call CopyFile(src_file_path, dst_file_path, Force:=Force)
    End If

    CreateBackupFile = dst_file_path
End Function
Private Function IFileSystemService_CreateBackupFile( _
        Optional ByVal SourceFilePath As String = "", _
        Optional ByVal DestinationDirectoryPath As String = "", _
        Optional ByVal BackupSuffix As String = "", _
        Optional ByVal Force As Boolean = False, _
        Optional ByVal Move As Boolean = False, _
        Optional ByVal FailIfSourceNotFound As Boolean = True) As String

    IFileSystemService_CreateBackupFile = CreateBackupFile( _
            SourceFilePath:=SourceFilePath, _
            DestinationDirectoryPath:=DestinationDirectoryPath, _
            BackupSuffix:=BackupSuffix, _
            Force:=Force, Move:=Move, FailIfSourceNotFound:=FailIfSourceNotFound)
End Function

Private Function pEndsWithPathSeparator(ByVal PathLikeString As String) As Boolean
    Dim last_char As String
    last_char = Right$(PathLikeString, 1)
    pEndsWithPathSeparator = (last_char = G_FS_PATH_SEP Or last_char = "/")
End Function

Private Function pResolveOperationDestinationPath(ByVal SourcePath As String, ByVal DestinationPath As String, ByVal EndsWithSep As Boolean) As String
    If EndsWithSep Then
        pResolveOperationDestinationPath = JoinPath(DestinationPath, GetLeafFromPath(SourcePath, IgnoreEndSep:=True))
    Else
        pResolveOperationDestinationPath = DestinationPath
    End If
End Function

Private Sub pValidateOperationDestination(ByVal SourcePath As String, ByVal DestinationPath As String, ByVal SourceIsDirectory As Boolean)
    If pPathsEqual(SourcePath, DestinationPath) Then
        Err.Raise vbObjectError + 1, "Class FileSystemService", "コピー/移動元とコピー/移動先が同一です。(" & SourcePath & ")"
    End If

    If SourceIsDirectory Then
        If pIsPathInside(DestinationPath, SourcePath) Then
            Err.Raise vbObjectError + 1, "Class FileSystemService", "コピー/移動先がコピー/移動元の配下です。(" & DestinationPath & ")"
        End If
        If pIsPathInside(SourcePath, DestinationPath) Then
            Err.Raise vbObjectError + 1, "Class FileSystemService", "コピー/移動先がコピー/移動元の親ディレクトリです。(" & DestinationPath & ")"
        End If
    End If
End Sub

Private Function pPathsEqual(ByVal Path1 As String, ByVal Path2 As String) As Boolean
    pPathsEqual = (StrComp(pNormalizeComparePath(Path1), pNormalizeComparePath(Path2), vbTextCompare) = 0)
End Function

Private Function pIsPathInside(ByVal ChildPath As String, ByVal ParentPath As String) As Boolean
    Dim child_path_norm As String
    child_path_norm = pNormalizeComparePath(ChildPath)
    Dim parent_path_norm As String
    parent_path_norm = pNormalizeComparePath(ParentPath)

    If Right$(parent_path_norm, 1) <> G_FS_PATH_SEP Then parent_path_norm = parent_path_norm & G_FS_PATH_SEP
    pIsPathInside = (Len(parent_path_norm) < Len(child_path_norm) And StrComp(Left$(child_path_norm, Len(parent_path_norm)), parent_path_norm, vbTextCompare) = 0)
End Function

Private Function pNormalizeComparePath(ByVal PathLikeString As String) As String
    Dim path_value As String
    path_value = GetAbsolutePath(PathLikeString)
    Dim root_value As String
    root_value = GetPathRoot(path_value)

    Do While Len(path_value) > Len(root_value) And Right$(path_value, 1) = G_FS_PATH_SEP
        path_value = Left$(path_value, Len(path_value) - 1)
    Loop

    pNormalizeComparePath = path_value
End Function
