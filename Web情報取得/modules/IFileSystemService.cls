VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileSystemService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#Interface

' #############################################################################
'!
'! @brief
'! ファイル システムへのアクセスを管理するサービスのインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IFileSystemService", Description:="このクラスはインターフェイス用です。"
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
End Function

'* OS ユーザー一時フォルダーのパスを取得します。
'*
'* @return OS ユーザー一時フォルダーの絶対パス。
'*
'* @details
'* OS が示すユーザー一時フォルダーの絶対 Windows パスを返します。末尾のパス区切り文字は付けません。
Public Function GetTemporaryDirectoryPath() As String
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
End Function

'* 指定されたディレクトリにあるファイルの絶対パス一覧を得ます。
'*
'* @param DirectoryPath ディレクトリのパス。
'* @param CatchRegExp [省略可] リストに含めるファイル名の正規表現。
'* @param IgnoreRegExp [省略可] リストに含めないファイル名の正規表現。
'* @param IgnoreCase [省略可] 正規表現の大小文字を無視するか。
'* @return 絶対パスの配列。
'*
'* @details
'* 指定されたディレクトリにあるファイルの絶対パス一覧を配列で得ます。該当なしの場合は 0 ベースの空配列を返します。
'* 戻り順は FileSystemObject / OS の列挙順のままで、ソートしません。
Public Function GetFileList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
End Function

'* 指定されたディレクトリにあるサブ ディレクトリの絶対パス一覧を得ます。
'*
'* @param DirectoryPath ディレクトリのパス。
'* @param CatchRegExp [省略可] リストに含めるサブ ディレクトリ名の正規表現。
'* @param IgnoreRegExp [省略可] リストに含めないサブ ディレクトリ名の正規表現。
'* @param IgnoreCase [省略可] 正規表現の大小文字を無視するか。
'* @return 絶対パスの配列。
'*
'* @details
'* 指定されたディレクトリにあるサブ ディレクトリの絶対パス一覧を配列で得ます。該当なしの場合は 0 ベースの空配列を返します。
'* 戻り順は FileSystemObject / OS の列挙順のままで、ソートしません。
Public Function GetDirectoryList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
End Function

'* 指定されたパスに実体が存在するかを確認します。
'*
'* @param PathLikeString パス文字列。
'* @return 実体が存在するか。
'*
'* @details
'* 指定されたパスに実体が存在するかを確認します。ファイルかディレクトリがある場合に True となります。
Public Function PathExists(ByVal PathLikeString As String) As Boolean
End Function

'* 指定されたパスにファイルが存在するかを確認します。
'*
'* @param FilePath ファイルのパス。
'* @return 実体が存在するか。
'*
'* @details
'* 指定されたパスにファイルが存在するかを確認します。
Public Function IsFile(ByVal FilePath As String) As Boolean
End Function

'* 指定されたパスにディレクトリが存在するかを確認します。
'*
'* @param DirectoryPath ディレクトリのパス。
'* @return 実体が存在するか。
'*
'* @details
'* 指定されたパスにディレクトリが存在するかを確認します。
Public Function IsDirectory(ByVal DirectoryPath As String) As Boolean
End Function

'* 最終更新日時を得ます
'*
'* @param PathLikeString パス文字列。
'* @return 最終更新日時。
'*
'* @details
'* 最終更新日時を得ます。
Public Function GetLastModified(ByVal PathLikeString As String) As Date
End Function

'* ファイルサイズをバイト単位で得ます。
'*
'* @param FilePath ファイルのパス。
'* @return ファイルサイズ。単位はバイトです。
'*
'* @details
'* 指定されたファイルのサイズをバイト単位で得ます。対象が存在しない、またはファイルではない場合はエラーです。
Public Function GetFileSize(ByVal FilePath As String) As Double
End Function

'* ディレクトリを作成します。
'*
'* @param DirectoryPath ディレクトリのパス。
'* @param Force [省略可] 作成先のパスにすでになにかが存在する場合に、上書き (削除) するか。
'* @param Recursive [省略可] 親ディレクトリが存在しない場合に、再帰的に作成するか。
'* @return 実際に作成したか。
'*
'* @details
'* 指定されたパスにディレクトリを作成します。すでにディレクトリが存在していた場合は何もせず、戻り値が False になります。
'* 実際にディレクトリを作成した場合は、戻り値が True になります。
Public Function CreateDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False, Optional ByVal Recursive As Boolean = False) As Boolean
End Function

'* ディレクトリを移動します。
'*
'* @param SourcePath 移動元ディレクトリのパス。
'* @param DestinationPath 移動先ディレクトリのパス。
'* @param Force [省略可] 移動先のパスにすでになにかが存在する場合に、上書きするか。
'*
'* @details
'* 指定されたパスにディレクトリを作成します。DestinationPath がパス区切り文字で終わる場合、そのディレクトリのサブ ディレクトリとして、現在のディレクトリ名のまま移動します。
'* Force:=True の場合は、CopyDirectory と同じ上書き規則で移動先へコピーしてから、移動元を削除します。既存の移動先ディレクトリ配下にある余剰ファイルは削除しません。
Public Sub MoveDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
End Sub

'* ディレクトリをコピーします。
'*
'* @param SourcePath コピー元ディレクトリのパス。
'* @param DestinationPath コピー先ディレクトリのパス。
'* @param Force [省略可] コピー先のパスにすでになにかが存在する場合に、上書きするか。
'*
'* @details
'* 指定されたパスにディレクトリを作成します。DestinationPath がパス区切り文字で終わる場合、そのディレクトリのサブ ディレクトリとして、現在のディレクトリ名のままコピーします。
'* Force:=True の場合は、既存ファイルを上書きします。既存のコピー先ディレクトリ配下にある余剰ファイルは削除しません。
Public Sub CopyDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
End Sub

'* ディレクトリを削除します。
'*
'* @param DirectoryPath ディレクトリのパス。
'* @param Force [省略可] 読み取り専用属性でも削除するか。
'* @return 実際に削除したか。
'*
'* @details
'* 指定されたパスにあるディレクトリを削除します。すでにディレクトリが存在していなかった場合は何もせず、戻り値が False になります。
'* 実際にディレクトリを削除した場合は、戻り値が True になります。
Public Function RemoveDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False) As Boolean
End Function

'* ファイルを移動します。
'*
'* @param SourcePath 移動元ファイルのパス。
'* @param DestinationPath 移動先ファイルのパス。
'* @param Force [省略可] 移動先のパスにすでになにかが存在する場合に、上書きするか。
'*
'* @details
'* 指定されたパスにファイルを作成します。DestinationPath がパス区切り文字で終わる場合、そのディレクトリ配下のファイルとして、現在のファイル名のまま移動します。
'* Force:=True の場合は、CopyFile と同じ上書き規則で移動先へコピーしてから、移動元を削除します。
Public Sub MoveFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
End Sub

'* ファイルをコピーします。
'*
'* @param SourcePath コピー元ファイルのパス。
'* @param DestinationPath コピー先ファイルのパス。
'* @param Force [省略可] コピー先のパスにすでになにかが存在する場合に、上書きするか。
'*
'* @details
'* 指定されたパスにファイルを作成します。DestinationPath がパス区切り文字で終わる場合、そのディレクトリ配下のファイルとして、現在のファイル名のままコピーします。
'* Force:=True の場合は、既存ファイルを上書きします。
Public Sub CopyFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
End Sub

'* ファイルを削除します。
'*
'* @param FilePath ファイルのパス。
'* @param Force [省略可] 読み取り専用属性でも削除するか。
'* @return 実際に削除したか。
'*
'* @details
'* 指定されたパスにあるファイルを削除します。すでにファイルが存在していなかった場合は何もせず、戻り値が False になります。
'* 実際にファイルを削除した場合は、戻り値が True になります。
Public Function RemoveFile(ByVal FilePath As String, Optional ByVal Force As Boolean = False) As Boolean
End Function

'* 指定されたフォルダー直下にある、指定された名前のファイルのうち、最新のものを取得します。
'*
'* @param DirectoryPath 検索対象のフォルダーのパス
'* @param CatchRegExp [省略可] チェック対象とするファイル名の正規表現。
'* @param IgnoreRegExp [省略可] チェック対象としないファイル名の正規表現。
'* @param IgnoreCase [省略可] 正規表現の大小文字を無視するか。
'* @return 最新のファイルのフルパス。該当するファイルがない場合は空文字列
'*
'* @details
'* 指定されたフォルダー内のファイルを検索し、指定されたパターンに一致するファイルの中で
'* 最も新しいもののフルパスを返します。走査順は FileSystemObject / OS の列挙順のままで、同一時刻の順序は保証しません。
Public Function GetNewestFile(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String
End Function

'* バックアップファイルを作成します。
'*
'* @param SourceFilePath [省略可] バックアップ対象のファイル パス (既定値は現在のワークブックのパス)
'* @param DestinationDirectoryPath [省略可] バックアップ先のディレクトリパス (既定値は対象ファイルのディレクトリ)
'* @param BackupSuffix [省略可] バックアップ ファイル名に追加するサフィックス (既定値はタイムスタンプ)
'* @param Force [省略可] 既存のバックアップ ファイルを強制的に上書きするかどうか
'* @param Move [省略可] バックアップ時に元のファイルを移動するかどうか (True の場合、移動)
'* @param FailIfSourceNotFound [省略可] 対象ファイルが存在しない場合にエラーを発生させるかどうか (True の場合、エラー)
'* @return 作成されたバックアップ ファイルのフルパス
'*
'* @details
'* 指定されたファイルのバックアップを作成します。バックアップ先のファイル名には元のファイル名に
'* サフィックスが追加されます。既存のファイルがある場合は `Force` の設定に基づいて処理されます。
'* `Force:=True` は既存のバックアップ ファイルを上書きします。
'* `Force:=False` かつ `BackupSuffix:=""` の既定サフィックスが衝突した場合は、既存バックアップのパスを返し、内容比較や `Move:=True` の移動は行いません。
'* `Force:=False` かつ `BackupSuffix` を明示した場合の衝突はエラーです。
Public Function CreateBackupFile( _
        Optional ByVal SourceFilePath As String = "", _
        Optional ByVal DestinationDirectoryPath As String = "", _
        Optional ByVal BackupSuffix As String = "", _
        Optional ByVal Force As Boolean = False, _
        Optional ByVal Move As Boolean = False, _
        Optional ByVal FailIfSourceNotFound As Boolean = True) As String
End Function
