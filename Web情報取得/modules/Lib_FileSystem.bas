Attribute VB_Name = "Lib_FileSystem"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! FileSystemService の共通インスタンスを管理するモジュールです。
'!
' #############################################################################

'* FileSystemService。ユニットテスト時にはテスト ダブルに置き換えてください。
Public FsSrv As IFileSystemService

'* FileSystemService を初期化します。
'*
'* @param Force [省略可] True の場合は差し替え済みサービスを本番サービスへ再生成します。
'*
'* @details
'* FsSrv が未設定の場合、または Force が True の場合に FileSystemService を生成します。テストで差し替え済みの場合は Force が False のときだけ維持します。
Public Sub InitializeFileSystemService(Optional ByVal Force As Boolean = False)
    If Force Or (FsSrv Is Nothing) Then Set FsSrv = New FileSystemService
End Sub
