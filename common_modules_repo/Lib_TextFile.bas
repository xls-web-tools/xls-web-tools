Attribute VB_Name = "Lib_TextFile"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! TextFileService の共通インスタンスを管理するモジュールです。
'!
' #############################################################################

'* TextFileService。ユニットテスト時にはテスト ダブルに置き換えてください。
Public TfSrv As ITextFileService

'* TextFileService を初期化します。
'*
'* @param Force [省略可] True の場合は差し替え済みサービスを本番サービスへ再生成します。
'*
'* @details
'* TfSrv が未設定の場合、または Force が True の場合に TextFileService を生成します。テストで差し替え済みの場合は Force が False のときだけ維持します。
Public Sub InitializeTextFileService(Optional ByVal Force As Boolean = False)
    If Force Or (TfSrv Is Nothing) Then Set TfSrv = New TextFileService
End Sub
