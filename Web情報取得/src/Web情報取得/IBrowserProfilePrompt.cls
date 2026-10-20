VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBrowserProfilePrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'#Interface

' #############################################################################
'!
'! @brief
'! 専用ブラウザプロファイル フォルダー作成確認のインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IBrowserProfilePrompt", Description:="このクラスはインターフェイス用です。"
End Sub

'* 専用ブラウザプロファイル フォルダーを作成して続行するか確認します。
'*
'* @param DirectoryPath 作成確認対象の解決済み絶対パス。
'* @return 作成して続行する場合 True。
'*
'* @details
'* 専用ブラウザプロファイル フォルダーを作成して続行するか確認します。
Public Function ConfirmCreateDirectory(ByVal DirectoryPath As String) As Boolean
End Function
