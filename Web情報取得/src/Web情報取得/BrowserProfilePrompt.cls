VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BrowserProfilePrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IBrowserProfilePrompt

' #############################################################################
'!
'! @brief
'! 専用ブラウザプロファイル フォルダー作成を利用者に確認します。
'!
' #############################################################################

'* 専用ブラウザプロファイル フォルダーを作成して続行するか確認します。
'*
'* @details
'* 専用ブラウザプロファイル フォルダーを作成して続行するか確認します。
Public Function ConfirmCreateDirectory(ByVal DirectoryPath As String) As Boolean
    Dim result_value As VbMsgBoxResult
    result_value = MsgBox( _
            "専用ブラウザプロファイル フォルダーが存在しません。" & vbCrLf & _
            "作成して処理を続行しますか?" & vbCrLf & vbCrLf & _
            DirectoryPath, _
            vbYesNo + vbQuestion + vbDefaultButton2, _
            "専用ブラウザプロファイル フォルダーの作成")

    ConfirmCreateDirectory = (result_value = vbYes)
End Function
Private Function IBrowserProfilePrompt_ConfirmCreateDirectory(ByVal DirectoryPath As String) As Boolean
    IBrowserProfilePrompt_ConfirmCreateDirectory = ConfirmCreateDirectory(DirectoryPath)
End Function
