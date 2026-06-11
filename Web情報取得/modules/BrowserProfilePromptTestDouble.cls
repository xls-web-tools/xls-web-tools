VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BrowserProfilePromptTestDouble"
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
'! 専用ブラウザプロファイル フォルダー作成確認のテストダブルです。
'!
' #############################################################################

'* スタブ設定と正常に成立した呼び出し履歴。
Public Store As TestDoubleBehaviorStore

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* 専用ブラウザプロファイル フォルダーを作成して続行するか確認します。
Public Function ConfirmCreateDirectory(ByVal DirectoryPath As String) As Boolean
    Call Store.RaiseIfError("ConfirmCreateDirectory", DirectoryPath)

    Dim result_value As Boolean
    result_value = Store.GetReturn("ConfirmCreateDirectory", DirectoryPath)
    Call Store.RecordCall("ConfirmCreateDirectory", DirectoryPath)

    ConfirmCreateDirectory = result_value
End Function
Private Function IBrowserProfilePrompt_ConfirmCreateDirectory(ByVal DirectoryPath As String) As Boolean
    IBrowserProfilePrompt_ConfirmCreateDirectory = ConfirmCreateDirectory(DirectoryPath)
End Function
