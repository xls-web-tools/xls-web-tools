VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonRunStateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! 1 回の GUI 実行で使用する共通実行状態を管理するクラスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Set DbgInfo = New DebugInformation
    Set ProgStat = New ProgressStatus
End Sub

'* 共通実行状態を明示的に初期化します。
'*
'* @details
'* DbgInfo と ProgStat に新規オブジェクトを割り当てます。
Public Sub Initialize()
    Set DbgInfo = New DebugInformation
    Set ProgStat = New ProgressStatus
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Call Clear
    On Error GoTo 0
End Sub

'* 共通実行状態を明示的に破棄します。
'*
'* @details
'* DbgInfo と ProgStat を Nothing に戻し、次の GUI 実行へ状態を持ち越さないようにします。
Public Sub Clear()
    Set DbgInfo = Nothing
    Set ProgStat = Nothing
End Sub
