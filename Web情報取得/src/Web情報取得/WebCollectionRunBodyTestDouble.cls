VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebCollectionRunBodyTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IWebCollectionRunBody

' #############################################################################
'!
'! @brief
'! ブラウザ取得 run body のテストダブルです。
'!
' #############################################################################

'* スタブ設定と正常に成立した呼び出し履歴。
Public Store As TestDoubleBehaviorStore

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* 一覧画面到達後の処理を実行します。
'*
'* @details
'* 一覧画面到達後の処理を実行します。
Public Sub RunAfterPreparation()
    Call Store.RaiseIfError("RunAfterPreparation")
    Call Store.RecordCall("RunAfterPreparation")
End Sub
Private Sub IWebCollectionRunBody_RunAfterPreparation()
    Call RunAfterPreparation
End Sub

'* エラー時に一覧画面到達後処理側の進捗表示を解除します。
'*
'* @details
'* エラー時に一覧画面到達後処理側の進捗表示を解除します。
Public Sub ClearProgressAfterError()
    Call Store.RaiseIfError("ClearProgressAfterError")
    Call Store.RecordCall("ClearProgressAfterError")
End Sub
Private Sub IWebCollectionRunBody_ClearProgressAfterError()
    Call ClearProgressAfterError
End Sub
