VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebCollectionRunBody"
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
'! ブラウザ取得 run lifecycle の一覧画面到達後処理です。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IWebCollectionRunBody", Description:="このクラスはインターフェイス用です。"
End Sub

'* 一覧画面到達後の処理を実行します。
Public Sub RunAfterPreparation()
End Sub

'* エラー時に一覧画面到達後処理側の進捗表示を解除します。
Public Sub ClearProgressAfterError()
End Sub
