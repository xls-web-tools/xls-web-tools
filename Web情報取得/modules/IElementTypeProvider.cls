VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IElementTypeProvider"
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
'! オブジェクトをコレクション内で扱う要素型を自己申告できることを示すインターフェイスです。
'!
'! 型付き要素コレクションは、このインターフェイスを実装するオブジェクトから ElementTypeKey を読み取り、そのオブジェクトをどの要素型契約として扱うかを決定します。
'! ElementTypeKey が実際に代入可能な型を表すことは、実装クラス側の責任です。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IElementTypeProvider", Description:="このクラスはインターフェイス用です。"
End Sub

'* 要素型契約名。
'*
'* @return 要素型契約名。
'*
'* @details
'* クラス モジュール名に使用可能な文字だけで構成された 31 文字以内の名前を返します。
'* 先頭文字は英字、2 文字目以降は英字、数字、アンダースコアを使用できます。
'* この名前が実際に代入可能な型を表すことは、実装クラス側の責任です。
Public Property Get ElementTypeKey() As String
End Property
