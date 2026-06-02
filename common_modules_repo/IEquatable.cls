VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEquatable"
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
'! このオブジェクトと他のオブジェクトが同一であるかを確認できることを示すインターフェイスです。
'!
'! すべてのプロパティが一致しているような同一性ではなく、付帯的な情報を除く重要な性質は一致しているというような同等性を確認する場合は、このインターフェイスを実装するのではなく、IsEquivalentTo というメンバ関数を持たせるようにしてください。
'! 集合の要素のような重複確認は IDuplicateCheckable を実装してください。
'!
'! 同等性に関しては、多態性を利用した場面が想定されていない (基本、同じコンクリート クラス同士での比較となる) ため、現状ではインターフェイスは用意していません。
'! 同一性も同等性も重複確認も、すべて同じ判断基準であれば、IsEquivalentTo() および IsDuplicateOf() から Equals() を呼び出すような実装にしてかまいません。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IEquatable", Description:="このクラスはインターフェイス用です。"
End Sub

'* このオブジェクトと他のオブジェクトが同一であるかを返します。
'*
'* @param OtherObject 他のオブジェクト。
'*
'* @details
'* このオブジェクトと他のオブジェクトが同一であるかを返します。
'* 特に理由がなければ、他のオブジェクトの型がこのオブジェクトと同じかチェックしたのち、
'* このオブジェクトと他のオブジェクトの GetIdentityString() を = で比較した結果を返します。
Public Function Equals(ByVal OtherObject As Object) As Boolean
End Function

'* 同一性をチェックするための文字列を返します。
'*
'* @details
'* 同一性をチェックするための文字列を返します。等しいと判断するオブジェクト間では同じ文字列を、異なると判断するオブジェクト間では異なる文字列を返すようにしてください。
'* 例えば、すべてのメンバ変数を連結した文字列を返すような実装をすることになります。
Public Function GetIdentityString() As String
End Function


