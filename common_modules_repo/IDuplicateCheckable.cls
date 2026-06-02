VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDuplicateCheckable"
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
'! 他のオブジェクトと重複しているかを確認できることを示すインターフェイスです。
'!
'! このオブジェクトが他のオブジェクトと重複しているかを GetKey() に基づきチェックします。
'! 特に理由がない限り、このオブジェクトの GetKey() と他のオブジェクトの GetKey() を呼び出し、文字列が一致するかを結果として返してください。
'!
'! すべてのプロパティが一致しているような同一性ではなく、付帯的な情報を除く重要な性質は一致しているというような同等性を確認する場合は、このインターフェイスを実装するのではなく、IsEquivalentTo というメンバ関数を持たせるようにしてください。
'! 集合の要素のような重複確認は IDuplicateCheckable を実装してください。
'!
'! 同等性に関しては、多態性を利用した場面が想定されていない (基本、同じコンクリート クラス同士での比較となる) ため、現状ではインターフェイスは用意していません。
'! 同一性も同等性も重複確認も、すべて同じ判断基準であれば、IsEquivalentTo() および IsDuplicateOf() から Equals() を呼び出すような実装にしてかまいません。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IDuplicateCheckable", Description:="このクラスはインターフェイス用です。"
End Sub

'* 他のオブジェクトとの重複を確認します。
'*
'* @param OtherObject 比較対象のオブジェクト。
'* @return 重複しているか否か。
'*
'* @details
'* このオブジェクトと他のオブジェクトが重複するかどうかを返します。
'* 特に理由がなければ、他のオブジェクトの型がこのオブジェクトと同じかチェックしたのち、
'* このオブジェクトと他のオブジェクトの GetKey() を = で比較した結果を返します。
Public Function IsDuplicateOf(ByVal OtherObject As Object) As Boolean
End Function

'* オブジェクトの重複を制御するための文字列表現を返します。
'*
'* @details
'* オブジェクトの重複を制御するための文字列表現を返します。
Public Function GetKey() As String
End Function
