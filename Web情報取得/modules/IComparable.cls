VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparable"
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
'! [インターフェイス] 比較可能なオブジェクトのインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IComparable", Description:="このクラスはインターフェイス用です。"
End Sub

'* オブジェクト間の比較を行います。
'*
'* @param OtherObject 比較対象のオブジェクト。
'* @return 比較結果を示す整数値。
'*
'* @details
'* オブジェクト間の比較を行い、比較結果を整数で返します｡
'* - 返り値が正の数：呼び出しオブジェクトが比較対象より大きい
'* - 返り値が 0：両オブジェクトが等しい
'* - 返り値が負の数：呼び出しオブジェクトが比較対象より小さい
'*
'* なお、ここでいう「等しい」とは「整列時に大きくも小さくもない」という意味合いであり、同一性や同等性とは必ずしも一致しない概念です。
'* 同一性、同等性に関しては IEquatable クラスの説明を参照してください。
Public Function CompareTo(ByVal OtherObject As Object) As Integer
End Function

'* OtherObject より大きいかどうかを判断します。
'*
'* @param OtherObject 比較対象のオブジェクト。
'* @return ブール値。True の場合、呼び出しオブジェクトが比較対象より大きい。
'*
'* @details
'* CompareTo 関数の戻り値が正の場合に True を返します。
Public Function IsGreaterThan(ByVal OtherObject As Object) As Boolean
End Function

'* OtherObject より小さいかどうかを判断します。
'*
'* @param OtherObject 比較対象のオブジェクト。
'* @return ブール値。True の場合、呼び出しオブジェクトが比較対象より小さい。
'*
'* @details
'* CompareTo 関数の戻り値が負の場合に True を返します。
Public Function IsLessThan(ByVal OtherObject As Object) As Boolean
End Function

'* OtherObject と等しいかどうかを判断します。
'*
'* @param OtherObject 比較対象のオブジェクト。
'* @return ブール値。True の場合、呼び出しオブジェクトが比較対象と等しい。
'*
'* @details
'* CompareTo 関数の戻り値が 0 の場合に True を返します。
'* ここでいう「等しい」とは「整列時に大きくも小さくもない」という意味合いであり、同一性や同等性とは必ずしも一致しない概念です。
Public Function IsEqualTo(ByVal OtherObject As Object) As Boolean
End Function
