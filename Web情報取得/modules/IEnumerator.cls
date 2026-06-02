VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnumerator"
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
'! コレクションを列挙するための列挙子のインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IEnumerator", Description:="このクラスはインターフェイス用です。"
End Sub

'* 現在の要素
Public Property Get Current() As Variant
End Property

'* 現在のインデックス
Public Property Get Index() As Long
End Property

'* 列挙対象のコレクション
Public Property Get Target() As Variant
End Property

'* 逆順かどうか
Public Property Get Descending() As Boolean
End Property

'* 読み取り専用かどうか
Public Property Get IsReadOnly() As Boolean
End Property

'* 次の要素に移動します。
'*
'* @return Boolean 次の要素が存在する場合は True、それ以外は False。
'*
'* @details
'* 現在の位置を次の要素に移動します。リストの終端に達した場合は False を返します。
'* 初期値はコレクションの先頭の 1 つ前であり、1 回呼び出すことで最初の要素を指すことになります。
Public Function MoveNext() As Boolean
End Function

'* 列挙をリセットします。
'*
'* @details
'* 列挙の現在位置をリストの先頭の 1 つ前に戻します。
Public Sub Reset()
End Sub

'* 指定した位置までスキップします。
'*
'* @param TargetIndex 移動先のインデックス。
'*
'* @details
'* 指定されたインデックスまで現在位置を進めます。現在位置より前やリストの範囲外を指定するとエラーを発生させます。
Public Sub SkipTo(ByVal TargetIndex As Long)
End Sub

'* 現在の要素を削除します。
'*
'* @return Variant 削除された要素。
'*
'* @details
'* 現在の要素を削除します。内部的には、対象コレクションの Remove(ItemIndex) メソッドが呼び出されます。
'* この操作を行うと、この列挙子が示す位置は一つ前に戻ります。
'* 順序を保証しないコレクションに対して使用した場合、この操作後の列挙子の動作は保証されません。
'* IsReadOnly が True の場合はエラーを発生させます。
Public Function Remove() As Variant
End Function

'* 現在の要素を更新します。
'*
'* @param ItemObject 更新する要素の値。
'*
'* @details
'* 現在の要素を更新します。内部的には、対象コレクションの Update(ItemIndex, ItemObject) メソッドが呼び出されます。
'* 配列列挙子など、実装が書き込みをサポートしない場合はエラーを発生させます。
'* 順序を保証しないコレクションに対して使用した場合、この操作後の列挙子の動作は保証されません。
'* IsReadOnly が True の場合はエラーを発生させます。
Public Sub Update(ByVal ItemObject As Variant)
End Sub

