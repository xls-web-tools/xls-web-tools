VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! Variant 配列を公開 API として最小限に包む低レベル配列ラッパーです。
'!
' #############################################################################

Private pArray() As Variant

'* 現在保持している配列の下限。
Public Property Get LowerBound() As Long
    LowerBound = LBound(pArray)
End Property

'* 現在保持している配列の上限。
Public Property Get UpperBound() As Long
    UpperBound = UBound(pArray)
End Property

'* このコレクションのインデクサ。境界確認は行わず、VBA の配列アクセスと同じエラーを露出します。
Public Property Get Item(ByVal ItemIndex As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If IsObject(pArray(ItemIndex)) Then
        Set Item = pArray(ItemIndex)
    Else
        Item = pArray(ItemIndex)
    End If
End Property
Public Property Let Item(ByVal ItemIndex As Long, ByVal ItemObject As Variant)
    Call Update(ItemIndex, ItemObject)
End Property
Public Property Set Item(ByVal ItemIndex As Long, ByVal ItemObject As Object)
    Call Update(ItemIndex, ItemObject)
End Property

Private Sub Class_Initialize()
    pArray = EmptyVariantArray()
End Sub

'Private Sub Class_Terminate()
'    On Error Resume Next
'    Debug.Print "Terminate: ArrayObject(0x" & Hex(ObjPtr(Me)) & ")"
'
'    Dim idx As Long
'    For idx = LBound(pArray) To UBound(pArray)
'        pArray(idx) = Empty
'    Next idx
'End Sub

'* 配列の境界を再設定します。
'*
'* @param LowerBound 新しい下限。
'* @param UpperBound 新しい上限。
'* @param PreserveItems 既存要素を保持するかどうか。
'*
'* @details
'* 生成直後は `EmptyVariantArray()` 由来の 0 ベース空配列 (`LBound = 0` / `UBound = -1`) を保持します。
'* `PreserveItems:=False` の場合は VBA の `ReDim` と同じ契約で動作します。
'* `PreserveItems:=True` の場合は VBA の `ReDim Preserve` と同じ契約で動作します。つまり既存配列の下限は変更できず、上限のみ変更できます。
Public Sub ReDimArray(ByVal LowerBound As Long, ByVal UpperBound As Long, Optional ByVal PreserveItems As Boolean = False)
    If PreserveItems Then
        ReDim Preserve pArray(LowerBound To UpperBound)
    Else
        ReDim pArray(LowerBound To UpperBound)
    End If
End Sub

'* 指定されたインデックスの要素を更新します。
'*
'* @param ItemIndex 要素のインデックス。
'* @param ItemObject 設定する要素。
'*
'* @details
'* 境界確認は行わず、VBA の配列アクセスと同じエラーを露出します。
'* オブジェクトは参照として設定し、それ以外の値は値として設定します。
Public Sub Update(ByVal ItemIndex As Long, ByVal ItemObject As Variant)
    If IsObject(ItemObject) Then
        Set pArray(ItemIndex) = ItemObject
    Else
        pArray(ItemIndex) = ItemObject
    End If
End Sub
