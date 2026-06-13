VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! ダイナミックな型とサイズのオブジェクトリストを管理するクラスです。
'! このクラスは、オブジェクトおよびプリミティブ型の動的配列を提供し、
'! 型の一貫性、オブジェクト操作、および拡張機能をサポートします。
'!
' #############################################################################

Private Const C_CHUNK_SIZE As Integer = 16

Private pDepth As Long
Private pRootArray As ArrayObject
Private pCount As Long
Private pTypeName As String
Private pElementTypeName As String
Private pHasItemTypeContract As Boolean
Private pRequireComparable As Boolean
Private pObjectKeyMode As Long
Private pIsObject As Boolean
Private pIsComparable As Boolean
Private pIsStringable As Boolean
Private pIsEquatable As Boolean
Private pIsDuplicateCheckable As Boolean
Private pNewEnumSnapshot As Collection
Private pNewEnumSnapshotVersion As Long
Private pEnumerationVersion As Long

'* リストの現在の件数。
Public Property Get Count() As Long
    Count = pCount
End Property

'* リスト内の要素の型。
Public Property Get ItemTypeName() As String
    ItemTypeName = pTypeName
End Property

'* リストに要素型契約が設定されているか否か。
Public Property Get HasItemTypeContract() As Boolean
    HasItemTypeContract = pHasItemTypeContract
End Property

'* Initialize で指定できる要素型契約名。
Public Property Get ElementTypeName() As String
    ElementTypeName = pElementTypeName
End Property

'* 要素に IComparable 実装を要求するか否か。
Public Property Get RequireComparable() As Boolean
    RequireComparable = pRequireComparable
End Property

'* オブジェクトの同一性/重複判定モード。
Public Property Get ObjectKeyMode() As Long
    ObjectKeyMode = pObjectKeyMode
End Property

'* 要素型がオブジェクト型か否か。
Public Property Get IsObjectItemType() As Boolean
    IsObjectItemType = pIsObject
End Property

'* 要素型が IComparable として扱われるか否か。
Public Property Get IsComparableItemType() As Boolean
    IsComparableItemType = pIsComparable
End Property

'* 要素型が IStringable として扱われるか否か。
Public Property Get IsStringableItemType() As Boolean
    IsStringableItemType = pIsStringable
End Property

'* 要素型が IEquatable として扱われるか否か。
Public Property Get IsEquatableItemType() As Boolean
    IsEquatableItemType = pIsEquatable
End Property

'* 要素型が IDuplicateCheckable として扱われるか否か。
Public Property Get IsDuplicateCheckableItemType() As Boolean
    IsDuplicateCheckableItemType = pIsDuplicateCheckable
End Property

'* 標準 For Each 用の列挙子を得ます。
'*
'* @return 標準 For Each が使用する列挙子。
'*
'* @details
'* 列挙開始時点の要素スナップショットを列挙します。
'* 標準 For Each の列挙中にこのリストを更新した場合、実行中の列挙の継続は保証されません。
'* 更新や削除を伴う列挙では GetEnumerator を使用してください。
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    If pNewEnumSnapshot Is Nothing Or pNewEnumSnapshotVersion <> pEnumerationVersion Then
        Set pNewEnumSnapshot = pCreateNewEnumSnapshot()
        pNewEnumSnapshotVersion = pEnumerationVersion
    End If

    Set NewEnum = pNewEnumSnapshot.[_NewEnum]
End Property

'* 指定されたインデックスの要素を返します。
'*
'* @param ItemIndex 要素のインデックス。
'* @return 指定されたインデックスにある要素。
'*
'* @details
'* 指定されたインデックスの要素を返します。
Public Function Item(ByVal ItemIndex As Long) As Variant
    Call pCheckIndex(ItemIndex)

    Dim index_arr() As Long
    Dim remain_index As Long
    Call pGetIndexArray(index_arr, remain_index, ItemIndex)

    Dim target_arr As ArrayObject
    Set target_arr = pRootArray

    Dim item_idx As Long
    For item_idx = LBound(index_arr) To UBound(index_arr) - 1
        Set target_arr = target_arr.Item(index_arr(item_idx))
    Next item_idx

    If pIsObject Then
        Set Item = target_arr.Item(index_arr(UBound(index_arr)))
    Else
        Item = target_arr.Item(index_arr(UBound(index_arr)))
    End If
End Function

Private Sub Class_Initialize()
    Call pInitializeCore
End Sub

Private Sub pInitializeCore()
    pDepth = 0
    pCount = 0
    pTypeName = ""
    pElementTypeName = ""
    pHasItemTypeContract = False
    pRequireComparable = False
    pObjectKeyMode = G_OBJECT_KEY_MODE_REFERENCE
    pIsObject = False
    pIsComparable = False
    pIsEquatable = False
    pIsDuplicateCheckable = False
    pIsStringable = False
    Set pRootArray = New ArrayObject
    Call pRootArray.ReDimArray(0, C_CHUNK_SIZE - 1)
    pEnumerationVersion = 0
    pNewEnumSnapshotVersion = -1
    Set pNewEnumSnapshot = Nothing
End Sub

'* リストの要素型契約を初期化します。
'*
'* @param ElementTypeName 要素型契約名。
'* @param RequireComparable 要素に IComparable 実装を要求するか否か。
'* @param ObjectKeyMode オブジェクトの同一性/重複判定モード。
'*
'* @details
'* オブジェクト型の要素だけを受け入れるリストとして初期化します。
'* ElementTypeName は、自己申告された ElementTypeKey または TypeName と比較されます。
Public Sub Initialize( _
        ByVal ElementTypeName As String, _
        Optional ByVal RequireComparable As Boolean = False, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE)

    If pCount <> 0 Or pHasItemTypeContract Then
        Err.Raise vbObjectError + 1, "Class ObjectList.Initialize", "初期化済みのリストを再初期化することはできません。"
    End If
    If Not IsValidElementTypeKey(ElementTypeName) Then
        Err.Raise vbObjectError + 1, "Class ObjectList.Initialize", "ElementTypeName はクラス モジュール名に使用可能な文字で指定してください。(" & ElementTypeName & ")"
    End If

    Call pCheckObjectKeyMode(ObjectKeyMode)
    pElementTypeName = ElementTypeName
    pTypeName = "Object@" & ElementTypeName
    pHasItemTypeContract = True
    pRequireComparable = RequireComparable
    pObjectKeyMode = ObjectKeyMode
    pIsObject = True
    pIsComparable = RequireComparable
    pIsEquatable = (ObjectKeyMode = G_OBJECT_KEY_MODE_I_EQUATABLE)
    pIsDuplicateCheckable = (ObjectKeyMode = G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE)
End Sub

'* 他のリストと同じ要素型状態で初期化します。
'*
'* @param Other 初期化元のリスト。
'*
'* @details
'* 他のリストが管理している要素型状態をこのリストへ引き継ぎます。
Public Sub InitializeFromOther(ByVal Other As ObjectList)
    If pCount <> 0 Or pTypeName <> "" Or pHasItemTypeContract Then
        Err.Raise vbObjectError + 1, "Class ObjectList.InitializeFromOther", "初期化済みのリストを再初期化することはできません。"
    End If

    pTypeName = Other.ItemTypeName
    pElementTypeName = Other.ElementTypeName
    pHasItemTypeContract = Other.HasItemTypeContract
    pRequireComparable = Other.RequireComparable
    pObjectKeyMode = Other.ObjectKeyMode
    pIsObject = Other.IsObjectItemType
    pIsComparable = Other.IsComparableItemType
    pIsStringable = Other.IsStringableItemType
    pIsEquatable = Other.IsEquatableItemType
    pIsDuplicateCheckable = Other.IsDuplicateCheckableItemType
End Sub

'* 他の集合と同じ要素型状態で初期化します。
'*
'* @param OtherSet 初期化元の集合。
'*
'* @details
'* 他の集合が管理している要素型状態をこのリストへ引き継ぎます。
Public Sub InitializeFromSet(ByVal OtherSet As ObjectSet)
    If pCount <> 0 Or pTypeName <> "" Or pHasItemTypeContract Then
        Err.Raise vbObjectError + 1, "Class ObjectList.InitializeFromSet", "初期化済みのリストを再初期化することはできません。"
    End If

    pTypeName = OtherSet.ItemTypeName
    pElementTypeName = OtherSet.ElementTypeName
    pHasItemTypeContract = OtherSet.HasItemTypeContract
    pRequireComparable = OtherSet.RequireComparable
    pObjectKeyMode = OtherSet.ObjectKeyMode
    pIsObject = OtherSet.IsObjectItemType
    pIsComparable = OtherSet.IsComparableItemType
    pIsStringable = OtherSet.IsStringableItemType
    pIsEquatable = OtherSet.IsEquatableItemType
    pIsDuplicateCheckable = OtherSet.IsDuplicateCheckableItemType
End Sub

Private Sub pCheckObjectKeyMode(ByVal ObjectKeyMode As Long)
    Select Case ObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE, G_OBJECT_KEY_MODE_I_EQUATABLE, G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
        Case Else
            Err.Raise vbObjectError + 1, "Class ObjectList.pCheckObjectKeyMode", "未対応のオブジェクト キー化モードです。(" & CStr(ObjectKeyMode) & ")"
    End Select
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    'Debug.Print "Terminate: ObjectList(0x" & Hex(ObjPtr(Me)) & ")"

    Dim item_idx As Long
    For item_idx = 0 To pRootArray.UpperBound - 1
        Call pRootArray.Update(item_idx, Empty)
    Next item_idx
End Sub

'* 指定されたインデックスの要素を更新します。
'*
'* @param ItemIndex 更新する要素のインデックス。
'* @param ItemObject 新しい要素。
'*
'* @details
'* 指定されたインデックスの要素を更新します。
Public Sub Update(ByVal ItemIndex As Long, ByVal ItemObject As Variant)
    Call pCheckIndex(ItemIndex)
    If pTypeName = "" Then
        Call pInitializeItemType(ItemObject)
    Else
        Call pCheckItemType(ItemObject)
    End If

    Call pUpdateCore(ItemIndex, ItemObject)
End Sub

'* リストを配列として返します。
'*
'* @return 要素の配列。
'*
'* @details
'* リストを配列として返します。要素数 0 の場合は、0 ベースの空配列を返します。
Public Function ConvertToArray() As Variant()
    If pCount = 0 Then
        ConvertToArray = EmptyVariantArray()
        Exit Function
    End If

    Dim result_value() As Variant
    Dim item_idx As Long

    ReDim result_value(0 To pCount - 1)

    For item_idx = LBound(result_value) To UBound(result_value)
        If pIsObject Then
            Set result_value(item_idx) = Item(item_idx)
        Else
            result_value(item_idx) = Item(item_idx)
        End If
    Next item_idx

    ConvertToArray = result_value
End Function

'* リストを文字列の配列として返します。
'*
'* @return 文字列の配列。
'*
'* @details
'* リストを文字列の配列として返します。要素数 0 の場合は、0 ベースの空配列を返します。
Public Function ConvertToStringArray() As String()
    If pCount = 0 Then
        ConvertToStringArray = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    Dim item_idx As Long
    Dim stringable_obj As IStringable

    ReDim result_value(0 To pCount - 1)

    For item_idx = LBound(result_value) To UBound(result_value)
        If pIsObject Then
            If Item(item_idx) Is Nothing Then
                result_value(item_idx) = GetTypedValueKey(Item(item_idx))
            Else
                If TypeOf Item(item_idx) Is IStringable Then
                    Set stringable_obj = Item(item_idx)
                    result_value(item_idx) = stringable_obj.ToString()
                Else
                    result_value(item_idx) = GetTypedValueKey(Item(item_idx))
                End If
            End If
        Else
            If IsError(Item(item_idx)) Then
                result_value(item_idx) = ExcelErrorToString(Item(item_idx))
            ElseIf IsArray(Item(item_idx)) Then
                result_value(item_idx) = GetTypedValueKey(Item(item_idx))
            Else
                result_value(item_idx) = CStr(Item(item_idx))
            End If
        End If
    Next item_idx

    ConvertToStringArray = result_value
End Function

'* 新しい要素をリストに追加します。
'*
'* @param ItemObject 追加する要素。
'*
'* @details
'* 新しい要素をリストに追加します。
Public Sub Add(ByVal ItemObject As Variant)
    If pTypeName = "" Then
        Call pInitializeItemType(ItemObject)
    Else
        Call pCheckItemType(ItemObject)
    End If

    Dim new_arr As ArrayObject
    Dim index_arr() As Long
    Dim remain_index As Long
    Call pGetIndexArray(index_arr, remain_index, pCount)
    Call pExpandIfNeed(remain_index)

    Dim target_arr As ArrayObject
    If 0 < remain_index Then
        Set target_arr = pGetChildArray(pRootArray, remain_index)
    Else
        Set target_arr = pRootArray
    End If

    Dim item_idx As Long
    For item_idx = LBound(index_arr) To UBound(index_arr) - 1
        Set target_arr = pGetChildArray(target_arr, index_arr(item_idx))
    Next item_idx

    Call target_arr.Update(index_arr(UBound(index_arr)), ItemObject)

    pCount = pCount + 1
    Call pIncrementEnumerationVersion
End Sub

Private Sub pIncrementEnumerationVersion()
    pEnumerationVersion = pEnumerationVersion + 1
End Sub

Private Function pGetChildArray(ByVal ParentArray As ArrayObject, ByVal ChildIndex As Long) As ArrayObject
    If IsObject(ParentArray.Item(ChildIndex)) Then
        Set pGetChildArray = ParentArray.Item(ChildIndex)
    Else
        Dim new_root As ArrayObject
        Set new_root = New ArrayObject
        Call new_root.ReDimArray(0, C_CHUNK_SIZE - 1)
        Call ParentArray.Update(ChildIndex, new_root)
        Set pGetChildArray = new_root
    End If
End Function

Private Sub pExpandIfNeed(ByVal RemainIndex As Long)
    If RemainIndex <= 0 Then Exit Sub

    Dim new_root As ArrayObject
    Set new_root = New ArrayObject
    Call new_root.ReDimArray(0, C_CHUNK_SIZE - 1)
    Call new_root.Update(0, pRootArray)
    Set pRootArray = new_root
    pDepth = pDepth + 1
End Sub

'* 配列の各要素をリストに追加します。
'*
'* @param ItemArray 追加する要素の配列。
'*
'* @details
'* 配列の各要素をリストに追加します。
Public Sub AddArray(ByVal ItemArray As Variant)
    If Not IsArray(ItemArray) Then Err.Raise vbObjectError + 1, "Class ObjectList.AddArray", "引数が配列ではありません"
    If IsEmptyArray(ItemArray) Then Exit Sub

    Dim item_obj As Variant

    For Each item_obj In ItemArray
        Call Add(item_obj)
    Next item_obj
End Sub

'* 他のリストの要素をこのリストに取り込みます。
'*
'* @param OtherList 取り込む要素を持つリスト。
'*
'* @details
'* 他のリストの要素をこのリストに取り込みます。
Public Sub AddOther(ByVal OtherList As ObjectList)
    If pTypeName = "" And OtherList.HasItemTypeContract Then
        Call InitializeFromOther(OtherList)
    Else
        Call pPrepareToAddObjectContract( _
                OtherList.HasItemTypeContract, _
                OtherList.ElementTypeName, _
                OtherList.RequireComparable, _
                OtherList.ObjectKeyMode)
    End If

    Dim item_idx As Long
    For item_idx = 0 To OtherList.Count - 1
        Call Me.Add(OtherList.Item(item_idx))
    Next item_idx
End Sub

'* 他のセットの要素をこのリストに取り込みます。
'*
'* @param OtherSet 取り込む要素を持つセット。
'*
'* @details
'* 他のセットの要素をこのリストに取り込みます。
Public Sub AddSet(ByVal OtherSet As ObjectSet)
    If pTypeName = "" And OtherSet.HasItemTypeContract Then
        Call InitializeFromSet(OtherSet)
    Else
        Call pPrepareToAddObjectContract( _
                OtherSet.HasItemTypeContract, _
                OtherSet.ElementTypeName, _
                OtherSet.RequireComparable, _
                OtherSet.ObjectKeyMode)
    End If

    Dim item_idx As Long
    For item_idx = 0 To OtherSet.Count - 1
        Call Me.Add(OtherSet.Item(item_idx))
    Next item_idx
End Sub

'* このリストの複製を生成します。
'*
'* @return 複製された新しいリスト。
'*
'* @details
'* このリストの複製を生成します。
'* リストの要素はもとのインスタンスがそのまま利用されます。(シャロー コピー)
Public Function CopyList() As ObjectList
    Dim item_idx As Long
    Dim new_list As ObjectList

    Set new_list = New ObjectList
    Call new_list.InitializeFromOther(Me)

    For item_idx = 0 To pCount - 1
        Call new_list.Add(Item(item_idx))
    Next item_idx

    Set CopyList = new_list
End Function

'* 指定されたインデックスの要素を削除します。
'*
'* @param ItemIndex 削除する要素のインデックス。
'* @return 削除された要素。
'*
'* @details
'* 指定されたインデックスの要素を削除します。
Public Function Remove(ByVal ItemIndex As Long) As Variant
    Dim item_idx As Long
    Dim old_obj As Variant

    Call pCheckIndex(ItemIndex)

    If pIsObject Then
        Set old_obj = Item(ItemIndex)
    Else
        old_obj = Item(ItemIndex)
    End If

    For item_idx = ItemIndex To pCount - 2
        Call pUpdateCore(item_idx, Item(item_idx + 1))
    Next item_idx

    Call pUpdateCore(pCount - 1, Empty)

    pCount = pCount - 1

    If pIsObject Then
        Set Remove = old_obj
    Else
        Remove = old_obj
    End If
End Function

'* このリストの内容をすべて削除します。
'*
'* @details
'* このリストの内容をすべて削除します。要素の型状態は保持します。
Public Sub RemoveAll()
    Dim item_idx As Long
    For item_idx = 0 To pCount - 1
        Call pUpdateCore(item_idx, Empty)
    Next item_idx

    pCount = 0
End Sub

'* このリストに登録されている、引数と同じ要素を削除します。
'*
'* @param ItemObject 削除対象の要素。
'* @return 削除が成功したかどうかを示すブール値。
'*
'* @details
'* このリストに登録されている、引数と同じ要素を削除します。
'* - このリストの要素が IDuplicateCheckable であれば、IsDuplicateOf の結果に基づいて処理します。
'* - このリストの要素が IEquatable であれば、Equals の結果に基づいて処理します。
'* - IEquotable かつ IDuplicateCheckable の場合は、IDuplicateCheckable としての扱いが優先されます。
Public Function RemoveItem(ByVal ItemObject As Variant) As Boolean
    Dim result_value As Boolean
    result_value = False

    If Not pCanSearchItem(ItemObject) Then
        RemoveItem = result_value
        Exit Function
    End If

    Dim item_idx As Long
    item_idx = 0

    Do While item_idx < pCount
        If pCheckRemoveItem(Item(item_idx), ItemObject) Then
            Call Remove(item_idx)
            result_value = True
        Else
            item_idx = item_idx + 1
        End If
    Loop

    RemoveItem = result_value
End Function

Private Function pCheckRemoveItem(ByVal ItemObject1 As Variant, ByVal ItemObject2 As Variant) As Boolean
    pCheckRemoveItem = pItemsEqual(ItemObject1, ItemObject2)
End Function

'* このリストに重複して登録されている要素を削除します。
'*
'* @return 重複削除が実行されたかどうかを示すブール値。
'*
'* @details
'* このリストに重複して登録されている要素を削除します。
Public Function RemoveDuplicate() As Boolean
    Dim item_idx As Long
    Dim search_idx As Long
    Dim is_removed As Boolean

    is_removed = False
    Do While item_idx < pCount
        Dim search_item As Variant
        If pIsObject Then
            Set search_item = Item(item_idx)
        Else
            search_item = Item(item_idx)
        End If

        search_idx = item_idx + 1
        Do While search_idx < pCount
            If pItemsEqual(search_item, Item(search_idx)) Then
                Call Remove(search_idx)
                is_removed = True
            Else
                search_idx = search_idx + 1
            End If
        Loop
        item_idx = item_idx + 1
    Loop

    RemoveDuplicate = is_removed
End Function

Private Function pItemsEqual(ByVal ItemObject1 As Variant, ByVal ItemObject2 As Variant) As Boolean
    If pIsObject Then
        If Not IsObject(ItemObject1) Or Not IsObject(ItemObject2) Then Exit Function

        If ItemObject1 Is Nothing Then
            pItemsEqual = (ItemObject2 Is Nothing)
            Exit Function
        End If

        If ItemObject2 Is Nothing Then Exit Function

        If ItemObject1 Is ItemObject2 Then
            pItemsEqual = True
            Exit Function
        End If

        If pIsDuplicateCheckable Then
            Dim duplicate_item As IDuplicateCheckable
            Set duplicate_item = ItemObject1
            pItemsEqual = duplicate_item.IsDuplicateOf(ItemObject2)
        ElseIf pIsEquatable Then
            Dim equatable_item As IEquatable
            Set equatable_item = ItemObject1
            pItemsEqual = equatable_item.Equals(ItemObject2)
        End If
    Else
        If IsObject(ItemObject1) Or IsObject(ItemObject2) Then Exit Function
        If IsArray(ItemObject1) Or IsArray(ItemObject2) Then
            If IsArray(ItemObject1) And IsArray(ItemObject2) Then
                pItemsEqual = (GetTypedValueKey(ItemObject1) = GetTypedValueKey(ItemObject2))
                Exit Function
            End If
            pItemsEqual = False
            Exit Function
        End If
        If IsError(ItemObject1) Or IsError(ItemObject2) Then
            If IsError(ItemObject1) And IsError(ItemObject2) Then
                pItemsEqual = pErrorValuesEqual(ItemObject1, ItemObject2)
                Exit Function
            End If
            pItemsEqual = False
            Exit Function
        End If
        pItemsEqual = (ItemObject1 = ItemObject2)
    End If
End Function

'* このリストの末尾に要素を追加します。
'*
'* @param ItemObject 追加する要素。
'*
'* @details
'* このリストの末尾に要素を追加します。Add(ItemObject) と同じです。
Public Sub PushItem(ByVal ItemObject As Variant)
    Call Me.Add(ItemObject)
End Sub

'* このリストの末尾から要素を取り出します。
'*
'* @return 取り出された要素。
'*
'* @details
'* このリストの末尾から要素を取り出します。Remove(Me.Count - 1) と同じです。
Public Function PopItem() As Variant
    If pIsObject Then
        Set PopItem = Me.Remove(Me.Count - 1)
    Else
        PopItem = Me.Remove(Me.Count - 1)
    End If
End Function

'* 等しいものを探しインデックスを返します。
'*
'* @param ItemObject 探す要素。
'* @param Optional StartIndex 探索を開始するインデックス。
'* @param Optional Reverse 逆順に探索するかどうか。
'* @return 見つかった要素のインデックス。見つからない場合は -1。
'*
'* @details
'* 等しいものを探しインデックスを返します。
Public Function GetIndexByItem(ByVal ItemObject As Variant, Optional ByVal StartIndex As Long = -1, Optional ByVal Reverse As Variant = False) As Long
    Dim result_value As Long
    result_value = -1

    If Not pCanSearchItem(ItemObject) Then
        GetIndexByItem = result_value
        Exit Function
    End If

    If Count <= 0 _
            Or pCount <= StartIndex And Not Reverse Then
        GetIndexByItem = result_value
        Exit Function
    End If

    If StartIndex < 0 Then
        If Reverse Then
            StartIndex = pCount - 1
        Else
            StartIndex = 0
        End If
    End If

    If pCount <= StartIndex And Reverse Then
        StartIndex = pCount - 1
    End If

    Dim idx_step As Long
    Dim idx_limit As Long

    If Reverse Then
        idx_step = -1
        idx_limit = 0
    Else
        idx_step = 1
        idx_limit = pCount - 1
    End If

    result_value = pGetIndexByItemCore(ItemObject, StartIndex, idx_limit, idx_step)

    GetIndexByItem = result_value
End Function

'* 等しいものが含まれるかを確認します。
'*
'* @param ItemObject 探す要素。
'* @return 含まれているかどうかを示すブール値。
'*
'* @details
'* 等しいものが含まれるかを確認します。
Public Function Exists(ByVal ItemObject As Variant) As Boolean
    Dim result_value As Boolean
    result_value = False

    If Not pCanSearchItem(ItemObject) Then
        Exists = result_value
        Exit Function
    End If

    Dim item_idx As Long
    item_idx = pGetIndexByItemCore(ItemObject, 0, pCount - 1, 1)
    If 0 <= item_idx Then
        result_value = True
    Else
        result_value = False
    End If

    Exists = result_value
End Function

'* このリストの中身をソートします。
'*
'* @param Optional Descending 降順にソートするかどうか。
'*
'* @details
'* このリストの中身をソートします。
'* 通常のプリミティブ値は VBA の大小比較で処理します。
'* 配列、Excel エラー値、非 IComparable オブジェクトはキー文字列で処理します。
'* IComparable は CompareTo の結果で処理します。
'* Nothing はすべての非 Nothing オブジェクトより小さいものとして処理します。
Public Sub Sort(Optional ByVal Descending As Boolean = False)
    Call pSortCore(0, pCount - 1, Descending)
End Sub

Private Sub pSortCore(ByVal MinIndex As Long, ByVal MaxIndex As Long, ByVal Descending As Boolean)
    If MinIndex >= MaxIndex Then Exit Sub ' MinIndex が MaxIndex 以降になる場合は関数を抜ける
    Dim item_idx As Long
    Dim pos_lng As Long: pos_lng = MinIndex

    Call Swap(MinIndex, Int((MinIndex + MaxIndex) \ 2)) '配列中央の値を基準値にして先頭と入れ替え
    For item_idx = MinIndex + 1 To MaxIndex
        If pShouldComeBefore(item_idx, MinIndex, Descending) Then ' MinIndex の要素を基準値として比較していく
            pos_lng = pos_lng + 1 '基準値より小さい数があるごとにpを1つ進める
            Call Swap(item_idx, pos_lng) 'Swap関数で入れ替えを行う。
        End If
    Next
    Call Swap(MinIndex, pos_lng) 'Swap関数で入れ替え

    Call pSortCore(MinIndex, pos_lng - 1, Descending) '基準値より小さいグループを再帰処理
    Call pSortCore(pos_lng + 1, MaxIndex, Descending) '基準値より大きいグループを再帰処理
End Sub

'* 指定したインデックスの要素を入れ替えます。
'*
'* @param ItemIndex1 入れ替える要素の一つ目のインデックス。
'* @param ItemIndex2 入れ替える要素の二つ目のインデックス。
'*
'* @details
'* 指定したインデックスの要素を入れ替えます。
Public Sub Swap(ByVal ItemIndex1 As Long, ByRef ItemIndex2 As Long)
    Dim tmp_item As Variant
    If pIsObject Then
        Set tmp_item = Me.Item(ItemIndex1)
    Else
        tmp_item = Me.Item(ItemIndex1)
    End If

    Call Update(ItemIndex1, Me.Item(ItemIndex2))
    Call Update(ItemIndex2, tmp_item)
End Sub

'* このリストの列挙子を得ます。
'*
'* @param Descending [省略可] 降順で列挙するか否か。
'* @param IsReadOnly [省略可] 読み取り専用にするか否か。
'* @return 列挙子オブジェクト。
'*
'* @details
'* このリストの列挙子を得ます。
Public Function GetEnumerator(Optional ByVal Descending As Boolean = False, Optional ByVal IsReadOnly As Boolean = False) As IEnumerator
    Dim result_value As Enumerator
    Set result_value = New Enumerator
    Call result_value.Initialize(Me, Descending:=Descending, IsReadOnly:=IsReadOnly)
    Set GetEnumerator = result_value
End Function

Private Sub pPrepareToAddObjectContract( _
        ByVal SourceHasItemTypeContract As Boolean, _
        ByVal SourceElementTypeName As String, _
        ByVal SourceRequireComparable As Boolean, _
        ByVal SourceObjectKeyMode As Long)

    If Not SourceHasItemTypeContract Then Exit Sub
    If SourceElementTypeName = "" Then Exit Sub

    If Not pHasItemTypeContract Then
        Call Initialize( _
                ElementTypeName:=SourceElementTypeName, _
                RequireComparable:=SourceRequireComparable, _
                ObjectKeyMode:=SourceObjectKeyMode)
        Exit Sub
    End If

    If StrComp(pElementTypeName, SourceElementTypeName, vbTextCompare) <> 0 _
            Or pRequireComparable <> SourceRequireComparable _
            Or pObjectKeyMode <> SourceObjectKeyMode Then
        Err.Raise vbObjectError + 1, "Class ObjectList.pPrepareToAddObjectContract", "取り込み元と取り込み先の要素型契約が異なります。"
    End If
End Sub

Private Sub pUpdateCore(ByVal ItemIndex As Long, ByVal ItemObject As Variant)
    Dim index_arr() As Long
    Dim remain_index As Long
    Call pGetIndexArray(index_arr, remain_index, ItemIndex)

    Dim target_arr As ArrayObject
    Set target_arr = pRootArray

    Dim item_idx As Long
    For item_idx = LBound(index_arr) To UBound(index_arr) - 1
        Set target_arr = target_arr.Item(index_arr(item_idx))
    Next item_idx

    Call target_arr.Update(index_arr(UBound(index_arr)), ItemObject)
    Call pIncrementEnumerationVersion
End Sub

Private Function pGetIndexByItemCore(ByVal ItemObject As Variant, ByVal StartIndex As Long, ByVal IndexLimit As Long, ByVal IndexStep As Long) As Long
    Dim item_idx As Long
    For item_idx = StartIndex To IndexLimit Step IndexStep
        If pItemsEqual(Me.Item(item_idx), ItemObject) Then
            pGetIndexByItemCore = item_idx
            Exit Function
        End If
    Next item_idx

    pGetIndexByItemCore = -1
End Function

Private Function pShouldComeBefore(ByVal ItemIndex1 As Long, ByVal ItemIndex2 As Long, ByVal Descending As Boolean) As Boolean
    Dim compare_result As Long
    compare_result = pCompareItems(ItemIndex1, ItemIndex2)

    If Descending Then
        pShouldComeBefore = (compare_result > 0)
    Else
        pShouldComeBefore = (compare_result < 0)
    End If
End Function

Private Function pCompareItems(ByVal ItemIndex1 As Long, ByVal ItemIndex2 As Long) As Long
    Dim result_value As Long
    If pIsObject Then
        result_value = pCompareObjectItems(ItemIndex1, ItemIndex2)
    Else
        result_value = pCompareNonObjectItems(ItemIndex1, ItemIndex2)
    End If

    pCompareItems = result_value
End Function

Private Function pCompareObjectItems(ByVal ItemIndex1 As Long, ByVal ItemIndex2 As Long) As Long
    Dim result_value As Long
    If Me.Item(ItemIndex1) Is Nothing Then
        If Me.Item(ItemIndex2) Is Nothing Then
            result_value = 0
        Else
            result_value = -1
        End If
        pCompareObjectItems = result_value
        Exit Function
    End If

    If Me.Item(ItemIndex2) Is Nothing Then
        result_value = 1
        pCompareObjectItems = result_value
        Exit Function
    End If

    If pIsComparable Then
        Dim comparable_obj As IComparable
        Set comparable_obj = Me.Item(ItemIndex1)

        Dim other_obj As Object
        Set other_obj = Me.Item(ItemIndex2)

        result_value = pNormalizeCompareResult(comparable_obj.CompareTo(other_obj))
    Else
        result_value = pCompareStringValues(pGetObjectSortKey(Me.Item(ItemIndex1)), pGetObjectSortKey(Me.Item(ItemIndex2)))
    End If

    pCompareObjectItems = result_value
End Function

Private Function pCompareNonObjectItems(ByVal ItemIndex1 As Long, ByVal ItemIndex2 As Long) As Long
    Dim result_value As Long
    If pShouldCompareBySortKey(Me.Item(ItemIndex1)) Or pShouldCompareBySortKey(Me.Item(ItemIndex2)) Then
        result_value = pCompareStringValues(GetTypedValueKey(Me.Item(ItemIndex1)), GetTypedValueKey(Me.Item(ItemIndex2)))
    Else
        result_value = pComparePrimitiveItems(Me.Item(ItemIndex1), Me.Item(ItemIndex2))
    End If

    pCompareNonObjectItems = result_value
End Function

Private Function pComparePrimitiveItems(ByVal ItemObject1 As Variant, ByVal ItemObject2 As Variant) As Long
    Dim result_value As Long
    If ItemObject1 < ItemObject2 Then
        result_value = -1
    ElseIf ItemObject1 > ItemObject2 Then
        result_value = 1
    Else
        result_value = 0
    End If

    pComparePrimitiveItems = result_value
End Function

Private Function pCompareStringValues(ByVal StringValue1 As String, ByVal StringValue2 As String) As Long
    Dim result_value As Long
    If StringValue1 < StringValue2 Then
        result_value = -1
    ElseIf StringValue1 > StringValue2 Then
        result_value = 1
    Else
        result_value = 0
    End If

    pCompareStringValues = result_value
End Function

Private Function pNormalizeCompareResult(ByVal CompareResult As Long) As Long
    Dim result_value As Long
    If CompareResult < 0 Then
        result_value = -1
    ElseIf CompareResult > 0 Then
        result_value = 1
    Else
        result_value = 0
    End If

    pNormalizeCompareResult = result_value
End Function

Private Function pShouldCompareBySortKey(ByVal ItemObject As Variant) As Boolean
    Dim result_value As Boolean
    If IsArray(ItemObject) Then
        result_value = True
    ElseIf IsError(ItemObject) Then
        result_value = True
    Else
        result_value = False
    End If

    pShouldCompareBySortKey = result_value
End Function

Private Function pGetObjectSortKey(ByVal ItemObject As Variant) As String
    Dim object_mode As Long
    If pIsDuplicateCheckable Then
        object_mode = G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
    ElseIf pIsEquatable Then
        object_mode = G_OBJECT_KEY_MODE_I_EQUATABLE
    Else
        object_mode = G_OBJECT_KEY_MODE_REFERENCE
    End If

    pGetObjectSortKey = GetTypedValueKey(ItemObject, object_mode, UseElementTypeKey:=True)
End Function

Private Function pCreateNewEnumSnapshot() As Collection
    Dim result_value As Collection
    Set result_value = New Collection

    Dim enum_obj As IEnumerator
    Set enum_obj = Me.GetEnumerator(IsReadOnly:=True)
    Do While enum_obj.MoveNext()
        Call result_value.Add(enum_obj.Current)
    Loop

    Set pCreateNewEnumSnapshot = result_value
End Function

Private Sub pCheckIndex(ByVal ItemIndex As Long)
    If pCount = 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ObjectList.pCheckIndex", Description:="インデックスが範囲外です。(index: " & ItemIndex & ", count: " & pCount & ")"
    End If

    If ItemIndex < 0 Or pCount <= ItemIndex Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ObjectList.pCheckIndex", Description:="インデックスが範囲外です。(index: " & ItemIndex & ", count: " & pCount & ")"
    End If
End Sub

Private Sub pGetIndexArray(ByRef IndexArray() As Long, ByRef RemainIndex As Long, ByVal ItemIndex As Long)
    Dim result_value() As Long
    ReDim result_value(0 To pDepth)

    Dim remain_index As Long
    remain_index = ItemIndex

    Dim item_idx As Long
    For item_idx = pDepth To 0 Step -1
        result_value(item_idx) = remain_index Mod C_CHUNK_SIZE
        remain_index = remain_index \ C_CHUNK_SIZE
    Next item_idx

    IndexArray = result_value
    RemainIndex = remain_index
End Sub

Private Sub pInitializeItemType(ByVal ItemObject As Variant)
    Call pCheckUnsupportedSpecialValue(ItemObject)

    Dim item_is_obj As Boolean
    item_is_obj = IsObject(ItemObject)

    If item_is_obj Then
        pIsObject = True
        If Not ItemObject Is Nothing Then
            pElementTypeName = pGetElementTypeName(ItemObject)
            pTypeName = "Object@" & pElementTypeName
            pHasItemTypeContract = True
            If TypeOf ItemObject Is IComparable Then
                pIsComparable = True
            End If
            If TypeOf ItemObject Is IEquatable Then
                pIsEquatable = True
                pObjectKeyMode = G_OBJECT_KEY_MODE_I_EQUATABLE
            End If
            If TypeOf ItemObject Is IDuplicateCheckable Then
                pIsDuplicateCheckable = True
                pIsEquatable = False
                pObjectKeyMode = G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            End If
            If TypeOf ItemObject Is IStringable Then
                pIsStringable = True
            End If
        End If
    Else
        If pIsObject Then
            Err.Raise vbObjectError + 1, "Class ObjectList.pInitializeItemType", "最初の要素はオブジェクト型でしたが、新しい要素はオブジェクト型ではありません。(" & pTypeName & ")"
        Else
            ' プリミティブ型
            pTypeName = GetTypeString(ItemObject)
            pHasItemTypeContract = True
        End If
    End If
End Sub

Private Function pGetElementTypeName(ByVal ItemObject As Object) As String
    Dim result_value As String
    If TypeOf ItemObject Is IElementTypeProvider Then
        Dim element_type_provider As IElementTypeProvider
        Set element_type_provider = ItemObject
        result_value = element_type_provider.ElementTypeKey
        If Not IsValidElementTypeKey(result_value) Then
            Err.Raise vbObjectError + 1, "Class ObjectList.pGetElementTypeName", "ElementTypeKey はクラス モジュール名に使用可能な文字で指定してください。(" & result_value & ")"
        End If
    Else
        result_value = TypeName(ItemObject)
    End If

    pGetElementTypeName = result_value
End Function

Private Sub pCheckObjectItemAbilities(ByVal ItemObject As Object)
    If pRequireComparable And Not TypeOf ItemObject Is IComparable Then
        Err.Raise vbObjectError + 1, "Class ObjectList.pCheckObjectItemAbilities", "IComparable ではないオブジェクトを RequireComparable=True のリストへ追加しようとしました。(" & pTypeName & ")"
    End If

    Select Case pObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE
        Case G_OBJECT_KEY_MODE_I_EQUATABLE
            If Not TypeOf ItemObject Is IEquatable Then
                Err.Raise vbObjectError + 1, "Class ObjectList.pCheckObjectItemAbilities", "IEquatable ではないオブジェクトを IEquatable モードのリストへ追加しようとしました。(" & pTypeName & ")"
            End If
        Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            If Not TypeOf ItemObject Is IDuplicateCheckable Then
                Err.Raise vbObjectError + 1, "Class ObjectList.pCheckObjectItemAbilities", "IDuplicateCheckable ではないオブジェクトを IDuplicateCheckable モードのリストへ追加しようとしました。(" & pTypeName & ")"
            End If
        Case Else
            Err.Raise vbObjectError + 1, "Class ObjectList.pCheckObjectItemAbilities", "未対応のオブジェクト キー化モードです。(" & CStr(pObjectKeyMode) & ")"
    End Select
End Sub

Private Function pCanSearchItem(ByVal ItemObject As Variant) As Boolean
    Dim result_value As Boolean
    result_value = False

    Call pCheckUnsupportedSpecialValue(ItemObject)

    If pTypeName <> "" Then
        Call pCheckItemType(ItemObject)
        result_value = True
        pCanSearchItem = result_value
        Exit Function
    End If

    If Not pIsObject Then
        pCanSearchItem = result_value
        Exit Function
    End If

    If Not IsObject(ItemObject) Then
        pCanSearchItem = result_value
        Exit Function
    End If

    If ItemObject Is Nothing Then
        result_value = True
    Else
        result_value = False
    End If

    pCanSearchItem = result_value
End Function

Private Sub pCheckItemType(ByVal ItemObject As Variant)
    Call pCheckUnsupportedSpecialValue(ItemObject)

    If pIsObject Then
        If Not IsObject(ItemObject) Then
            Err.Raise vbObjectError + 1, "Class ObjectList.pCheckItemType", "最初の要素はオブジェクト型でしたが、新しい要素はオブジェクト型ではありません。(" & pTypeName & ")"
        End If
        If ItemObject Is Nothing Then Exit Sub
    End If

    Dim item_type_name As String
    If pIsObject Then
        item_type_name = GetTypeString(ItemObject, UseElementTypeKey:=True)
    Else
        item_type_name = GetTypeString(ItemObject)
    End If

    If StrComp(item_type_name, pTypeName, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1, "Class ObjectList.pCheckItemType", "最初に設定された要素と型が異なります。(list: " & pTypeName & ", item: " & item_type_name & ")"
    End If

    If pIsObject Then
        Call pCheckObjectItemAbilities(ItemObject)
    End If
End Sub

Private Sub pCheckUnsupportedSpecialValue(ByVal ItemObject As Variant)
    If IsObject(ItemObject) Then Exit Sub
    If IsArray(ItemObject) Then Exit Sub
    If IsEmpty(ItemObject) Then Err.Raise vbObjectError + 1, "Class ObjectList.pCheckUnsupportedSpecialValue", "Empty は要素として使用できません。"
    If IsNull(ItemObject) Then Err.Raise vbObjectError + 1, "Class ObjectList.pCheckUnsupportedSpecialValue", "Null は要素として使用できません。"
End Sub

Private Function pErrorValuesEqual(ByVal ErrorValue1 As Variant, ByVal ErrorValue2 As Variant) As Boolean
    Dim error_values As Dictionary
    Set error_values = New Dictionary
    Call error_values.Add(ErrorValue1, True)
    pErrorValuesEqual = error_values.Exists(ErrorValue2)
End Function
