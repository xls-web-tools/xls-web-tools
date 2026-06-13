VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 任意のオブジェクトまたは値を格納し、効率的にアクセスできるようにするための集合を提供します。
'! 各要素はユニークなキーによって管理されます。
'!
' #############################################################################

Private pDict As Dictionary
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

'* 集合内の要素の総数。
Public Property Get Count() As Long
    Count = pDict.Count
End Property

'* 集合内の要素の型。
Public Property Get ItemTypeName() As String
    ItemTypeName = pTypeName
End Property

'* 集合に要素型契約が設定されているか否か。
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
'* 標準 For Each の列挙中にこの集合を更新した場合、実行中の列挙の継続は保証されません。
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

'* このコレクションのインデクサ。読み取り専用です。指定されたインデックスにある要素を返します。要素は内部的にキーに基づいてアクセスされます。
'*
'* @param ItemIndex 要素のインデックス。
Public Property Get Item(ByVal ItemIndex As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If pIsObject Then
        Set Item = pDict(pDict.Keys(ItemIndex))
    Else
        Item = pDict(pDict.Keys(ItemIndex))
    End If
End Property

Private Sub Class_Initialize()
    Set pDict = New Dictionary
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
    pEnumerationVersion = 0
    pNewEnumSnapshotVersion = -1
    Set pNewEnumSnapshot = Nothing
End Sub

'* 集合の要素型契約を初期化します。
'*
'* @param ElementTypeName 要素型契約名。
'* @param RequireComparable 要素に IComparable 実装を要求するか否か。
'* @param ObjectKeyMode オブジェクトの同一性/重複判定モード。
'*
'* @details
'* オブジェクト型の要素だけを受け入れる集合として初期化します。
'* ElementTypeName は、自己申告された ElementTypeKey または TypeName と比較されます。
Public Sub Initialize( _
        ByVal ElementTypeName As String, _
        Optional ByVal RequireComparable As Boolean = False, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE)

    If pDict.Count <> 0 Or pHasItemTypeContract Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.Initialize", "初期化済みの集合を再初期化することはできません。"
    End If
    If Not IsValidElementTypeKey(ElementTypeName) Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.Initialize", "ElementTypeName はクラス モジュール名に使用可能な文字で指定してください。(" & ElementTypeName & ")"
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

'* 他の集合と同じ要素型状態で初期化します。
'*
'* @param Other 初期化元の集合。
'*
'* @details
'* 他の集合が管理している要素型状態をこの集合へ引き継ぎます。
Public Sub InitializeFromOther(ByVal Other As ObjectSet)
    If pDict.Count <> 0 Or pTypeName <> "" Or pHasItemTypeContract Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.InitializeFromOther", "初期化済みの集合を再初期化することはできません。"
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

'* 他のリストと同じ要素型状態で初期化します。
'*
'* @param OtherList 初期化元のリスト。
'*
'* @details
'* 他のリストが管理している要素型状態をこの集合へ引き継ぎます。
Public Sub InitializeFromList(ByVal OtherList As ObjectList)
    If pDict.Count <> 0 Or pTypeName <> "" Or pHasItemTypeContract Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.InitializeFromList", "初期化済みの集合を再初期化することはできません。"
    End If

    pTypeName = OtherList.ItemTypeName
    pElementTypeName = OtherList.ElementTypeName
    pHasItemTypeContract = OtherList.HasItemTypeContract
    pRequireComparable = OtherList.RequireComparable
    pObjectKeyMode = OtherList.ObjectKeyMode
    pIsObject = OtherList.IsObjectItemType
    pIsComparable = OtherList.IsComparableItemType
    pIsStringable = OtherList.IsStringableItemType
    pIsEquatable = OtherList.IsEquatableItemType
    pIsDuplicateCheckable = OtherList.IsDuplicateCheckableItemType
End Sub

Private Sub pCheckObjectKeyMode(ByVal ObjectKeyMode As Long)
    Select Case ObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE, G_OBJECT_KEY_MODE_I_EQUATABLE, G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
        Case Else
            Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckObjectKeyMode", "未対応のオブジェクト キー化モードです。(" & CStr(ObjectKeyMode) & ")"
    End Select
End Sub

'* 指定されたインデックスの要素を新しいオブジェクトで更新します。
'*
'* @param ItemIndex 更新する要素のインデックス。
'* @param ItemObject 新しい要素。
'*
'* @details
'* 指定されたインデックスの要素を新しいオブジェクトで更新します。キーが変更される場合はエラーが発生します。
Public Sub Update(ByVal ItemIndex As Long, ByVal ItemObject As Variant)
    Call pCheckItemType(ItemObject)

    Dim old_key As Variant
    If IsObject(pDict.Keys(ItemIndex)) Then
        Set old_key = pDict.Keys(ItemIndex)
    Else
        old_key = pDict.Keys(ItemIndex)
    End If

    Dim new_key As Variant
    Call pGetKey(new_key, ItemObject)

    If Not pKeyEquals(old_key, new_key) Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.Update", "元のオブジェクトと、変更後のオブジェクトのキーが違います。変更が必要な場合は、元の要素を Remove 後に新しい要素を Add してください。(old_key: " & old_key & " / new_key: " & new_key & ")"
        Exit Sub
    End If

    If pIsObject Then
        Set pDict(old_key) = ItemObject
    Else
        pDict(old_key) = ItemObject
    End If
    Call pIncrementEnumerationVersion
End Sub

Private Function pKeyEquals(ByVal ItemKey1 As Variant, ByVal ItemKey2 As Variant) As Boolean
    Dim result_value As Boolean
    result_value = (CStr(ItemKey1) = CStr(ItemKey2))

    pKeyEquals = result_value
End Function

'* 集合を配列として返します。
'*
'* @return 配列形式の要素。
'*
'* @details
'* 集合のすべての要素を配列形式で返します。要素数 0 の場合は、0 ベースの空配列を返します。
Public Function ConvertToArray() As Variant()
    If pDict.Count < 1 Then
        ConvertToArray = EmptyVariantArray()
        Exit Function
    End If

    Dim result_value() As Variant
    Dim item_idx As Long

    ReDim result_value(0 To pDict.Count - 1)

    item_idx = LBound(result_value)
    Dim dict_key As Variant
    For Each dict_key In pDict
        If pIsObject Then
            Set result_value(item_idx) = pDict(dict_key)
        Else
            result_value(item_idx) = pDict(dict_key)
        End If

        item_idx = item_idx + 1
    Next dict_key

    ConvertToArray = result_value
End Function

'* 集合を文字列の配列として返します。
'*
'* @return 文字列形式の配列。
'*
'* @details
'* 集合のすべての要素を文字列形式で配列にして返します。要素数 0 の場合は、0 ベースの空配列を返します。
Public Function ConvertToStringArray() As String()
    If pDict.Count < 1 Then
        ConvertToStringArray = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    Dim item_idx As Long
    Dim stringable_obj As IStringable
    Dim dict_item_obj As Object

    ReDim result_value(0 To pDict.Count - 1)

    item_idx = LBound(result_value)
    Dim dict_key As Variant
    For Each dict_key In pDict
        If pIsObject Then
            Set dict_item_obj = pDict(dict_key)
            If dict_item_obj Is Nothing Then
                result_value(item_idx) = GetTypedValueKey(dict_item_obj)
            Else
                If TypeOf dict_item_obj Is IStringable Then
                    Set stringable_obj = dict_item_obj
                    result_value(item_idx) = stringable_obj.ToString()
                Else
                    result_value(item_idx) = GetTypedValueKey(dict_item_obj)
                End If
            End If
        Else
            If IsError(pDict(dict_key)) Then
                result_value(item_idx) = ExcelErrorToString(pDict(dict_key))
            ElseIf IsArray(pDict(dict_key)) Then
                result_value(item_idx) = GetTypedValueKey(pDict(dict_key))
            Else
                result_value(item_idx) = CStr(pDict(dict_key))
            End If
        End If
        item_idx = item_idx + 1
    Next dict_key

    ConvertToStringArray = result_value
End Function

'* 新しい要素を集合に追加します。
'*
'* @param ItemObject 追加する要素。
'* @param Optional ErrorIfExists 既に存在する場合にエラーを発生させるかどうか。
'* @return 要素が追加されたかどうかを示すブール値。
'*
'* @details
'* 新しい要素を集合に追加します。オプションで、要素が既に存在する場合の挙動を指定できます。
Public Function Add(ByVal ItemObject As Variant, Optional ByVal ErrorIfExists As Boolean = True) As Boolean
    If pTypeName = "" Then
        Call pInitializeItemType(ItemObject)
    Else
        Call pCheckItemType(ItemObject)
    End If

    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)

    If Not ErrorIfExists Then
        If pDict.Exists(key_var) Then
            Add = False
            Exit Function
        End If
    End If

    Call pDict.Add(key_var, ItemObject)
    Call pIncrementEnumerationVersion
    Add = True
End Function

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
            Err.Raise vbObjectError + 1, "Class ObjectSet.pInitializeItemType", "最初の要素はオブジェクト型でしたが、新しい要素はオブジェクト型ではありません。(" & pTypeName & ")"
        Else
            ' プリミティブ型
            pTypeName = GetTypeString(ItemObject)
            pHasItemTypeContract = True
        End If
    End If
End Sub

'* 配列の各要素を集合に追加します。
'*
'* @param ItemArray 追加する要素の配列。
'* @param Optional ErrorIfExists 既に存在する場合にエラーを発生させるかどうか。
'*
'* @details
'* 配列の各要素を集合に追加します。オプションで、要素が既に存在する場合の挙動を指定できます。
Public Sub AddArray(ByVal ItemArray As Variant, Optional ByVal ErrorIfExists As Boolean = True)
    If Not IsArray(ItemArray) Then Err.Raise vbObjectError + 1, "Class ObjectSet.AddArray", "引数が配列ではありません"
    If IsEmptyArray(ItemArray) Then Exit Sub

    Dim key_var As Variant
    Dim set_item_obj As IDuplicateCheckable
    Dim item_obj As Variant

    For Each item_obj In ItemArray
        Call Add(item_obj, ErrorIfExists)
    Next item_obj
End Sub

'* 他の集合の要素をこの集合に取り込みます。
'*
'* @param OtherSet 取り込む要素を持つ他の集合。
'* @param Optional ErrorIfExists 既に存在する場合にエラーを発生させるかどうか。
'*
'* @details
'* 他の集合の要素をこの集合に取り込みます。オプションで、要素が既に存在する場合の挙動を指定できます。
Public Sub AddOther(ByVal OtherSet As ObjectSet, Optional ByVal ErrorIfExists As Boolean = True)
    If pTypeName = "" And OtherSet.HasItemTypeContract Then
        Call InitializeFromOther(OtherSet)
    Else
        Call pPrepareToAddObjectContract( _
                OtherSet.HasItemTypeContract, _
                OtherSet.ElementTypeName, _
                OtherSet.RequireComparable, _
                OtherSet.ObjectKeyMode)
    End If

    Dim item_idx As Long
    For item_idx = 0 To OtherSet.Count - 1
        Call Me.Add(OtherSet.Item(item_idx), ErrorIfExists:=ErrorIfExists)
    Next item_idx
End Sub

'* 他のリストの要素をこの集合に取り込みます。
'*
'* @param OtherList 取り込む要素を持つ他のリスト。
'* @param Optional ErrorIfExists 既に存在する場合にエラーを発生させるかどうか。
'*
'* @details
'* 他のリストの要素をこの集合に取り込みます。オプションで、要素が既に存在する場合の挙動を指定できます。
Public Sub AddList(ByVal OtherList As ObjectList, Optional ByVal ErrorIfExists As Boolean = True)
    If pTypeName = "" And OtherList.HasItemTypeContract Then
        Call InitializeFromList(OtherList)
    Else
        Call pPrepareToAddObjectContract( _
                OtherList.HasItemTypeContract, _
                OtherList.ElementTypeName, _
                OtherList.RequireComparable, _
                OtherList.ObjectKeyMode)
    End If

    Dim item_idx As Long
    For item_idx = 0 To OtherList.Count - 1
        Call Me.Add(OtherList.Item(item_idx), ErrorIfExists)
    Next item_idx
End Sub

'* この集合の複製を作成します。
'*
'* @return 複製された新しい集合。
'*
'* @details
'* この集合の複製を作成します。
Public Function CopySet() As ObjectSet
    Dim result_value As ObjectSet
    Set result_value = New ObjectSet
    Call result_value.InitializeFromOther(Me)

    Dim key_var As Variant
    For Each key_var In pDict.Keys
        Call result_value.Add(pDict(key_var))
    Next key_var

    Set CopySet = result_value
End Function

'* 指定されたインデックスの要素を削除します。
'*
'* @param Index 削除する要素のインデックス。
'* @return 削除された要素。
'*
'* @details
'* 指定されたインデックスの要素を削除します。
Public Function Remove(ByVal Index As Long) As Variant
    Dim old_key As Variant
    Dim old_obj As Variant

    If pIsObject Then
        Set old_obj = pDict.Item(pDict.Keys(Index))
    Else
        old_obj = pDict.Item(pDict.Keys(Index))
    End If

    If pIsObject Then
        Set Remove = RemoveItem(old_obj)
    Else
        Remove = RemoveItem(old_obj)
    End If
End Function

'* 指定された要素を削除します。
'*
'* @param ItemObject 削除する要素。
'* @param Optional IgnoreNotExists 存在しない場合に無視するかどうか。
'* @return 削除された要素。
'*
'* @details
'* 指定された要素と重複と判断される要素を削除します。オプションで、要素が存在しない場合の挙動を指定できます。
Public Function RemoveItem(ByVal ItemObject As Variant, Optional ByVal IgnoreNotExists As Boolean = True) As Variant
    If Not pCanSearchItem(ItemObject) Then
        If Not IgnoreNotExists Then
            Err.Raise vbObjectError + 1, "Class ObjectSet.RemoveItem", "存在しない要素を削除しようとしました。"
        End If
        If pIsObject Or IsObject(ItemObject) Then
            Set RemoveItem = Nothing
            Exit Function
        Else
            RemoveItem = Empty
            Exit Function
        End If
    End If

    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)

    If Not pDict.Exists(key_var) Then
        If Not IgnoreNotExists Then
            Err.Raise vbObjectError + 1, "Class ObjectSet.RemoveItem", "存在しない要素を削除しようとしました。"
        End If
        If pIsObject Or IsObject(ItemObject) Then
            Set RemoveItem = Nothing
            Exit Function
        Else
            RemoveItem = Empty
            Exit Function
        End If
    End If

    If pIsObject Then
        Dim result_obj As Object
        Set result_obj = pDict(key_var)
        Call pDict.Remove(key_var)
        Call pIncrementEnumerationVersion
        Set RemoveItem = result_obj
        Exit Function
    End If

    Dim result_value As Variant
    result_value = pDict(key_var)
    Call pDict.Remove(key_var)
    Call pIncrementEnumerationVersion

    RemoveItem = result_value
End Function

'* 集合のすべての要素を削除します。
'*
'* @details
'* 集合のすべての要素を削除します。要素の型状態は保持します。
Public Sub RemoveAll()
    If pDict.Count = 0 Then Exit Sub
    Call pDict.RemoveAll
    Call pIncrementEnumerationVersion
End Sub

'* 引数の要素がすでに含まれているかを確認します。
'*
'* @param ItemObject 確認する要素。
'* @return 要素が含まれているかどうかを示すブール値。
'*
'* @details
'* 指定されたオブジェクトと重複している要素が含まれるかを確認します。
Public Function Exists(ByVal ItemObject As Variant) As Boolean
    Dim result_value As Boolean
    result_value = False

    If Not pCanSearchItem(ItemObject) Then
        Exists = result_value
        Exit Function
    End If

    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)
    result_value = pDict.Exists(key_var)

    Exists = result_value
End Function

'* 引数と重複するとみなされる要素を返します。
'*
'* @param ItemObject 確認する要素。
'* @return 重複するとみなされる要素。
'*
'* @details
'* 引数と重複するとみなされる要素を返します。要素が存在しない場合は Nothing を返します。
Public Function FindItem(ByVal ItemObject As Variant) As Variant
    If Not pCanSearchItem(ItemObject) Then
        If pIsObject Or IsObject(ItemObject) Then
            Set FindItem = Nothing
            Exit Function
        Else
            FindItem = Empty
            Exit Function
        End If
    End If

    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)

    If Not pDict.Exists(key_var) Then
        If pIsObject Or IsObject(ItemObject) Then
            Set FindItem = Nothing
            Exit Function
        Else
            FindItem = Empty
            Exit Function
        End If
    End If

    If pIsObject Then
        Dim result_obj As Object
        Set result_obj = pDict(key_var)
        Set FindItem = result_obj
        Exit Function
    End If

    Dim result_value As Variant
    result_value = pDict(key_var)

    FindItem = result_value
End Function

'* この集合の中身をソートします。
'*
'* @param Optional Descending 降順にソートするかどうか。
'*
'* @details
'* この集合の中身をソートします。
'* 内部的に ObjectList.Sort() を利用するため、挙動の詳細はそちらを参照してください。
Public Sub Sort(Optional ByVal Descending As Boolean = False)
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.InitializeFromSet(Me)
    Call obj_list.AddSet(Me)
    Call obj_list.Sort(Descending:=Descending)
    Call Me.RemoveAll
    Call Me.AddList(obj_list)
End Sub

'* この集合の列挙子を得ます。
'*
'* @param Descending [省略可] 降順で列挙するか否か。
'* @param IsReadOnly [省略可] 読み取り専用にするか否か。
'* @return 列挙子オブジェクト。
'*
'* @details
'* この集合の列挙子を得ます。
Public Function GetEnumerator(Optional ByVal Descending As Boolean = False, Optional ByVal IsReadOnly As Boolean = False) As IEnumerator
    Dim result_value As Enumerator
    Set result_value = New Enumerator
    Call result_value.Initialize(Me, Descending:=Descending, IsReadOnly:=IsReadOnly)
    Set GetEnumerator = result_value
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

Private Sub pIncrementEnumerationVersion()
    pEnumerationVersion = pEnumerationVersion + 1
End Sub

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
        Err.Raise vbObjectError + 1, "Class ObjectSet.pPrepareToAddObjectContract", "取り込み元と取り込み先の要素型契約が異なります。"
    End If
End Sub

Private Function pGetElementTypeName(ByVal ItemObject As Object) As String
    Dim result_value As String
    If TypeOf ItemObject Is IElementTypeProvider Then
        Dim element_type_provider As IElementTypeProvider
        Set element_type_provider = ItemObject
        result_value = element_type_provider.ElementTypeKey
        If Not IsValidElementTypeKey(result_value) Then
            Err.Raise vbObjectError + 1, "Class ObjectSet.pGetElementTypeName", "ElementTypeKey はクラス モジュール名に使用可能な文字で指定してください。(" & result_value & ")"
        End If
    Else
        result_value = TypeName(ItemObject)
    End If

    pGetElementTypeName = result_value
End Function

Private Sub pCheckObjectItemAbilities(ByVal ItemObject As Object)
    If pRequireComparable And Not TypeOf ItemObject Is IComparable Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckObjectItemAbilities", "IComparable ではないオブジェクトを RequireComparable=True の集合へ追加しようとしました。(" & pTypeName & ")"
    End If

    Select Case pObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE
        Case G_OBJECT_KEY_MODE_I_EQUATABLE
            If Not TypeOf ItemObject Is IEquatable Then
                Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckObjectItemAbilities", "IEquatable ではないオブジェクトを IEquatable モードの集合へ追加しようとしました。(" & pTypeName & ")"
            End If
        Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            If Not TypeOf ItemObject Is IDuplicateCheckable Then
                Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckObjectItemAbilities", "IDuplicateCheckable ではないオブジェクトを IDuplicateCheckable モードの集合へ追加しようとしました。(" & pTypeName & ")"
            End If
        Case Else
            Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckObjectItemAbilities", "未対応のオブジェクト キー化モードです。(" & CStr(pObjectKeyMode) & ")"
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
            Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckItemType", "最初の要素はオブジェクト型でしたが、新しい要素はオブジェクト型ではありません。(" & pTypeName & ")"
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
        Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckItemType", "最初に設定された要素と型が異なります。(list: " & pTypeName & ", item: " & item_type_name & ")"
    End If

    If pIsObject Then
        Call pCheckObjectItemAbilities(ItemObject)
    End If
End Sub

Private Sub pGetKey(ByRef ObjectKey As Variant, ByVal ObjectItem As Variant)
    Call pCheckUnsupportedSpecialValue(ObjectItem)

    ObjectKey = GetTypedValueKey(ObjectItem, pObjectKeyMode, UseElementTypeKey:=True)
End Sub

Private Sub pCheckUnsupportedSpecialValue(ByVal ItemObject As Variant)
    If IsObject(ItemObject) Then Exit Sub
    If IsArray(ItemObject) Then Exit Sub
    If IsEmpty(ItemObject) Then Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckUnsupportedSpecialValue", "Empty は要素として使用できません。"
    If IsNull(ItemObject) Then Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckUnsupportedSpecialValue", "Null は要素として使用できません。"
End Sub

Private Function pErrorValuesEqual(ByVal ErrorValue1 As Variant, ByVal ErrorValue2 As Variant) As Boolean
    Dim error_values As Dictionary
    Set error_values = New Dictionary
    Call error_values.Add(ErrorValue1, True)
    pErrorValuesEqual = error_values.Exists(ErrorValue2)
End Function
