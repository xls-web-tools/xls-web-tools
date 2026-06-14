VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! キーで要素を参照できる型付きオブジェクト辞書を管理するクラスです。
'!
' #############################################################################

Private pKeyToIndex As Object
Private pKeys As Collection
Private pValues As ObjectList
Private pHasEverAdded As Boolean
Private pNewEnumSnapshot As Collection
Private pNewEnumSnapshotVersion As Long
Private pEnumerationVersion As Long

'* 辞書の現在の件数。
Public Property Get Count() As Long
    Count = pValues.Count
End Property

'* 文字列 key の比較モード。
Public Property Get CompareMode() As Long
    CompareMode = pKeyToIndex.CompareMode
End Property
Public Property Let CompareMode(ByVal NewValue As Long)
    If pHasEverAdded Then
        Err.Raise vbObjectError + 1, "Class ObjectDictionary.CompareMode", "要素追加後に CompareMode を変更することはできません。"
    End If

    pKeyToIndex.CompareMode = NewValue
End Property

'* 標準 For Each 用の key 列挙子を得ます。
'*
'* @return 標準 For Each が使用する列挙子。
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    If pNewEnumSnapshot Is Nothing Or pNewEnumSnapshotVersion <> pEnumerationVersion Then
        Set pNewEnumSnapshot = pCreateNewEnumSnapshot()
        pNewEnumSnapshotVersion = pEnumerationVersion
    End If

    Set NewEnum = pNewEnumSnapshot.[_NewEnum]
End Property

'* value を key 順で保持する ObjectList copy を返します。
'*
'* @return value の ObjectList copy。
Public Property Get Items() As ObjectList
    Set Items = pValues.CopyList()
End Property

'* このコレクションのインデクサ。参照時は指定されたキーに対応する要素を返します。代入時は AddOrUpdate と同様に動作します。
'*
'* @param Key 参照するキー。
Public Property Get Item(ByVal Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    If Not pKeyToIndex.Exists(Key) Then
        Err.Raise vbObjectError + 1, "Class ObjectDictionary.Item", "指定されたキーは存在しません。(" & pKeyToDisplayString(Key) & ")"
    End If

    Dim item_idx As Long
    item_idx = CLng(pKeyToIndex.Item(Key))

    If pValues.IsObjectItemType Then
        Set Item = pValues.Item(item_idx)
    Else
        Item = pValues.Item(item_idx)
    End If
End Property
Public Property Let Item(ByVal Key As Variant, ByVal Value As Variant)
    Call AddOrUpdate(Key, Value)
End Property
Public Property Set Item(ByVal Key As Variant, ByVal Value As Object)
    Call AddOrUpdate(Key, Value)
End Property

Private Sub Class_Initialize()
    Set pKeyToIndex = CreateObject("Scripting.Dictionary")
    Set pKeys = New Collection
    Set pValues = New ObjectList
    pHasEverAdded = False
    Set pNewEnumSnapshot = Nothing
    pNewEnumSnapshotVersion = -1
    pEnumerationVersion = 0
End Sub

'* 辞書の要素型契約を初期化します。
'*
'* @param ElementTypeName 要素型契約名。
'* @param RequireComparable 要素に IComparable 実装を要求するか否か。
'* @param ObjectKeyMode オブジェクトの同一性/重複判定モード。
Public Sub Initialize( _
        ByVal ElementTypeName As String, _
        Optional ByVal RequireComparable As Boolean = False, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE)

    Call pValues.Initialize( _
            ElementTypeName:=ElementTypeName, _
            RequireComparable:=RequireComparable, _
            ObjectKeyMode:=ObjectKeyMode)
End Sub

'* 他の辞書と同じ value 要素型状態と key 比較モードで初期化します。
'*
'* @param Other 初期化元の辞書。
'*
'* @details
'* 他の辞書が管理している value 要素型状態と key 比較モードをこの辞書へ引き継ぎます。
Public Sub InitializeFromOther(ByVal Other As ObjectDictionary)
    Call pValues.InitializeFromOther(Other.Items)
    Me.CompareMode = Other.CompareMode
End Sub

'* 新しい key-value pair を辞書に追加します。
'*
'* @param Key 追加するキー。
'* @param Value 追加する値。
Public Sub Add(ByVal Key As Variant, ByVal Value As Variant)
    If pKeyToIndex.Exists(Key) Then
        Err.Raise vbObjectError + 1, "Class ObjectDictionary.Add", "指定されたキーは既に存在します。(" & pKeyToDisplayString(Key) & ")"
    End If

    Dim new_index As Long
    new_index = pValues.Count

    Dim key_registered As Boolean
    Dim key_stored As Boolean

    On Error GoTo ADD_ERROR
    Call pKeyToIndex.Add(Key, new_index)
    key_registered = True
    Call pAddKey(Key)
    key_stored = True
    Call pValues.Add(Value)
    pHasEverAdded = True
    Call pIncrementEnumerationVersion
    On Error GoTo 0

    Exit Sub

ADD_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    On Error Resume Next
    If key_registered Then Call pKeyToIndex.Remove(Key)
    If key_stored And 0 < pKeys.Count Then Call pKeys.Remove(pKeys.Count)
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Sub

Private Sub pAddKey(ByVal Key As Variant)
    If IsObject(Key) Then
        Dim key_obj As Object
        Set key_obj = Key
        Call pKeys.Add(key_obj)
    Else
        Call pKeys.Add(Key)
    End If
End Sub

'* 他の辞書の key-value pair をこの辞書に取り込みます。
'*
'* @param OtherDictionary 取り込む key-value pair を持つ他の辞書。
'* @param ErrorIfExists [省略可] key が既に存在する場合にエラーを発生させるか否か。
'* @param UpdateIfExists [省略可] key が既に存在する場合に value を更新するか否か。
'*
'* @details
'* 他の辞書の key-value pair をこの辞書に取り込みます。UpdateIfExists が True の場合、ErrorIfExists の値に関わらず既存 key の value を更新します。
Public Sub AddOther( _
        ByVal OtherDictionary As ObjectDictionary, _
        Optional ByVal ErrorIfExists As Boolean = True, _
        Optional ByVal UpdateIfExists As Boolean = False)

    If pCanInitializeFromOther Then Call InitializeFromOther(OtherDictionary)
    If ErrorIfExists And Not UpdateIfExists Then Call pCheckNoExistingKeys(OtherDictionary)

    Dim other_keys() As Variant
    other_keys = OtherDictionary.Keys

    If IsEmptyArray(other_keys) Then Exit Sub

    Dim dict_key As Variant
    For Each dict_key In other_keys
        If Me.Exists(dict_key) Then
            If UpdateIfExists Then
                Call Me.Update(dict_key, OtherDictionary.Item(dict_key))
            ElseIf ErrorIfExists Then
                Err.Raise vbObjectError + 1, "Class ObjectDictionary.AddOther", "指定されたキーは既に存在します。(" & pKeyToDisplayString(dict_key) & ")"
            End If
        Else
            Call Me.Add(dict_key, OtherDictionary.Item(dict_key))
        End If
    Next dict_key
End Sub

Private Function pCanInitializeFromOther() As Boolean
    pCanInitializeFromOther = (pValues.Count = 0 And pValues.ItemTypeName = "" And Not pValues.HasItemTypeContract)
End Function

Private Sub pCheckNoExistingKeys(ByVal OtherDictionary As ObjectDictionary)
    Dim other_keys() As Variant
    other_keys = OtherDictionary.Keys

    If IsEmptyArray(other_keys) Then Exit Sub

    Dim dict_key As Variant
    For Each dict_key In other_keys
        If Me.Exists(dict_key) Then
            Err.Raise vbObjectError + 1, "Class ObjectDictionary.AddOther", "指定されたキーは既に存在します。(" & pKeyToDisplayString(dict_key) & ")"
        End If
    Next dict_key
End Sub

'* 指定されたキーが存在するか否かを返します。
'*
'* @param Key 確認するキー。
'* @return キーが存在する場合は True。
Public Function Exists(ByVal Key As Variant) As Boolean
    Exists = pKeyToIndex.Exists(Key)
End Function

'* 指定されたキーに対応する要素を更新します。
'*
'* @param Key 更新するキー。
'* @param Value 新しい値。
Public Sub Update(ByVal Key As Variant, ByVal Value As Variant)
    If Not pKeyToIndex.Exists(Key) Then
        Err.Raise vbObjectError + 1, "Class ObjectDictionary.Update", "指定されたキーは存在しません。(" & pKeyToDisplayString(Key) & ")"
    End If

    Dim item_idx As Long
    item_idx = CLng(pKeyToIndex.Item(Key))
    Call pValues.Update(item_idx, Value)
End Sub

'* 指定されたキーが存在しない場合は追加し、存在する場合は値を更新します。
'*
'* @param Key 追加または更新するキー。
'* @param Value 追加または更新する値。
Public Sub AddOrUpdate(ByVal Key As Variant, ByVal Value As Variant)
    If pKeyToIndex.Exists(Key) Then
        Call Update(Key, Value)
    Else
        Call Add(Key, Value)
    End If
End Sub

'* 指定されたキーに対応する要素を削除します。
'*
'* @param Key 削除するキー。
'* @return 削除された要素。
Public Function Remove(ByVal Key As Variant) As Variant
    If Not pKeyToIndex.Exists(Key) Then
        Err.Raise vbObjectError + 1, "Class ObjectDictionary.Remove", "指定されたキーは存在しません。(" & pKeyToDisplayString(Key) & ")"
    End If

    Dim item_idx As Long
    item_idx = CLng(pKeyToIndex.Item(Key))

    Dim removed_value As Variant
    If pValues.IsObjectItemType Then
        Set removed_value = pValues.Remove(item_idx)
    Else
        removed_value = pValues.Remove(item_idx)
    End If

    Call pKeyToIndex.Remove(Key)
    Call pKeys.Remove(item_idx + 1)
    Call pRebuildIndexesFrom(item_idx)
    Call pIncrementEnumerationVersion

    If IsObject(removed_value) Then
        Set Remove = removed_value
    Else
        Remove = removed_value
    End If
End Function

Private Sub pRebuildIndexesFrom(ByVal StartIndex As Long)
    Dim key_idx As Long
    For key_idx = StartIndex To pKeys.Count - 1
        Dim key_value As Variant
        If IsObject(pKeys.Item(key_idx + 1)) Then
            Set key_value = pKeys.Item(key_idx + 1)
        Else
            key_value = pKeys.Item(key_idx + 1)
        End If

        pKeyToIndex.Item(key_value) = key_idx
    Next key_idx
End Sub

'* すべての key-value pair を削除します。
'*
'* @details
'* value の要素型契約と CompareMode は維持します。
Public Sub RemoveAll()
    Call pKeyToIndex.RemoveAll
    Set pKeys = New Collection
    Call pValues.RemoveAll
    Call pIncrementEnumerationVersion
End Sub

'* key を追加順の 0-based 配列として返します。
'*
'* @return key の配列。
Public Function Keys() As Variant()
    If pKeys.Count = 0 Then
        Keys = EmptyVariantArray()
        Exit Function
    End If

    Dim result_value() As Variant
    ReDim result_value(0 To pKeys.Count - 1)

    Dim key_idx As Long
    For key_idx = 1 To pKeys.Count
        If IsObject(pKeys.Item(key_idx)) Then
            Set result_value(key_idx - 1) = pKeys.Item(key_idx)
        Else
            result_value(key_idx - 1) = pKeys.Item(key_idx)
        End If
    Next key_idx

    Keys = result_value
End Function

'* value を key 順の 0-based 配列として返します。
'*
'* @return value の配列。
Public Function ConvertToArray() As Variant()
    ConvertToArray = pValues.ConvertToArray()
End Function

'* value を key 順の 0-based 文字列配列として返します。
'*
'* @return value の文字列配列。
Public Function ConvertToStringArray() As String()
    ConvertToStringArray = pValues.ConvertToStringArray()
End Function

Private Sub pIncrementEnumerationVersion()
    pEnumerationVersion = pEnumerationVersion + 1
End Sub

Private Function pCreateNewEnumSnapshot() As Collection
    Dim result_value As Collection
    Set result_value = New Collection

    Dim key_idx As Long
    For key_idx = 1 To pKeys.Count
        If IsObject(pKeys.Item(key_idx)) Then
            Dim key_obj As Object
            Set key_obj = pKeys.Item(key_idx)
            Call result_value.Add(key_obj)
        Else
            Call result_value.Add(pKeys.Item(key_idx))
        End If
    Next key_idx

    Set pCreateNewEnumSnapshot = result_value
End Function

Private Function pKeyToDisplayString(ByVal Key As Variant) As String
    On Error GoTo TO_STRING_ERROR

    If IsObject(Key) Then
        pKeyToDisplayString = TypeName(Key) & "@0x" & Hex(ObjPtr(Key))
    ElseIf IsArray(Key) Then
        pKeyToDisplayString = "Array"
    Else
        pKeyToDisplayString = CStr(Key)
    End If

    Exit Function

TO_STRING_ERROR:
    pKeyToDisplayString = TypeName(Key)
End Function
