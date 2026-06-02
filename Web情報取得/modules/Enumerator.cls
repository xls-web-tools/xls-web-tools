VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IEnumerator

' #############################################################################
'!
'! @brief
'! コレクションを列挙するための列挙子のクラスです。
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pList As Object
Private pArray As Variant
Private pDescending As Boolean
Private pIsReadOnly As Boolean
Private pIsArray As Boolean
Private pIndex As Long
Private pMinIndex As Long

Public Property Get Descending() As Boolean
    Call pCheckInit
    Descending = pDescending
End Property
Private Property Get IEnumerator_Descending() As Boolean
    IEnumerator_Descending = Descending
End Property

'* 読み取り専用かどうか
Public Property Get IsReadOnly() As Boolean
    Call pCheckInit
    IsReadOnly = pIsReadOnly
End Property
Private Property Get IEnumerator_IsReadOnly() As Boolean
    IEnumerator_IsReadOnly = IsReadOnly
End Property

' 現在の要素
Public Property Get Current() As Variant
    Call pCheckInit
    Call pCheckEndOfList

    If pIsArray Then
        If IsObject(pArray(pIndex)) Then
            Set Current = pArray(pIndex)
        Else
            Current = pArray(pIndex)
        End If
    Else
        If IsObject(pList.Item(pIndex)) Then
            Set Current = pList.Item(pIndex)
        Else
            Current = pList.Item(pIndex)
        End If
    End If
End Property
Private Property Get IEnumerator_Current() As Variant
    If IsObject(Current) Then
        Set IEnumerator_Current = Current
    Else
        IEnumerator_Current = Current
    End If
End Property

' 現在のインデックス
Public Property Get Index() As Long
    Call pCheckInit
    Index = pIndex
End Property
Private Property Get IEnumerator_Index() As Long
    IEnumerator_Index = Index
End Property

' 列挙対象のコレクション
Public Property Get Target() As Variant
    Call pCheckInit
    If pIsArray Then
        Target = pArray
    Else
        Set Target = pList
    End If
End Property
Private Property Let Target(ByVal NewValue As Variant)
    If Not IsArray(NewValue) Then
        Err.Raise vbObjectError + 1, "Class Enumerator", "配列ではありません。"
    End If

    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Call GetArrayBounds(lbound_arr, ubound_arr, NewValue)
    If LBound(lbound_arr) <> UBound(lbound_arr) Then
        Err.Raise vbObjectError + 1, "Class Enumerator", "配列列挙子は 1 次元配列のみ対応しています。"
    End If

    pArray = NewValue
    pMinIndex = lbound_arr(LBound(lbound_arr))
    pIsArray = True
End Property
Private Property Set Target(ByVal NewValue As Object)
    Set pList = NewValue
    pMinIndex = 0
End Property

Private Property Get IEnumerator_Target() As Variant
    If IsObject(Target) Then
        Set IEnumerator_Target = Target
    Else
        IEnumerator_Target = Target
    End If
End Property

'* このオブジェクトを初期化します。
'*
'* @param TargetCollection 列挙対象のコレクションまたは 1 次元配列。Nothing と未初期化配列、多次元配列はエラーです。
'* @param Descending [省略可] 降順で列挙するか否か。
'* @param IsReadOnly [省略可] 読み取り専用にするか否か。
'*
'* @details
'* このオブジェクトを初期化します。オブジェクトは Count / Item を持つものとして扱います。
Public Sub Initialize(ByVal TargetCollection As Variant, Optional ByVal Descending As Boolean = False, Optional ByVal IsReadOnly As Boolean = False)
    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class Enumerator", "初期化済みです。"

    If IsObject(TargetCollection) Then
        If TargetCollection Is Nothing Then Err.Raise vbObjectError + 1, "Class Enumerator", "列挙対象が Nothing です。"
        Set Target = TargetCollection
    Else
        Target = TargetCollection
    End If

    pDescending = Descending
    pIsReadOnly = IsReadOnly

    pIsInitialized = True

    Call Reset
End Sub

' 次の要素に移動します。
'
' @return Boolean 次の要素が存在する場合は True、それ以外は False。
'
' @details
' 現在の位置を次の要素に移動します。リストの終端に達した場合は False を返します。
' 初期値はコレクションの先頭の 1 つ前であり、1 回呼び出すことで最初の要素を指すことになります。
Public Function MoveNext() As Boolean
    Call pCheckInit

    If Descending Then
        If pMinIndex < pIndex Then
            pIndex = pIndex - 1
            MoveNext = True
        Else
            pIndex = pMinIndex - 1
            MoveNext = False
        End If
    Else
        If pIndex < pGetCount() - 1 Then
            pIndex = pIndex + 1
            MoveNext = True
        Else
            pIndex = pGetCount()
            MoveNext = False
        End If
    End If
End Function
Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = MoveNext()
End Function

' 列挙をリセットします。
'
' @details
' 列挙の現在位置をリストの先頭の 1 つ前 (逆順のときはリストの末尾の 1 つ後) に戻します。
Public Sub Reset()
    Call pCheckInit
    If Descending Then
        pIndex = pGetCount
    Else
        pIndex = pMinIndex - 1
    End If
End Sub
Private Sub IEnumerator_Reset()
    Call Reset
End Sub

' 指定した位置までスキップします。
'
' @param TargetIndex 移動先のインデックス。
'
' @details
' 指定されたインデックスまで現在位置を進めます。現在位置より前やリストの範囲外を指定するとエラーを発生させます。
Public Sub SkipTo(ByVal TargetIndex As Long)
    Call pCheckInit
    If TargetIndex < pMinIndex Or pGetCount() <= TargetIndex Then Err.Raise vbObjectError + 1, "Class Enumerator", "移動先が範囲外です。(min: " & pMinIndex & " / max: " & (pGetCount() - 1) & " / target: " & TargetIndex & ")"
    If Descending Then
        If pIndex < TargetIndex Then Err.Raise vbObjectError + 1, "Class Enumerator", "現在の位置より前には移動できません。(current: " & pIndex & " / target: " & TargetIndex & ")"
    Else
        If TargetIndex < pIndex Then Err.Raise vbObjectError + 1, "Class Enumerator", "現在の位置より前には移動できません。(current: " & pIndex & " / target: " & TargetIndex & ")"
    End If
    pIndex = TargetIndex
End Sub
Private Sub IEnumerator_SkipTo(ByVal TargetIndex As Long)
    Call SkipTo(TargetIndex)
End Sub

' 現在の要素を削除します。
'
' @return Variant 削除された要素。
'
' @details
' 現在の要素を削除します。内部的には、対象コレクションの Remove(ItemIndex) メソッドが呼び出されます。
' この操作を行うと、この列挙子が示す位置は一つ前に戻ります。
' 順序を保証しないコレクションに対して使用した場合、この操作後の列挙子の動作は保証されません。
Public Function Remove() As Variant
    Call pCheckInit
    Call pCheckEndOfList
    Call pCheckWritable

    If pIsArray Then
        Err.Raise vbObjectError + 1, "Class Enumerator", "配列では使用できません。"
    Else
        If IsObject(pList.Item(pIndex)) Then
            Set Remove = pList.Remove(pIndex)
        Else
            Remove = pList.Remove(pIndex)
        End If
    End If

    If Not Descending Then
        pIndex = pIndex - 1
    End If
End Function
Private Function IEnumerator_Remove() As Variant
    ' オブジェクトとプリミティブ型の扱いを統一できないため、コピペで対応
    Call pCheckInit
    Call pCheckEndOfList
    Call pCheckWritable

    If pIsArray Then
        Err.Raise vbObjectError + 1, "Class Enumerator", "配列では使用できません。"
    Else
        If IsObject(pList.Item(pIndex)) Then
            Set IEnumerator_Remove = pList.Remove(pIndex)
        Else
            IEnumerator_Remove = pList.Remove(pIndex)
        End If
    End If

    If Not Descending Then
        pIndex = pIndex - 1
    End If
End Function

' 現在の要素を更新します。
'
' @param ItemObject 更新する要素の値。
'
' @details
' 現在の要素を更新します。内部的には、対象コレクションの Update(ItemIndex, ItemObject) メソッドが呼び出されます。
' 配列列挙子では呼び出し元配列へ更新を反映できないため、この操作は使用できません。
' 順序を保証しないコレクションに対して使用した場合、この操作後の列挙子の動作は保証されません。
Public Sub Update(ByVal ItemObject As Variant)
    Call pCheckInit
    Call pCheckEndOfList
    Call pCheckWritable

    If pIsArray Then
        Err.Raise vbObjectError + 1, "Class Enumerator", "配列では使用できません。"
    Else
        Call pList.Update(pIndex, ItemObject)
    End If
End Sub
Private Sub IEnumerator_Update(ByVal ItemObject As Variant)
    Call Update(ItemObject)
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class Enumerator", "未初期化状態です。"
End Sub

Private Sub pCheckWritable()
    If pIsReadOnly Then Err.Raise vbObjectError + 1, "Class Enumerator", "読み取り専用の列挙子では要素を更新または削除できません。"
End Sub

Private Function pGetCount() As Long
    If pIsArray Then
        If IsEmptyArray(pArray) Then
            pGetCount = 0
        Else
            pGetCount = UBound(pArray) + 1
        End If
    Else
        pGetCount = pList.Count
    End If
End Function

Private Sub pCheckEndOfList()
    Dim count_var As Long
    count_var = pGetCount()

    If pIndex < pMinIndex Or count_var <= pIndex Then
        Err.Raise vbObjectError + 1, "Class Enumerator", "現在の要素がありません。(index: " & pIndex & " / min: " & pMinIndex & " / max: " & (count_var - 1) & ")"
    End If
End Sub
