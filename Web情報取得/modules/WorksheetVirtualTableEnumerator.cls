VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetVirtualTableEnumerator"
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
'! WorksheetVirtualTable の行を列挙するための列挙子のクラスです。
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pDescending As Boolean
Private pTable As WorksheetVirtualTable
Private pLength As Long
Private pIndex As Long

'* 逆順で列挙するかどうか
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
    IsReadOnly = True
End Property
Private Property Get IEnumerator_IsReadOnly() As Boolean
    IEnumerator_IsReadOnly = IsReadOnly
End Property

'* 現在位置の行 ObjectDictionary
Public Property Get Current() As Variant
    Call pCheckInit
    Call pCheckEndOfList

    Set Current = pTable.Item(pIndex)
End Property
Private Property Get IEnumerator_Current() As Variant
    Set IEnumerator_Current = Current
End Property

'* 現在位置の 0-based index
Public Property Get Index() As Long
    Call pCheckInit
    Index = pIndex
End Property
Private Property Get IEnumerator_Index() As Long
    IEnumerator_Index = Index
End Property

'* 列挙対象の仮想表
Public Property Get Target() As Variant
    Call pCheckInit
    Set Target = pTable
End Property
Private Property Get IEnumerator_Target() As Variant
    Set IEnumerator_Target = Target
End Property

'* この列挙子を初期化します。
'*
'* @param TargetTable 列挙対象の `WorksheetVirtualTable`。
'* @param Descending [省略可] 降順で列挙するか否か。
'*
'* @details
'* `TargetTable` が Nothing の場合、または初期化済みの列挙子を再初期化した場合はエラーになります。
Public Sub Initialize( _
        ByVal TargetTable As WorksheetVirtualTable, _
        Optional ByVal Descending As Boolean = False)

    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator", "初期化済みです。"
    If TargetTable Is Nothing Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator", "列挙対象が Nothing です。"

    Set pTable = TargetTable
    pLength = TargetTable.RowCount
    pDescending = Descending
    pIsInitialized = True

    Call Reset
End Sub

'* 次の要素に移動します。
'*
'* @return 次の要素が存在する場合は True、それ以外は False。
'*
'* @details
'* 初期位置は先頭の 1 つ前です。降順の場合は末尾の 1 つ後から開始します。
Public Function MoveNext() As Boolean
    Call pCheckInit
    If Descending Then
        If 0 < pIndex Then
            pIndex = pIndex - 1
            MoveNext = True
        Else
            pIndex = -1
            MoveNext = False
        End If
    Else
        If pIndex < pLength - 1 Then
            pIndex = pIndex + 1
            MoveNext = True
        Else
            pIndex = pLength
            MoveNext = False
        End If
    End If
End Function
Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = MoveNext()
End Function

'* 列挙をリセットします。
'*
'* @details
'* 現在位置を、昇順では先頭の 1 つ前、降順では末尾の 1 つ後へ戻します。
Public Sub Reset()
    Call pCheckInit
    If Descending Then
        pIndex = pLength
    Else
        pIndex = -1
    End If
End Sub
Private Sub IEnumerator_Reset()
    Call Reset
End Sub

'* 指定した 0-based index までスキップします。
'*
'* @param TargetIndex 移動先の 0-based index。
'*
'* @details
'* 範囲外の index、または現在位置より前へ戻る index を指定した場合はエラーになります。
Public Sub SkipTo(ByVal TargetIndex As Long)
    Call pCheckInit
    If TargetIndex < 0 Or pLength <= TargetIndex Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator", "移動先が範囲外です。(min: 0 / max: " & (pLength - 1) & " / target: " & TargetIndex & ")"
    If Descending Then
        If pIndex < TargetIndex Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator", "現在の位置より前には移動できません。(current: " & pIndex & " / target: " & TargetIndex & ")"
    Else
        If TargetIndex < pIndex Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator", "現在の位置より前には移動できません。(current: " & pIndex & " / target: " & TargetIndex & ")"
    End If
    pIndex = TargetIndex
End Sub
Private Sub IEnumerator_SkipTo(ByVal TargetIndex As Long)
    Call SkipTo(TargetIndex)
End Sub

'* 現在の要素の削除を試行します。
'*
'* @return 削除された要素。
'*
'* @details
'* `WorksheetVirtualTableEnumerator` は読み取り専用のため、常にエラーになります。
Public Function Remove() As Variant
    Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator", "このコレクションは読み取り専用です。要素の削除に対応していません。"
End Function
Private Function IEnumerator_Remove() As Variant
    Set IEnumerator_Remove = Remove()
End Function

'* 現在の要素の更新を試行します。
'*
'* @param ItemObject 更新する要素。
'*
'* @details
'* `WorksheetVirtualTableEnumerator` は読み取り専用のため、常にエラーになります。
Public Sub Update(ByVal ItemObject As Variant)
    Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator", "このコレクションは読み取り専用です。要素の更新に対応していません。"
End Sub
Private Sub IEnumerator_Update(ByVal ItemObject As Variant)
    Call Update(ItemObject)
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator", "未初期化状態です。"
End Sub

Private Sub pCheckEndOfList()
    If pIndex < 0 Or pLength <= pIndex Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator", "現在の要素がありません。(index: " & pIndex & " / min: 0 / max: " & (pLength - 1) & ")"
    End If
End Sub
