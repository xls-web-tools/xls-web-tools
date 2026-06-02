VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetRangeBoundsEnumerator"
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
'! WorksheetRangeBounds の行、列、セルを列挙するための列挙子のクラスです。
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pDescending As Boolean
Private pEnumerationMode As Long
Private pRangeBounds As WorksheetRangeBounds
Private pLength As Long
Private pIndex As Long

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

'* 列挙方式。
Public Property Get EnumerationMode() As Long
    Call pCheckInit
    EnumerationMode = pEnumerationMode
End Property

' 現在の要素
Public Property Get Current() As Variant
    Call pCheckInit
    Call pCheckEndOfList

    Select Case pEnumerationMode
     Case G_RANGE_ENUM_MODE_CELLS_HORIZONTAL
        Set Current = pRangeBounds.Item(pIndex, ColumnDirection:=False)
     Case G_RANGE_ENUM_MODE_CELLS_VERTICAL
        Set Current = pRangeBounds.Item(pIndex, ColumnDirection:=True)
     Case G_RANGE_ENUM_MODE_COLUMNS
        Set Current = pRangeBounds.GetColumn(pIndex + 1)
     Case Else
        Set Current = pRangeBounds.GetRow(pIndex + 1)
    End Select
End Property
Private Property Get IEnumerator_Current() As Variant
    Set IEnumerator_Current = Current
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
    Set Target = pRangeBounds
End Property
Private Property Set Target(ByVal NewValue As Object)
    Set pRangeBounds = NewValue

    ' 長さを計算
    Call pCalculateLength
End Property

Private Property Get IEnumerator_Target() As Variant
    Set IEnumerator_Target = Target
End Property

'* このオブジェクトを初期化します。
'*
'* @param TargetCollection 列挙対象の WorksheetRangeBounds。Nothing はエラーです。
'* @param EnumerationMode [省略可] 列挙方式。
'* @param Descending [省略可] 降順で列挙するか否か。
'*
'* @details
'* このオブジェクトを初期化します。TargetCollection は WorksheetRangeBounds.GetEnumerator から渡される前提です。
Public Sub Initialize( _
        ByVal TargetCollection As WorksheetRangeBounds, _
        Optional ByVal EnumerationMode As Long = G_RANGE_ENUM_MODE_ROWS, _
        Optional ByVal Descending As Boolean = False)

    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator", "初期化済みです。"
    If TargetCollection Is Nothing Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator", "列挙対象が Nothing です。"

    Set Target = TargetCollection
    Call pSetEnumerationMode(EnumerationMode)
    pDescending = Descending

    pIsInitialized = True

    Call Reset
End Sub

Private Sub pSetEnumerationMode(ByVal NewValue As Long)
    Call pCheckEnumerationMode(NewValue)
    pEnumerationMode = NewValue

    ' 設定に合わせて長さを再計算
    Call pCalculateLength
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

' 列挙をリセットします。
'
' @details
' 列挙の現在位置をリストの先頭の 1 つ前 (逆順のときはリストの末尾の 1 つ後) に戻します。
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

' 指定した位置までスキップします。
'
' @param TargetIndex 移動先のインデックス。
'
' @details
' 指定されたインデックスまで現在位置を進めます。現在位置より前やリストの範囲外を指定するとエラーを発生させます。
Public Sub SkipTo(ByVal TargetIndex As Long)
    Call pCheckInit
    If TargetIndex < 0 Or pLength <= TargetIndex Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator", "移動先が範囲外です。(min: 0 / max: " & (pLength - 1) & " / target: " & TargetIndex & ")"
    If Descending Then
        If pIndex < TargetIndex Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator", "現在の位置より前には移動できません。(current: " & pIndex & " / target: " & TargetIndex & ")"
    Else
        If TargetIndex < pIndex Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator", "現在の位置より前には移動できません。(current: " & pIndex & " / target: " & TargetIndex & ")"
    End If
    pIndex = TargetIndex
End Sub
Private Sub IEnumerator_SkipTo(ByVal TargetIndex As Long)
    Call SkipTo(TargetIndex)
End Sub

' 現在の要素の削除を試行します。
'
' @return Variant 削除された要素。
'
' @details
' WorksheetRangeBoundsEnumerator は読み取り専用です。
' このメソッドは常にエラーを発生させます。
Public Function Remove() As Variant
    Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator", "このコレクションは読み取り専用です。要素の削除に対応していません。"
End Function
Private Function IEnumerator_Remove() As Variant
    Set IEnumerator_Remove = Remove()
End Function

' 現在の要素の更新を試行します。
'
' @param ItemObject 更新する要素の値。
'
' @details
' WorksheetRangeBoundsEnumerator は読み取り専用です。
' このメソッドは常にエラーを発生させます。
Public Sub Update(ByVal ItemObject As Variant)
    Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator", "このコレクションは読み取り専用です。要素の更新に対応していません。"
End Sub
Private Sub IEnumerator_Update(ByVal ItemObject As Variant)
    Call Update(ItemObject)
End Sub

Private Sub pCheckEnumerationMode(ByVal TargetEnumerationMode As Long)
    Select Case TargetEnumerationMode
     Case G_RANGE_ENUM_MODE_ROWS
     Case G_RANGE_ENUM_MODE_COLUMNS
     Case G_RANGE_ENUM_MODE_CELLS_HORIZONTAL
     Case G_RANGE_ENUM_MODE_CELLS_VERTICAL
     Case Else
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator", "列挙方式が不正です。(" & TargetEnumerationMode & ")"
    End Select
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator", "未初期化状態です。"
End Sub

Private Sub pCalculateLength()
    If pRangeBounds Is Nothing Then Exit Sub
    If pRangeBounds.IsEmpty Then
        pLength = 0
        Exit Sub
    End If
    Select Case pEnumerationMode
     Case G_RANGE_ENUM_MODE_CELLS_HORIZONTAL, G_RANGE_ENUM_MODE_CELLS_VERTICAL
        pLength = pRangeBounds.Count
     Case G_RANGE_ENUM_MODE_COLUMNS
        pLength = pRangeBounds.ColumnCount
     Case Else
        pLength = pRangeBounds.RowCount
    End Select
End Sub

Private Sub pCheckEndOfList()
    If pIndex < 0 Or pLength <= pIndex Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator", "現在の要素がありません。(index: " & pIndex & " / min: 0 / max: " & (pLength - 1) & ")"
    End If
End Sub
