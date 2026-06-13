VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetVirtualTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 複数の WorksheetRangeBounds を header と相対行で対応付ける仮想表です。
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pHeaders() As String
Private pColumnRangeList As ObjectList
Private pCount As Long
Private pHeaderCompareMode As VbCompareMethod
Private pHeaderRowCount As Long
Private pNewEnumSnapshot As Collection

'* 仮想表の行数
Public Property Get Count() As Long
    Call pCheckInit
    Count = pCount
End Property

'* 仮想表の行数
Public Property Get RowCount() As Long
    RowCount = Count
End Property

'* 解決済み header 名の配列 copy
Public Property Get Headers() As String()
    Call pCheckInit
    Headers = pCopyHeaders()
End Property

'* 標準 For Each 用の列挙子
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Call pCheckInit
    If pNewEnumSnapshot Is Nothing Then Set pNewEnumSnapshot = pCreateNewEnumSnapshot()

    Set NewEnum = pNewEnumSnapshot.[_NewEnum]
End Property

'* 指定された 0-based index の行を返します。
'*
'* @param ItemIndex 取得する行の 0-based index。
'* @return header を key、対応する 1 行分の `WorksheetRangeBounds` を値に持つ `ObjectDictionary`。
'*
'* @details
'* 短い field の不足行は、元 field の幅を維持した RowCount = 0 の空範囲として返します。
Public Function Item(ByVal ItemIndex As Long) As ObjectDictionary
    Call pCheckInit
    Call pCheckItemIndex(ItemIndex)

    Dim result_value As ObjectDictionary
    Set result_value = New_ObjectDictionary("WorksheetRangeBounds")
    result_value.CompareMode = pHeaderCompareMode

    Dim col_idx As Long
    For col_idx = 0 To pColumnRangeList.Count - 1
        Dim row_bounds As WorksheetRangeBounds
        Set row_bounds = pGetFieldRowBounds(pColumnRangeList.Item(col_idx), ItemIndex)
        Call result_value.Add(pHeaders(col_idx), row_bounds)
    Next col_idx

    Set Item = result_value
End Function

Private Function pGetFieldRowBounds(ByVal FieldRange As WorksheetRangeBounds, ByVal ItemIndex As Long) As WorksheetRangeBounds
    Dim field_row_index As Long
    field_row_index = ItemIndex + pHeaderRowCount + 1

    If field_row_index <= FieldRange.RowCount Then
        Set pGetFieldRowBounds = FieldRange.GetRow(field_row_index)
        Exit Function
    End If

    Set pGetFieldRowBounds = FieldRange.TransformAbsolute(Row:=FieldRange.Row + field_row_index - 1, FinishRow:=FieldRange.Row + field_row_index - 2)
End Function

Private Sub pCheckItemIndex(ByVal ItemIndex As Long)
    If ItemIndex < 0 Or pCount <= ItemIndex Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "行 index が範囲外です。(" & CStr(ItemIndex) & ")"
    End If
End Sub

'* 仮想表を初期化します。
'*
'* @param ColumnRangeList 仮想表の field として扱う `WorksheetRangeBounds` の `ObjectList`。
'* @param HeaderNames [省略可] header 名の 1 次元文字列配列。
'* @param TreatFirstRowAsHeader [省略可] True の場合は各 field 範囲の 1 行目を header として読み取ります。
'* @param HeaderCompareMode [省略可] header 重複判定と行 `ObjectDictionary` の比較モード。
'* @param HeaderJoinDelimiter [省略可] 複数列 field の header セルを連結するときの区切り文字。
'*
'* @details
'* `ColumnRangeList` は `ObjectList("WorksheetRangeBounds")` の型契約を持つ必要があります。
'* `TreatFirstRowAsHeader=True` の場合、`HeaderNames` は無視されます。
'* header 件数不一致、重複 header、または `HeaderNames` の型不一致は初期化エラーになります。
Public Sub Initialize( _
        ByVal ColumnRangeList As ObjectList, _
        Optional ByVal HeaderNames As Variant, _
        Optional ByVal TreatFirstRowAsHeader As Boolean = False, _
        Optional ByVal HeaderCompareMode As VbCompareMethod = vbBinaryCompare, _
        Optional ByVal HeaderJoinDelimiter As String = "")

    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "すでに初期化済みです。"
    Call pValidateColumnRangeList(ColumnRangeList)
    If TreatFirstRowAsHeader Then
        pHeaderRowCount = 1
        pHeaders = pReadHeaderNames(ColumnRangeList, HeaderJoinDelimiter)
    Else
        pHeaderRowCount = 0
        If IsMissing(HeaderNames) Then
            Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "HeaderNames が指定されていません。"
        End If
        pHeaders = pGetHeaderNamesArray(HeaderNames)
    End If

    If pHeaderCount() <> ColumnRangeList.Count Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "HeaderNames と ColumnRangeList の件数が一致しません。"
    End If

    pHeaderCompareMode = HeaderCompareMode
    Call pValidateDuplicateHeaders

    Set pColumnRangeList = ColumnRangeList.CopyList()
    pCount = pCalculateRowCount()
    pIsInitialized = True
End Sub

'* WorksheetRangeBounds から仮想表を初期化します。
'*
'* @param TableRange 仮想表の列に分割する WorksheetRangeBounds。
'* @param HeaderNames [省略可] header 名の 1 次元文字列配列。
'* @param TreatFirstRowAsHeader [省略可] True の場合は各 field 範囲の 1 行目を header として読み取ります。
'* @param HeaderCompareMode [省略可] header 重複判定と行 ObjectDictionary の比較モード。
'* @param HeaderJoinDelimiter [省略可] 複数列 field の header セルを連結するときの区切り文字。
'*
'* @details
'* TableRange.GetColumns() で列単位の WorksheetRangeBounds 一覧を取得し、既存の Initialize に委譲します。
Public Sub InitializeFromRangeBounds( _
        ByVal TableRange As WorksheetRangeBounds, _
        Optional ByVal HeaderNames As Variant, _
        Optional ByVal TreatFirstRowAsHeader As Boolean = False, _
        Optional ByVal HeaderCompareMode As VbCompareMethod = vbBinaryCompare, _
        Optional ByVal HeaderJoinDelimiter As String = "")

    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "すでに初期化済みです。"
    If TableRange Is Nothing Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "TableRange が指定されていません。"

    Call Initialize( _
            ColumnRangeList:=TableRange.GetColumns(), _
            HeaderNames:=HeaderNames, _
            TreatFirstRowAsHeader:=TreatFirstRowAsHeader, _
            HeaderCompareMode:=HeaderCompareMode, _
            HeaderJoinDelimiter:=HeaderJoinDelimiter)
End Sub

Private Sub pValidateColumnRangeList(ByVal ColumnRangeList As ObjectList)
    If ColumnRangeList Is Nothing Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "ColumnRangeList が指定されていません。"
    End If
    If Not ColumnRangeList.HasItemTypeContract Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "ColumnRangeList には WorksheetRangeBounds の型契約が必要です。"
    End If
    If StrComp(ColumnRangeList.ElementTypeName, "WorksheetRangeBounds", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "ColumnRangeList の要素型契約は WorksheetRangeBounds を指定してください。"
    End If
End Sub

Private Function pGetHeaderNamesArray(ByVal HeaderNames As Variant) As String()
    Call pValidateHeaderNamesType(HeaderNames)

    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Call GetArrayBounds(lbound_arr, ubound_arr, HeaderNames)
    If LBound(lbound_arr) <> UBound(lbound_arr) Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "HeaderNames には 1 次元の文字列配列を指定してください。"
    End If

    If IsEmptyArray(HeaderNames) Then
        pGetHeaderNamesArray = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    ReDim result_value(0 To ubound_arr(LBound(ubound_arr)) - lbound_arr(LBound(lbound_arr)))

    Dim src_idx As Long
    Dim dst_idx As Long
    dst_idx = 0
    For src_idx = lbound_arr(LBound(lbound_arr)) To ubound_arr(LBound(ubound_arr))
        result_value(dst_idx) = CStr(HeaderNames(src_idx))
        dst_idx = dst_idx + 1
    Next src_idx

    pGetHeaderNamesArray = result_value
End Function

Private Function pReadHeaderNames(ByVal ColumnRangeList As ObjectList, ByVal HeaderJoinDelimiter As String) As String()
    If ColumnRangeList.Count = 0 Then
        pReadHeaderNames = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    ReDim result_value(0 To ColumnRangeList.Count - 1)

    Dim col_idx As Long
    For col_idx = 0 To ColumnRangeList.Count - 1
        result_value(col_idx) = pBuildHeaderName(ColumnRangeList.Item(col_idx), HeaderJoinDelimiter)
    Next col_idx

    pReadHeaderNames = result_value
End Function

Private Sub pValidateDuplicateHeaders()
    Dim header_dict As Object
    Set header_dict = CreateObject("Scripting.Dictionary")
    header_dict.CompareMode = pHeaderCompareMode

    Dim header_idx As Long
    For header_idx = LBound(pHeaders) To UBound(pHeaders)
        If header_dict.Exists(pHeaders(header_idx)) Then
            Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "header が重複しています。(" & pHeaders(header_idx) & ")"
        End If
        Call header_dict.Add(pHeaders(header_idx), True)
    Next header_idx
End Sub

Private Function pCalculateRowCount() As Long
    Dim result_value As Long
    result_value = 0

    Dim col_idx As Long
    For col_idx = 0 To pColumnRangeList.Count - 1
        Dim field_range As WorksheetRangeBounds
        Set field_range = pColumnRangeList.Item(col_idx)

        Dim field_row_count As Long
        field_row_count = field_range.RowCount - pHeaderRowCount
        If field_row_count < 0 Then field_row_count = 0
        If result_value < field_row_count Then result_value = field_row_count
    Next col_idx

    pCalculateRowCount = result_value
End Function

Private Function pHeaderCount() As Long
    If IsEmptyArray(pHeaders) Then
        pHeaderCount = 0
    Else
        pHeaderCount = UBound(pHeaders) - LBound(pHeaders) + 1
    End If
End Function

'* 指定された行 index の行を返します。
'*
'* @param RowIndex 行 index。正数は上からの 1 オリジン、負数は下からの 1 オリジン。0 は範囲外エラー。
'* @return header を key、1 行分の `WorksheetRangeBounds` を値に持つ `ObjectDictionary`。
'*
'* @details
'* `WorksheetRangeBounds.GetRow` と同じ行 index 規則で行を取得します。
Public Function GetRow(ByVal RowIndex As Long) As ObjectDictionary
    Call pCheckInit
    If RowIndex = 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "行 index が範囲外です。(0)"
    End If

    Dim item_idx As Long
    If 0 < RowIndex Then
        item_idx = RowIndex - 1
    Else
        item_idx = pCount + RowIndex
    End If

    Set GetRow = Item(item_idx)
End Function

'* この仮想表の列挙子を返します。
'*
'* @param Descending [省略可] 降順で列挙するか否か。
'* @return 行 `ObjectDictionary` を返す列挙子。
'*
'* @details
'* 返される列挙子は読み取り専用です。`IEnumerator.Index` は 0-based です。
Public Function GetEnumerator(Optional ByVal Descending As Boolean = False) As IEnumerator
    Call pCheckInit

    Dim result_value As WorksheetVirtualTableEnumerator
    Set result_value = New WorksheetVirtualTableEnumerator
    Call result_value.Initialize(Me, Descending:=Descending)

    Set GetEnumerator = result_value
End Function

Private Sub pValidateHeaderNamesType(ByVal HeaderNames As Variant)
    If Not IsArray(HeaderNames) Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "HeaderNames には 1 次元の文字列配列を指定してください。"
    End If

    Dim header_var_type As Long
    header_var_type = VarType(HeaderNames)
    If (header_var_type And vbArray) = 0 Or header_var_type - vbArray <> vbString Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "HeaderNames には 1 次元の文字列配列を指定してください。"
    End If
End Sub

Private Function pBuildHeaderName(ByVal FieldRange As WorksheetRangeBounds, ByVal HeaderJoinDelimiter As String) As String
    If FieldRange.RowCount = 0 Then
        pBuildHeaderName = ""
        Exit Function
    End If

    Dim header_row As WorksheetRangeBounds
    Set header_row = FieldRange.GetRow(1)

    Dim header_parts() As String
    ReDim header_parts(0 To header_row.ColumnCount - 1)

    Dim col_idx As Long
    For col_idx = 1 To header_row.ColumnCount
        header_parts(col_idx - 1) = WsSrv.ReadCell(header_row.GetCell(1, col_idx))
    Next col_idx

    pBuildHeaderName = Join(header_parts, HeaderJoinDelimiter)
End Function

Private Function pCopyHeaders() As String()
    If pHeaderCount() = 0 Then
        pCopyHeaders = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    ReDim result_value(LBound(pHeaders) To UBound(pHeaders))

    Dim item_idx As Long
    For item_idx = LBound(pHeaders) To UBound(pHeaders)
        result_value(item_idx) = pHeaders(item_idx)
    Next item_idx

    pCopyHeaders = result_value
End Function

Private Function pCreateNewEnumSnapshot() As Collection
    Dim result_value As Collection
    Set result_value = New Collection

    Dim enum_obj As IEnumerator
    Set enum_obj = Me.GetEnumerator()
    Do While enum_obj.MoveNext()
        Call result_value.Add(enum_obj.Current)
    Loop

    Set pCreateNewEnumSnapshot = result_value
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable", "未初期化状態です。"
End Sub
