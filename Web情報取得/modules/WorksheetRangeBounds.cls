VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetRangeBounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IStringable
Implements IEquatable

' #############################################################################
'!
'! @brief
'! Excel の範囲情報を保持するクラス モジュールです。
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pRow As Long
Private pColumn As Long
Private pFinishRow As Long
Private pFinishColumn As Long
Private pWorkbookName As String
Private pWorksheetName As String
Private pNewEnumSnapshot As Collection

'* 行番号
Public Property Get Row() As Long
    Call pCheckInit
    Row = pRow
End Property

'* 列番号
Public Property Get Column() As Long
    Call pCheckInit
    Column = pColumn
End Property

'* 最後の行番号
Public Property Get FinishRow() As Long
    Call pCheckInit
    FinishRow = pFinishRow
End Property

'* 最後の列番号
Public Property Get FinishColumn() As Long
    Call pCheckInit
    FinishColumn = pFinishColumn
End Property

'* ワークシート名
Public Property Get WorksheetName() As String
    Call pCheckInit
    WorksheetName = pWorksheetName
End Property

'* ワークブック名
Public Property Get WorkbookName() As String
    Call pCheckInit
    WorkbookName = pWorkbookName
End Property

'* 行数
Public Property Get RowCount() As Long
    Call pCheckInit
    Dim result_value As Long
    result_value = pFinishRow - pRow + 1
    If 0 < result_value Then
        RowCount = result_value
    Else
        RowCount = 0
    End If
End Property

'* 列数
Public Property Get ColumnCount() As Long
    Call pCheckInit
    Dim result_value As Long
    result_value = pFinishColumn - pColumn + 1
    If 0 < result_value Then
        ColumnCount = result_value
    Else
        ColumnCount = 0
    End If
End Property

'* セル数
Public Property Get Count() As Long
    Call pCheckInit

    Count = RowCount * ColumnCount

'    Dim row_count As Long, col_count As Long
'    row_count = RowCount
'    col_count = ColumnCount
'
'    If row_count = 0 And col_count = 0 Then
'        Count = 0
'    Else
'        Count = MaxLng(RowCount, 1) * MaxLng(ColumnCount, 1)
'    End If
End Property

'* 空かどうか
Public Property Get IsEmpty() As Boolean
    Call pCheckInit
    If pFinishRow < pRow Or pFinishColumn < pColumn Then
        IsEmpty = True
    Else
        IsEmpty = False
    End If
End Property

'* シート全体かどうか
Public Property Get IsEntireSheet() As Boolean
    Call pCheckInit
    If IsEntireRow And IsEntireColumn Then
        IsEntireSheet = True
    Else
        IsEntireSheet = False
    End If
End Property

'* 行全体 (1:2 のような範囲) かどうか
Public Property Get IsEntireRow() As Boolean
    Call pCheckInit
    If pColumn = 1 And pFinishColumn = G_COL_MAX Then
        IsEntireRow = True
    Else
        IsEntireRow = False
    End If
End Property

'* 列全体 (A:B のような範囲) かどうか
Public Property Get IsEntireColumn() As Boolean
    Call pCheckInit
    If pRow = 1 And pFinishRow = G_ROW_MAX Then
        IsEntireColumn = True
    Else
        IsEntireColumn = False
    End If
End Property

'* 単一セルかどうか。Area でない非空範囲と同義です。
Public Property Get IsCell() As Boolean
    Call pCheckInit
    If Not IsEmpty And pRow = pFinishRow And pColumn = pFinishColumn Then
        IsCell = True
    Else
        IsCell = False
    End If
End Property

'* Area かどうか。単一セルを除く、連続した複数セル範囲を Area として扱います。
Public Property Get IsArea() As Boolean
    Call pCheckInit
    If Not IsEmpty And Not IsCell() Then
        IsArea = True
    Else
        IsArea = False
    End If
End Property

'* 幅が 1 行の範囲 (1:1 や A1:B1) かどうか。単一セルも True となります。
Public Property Get IsOneRow() As Boolean
    Call pCheckInit
    If IsEmpty Then
        IsOneRow = False
        Exit Property
    End If

    If pRow = pFinishRow Then
        IsOneRow = True
    Else
        IsOneRow = False
    End If
End Property

'* 幅が 1 列の範囲 (A:A や A1:A2) かどうか。単一セルも True となります。
Public Property Get IsOneColumn() As Boolean
    Call pCheckInit
    If IsEmpty Then
        IsOneColumn = False
        Exit Property
    End If

    If pColumn = pFinishColumn Then
        IsOneColumn = True
    Else
        IsOneColumn = False
    End If
End Property

'* 1 行だけで構成される Area かどうか。単一セルは False となります。
Public Property Get IsOneRowArea() As Boolean
    Call pCheckInit
    If IsArea And IsOneRow Then
        IsOneRowArea = True
    Else
        IsOneRowArea = False
    End If
End Property

'* 1 列だけで構成される Area かどうか。単一セルは False となります。
Public Property Get IsOneColumnArea() As Boolean
    Call pCheckInit
    If IsArea And IsOneColumn Then
        IsOneColumnArea = True
    Else
        IsOneColumnArea = False
    End If
End Property

'* 標準 For Each 用の列挙子を得ます。
'*
'* @return 標準 For Each が使用する列挙子。
'*
'* @details
'* 列挙開始時点のセル要素スナップショットを、行方向かつ昇順で列挙します。
'* 標準 For Each の列挙方式は固定であり、GetEnumerator の引数指定とは独立しています。
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    If pNewEnumSnapshot Is Nothing Then
        Set pNewEnumSnapshot = pCreateNewEnumSnapshot()
    End If

    Set NewEnum = pNewEnumSnapshot.[_NewEnum]
End Property

'* 番号を指定してセル範囲を取得します。
'*
'* @param ItemIndex 番号。
'* @param ColumnDirection [省略可] 列 (縦) 方向に走査するか。デフォルトは False で、行 (横) 方向に走査します。
'*
'* @details
'* 番号を指定してセル範囲を取得します。
Public Function Item(ByVal ItemIndex As Long, Optional ByVal ColumnDirection As Boolean = False) As WorksheetRangeBounds
    Call pCheckInit
    If ItemIndex < 0 Then
        Err.Raise vbObjectError, "Class WorksheetRangeBounds", "インデックスが範囲外です。"
    End If

    If ColumnDirection Then
        Set Item = pGetCellVertical(ItemIndex)
    Else
        Set Item = pGetCellHorizontal(ItemIndex)
    End If
End Function

Private Function pGetCellVertical(ByVal ItemIndex As Long) As WorksheetRangeBounds
    If IsEmpty Then
        Err.Raise vbObjectError, "Class WorksheetRangeBounds", "範囲が空です。"
    End If

    Dim row_count As Long, row_idx As Long, col_idx As Long
    row_count = RowCount
    row_idx = ItemIndex Mod row_count
    col_idx = (ItemIndex - row_idx) / row_count

    If col_idx < ColumnCount Then
        Set pGetCellVertical = GetCell(row_idx + 1, col_idx + 1)
    Else
        Err.Raise vbObjectError, "Class WorksheetRangeBounds", "インデックスが範囲外です。"
    End If
End Function

Private Function pGetCellHorizontal(ByVal ItemIndex As Long) As WorksheetRangeBounds
    If IsEmpty Then
        Err.Raise vbObjectError, "Class WorksheetRangeBounds", "範囲が空です。"
    End If

    Dim col_count As Long, row_idx As Long, col_idx As Long
    col_count = ColumnCount
    col_idx = ItemIndex Mod col_count
    row_idx = (ItemIndex - col_idx) / col_count

    If row_idx < RowCount Then
        Set pGetCellHorizontal = GetCell(row_idx + 1, col_idx + 1)
    Else
        Err.Raise vbObjectError, "Class WorksheetRangeBounds", "インデックスが範囲外です。"
    End If
End Function

'* 各値を設定します。
'*
'* @param RowIndex 行番号。
'* @param ColumnIndex 列番号。
'* @param FinishRowIndex 最後の行番号。
'* @param FinishColumnIndex 最後の列番号。
'* @param SheetName ワークシート名。
'* @param BookName ワークブック名。
'*
'* @details
'* 各値を設定します。
'*
'* 使用例:
'* @code
'* ' 1 セル (B5)
'* Dim cell_bounds As WorksheetRangeBounds
'* Set cell_bounds = New WorksheetRangeBounds
'* Call cell_bounds.Initialize(Row:=5, Column:=2)
'*
'* ' 1 行全体 (5:5)
'* Dim row_bounds As WorksheetRangeBounds
'* Set row_bounds = New WorksheetRangeBounds
'* Call row_bounds.Initialize(Row:=5)
'*
'* ' 2 列全体 (B:C)
'* Dim two_columns_bounds As WorksheetRangeBounds
'* Set two_columns_bounds = New WorksheetRangeBounds
'* Call two_columns_bounds.Initialize(Column:=2, FinishColumn:=3)
'*
'* ' 表全体 (B2:E20)
'* Dim table_bounds As WorksheetRangeBounds
'* Set table_bounds = New WorksheetRangeBounds
'* Call table_bounds.Initialize(Row:=2, Column:=2, FinishRow:=20, FinishColumn:=5)
'* @endcode
Public Sub Initialize( _
        Optional ByVal Row As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Column As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishColumn As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Sheet As String = "Sheet1", _
        Optional ByVal Book As String = "")

    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "すでに初期化済みです。"

'    If 0 < FinishRow <> G_OMIT_CELL_INDEX And FinishRow < Row Then
'        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "開始行と終了行が逆転しています。(" & Row & ", " & FinishRow & ")"
'        Exit Sub
'    End If
'
'    If 0 < FinishColumn And FinishColumn < Column Then
'        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "開始列と終了列が逆転しています。(" & Column & ", " & FinishColumn & ")"
'        Exit Sub
'    End If

    If Sheet = "" Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "ワークシート名は省略できません。"
        Exit Sub
    End If

    If Book = "" Then
        Book = pGetDefaultWorkbookName()
    End If

    Call pWellFormBounds(Row, Column, FinishRow, FinishColumn)

    pRow = Row
    pColumn = Column
    pFinishRow = FinishRow
    pFinishColumn = FinishColumn
    pWorksheetName = Sheet
    pWorkbookName = Book
    Set pNewEnumSnapshot = Nothing
    pIsInitialized = True
End Sub

'* Excel アドレス文字列から初期化します。
'*
'* @param AddressString 初期化に使用する Excel アドレス文字列。
'*
'* @details
'* A1 形式の単一矩形範囲だけを扱います。複数範囲は WorksheetRangeBounds で表現できないためエラーになります。
Public Sub InitializeFromAddress(ByVal AddressString As String)
    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "すでに初期化済みです。"

    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, AddressString)

    If folder_path <> "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WorksheetRangeBounds", Description:="フォルダパス付きアドレスは扱えません。(" & AddressString & ")"
    End If

    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, cell_address)

    If sheet_name = "" Then
        sheet_name = "Sheet1"
    End If

    Call Initialize( _
            Row:=start_row, _
            Column:=start_col, _
            FinishRow:=finish_row, _
            FinishColumn:=finish_col, _
            Sheet:=sheet_name, _
            Book:=book_name)
End Sub

Private Function pGetDefaultWorkbookName() As String
    Call pRequireWorkbookService("pGetDefaultWorkbookName")
    pGetDefaultWorkbookName = WbSrv.GetThisWorkbookName()
End Function

Private Sub pWellFormBounds(ByRef Row As Long, ByRef Column As Long, ByRef FinishRow As Long, ByRef FinishColumn As Long)
    If Row = G_OMIT_CELL_INDEX Then
        Row = 1
        If FinishRow = G_OMIT_CELL_INDEX Then
            FinishRow = G_ROW_MAX
        End If
    ElseIf Row < 1 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "開始行は 1 以上を指定してください。(" & Row & ")"
    ElseIf G_ROW_MAX < Row Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "開始行が " & G_ROW_MAX & " を超過しています。(" & Row & ")"
    End If

    If Column = G_OMIT_CELL_INDEX Then
        Column = 1
        If FinishColumn = G_OMIT_CELL_INDEX Then
            FinishColumn = G_COL_MAX
        End If
    ElseIf Column < 1 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "開始列は 1 以上を指定してください。(" & Column & ")"
    ElseIf G_COL_MAX < Column Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "開始列が " & G_COL_MAX & " を超過しています。(" & Column & ")"
    End If

    If FinishRow = G_OMIT_CELL_INDEX Then
        FinishRow = Row
    ElseIf FinishRow < 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "終了行に負数は指定できません。(" & FinishRow & ")"
    ElseIf FinishRow < Row Then
        FinishRow = 0
    End If

    If FinishColumn = G_OMIT_CELL_INDEX Then
        FinishColumn = Column
    ElseIf FinishColumn < 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "終了列に負数は指定できません。(" & FinishColumn & ")"
    ElseIf FinishColumn < Column Then
        FinishColumn = 0
    End If

    If G_ROW_MAX < FinishRow Then FinishRow = G_ROW_MAX
    If G_COL_MAX < FinishColumn Then FinishColumn = G_COL_MAX

'    If FinishRow = 0 Or FinishColumn = 0 Then
'        FinishRow = 0
'        FinishColumn = 0
'    End If
End Sub

'* 最上部の行からの行インデックスを指定して、1 行の範囲を得ます。
'*
'* @param RowIndex 行インデックス。正数は上からの 1 オリジン、負数は下からの 1 オリジン。0 は範囲外エラー。
'* @return 指定された範囲 (1 行)
'*
'* @details
'* 正数では最上部、負数では最下部から見た行インデックスを指定して、1 行の範囲を得ます。
Public Function GetRow(ByVal RowIndex As Long) As WorksheetRangeBounds
    Call pCheckInit
    Dim target_row As Long: target_row = pGetTargetRow(RowIndex)

    Dim result_value As WorksheetRangeBounds
    Set result_value = New_RangeBounds(Row:=target_row, Column:=pColumn, FinishRow:=target_row, FinishColumn:=pFinishColumn, Sheet:=pWorksheetName, Book:=pWorkbookName)
    Set GetRow = result_value
End Function

'* 最も左の列からの列インデックスを指定して、1 列の範囲を得ます。
'*
'* @param ColumnIndex 列インデックス。正数は左からの 1 オリジン、負数は右からの 1 オリジン。0 は範囲外エラー。
'* @return 指定された範囲 (1 列)
'*
'* @details
'* 正数では最も左、負数では最も右から見た列インデックスを指定して、1 列の範囲を得ます。
Public Function GetColumn(ByVal ColumnIndex As Long) As WorksheetRangeBounds
    Call pCheckInit
    Dim target_col As Long: target_col = pGetTargetColumn(ColumnIndex)

    Dim result_value As WorksheetRangeBounds
    Set result_value = New_RangeBounds(Row:=pRow, Column:=target_col, FinishRow:=pFinishRow, FinishColumn:=target_col, Sheet:=pWorksheetName, Book:=pWorkbookName)
    Set GetColumn = result_value
End Function

'* 最も左上のセルからの行インデックスおよび列インデックスを指定して、1 セルの範囲を得ます。
'*
'* @param RowIndex 行インデックス。正数は上からの 1 オリジン、負数は下からの 1 オリジン。0 は範囲外エラー。
'* @param ColumnIndex 列インデックス。正数は左からの 1 オリジン、負数は右からの 1 オリジン。0 は範囲外エラー。
'* @return 指定された範囲 (1 セル)
'*
'* @details
'* 正数では左上、負数では右下を基準にした行インデックスおよび列インデックスを指定して、1 セルの範囲を得ます。
Public Function GetCell(ByVal RowIndex As Long, ByVal ColumnIndex As Long) As WorksheetRangeBounds
    Call pCheckInit
    Dim target_row As Long: target_row = pGetTargetRow(RowIndex)
    Dim target_col As Long: target_col = pGetTargetColumn(ColumnIndex)

    Dim result_value As WorksheetRangeBounds
    Set result_value = New_RangeBounds(Row:=target_row, Column:=target_col, FinishRow:=target_row, FinishColumn:=target_col, Sheet:=pWorksheetName, Book:=pWorkbookName)
    Set GetCell = result_value
End Function

'* この範囲を変形した新しい範囲を得ます。
'*
'* @param Row 行番号。
'* @param Column 列番号。
'* @param FinishRow 最後の行番号。
'* @param FinishColumn 最後の列番号。
'* @return 変形した新しいオブジェクト
'*
'* @details
'* この範囲を変形した新しい範囲を得ます。
Public Function TransformAbsolute( _
        Optional ByVal Row As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Column As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishColumn As Long = G_OMIT_CELL_INDEX) As WorksheetRangeBounds
    Call pCheckInit

    Dim s_row As Long, f_row As Long
    Dim s_col As Long, f_col As Long

    If Row = G_OMIT_CELL_INDEX Then
        s_row = pRow
    ElseIf Row < 1 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "開始行は 1 以上を指定してください。(" & Row & ")"
    ElseIf G_ROW_MAX < Row Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "開始行が " & G_ROW_MAX & " を超過しています。(" & Row & ")"
    Else
        s_row = Row
    End If

    If FinishRow = G_OMIT_CELL_INDEX Then
        f_row = pFinishRow
    ElseIf FinishRow < 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "終了行に負数は指定できません。(" & FinishRow & ")"
    Else
        f_row = FinishRow
    End If

    If Column = G_OMIT_CELL_INDEX Then
        s_col = pColumn
    ElseIf Column < 1 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "開始列は 1 以上を指定してください。(" & Column & ")"
    ElseIf G_COL_MAX < Column Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "開始列が " & G_COL_MAX & " を超過しています。(" & Column & ")"
    Else
        s_col = Column
    End If

    If FinishColumn = G_OMIT_CELL_INDEX Then
        f_col = pFinishColumn
    ElseIf FinishColumn < 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "終了列に負数は指定できません。(" & FinishColumn & ")"
    Else
        f_col = FinishColumn
    End If

    Dim result_value As WorksheetRangeBounds
    Set result_value = New_RangeBounds(Row:=s_row, Column:=s_col, FinishRow:=f_row, FinishColumn:=f_col, Sheet:=pWorksheetName, Book:=pWorkbookName)
    Set TransformAbsolute = result_value
End Function

'* この範囲を変形した新しい範囲を得ます。
'*
'* @param AddRow 追加する行数。
'* @param AddColumn 追加する列数。
'* @return 変形した新しいオブジェクト
'*
'* @details
'* この範囲の開始位置を維持したまま、行数と列数を増減した新しい範囲を得ます。
Public Function Transform( _
        Optional ByVal AddRow As Long = 0, _
        Optional ByVal AddColumn As Long = 0) As WorksheetRangeBounds
    Call pCheckInit

    Dim new_finish_row As Long
    Dim new_finish_col As Long
    If AddRow = 0 Then
        new_finish_row = pFinishRow
    Else
        Dim new_finish_row_value As Double
        If pFinishRow = 0 Then
            new_finish_row_value = CDbl(pRow) + CDbl(AddRow) - 1#
            If new_finish_row_value < 0# Then
                new_finish_row = 0
            ElseIf G_ROW_MAX < new_finish_row_value Then
                Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "変形後の行が、" & G_ROW_MAX & " を超過します。(row: " & pRow & ", add: " & AddRow & ")"
            Else
                new_finish_row = CLng(new_finish_row_value)
            End If
        Else
            new_finish_row_value = CDbl(pFinishRow) + CDbl(AddRow)
            If new_finish_row_value < 0# Then
                new_finish_row = 0
            ElseIf G_ROW_MAX < new_finish_row_value Then
                Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "変形後の行が、" & G_ROW_MAX & " を超過します。(row: " & pFinishRow & ", add: " & AddRow & ")"
            Else
                new_finish_row = CLng(new_finish_row_value)
            End If
        End If
    End If

    If AddColumn = 0 Then
        new_finish_col = pFinishColumn
    Else
        Dim new_finish_col_value As Double
        If pFinishColumn = 0 Then
            new_finish_col_value = CDbl(pColumn) + CDbl(AddColumn) - 1#
            If new_finish_col_value < 0# Then
                new_finish_col = 0
            ElseIf G_COL_MAX < new_finish_col_value Then
                Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "変形後の列が、" & G_COL_MAX & " を超過します。(col: " & pColumn & ", add: " & AddColumn & ")"
            Else
                new_finish_col = CLng(new_finish_col_value)
            End If
        Else
            new_finish_col_value = CDbl(pFinishColumn) + CDbl(AddColumn)
            If new_finish_col_value < 0# Then
                new_finish_col = 0
            ElseIf G_COL_MAX < new_finish_col_value Then
                Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "変形後の列が、" & G_COL_MAX & " を超過します。(col: " & pFinishColumn & ", add: " & AddColumn & ")"
            Else
                new_finish_col = CLng(new_finish_col_value)
            End If
        End If
    End If

    Dim result_value As WorksheetRangeBounds
    Set result_value = New_RangeBounds(Row:=pRow, Column:=pColumn, FinishRow:=new_finish_row, FinishColumn:=new_finish_col, Sheet:=pWorksheetName, Book:=pWorkbookName)
    Set Transform = result_value
End Function

'* この範囲をシフトした新しい範囲を得ます。
'*
'* @param ShiftRow 行シフト
'* @param ShiftColumn 列シフト
'* @return シフトした新しいオブジェクト。
'*
'* @details
'* このオブジェクトの複製を生成します。
'* この範囲をシフトした新しい範囲を得ます。
Public Function Shift(Optional ByVal Row As Long = 0, Optional ByVal Column As Long = 0) As WorksheetRangeBounds
    Call pCheckInit
    Dim new_s_row As Long, new_s_col As Long
    Dim new_f_row As Long, new_f_col As Long
    Dim new_s_row_value As Double, new_s_col_value As Double
    Dim new_f_row_value As Double, new_f_col_value As Double

    new_s_row_value = CDbl(pRow) + CDbl(Row)
    new_s_col_value = CDbl(pColumn) + CDbl(Column)
    If pFinishRow = 0 Then
        new_f_row_value = 0#
    Else
        new_f_row_value = CDbl(pFinishRow) + CDbl(Row)
    End If
    If pFinishColumn = 0 Then
        new_f_col_value = 0#
    Else
        new_f_col_value = CDbl(pFinishColumn) + CDbl(Column)
    End If

    If new_s_row_value < 1# Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "シフト後の行が、1 未満になります。(row: " & pRow & ", shift: " & Row & ")"
    If new_s_col_value < 1# Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "シフト後の列が、1 未満になります。(col: " & pColumn & ", shift: " & Column & ")"
    If new_f_row_value <> 0# And new_f_row_value < 1# Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "シフト後の行が、1 未満になります。(row: " & pFinishRow & ", shift: " & Row & ")"
    If new_f_col_value <> 0# And new_f_col_value < 1# Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "シフト後の列が、1 未満になります。(col: " & pFinishColumn & ", shift: " & Column & ")"
    If G_ROW_MAX < new_s_row_value Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "シフト後の行が、" & G_ROW_MAX & " を超過します。(row: " & pRow & ", shift: " & Row & ")"
    If G_ROW_MAX < new_f_row_value Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "シフト後の行が、" & G_ROW_MAX & " を超過します。(row: " & pFinishRow & ", shift: " & Row & ")"
    If G_COL_MAX < new_s_col_value Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "シフト後の列が、" & G_COL_MAX & " を超過します。(col: " & pColumn & ", shift: " & Column & ")"
    If G_COL_MAX < new_f_col_value Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "シフト後の列が、" & G_COL_MAX & " を超過します。(col: " & pFinishColumn & ", shift: " & Column & ")"

    new_s_row = CLng(new_s_row_value)
    new_s_col = CLng(new_s_col_value)
    new_f_row = CLng(new_f_row_value)
    new_f_col = CLng(new_f_col_value)

    Dim result_value As WorksheetRangeBounds
    Set result_value = New WorksheetRangeBounds

    Call result_value.Initialize( _
            Row:=new_s_row, _
            Column:=new_s_col, _
            FinishRow:=new_f_row, _
            FinishColumn:=new_f_col, _
            Sheet:=pWorksheetName, Book:=pWorkbookName)

    Set Shift = result_value
End Function

'* 2 つの範囲が重なる範囲を返します。
'*
'* @param 他の範囲。
'* @return 重なる範囲。
'*
'* @details
'* この範囲と他の範囲が重なる範囲を示す、新たなオブジェクトを返します。
'* 同じワークシート内の範囲でない場合は、エラーになります。
Public Function Intersect(ByVal OtherRangeBounds As WorksheetRangeBounds) As WorksheetRangeBounds
    Call pCheckInit

    If StrComp(pWorksheetName, OtherRangeBounds.WorksheetName, vbTextCompare) <> 0 Or StrComp(pWorkbookName, OtherRangeBounds.WorkbookName, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "ワークシートが異なります。(this: " & pWorksheetName & "@" & pWorkbookName & " / other: " & OtherRangeBounds.WorksheetName & "@" & OtherRangeBounds.WorkbookName & ")"
    End If

    Dim s_row As Long, f_row As Long
    Dim s_col As Long, f_col As Long
    s_row = MaxLng(pRow, OtherRangeBounds.Row)
    f_row = MinLng(pFinishRow, OtherRangeBounds.FinishRow)
    s_col = MaxLng(pColumn, OtherRangeBounds.Column)
    f_col = MinLng(pFinishColumn, OtherRangeBounds.FinishColumn)

'    If s_row <= f_row And s_col <= f_col Then
        Dim result_value As WorksheetRangeBounds
        Set result_value = New_RangeBounds(Row:=s_row, Column:=s_col, FinishRow:=f_row, FinishColumn:=f_col, Sheet:=pWorksheetName, Book:=pWorkbookName)
        Set Intersect = result_value
'    Else
'        Set Intersect = New_RangeBounds( _
'                Row:=MaxLng(MinLng(s_row, f_row, pFinishRow), pRow), _
'                Column:=MaxLng(MinLng(s_col, f_col, pFinishColumn), pColumn), _
'                Sheet:=pWorksheetName, Book:=pWorkbookName)
'    End If
End Function

'* このオブジェクトの複製を生成します。
'*
'* @return 複製された新しいオブジェクト。
'*
'* @details
'* このオブジェクトの複製を生成します。
Public Function CopyObject() As WorksheetRangeBounds
    Call pCheckInit
    Dim result_value As WorksheetRangeBounds
    Set result_value = New WorksheetRangeBounds
    Call result_value.Initialize(pRow, pColumn, pFinishRow, pFinishColumn, pWorksheetName, pWorkbookName)
    Set CopyObject = result_value
End Function

'* アドレス表記文字列を得ます。
'*
'* @param IsAbsoluteStartRow 開始行を絶対参照とする場合は True (既定値は False)
'* @param IsAbsoluteStartColumn 開始列を絶対参照とする場合は True (既定値は False)
'* @param IsAbsoluteFinishRow 終了行を絶対参照とする場合は True (既定値は False)
'* @param IsAbsoluteFinishColumn 終了列を絶対参照とする場合は True (既定値は False)
'* @param AddressType アドレス形式 ("A1" または "R1C1") (既定値は "A1")
'* @return アドレス文字列
'*
'* @details
'* この範囲のアドレス表記文字列で生成します。
Public Function ToString( _
        Optional ByVal IsAbsoluteStartRow As Boolean = False, _
        Optional ByVal IsAbsoluteStartColumn As Boolean = False, _
        Optional ByVal IsAbsoluteFinishRow As Boolean = False, _
        Optional ByVal IsAbsoluteFinishColumn As Boolean = False, _
        Optional ByVal AddressType As String = "A1", _
        Optional ByVal CellOnly As Boolean = False) As String
    If Not pIsInitialized Then
        ToString = "UNINITIALIZED(" & Hex(ObjPtr(Me)) & ")"
        Exit Function
    End If
    Dim book_name As String
    Dim sheet_name As String
    If CellOnly Then
        book_name = ""
        sheet_name = ""
    Else
        book_name = WorkbookName
        sheet_name = WorksheetName
    End If

    If IsEmpty Then
        If IsEntireRow Then
            ToString = RangeAddress( _
                    StartRow:=Row, _
                    FinishRow:=Row, _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name) & ":EMPTY"
        ElseIf IsEntireColumn Then
            ToString = RangeAddress( _
                    StartColumn:=Column, _
                    FinishColumn:=Column, _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name) & ":EMPTY"
        Else
            ToString = RangeAddress( _
                    StartRow:=Row, _
                    StartColumn:=Column, _
                    FinishRow:=MaxLng(FinishRow, Row), _
                    FinishColumn:=MaxLng(FinishColumn, Column), _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name) & ":EMPTY(R" & RowCount & ",C" & ColumnCount & ")"
        End If
    Else
        If IsEntireRow Then
            ToString = RangeAddress( _
                    StartRow:=Row, _
                    FinishRow:=FinishRow, _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name)
        ElseIf IsEntireColumn Then
            ToString = RangeAddress( _
                    StartColumn:=Column, _
                    FinishColumn:=FinishColumn, _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name)
        Else
            ToString = RangeAddress( _
                    StartRow:=Row, _
                    StartColumn:=Column, _
                    FinishRow:=FinishRow, _
                    FinishColumn:=FinishColumn, _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name)
        End If
    End If
End Function
Private Function IStringable_ToString() As String
    IStringable_ToString = ToString()
End Function

'
' このオブジェクトと他のオブジェクトが等しいかを返します。
'
Public Function Equals(ByVal OtherObject As Object) As Boolean
    Call pCheckInit

    If Not TypeOf OtherObject Is WorksheetRangeBounds Then
        Equals = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = OtherObject
    If Me.GetIdentityString() = range_bounds.GetIdentityString() Then
        Equals = True
    Else
        Equals = False
    End If
End Function
Private Function IEquatable_Equals(ByVal OtherObject As Object) As Boolean
    IEquatable_Equals = Equals(OtherObject)
End Function

'
' このオブジェクトの同一性を判断するための文字列表現を返します。
'
Public Function GetIdentityString() As String
    Call pCheckInit
    GetIdentityString = LCase$(ToString(CellOnly:=False))
End Function
Private Function IEquatable_GetIdentityString() As String
    IEquatable_GetIdentityString = GetIdentityString()
End Function

'* この範囲の列挙子を得ます。
'*
'* @param EnumerationMode [省略可] 列挙方式。既定値は行単位です。
'* @param Descending [省略可] 降順で列挙するか否か。
'* @return 列挙子オブジェクト。
'*
'* @details
'* この範囲の列挙子を得ます。列挙方式はこの呼び出しに指定した引数だけで決まります。
Public Function GetEnumerator( _
        Optional ByVal EnumerationMode As Long = G_RANGE_ENUM_MODE_ROWS, _
        Optional ByVal Descending As Boolean = False) As IEnumerator
    Call pCheckInit
    Dim result_value As WorksheetRangeBoundsEnumerator
    Set result_value = New WorksheetRangeBoundsEnumerator
    Call result_value.Initialize(Me, EnumerationMode:=EnumerationMode, Descending:=Descending)
    Set GetEnumerator = result_value
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "未初期化状態です。"
End Sub

Private Function pGetTargetRow(ByVal RowIndex As Long) As Long
    Dim row_count_value As Long
    row_count_value = RowCount

    If RowIndex = 0 Or row_count_value < RowIndex Or RowIndex < -row_count_value Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "行範囲を超えています。(index:" & RowIndex & ", start: " & pRow & ", finish: " & pFinishRow & ")"
    End If

    Dim target_row_value As Double
    If 0 < RowIndex Then
        target_row_value = CDbl(pRow) + CDbl(RowIndex) - 1#
    Else
        target_row_value = CDbl(pFinishRow) + CDbl(RowIndex) + 1#
    End If

    Dim result_value As Long
    result_value = CLng(target_row_value)

    pGetTargetRow = result_value
End Function

Private Function pGetTargetColumn(ByVal ColumnIndex As Long) As Long
    Dim column_count_value As Long
    column_count_value = ColumnCount

    If ColumnIndex = 0 Or column_count_value < ColumnIndex Or ColumnIndex < -column_count_value Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "列範囲を超えています。(index:" & ColumnIndex & ", start: " & pColumn & ", finish: " & pFinishColumn & ")"
    End If

    Dim target_column_value As Double
    If 0 < ColumnIndex Then
        target_column_value = CDbl(pColumn) + CDbl(ColumnIndex) - 1#
    Else
        target_column_value = CDbl(pFinishColumn) + CDbl(ColumnIndex) + 1#
    End If

    Dim result_value As Long
    result_value = CLng(target_column_value)

    pGetTargetColumn = result_value
End Function

Private Sub pRequireWorkbookService(ByVal MethodName As String)
    If WbSrv Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WorksheetRangeBounds" & "." & MethodName, Description:="InitializeCommonService が呼ばれていません。"
    End If
End Sub

Private Function pCreateNewEnumSnapshot() As Collection
    Dim result_value As Collection
    Set result_value = New Collection

    Dim enum_obj As IEnumerator
    Set enum_obj = Me.GetEnumerator(EnumerationMode:=G_RANGE_ENUM_MODE_CELLS_HORIZONTAL, Descending:=False)
    Do While enum_obj.MoveNext()
        Call result_value.Add(enum_obj.Current)
    Loop

    Set pCreateNewEnumSnapshot = result_value
End Function
