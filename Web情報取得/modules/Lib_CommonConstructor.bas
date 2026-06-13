Attribute VB_Name = "Lib_CommonConstructor"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! 共通クラス モジュールのコンストラクター相当関数をまとめた標準モジュールです。
'!
' #############################################################################

'* WorksheetRangeBounds インスタンスを新規作成します。
'*
'* @param RowIndex [省略可] 行番号。
'* @param ColumnIndex [省略可] 列番号。
'* @param FinishRowIndex [省略可] 最後の行番号。
'* @param FinishColumnIndex [省略可] 最後の列番号。
'* @param SheetName [省略可] ワークシート名。
'* @param BookName [省略可] ワークブック名。
'* @return 新しい WorksheetRangeBounds インスタンス
'*
'* @details
'* WorksheetRangeBounds インスタンスを新規作成します。
'*
'* 使用例:
'* @code
'* ' 1 セル (B5)
'* Dim cell_bounds As WorksheetRangeBounds
'* Set cell_bounds = New_RangeBounds(Row:=5, Column:=2)
'*
'* ' 1 行全体 (5:5)
'* Dim row_bounds As WorksheetRangeBounds
'* Set row_bounds = New_RangeBounds(Row:=5)
'*
'* ' 2 列全体 (B:C)
'* Dim two_columns_bounds As WorksheetRangeBounds
'* Set two_columns_bounds = New_RangeBounds(Column:=2, FinishColumn:=3)
'*
'* ' 表全体 (B2:E20)
'* Dim table_bounds As WorksheetRangeBounds
'* Set table_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=20, FinishColumn:=5)
'* @endcode
Public Function New_RangeBounds( _
        Optional ByVal Row As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Column As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishColumn As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Sheet As String = "Sheet1", _
        Optional ByVal Book As String = "") As WorksheetRangeBounds

    Dim result_value As WorksheetRangeBounds
    Set result_value = New WorksheetRangeBounds

    Call result_value.Initialize( _
            Row:=Row, _
            Column:=Column, _
            FinishRow:=FinishRow, _
            FinishColumn:=FinishColumn, _
            Sheet:=Sheet, _
            Book:=Book)

    Set New_RangeBounds = result_value
End Function

'* ObjectList インスタンスを新規作成し、要素型契約を設定します。
'*
'* @param ElementTypeName 要素型契約名。
'* @param RequireComparable 要素に IComparable 実装を要求するか否か。
'* @param ObjectKeyMode オブジェクトの同一性/重複判定モード。
'* @return 初期化済みの ObjectList インスタンス。
'*
'* @details
'* New_ 系の処理は薄いファクトリに留め、実処理は ObjectList.Initialize に委譲します。
Public Function New_ObjectList( _
        ByVal ElementTypeName As String, _
        Optional ByVal RequireComparable As Boolean = False, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE) As ObjectList

    Dim result_value As ObjectList
    Set result_value = New ObjectList
    Call result_value.Initialize( _
            ElementTypeName:=ElementTypeName, _
            RequireComparable:=RequireComparable, _
            ObjectKeyMode:=ObjectKeyMode)

    Set New_ObjectList = result_value
End Function

'* ObjectSet インスタンスを新規作成し、要素型契約を設定します。
'*
'* @param ElementTypeName 要素型契約名。
'* @param RequireComparable 要素に IComparable 実装を要求するか否か。
'* @param ObjectKeyMode オブジェクトの同一性/重複判定モード。
'* @return 初期化済みの ObjectSet インスタンス。
'*
'* @details
'* New_ 系の処理は薄いファクトリに留め、実処理は ObjectSet.Initialize に委譲します。
Public Function New_ObjectSet( _
        ByVal ElementTypeName As String, _
        Optional ByVal RequireComparable As Boolean = False, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE) As ObjectSet

    Dim result_value As ObjectSet
    Set result_value = New ObjectSet
    Call result_value.Initialize( _
            ElementTypeName:=ElementTypeName, _
            RequireComparable:=RequireComparable, _
            ObjectKeyMode:=ObjectKeyMode)

    Set New_ObjectSet = result_value
End Function

'* ObjectDictionary インスタンスを新規作成し、必要に応じて要素型契約を設定します。
'*
'* @param ElementTypeName [省略可] 要素型契約名。
'* @param RequireComparable 要素に IComparable 実装を要求するか否か。
'* @param ObjectKeyMode オブジェクトの同一性/重複判定モード。
'* @return 新しい ObjectDictionary インスタンス。
'*
'* @details
'* ElementTypeName を省略した場合、最初に追加された要素から ObjectList と同じ型推論を行います。
Public Function New_ObjectDictionary( _
        Optional ByVal ElementTypeName As String = "", _
        Optional ByVal RequireComparable As Boolean = False, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE) As ObjectDictionary

    Dim result_value As ObjectDictionary
    Set result_value = New ObjectDictionary
    If ElementTypeName <> "" Then
        Call result_value.Initialize( _
                ElementTypeName:=ElementTypeName, _
                RequireComparable:=RequireComparable, _
                ObjectKeyMode:=ObjectKeyMode)
    End If

    Set New_ObjectDictionary = result_value
End Function

'* Excel アドレス文字列から WorksheetRangeBounds インスタンスを新規作成します。
'*
'* @param AddressString 初期化に使用する Excel アドレス文字列。
'* @return 新しい WorksheetRangeBounds インスタンス
'*
'* @details
'* New_ 系の処理は薄いファクトリに留め、実処理は WorksheetRangeBounds.InitializeFromAddress に委譲します。
Public Function New_RangeBoundsFromAddress(ByVal AddressString As String) As WorksheetRangeBounds
    Dim result_value As WorksheetRangeBounds
    Set result_value = New WorksheetRangeBounds
    Call result_value.InitializeFromAddress(AddressString)

    Set New_RangeBoundsFromAddress = result_value
End Function

'* WorksheetVirtualTable インスタンスを新規作成します。
'*
'* @param ColumnRangeList 仮想表の列として扱う WorksheetRangeBounds 一覧。
'* @param HeaderNames [省略可] header 名の文字列配列。
'* @param TreatFirstRowAsHeader True の場合は各範囲の 1 行目を header として読み取ります。
'* @param HeaderCompareMode header の重複判定と行 Dictionary の比較モード。
'* @param HeaderJoinDelimiter 複数列範囲の header セルを連結するときの区切り文字。
'* @return 初期化済みの WorksheetVirtualTable インスタンス。
'*
'* @details
'* New_ 系の処理は薄いファクトリに留め、実処理は WorksheetVirtualTable.Initialize に委譲します。
Public Function New_WorksheetVirtualTable( _
        ByVal ColumnRangeList As ObjectList, _
        Optional ByVal HeaderNames As Variant, _
        Optional ByVal TreatFirstRowAsHeader As Boolean = False, _
        Optional ByVal HeaderCompareMode As VbCompareMethod = vbBinaryCompare, _
        Optional ByVal HeaderJoinDelimiter As String = "") As WorksheetVirtualTable

    Dim result_value As WorksheetVirtualTable
    Set result_value = New WorksheetVirtualTable
    Call result_value.Initialize( _
            ColumnRangeList:=ColumnRangeList, _
            HeaderNames:=HeaderNames, _
            TreatFirstRowAsHeader:=TreatFirstRowAsHeader, _
            HeaderCompareMode:=HeaderCompareMode, _
            HeaderJoinDelimiter:=HeaderJoinDelimiter)

    Set New_WorksheetVirtualTable = result_value
End Function

'* WorksheetRangeBounds から WorksheetVirtualTable インスタンスを新規作成します。
'*
'* @param TableRange 仮想表の列に分割する WorksheetRangeBounds。
'* @param HeaderNames [省略可] header 名の文字列配列。
'* @param TreatFirstRowAsHeader True の場合は各範囲の 1 行目を header として読み取ります。
'* @param HeaderCompareMode header の重複判定と行 Dictionary の比較モード。
'* @param HeaderJoinDelimiter 複数列範囲の header セルを連結するときの区切り文字。
'* @return 初期化済みの WorksheetVirtualTable インスタンス。
'*
'* @details
'* New_ 系の処理は薄いファクトリに留め、実処理は WorksheetVirtualTable.InitializeFromRangeBounds に委譲します。
Public Function New_WorksheetVirtualTableFromRangeBounds( _
        ByVal TableRange As WorksheetRangeBounds, _
        Optional ByVal HeaderNames As Variant, _
        Optional ByVal TreatFirstRowAsHeader As Boolean = False, _
        Optional ByVal HeaderCompareMode As VbCompareMethod = vbBinaryCompare, _
        Optional ByVal HeaderJoinDelimiter As String = "") As WorksheetVirtualTable

    Dim result_value As WorksheetVirtualTable
    Set result_value = New WorksheetVirtualTable
    Call result_value.InitializeFromRangeBounds( _
            TableRange:=TableRange, _
            HeaderNames:=HeaderNames, _
            TreatFirstRowAsHeader:=TreatFirstRowAsHeader, _
            HeaderCompareMode:=HeaderCompareMode, _
            HeaderJoinDelimiter:=HeaderJoinDelimiter)

    Set New_WorksheetVirtualTableFromRangeBounds = result_value
End Function
