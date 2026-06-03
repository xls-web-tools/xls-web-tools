Attribute VB_Name = "Lib_Common"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! 基本的な関数などをまとめた標準モジュールです。
'! 他のツールとも共用されるため、このツールで使用しないものも含まれます。
'!
' #############################################################################

'* WorkbookService。ユニットテキスト時にはテスト ダブルに置き換えてください。
Public WbSrv As IWorkbookService

'* WorksheetService。ユニットテキスト時にはテスト ダブルに置き換えてください。
Public WsSrv As IWorksheetService

'* DebugInformation。GUI 実行中のデバッグ情報を保持する共通実行状態です。
Public DbgInfo As DebugInformation

'* ProgressStatus。GUI 実行中の進捗状態を保持する共通実行状態です。
Public ProgStat As ProgressStatus


'* 未設定の日付を表す
Public Const G_DATE_NULL As Date = #12/31/1899#

'* 有効な日付の最小値
Public Const G_DATE_MINIMUM As Date = #3/1/1900#

'* 行番号の最大値
Public Const G_ROW_MAX As Long = 1048576

'* 列番号の最大値
Public Const G_COL_MAX As Long = 16384

'* 行番号および列番号の省略値
Public Const G_OMIT_CELL_INDEX As Long = -2147483648#

'* WorksheetService.SetAlignment で設定を変更しないことを表す値
Public Const G_ALIGNMENT_NO_CHANGE As Long = -2147483648#

'* ファイルシステムのパス区切り文字
Public Const G_FS_PATH_SEP As String = "\"

'* GetTypedValueKey / GetValueKey でオブジェクト参照をキー化することを表す値
Public Const G_OBJECT_KEY_MODE_REFERENCE As Long = 0

'* GetTypedValueKey / GetValueKey で IEquatable.GetIdentityString をキー化することを表す値
Public Const G_OBJECT_KEY_MODE_I_EQUATABLE As Long = 1

'* GetTypedValueKey / GetValueKey で IDuplicateCheckable.GetKey をキー化することを表す値
Public Const G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE As Long = 2

'* WorksheetRangeBounds を行単位で列挙することを表す値
Public Const G_RANGE_ENUM_MODE_ROWS As Long = 0

'* WorksheetRangeBounds を列単位で列挙することを表す値
Public Const G_RANGE_ENUM_MODE_COLUMNS As Long = 1

'* WorksheetRangeBounds を行方向のセル順で列挙することを表す値
Public Const G_RANGE_ENUM_MODE_CELLS_HORIZONTAL As Long = 2

'* WorksheetRangeBounds を列方向のセル順で列挙することを表す値
Public Const G_RANGE_ENUM_MODE_CELLS_VERTICAL As Long = 3

Private Const C_LONG_MAX As Long = 2147483647#
Private Const C_LONG_MIN As Long = -2147483648#
Private Const C_INTEGER_MAX As Long = 32767
Private Const C_INTEGER_MIN As Long = -32768
Private Const C_A1_TOKEN_CELL As Long = 1
Private Const C_A1_TOKEN_ROW As Long = 2
Private Const C_A1_TOKEN_COLUMN As Long = 3

' #############################################################################
'
' 独自ワークシート関数
'
' #############################################################################

''* 新しい Excel で存在するワークシート関数 TEXTSPLIT の独自実装。
''*
''* @param Expression 分割するテキスト。
''* @param ColumnDelimiter 列の区切り文字。(一般的な CSV では 「,」)
''* @param RowDelimiter 行の区切り文字。(一般的な CSV では 改行記号) (省略可能)
''* @param IgnoreEmpty 連続する区切り記号を無視かどうか。無視するには True を指定します。 既定値は False で、PadWith で埋められます。
''* @param MatchMode (未実装の機能)
''* @param PadWith 連続する区切り記号の場合の値。既定値は #N/A。
''* @return 分割した結果の配列。
''*
''* @details
''* TEXTSPLIT がないバージョンの Excel のための実装です。詳細については Microsoft の TEXTSPLIT の説明を参照してください。
''*
''* @see https://support.microsoft.com/ja-jp/office/textsplit-%E9%96%A2%E6%95%B0-b1ca414e-4c21-4ca0-b1b7-bdecace8a6e7
'Function TEXTSPLIT( _
'        ByVal Expression As String, ByVal ColumnDelimiter As String, Optional ByVal RowDelimiter As String, _
'        Optional ByVal IgnoreEmpty As Boolean = False, Optional ByVal MatchMode As Integer = 0, Optional ByVal PadWith As Variant = Nothing) As Variant()
'
'    If MatchMode <> 0 Then
'        ' MatchMode 1 (IgnoreCase) は未実装
'        TEXTSPLIT = CVErr(xlErrValue)
'        Exit Function
'    End If
'
'    If IsObject(PadWith) Then
'        If PadWith Is Nothing Then
'            PadWith = CVErr(xlErrNA)
'        End If
'    End If
'
'    Dim rows_arr() As String
'    rows_arr = Split(Expression, RowDelimiter)
'
'    Dim max_col As Long
'    Dim rows_list As ObjectList
'    Set rows_list = New ObjectList
'    Dim row_item As Variant 'String
'    For Each row_item In rows_arr
'        If row_item <> "" Or Not IgnoreEmpty Then
'            Dim cols_arr() As String
'            cols_arr = Split(row_item, ColumnDelimiter)
'
'            Dim cols_list As ObjectList
'            Set cols_list = New ObjectList
'
'            Dim col_item As Variant 'String
'            For Each col_item In cols_arr
'                If col_item <> "" Or Not IgnoreEmpty Then
'                    Call cols_list.Add(col_item)
'                End If
'            Next col_item
'
'            If 0 < cols_list.Count Or Not IgnoreEmpty Then
'                Call rows_list.Add(cols_list)
'                max_col = MaxLng(max_col, cols_list.Count)
'            End If
'        End If
'    Next row_item
'
'    Dim result_value() As Variant
'    ReDim result_value(0 To rows_list.Count - 1, 0 To max_col - 1)
'
'    Dim row_idx As Long
'    Dim col_idx As Long
'    For row_idx = 0 To UBound(result_value, 1) Step 1
'        For col_idx = 0 To UBound(result_value, 2) Step 1
'            Set cols_list = rows_list.Item(row_idx)
'            If col_idx < cols_list.Count Then
'                result_value(row_idx, col_idx) = cols_list.Item(col_idx)
'            Else
'                result_value(row_idx, col_idx) = PadWith
'            End If
'        Next col_idx
'    Next row_idx
'
'    TEXTSPLIT = result_value
'End Function

''* 新しい Excel で存在するワークシート関数 TEXTJOIN の独自実装。
''*
''* @param Delimiter 区切り文字。
''* @param IgnoreEmpty 空文字列を無視かどうか。True の場合、空文字列は無視されます。False の場合、無視されず区切り文字が連続します。
''* @param Expression1 1 番目の文字列
''* @param Expressions 2 番目以降の文字列。(可変長引数)
''*
''* @return 結合した文字列。
''*
''* @details
''* TEXTJOIN がないバージョンの Excel のための実装です。詳細については Microsoft の TEXTJOIN の説明を参照してください。
''*
''* @see https://support.microsoft.com/ja-jp/office/textjoin-%E9%96%A2%E6%95%B0-357b449a-ec91-49d0-80c3-0e8fc845691c
'Function TEXTJOIN(ByVal Delimiter As String, ByVal IgnoreEmpty As Boolean, ByVal Expression1 As Variant, ParamArray Expressions() As Variant) As Variant
'    Dim result_value As String
'    Dim is_first As Boolean
'
'    is_first = True
'
'    On Error Resume Next
'    Call pTextJoinCore(result_value, is_first, Delimiter, Expression1, IgnoreEmpty)
'    If Err.Number <> 0 Then
'        TEXTJOIN = CVErr(xlErrNA)
'        Exit Function
'    End If
'
'    Dim param_item As Variant
'    For Each param_item In Expressions
'        Call pTextJoinCore(result_value, is_first, Delimiter, param_item, IgnoreEmpty)
'        If Err.Number <> 0 Then
'            TEXTJOIN = CVErr(xlErrNA)
'            Exit Function
'        End If
'    Next param_item
'
'    TEXTJOIN = result_value
'End Function

'Sub pTextJoinCore(ByRef ResultString As String, ByRef IsFirst As Boolean, Delimiter As String, Expression As Variant, IgnoreEmpty As Boolean)
'    If TypeOf Expression Is Range Or IsArray(Expression) Then
'        ' Expression が Range か配列の場合
'        Dim cell_item As Range
'        For Each cell_item In Expression
'            Call pTextJoinItemCore(ResultString, IsFirst, Delimiter, cell_item.Value, IgnoreEmpty)
'        Next cell_item
''    ElseIf IsArray(Expression) Then
''        ' Expression が配列の場合
''        If LBound(Expression) <= UBound(Expression) Then
''            Dim item_idx As Long
''            For item_idx = LBound(Expression) To UBound(Expression)
''                Call pTextJoinItemCore(ResultString, IsFirst, Delimiter, Expression(item_idx), IgnoreEmpty)
''            Next item_idx
''        ElseIf Not IgnoreEmpty Then
''            ResultString = ResultString & Delimiter
''        End If
'    Else
'        ' Range や配列でない場合
'        Call pTextJoinItemCore(ResultString, IsFirst, Delimiter, Expression, IgnoreEmpty)
'    End If
'End Sub

'Private Sub pTextJoinItemCore(ByRef ResultString As String, ByRef IsFirst As Boolean, Delimiter As String, Expression As Variant, IgnoreEmpty As Boolean)
'    If IsError(Expression) Then Err.Raise xlErrNA, "TEXTJOIN", "結合対象にエラー値が含まれます"
'    If Expression <> "" Or Not IgnoreEmpty Then
'        If IsFirst Then
'            ResultString = Expression
'            IsFirst = False
'        Else
'            ResultString = ResultString & Delimiter & Expression
'        End If
'    End If
'End Sub

' #############################################################################
'
' 共通クラス モジュールのコンストラクターなど
'
' #############################################################################

'* WorkbookService を初期化します。
'*
'* @param Force [省略可] True の場合は差し替え済みサービスを本番サービスへ再生成します。
'*
'* @details
'* WbSrv が未設定の場合、または Force が True の場合に WorkbookService を生成します。テストで差し替え済みの場合は Force が False のときだけ維持します。
Public Sub InitializeWorkbookService(Optional ByVal Force As Boolean = False)
    If Force Or (WbSrv Is Nothing) Then Set WbSrv = New WorkbookService
End Sub

'* WorksheetService を初期化します。
'*
'* @param Force [省略可] True の場合は差し替え済みサービスを本番サービスへ再生成します。
'*
'* @details
'* WsSrv が未設定の場合、または Force が True の場合に WorksheetService を生成します。テストで差し替え済みの場合は Force が False のときだけ維持します。
Public Sub InitializeWorksheetService(Optional ByVal Force As Boolean = False)
    If Force Or (WsSrv Is Nothing) Then Set WsSrv = New WorksheetService
End Sub

'* UDF から安全に利用できる共通サービスを初期化します。
'*
'* @param Force [省略可] True の場合は差し替え済みサービスを本番サービスへ再生成します。
'*
'* @details
'* Excel ワークシート関数の再計算中に Application.Run へ到達しないよう、WorkbookService と WorksheetService だけを初期化します。FileSystemService と TextFileService は初期化しません。
Public Sub InitializeUdfCommonService(Optional ByVal Force As Boolean = False)
    Call InitializeWorkbookService(Force)
    Call InitializeWorksheetService(Force)
End Sub

'* 共通的なサービスを初期化します。
'*
'* @param Force [省略可] True の場合は差し替え済みサービスを本番サービスへ再生成します。
'*
'* @details
'* 共通的なサービスを初期化します。初期化対象は以下の通りです。
'*
'* * WorkbookService
'* * WorksheetService
'* * FileSystemService (インポート済みの場合)
'* * TextFileService (インポート済みの場合)
Public Sub InitializeCommonService(Optional ByVal Force As Boolean = False)
    Call InitializeWorkbookService(Force)
    Call InitializeWorksheetService(Force)
    Call pTryInitializeOptionalCommonService("InitializeFileSystemService", Force)
    Call pTryInitializeOptionalCommonService("InitializeTextFileService", Force)
End Sub

Private Sub pTryInitializeOptionalCommonService(ByVal InitializerName As String, ByVal Force As Boolean)
    On Error Resume Next
    Err.Clear
    Call Application.Run("'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & InitializerName, Force)

    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext
    Err.Clear
    On Error GoTo 0

    If err_num = 0 Then Exit Sub
    If err_num = 1004 And 0 < InStr(1, err_desc, InitializerName, vbTextCompare) Then Exit Sub

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Sub

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


' #############################################################################
'
' エラー ハンドリング
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' エラー情報関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* エラー情報を処理し、必要に応じて情報を補完または蓄積するためのサブルーチン。
'*
'* @param ErrNumber [出力] エラー番号を格納する変数。既定値は 0。Append を True にしても、ErrNumber は上書きされます。(省略可能)
'* @param ErrSource [出力] エラーの発生源を格納する文字列変数。(省略可能)
'* @param ErrDescription [出力] エラーの説明を格納する文字列変数。(省略可能)
'* @param Supplementation エラーの説明に付加する補足情報。既定値は空文字列。
'* @param ErrClear エラーをクリアするかどうか。True の場合、エラー情報はクリアされます。既定値は True。
'* @param Append 複数のエラー情報を結合するかどうか。True の場合、既存の情報に追加されます。既定値は False。
'*
'* @details
'* このサブルーチンは、VBA の標準的なエラー情報をカスタマイズした形で処理します。
'* 特に、エラー番号、発生源、説明を取得し、補足情報を付加する機能があります。
'* 複数のエラーが発生した場合、Append パラメータを使用して情報を連結できます。
'*
'* 使用例:
'* @code
'* Dim ErrNum As Long
'* Dim ErrSrc As String
'* Dim ErrDesc As String
'*
'* On Error Resume Next
'*
'* ''//故意にエラーを発生させる例 (ゼロ除算)
'* Debug.Print 1 / 0
'*
'* Call HandleError(ErrNum, ErrSrc, ErrDesc, "補足情報 1", Append:=True)
'*
'* ''//故意にエラーを発生させる例 (未定義サブ プロシージャ呼び出し)
'* Call NonExistentProcedure
'*
'* Call HandleError(ErrNum, ErrSrc, ErrDesc, "補足情報 2", Append:=True)
'*
'* On Error GoTo 0
'*
'* Err.Raise ErrNum, ErrSrc, ErrDesc
'* @endcode
Public Sub HandleError( _
        Optional ByRef ErrNumber As Long = 0, _
        Optional ByRef ErrSource As String = "", _
        Optional ByRef ErrDescription As String = "", _
        Optional ByVal Supplementation As String = "", _
        Optional ByVal ErrClear As Boolean = True, _
        Optional ByVal Append As Boolean = False)

    If Err.Number <> 0 Then
        ErrNumber = Err.Number
        If Append Then
            ErrSource = ErrSource & Err.Source & vbCrLf
            If Supplementation <> "" Then
                ErrDescription = ErrDescription & Err.Description & " [ " & Supplementation & " ]" & vbCrLf
            Else
                ErrDescription = ErrDescription & Err.Description & vbCrLf
            End If
        Else
            ErrSource = Err.Source
            If Supplementation <> "" Then
                ErrDescription = Err.Description & " [ " & Supplementation & " ]"
            Else
                ErrDescription = Err.Description
            End If
        End If

        If ErrClear Then
            Err.Clear
        End If
    End If
End Sub


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel エラー値関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Excel エラー値を Excel 表示文字列へ変換します。
'*
'* @param ErrorValue Excel エラー値。
'* @return Excel エラー表示文字列。
'*
'* @details
'* `CVErr(...)` で表される Excel エラー値を、セル表示や数式リテラルで使える文字列へ変換します。
'* 未知のエラー値はエラー番号を含む `#ERRNO_xxx!` として返します。
Public Function ExcelErrorToString(ByVal ErrorValue As Variant) As String
    If Not IsError(ErrorValue) Then
        Err.Raise vbObjectError + 1, "Function ExcelErrorToString", "Excel エラー値ではありません。(" & TypeName(ErrorValue) & ")"
    End If

    Dim result_value As String
    Select Case ErrorValue
     Case CVErr(xlErrDiv0)
        result_value = "#DIV/0!"
     Case CVErr(xlErrNA)
        result_value = "#N/A"
     Case CVErr(xlErrName)
        result_value = "#NAME?"
     Case CVErr(xlErrNull)
        result_value = "#NULL!"
     Case CVErr(xlErrNum)
        result_value = "#NUM!"
     Case CVErr(xlErrRef)
        result_value = "#REF!"
     Case CVErr(xlErrValue)
        result_value = "#VALUE!"
     Case CVErr(xlErrGettingData)
        result_value = "#GETTING_DATA"
     Case CVErr(xlErrSpill)
        result_value = "#SPILL!"
     Case CVErr(xlErrConnect)
        result_value = "#CONNECT!"
     Case CVErr(xlErrBlocked)
        result_value = "#BLOCKED!"
     Case CVErr(xlErrUnknown)
        result_value = "#UNKNOWN!"
     Case CVErr(xlErrField)
        result_value = "#FIELD!"
     Case CVErr(xlErrCalc)
        result_value = "#CALC!"
     Case Else
        result_value = "#ERRNO_" & CStr(CLng(ErrorValue)) & "!"
    End Select

    ExcelErrorToString = result_value
End Function

'* Excel エラー表示文字列を Excel エラー値へ変換します。
'*
'* @param Expression Excel エラー表示文字列。
'* @param ErrorValue [出力] 変換できた場合の Excel エラー値。
'* @return 変換できた場合は True、それ以外は False。
'*
'* @details
'* 既知の Excel エラー表示文字列だけを `CVErr(...)` へ変換します。大小文字は区別しません。
'* 未知の文字列では `ErrorValue` に `Empty` を明示代入して False を返します。
Public Function TryConvertExcelErrorStringToCVErr( _
        ByVal Expression As String, _
        ByRef ErrorValue As Variant) As Boolean

    Dim result_value As Boolean
    result_value = True

    Select Case UCase$(Expression)
     Case "#DIV/0!"
        ErrorValue = CVErr(xlErrDiv0)
     Case "#N/A"
        ErrorValue = CVErr(xlErrNA)
     Case "#NAME?"
        ErrorValue = CVErr(xlErrName)
     Case "#NULL!"
        ErrorValue = CVErr(xlErrNull)
     Case "#NUM!"
        ErrorValue = CVErr(xlErrNum)
     Case "#REF!"
        ErrorValue = CVErr(xlErrRef)
     Case "#VALUE!"
        ErrorValue = CVErr(xlErrValue)
     Case "#GETTING_DATA"
        ErrorValue = CVErr(xlErrGettingData)
     Case "#SPILL!"
        ErrorValue = CVErr(xlErrSpill)
     Case "#CONNECT!"
        ErrorValue = CVErr(xlErrConnect)
     Case "#BLOCKED!"
        ErrorValue = CVErr(xlErrBlocked)
     Case "#UNKNOWN!"
        ErrorValue = CVErr(xlErrUnknown)
     Case "#FIELD!"
        ErrorValue = CVErr(xlErrField)
     Case "#CALC!"
        ErrorValue = CVErr(xlErrCalc)
     Case Else
        ErrorValue = Empty
        result_value = False
    End Select

    TryConvertExcelErrorStringToCVErr = result_value
End Function


' #############################################################################
'
' GUI 関連
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' ボタン関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* ワークシートへボタンを作成します。
'*
'* @param TargetSheet 対象のワークシート。
'* @param RowIndex ボタンを追加する位置のセルの行番号。
'* @param ColumnIndex ボタンを追加する位置のセルの列番号。
'* @param Caption ボタンの表示文字列。
'* @param OnAction クリック時に実行するサブ プロシージャ。
'* @param OnAction [省略可]ボタンのコンポーネント名 (Name プロパティ)。
'* @return 追加したボタン オブジェクト
'*
'* @details
'* ワークシートへボタンを作成します。
Public Function AddButton( _
    ByVal TargetSheet As Worksheet, _
    ByVal RowIndex As Long, _
    ByVal ColumnIndex As Long, _
    ByVal Caption As String, _
    ByVal OnAction As String, _
    Optional ByVal Name As String = "") As Shape

    Dim tgt_rng As Range
    Set tgt_rng = TargetSheet.Cells(RowIndex, ColumnIndex)

    Dim shp_btn As Shape
    Set shp_btn = TargetSheet.Shapes.AddShape( _
                 Type:=msoShapeRectangle, _
                 Left:=tgt_rng.Left, _
                 Top:=tgt_rng.Top, _
                 Width:=tgt_rng.Width, _
                 Height:=tgt_rng.Height)

    If Name <> "" Then
        shp_btn.Name = Name
    End If

    shp_btn.OnAction = OnAction

    With shp_btn.TextFrame2
        .TextRange.Text = Caption

        With .TextRange.Font
            .NameComplexScript = "Yu Gothic UI Semibold"
            .NameFarEast = "Yu Gothic UI Semibold"
            .Name = "Yu Gothic UI Semibold"
            .Size = 9
            .Fill.ForeColor.RGB = RGB(64, 64, 64)
        End With

        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
    End With

    shp_btn.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp_btn.TextFrame.VerticalAlignment = xlVAlignCenter

    shp_btn.Fill.ForeColor.RGB = RGB(192, 192, 192)
    shp_btn.Fill.Solid
    shp_btn.Line.ForeColor.RGB = RGB(64, 64, 64)
    shp_btn.Line.Weight = 1

    shp_btn.Shadow.Visible = msoFalse
    shp_btn.ThreeD.Visible = False

    Set AddButton = shp_btn
End Function

'* ワークシートのボタンを削除します。
'*
'* @param TargetSheet 対象のワークシート。
'* @param Name ボタンのコンポーネント名 (Name プロパティ)。
'*
'* @details
'* ワークシート上のボタンを、コンポーネント名を指定して削除します。
Public Sub DeleteButton(ByVal TargetSheet As Worksheet, ByVal Name As String)

    Dim shape_obj As Shape
    Dim is_deleted As Boolean: is_deleted = False

    For Each shape_obj In TargetSheet.Shapes
        If shape_obj.Name = Name Then
            shape_obj.Delete
            is_deleted = True
            Exit For
        End If
    Next shape_obj

    If Not is_deleted Then
        Err.Raise vbObjectError + 1, "Sub DeleteButton", "ボタンが見つかりませんでした。(" & Name & ")"
    End If

End Sub

'* ワークシートのすべてのボタンを削除します。
'*
'* @param TargetSheet 対象のワークシート。
'*
'* @details
'* ワークシート上のボタンを、すべて削除します。
Public Sub ClearButton(ByVal TargetSheet As Worksheet)
    Dim item_idx As Long
    ' 後ろから削除しないとループ中に変更が起きた際に不具合が出ることがあるので
    ' For i = Shapes.Count To 1 Step -1 の形が安全です
    For item_idx = TargetSheet.Shapes.Count To 1 Step -1

        If TargetSheet.Shapes(item_idx).OnAction <> "" Then
            ' OnActionが設定されている → 「図形ボタン」と判断して削除
            TargetSheet.Shapes(item_idx).Delete
        End If

    Next item_idx

End Sub


' #############################################################################
'
' クリップボード関連
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' クリップボード操作
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* クリップボードへ文字列を送信します。
'*
'* @param SourceText クリップボードにコピーする文字列。
'*
'* @details
'* VBA の `Forms.TextBox.1` オブジェクトを使用して、指定された文字列をクリップボードにコピーします。
Public Sub SetClipboard(ByVal SourceText As String)
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = SourceText
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
End Sub

'* クリップボードから文字列を取得します。
'*
'* @return クリップボードに格納されている文字列。クリップボードに文字列が含まれていない場合、空文字列を返します。
'*
'* @details
'* VBA の `Forms.TextBox.1` オブジェクトを使用して、クリップボードの内容を取得します。
Public Function GetClipboard() As String
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        If .CanPaste = True Then .Paste
        GetClipboard = .Text
    End With
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 数式貼り付け
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 可能な限り書式を除いた貼り付けを実行します。
'*
'* @details
'* クリップボードの内容を、以下の順序で貼り付けを試みます:
'* 1. 数式として貼り付け (`Paste:=xlPasteFormulas`)。
'* 2. テキストのみを貼り付け (`NoHTMLFormatting:=True`)。
'* 3. いずれも失敗した場合、標準の貼り付けを実行。
Public Sub PasteFormulas()
    On Error Resume Next

    Call Selection.PasteSpecial(Paste:=xlPasteFormulas, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
    If Err.Number = 0 Then Exit Sub
    Err.Clear

    Call ActiveSheet.PasteSpecial(Format:="HTML", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True)
    If Err.Number = 0 Then Exit Sub
    Err.Clear

    Call ActiveSheet.Paste
End Sub


' #############################################################################
'
' ファイル操作関連
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' ファイルシステム パス文字列関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* パス内の特殊文字を全角文字に置換します。
'*
'* @param Path 置換対象のパス文字列
'* @return 特殊文字を全角文字に置換したパス文字列
'*
'* @details
'* ファイルシステムで使用できない特殊文字を全角文字に置換します。
'* 以下の文字が置換されます: `\`, `/`, `:`, `*`, `?`, `"`, `<`, `>`, `|`。
Public Function ReplaceSpecialCharacterOnFileSystemPath(ByVal Path As String) As String
    Dim result_value As String

    result_value = Path

    result_value = Replace(result_value, "\", "＼")
    result_value = Replace(result_value, "/", "／")
    result_value = Replace(result_value, ":", "：")
    result_value = Replace(result_value, "*", "＊")
    result_value = Replace(result_value, "?", "？")
    result_value = Replace(result_value, """", "″")
    result_value = Replace(result_value, "<", "＜")
    result_value = Replace(result_value, ">", "＞")
    result_value = Replace(result_value, "|", "｜")

    ReplaceSpecialCharacterOnFileSystemPath = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' パス結合
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* パスを結合します。
'*
'* @param Path1 最初のパス文字列
'* @param Path2 結合する2番目のパス文字列
'* @param Paths 可変長の追加パス文字列
'* @return 結合されたパス文字列
'*
'* @details
'* 指定された複数の文字列をパス区切り文字で結合します。
'* 追加のパス文字列を可変長引数として渡すことができます。
'*
'* @note
'* JoinPath は文字列結合 API です。Path2 以降に絶対パスまたは URL を渡しても、先行パスを置き換えず通常のパス要素として結合します。
'* Path1 が URL の場合、Path2 以降は URL 用の区切り文字へ正規化しますが、Path1 側の query / fragment は移動しません。
'* そのため Path1 に ? または # が含まれる場合、後続パスは URL サフィックスの後ろへ結合されます。
Public Function JoinPath(ByVal Path1 As String, ByVal Path2 As String, ParamArray Paths() As Variant) As String
    Dim result_value As String

    result_value = pJoinPathCore(Path1, Path2)

    If UBound(Paths) = -1 Then
        JoinPath = result_value
        Exit Function
    End If

    Dim item_idx As Long
    For item_idx = LBound(Paths) To UBound(Paths)
        result_value = pJoinPathCore(result_value, Paths(item_idx))
    Next item_idx

    JoinPath = result_value
End Function

Private Function pJoinPathCore(ByVal Path1 As String, ByVal Path2 As String) As String
    If IsUrlPath(Path1) Then
        pJoinPathCore = pJoinPathCoreWithSep(Path1, pNormalizeUrlRelativePathForJoin(Path2), "/")
    Else
        pJoinPathCore = pJoinPathCoreWithSep(pNormalizeWindowsPathSeparators(Path1), pNormalizeWindowsPathSeparators(Path2), G_FS_PATH_SEP)
    End If
End Function

Private Function pJoinPathCoreWithSep(ByVal Path1 As String, ByVal Path2 As String, ByVal PathSep As String) As String
    If EndsWith(Path1, PathSep) Then
        If StartsWith(Path2, PathSep) Then
            pJoinPathCoreWithSep = Left$(Path1, Len(Path1) - 1) & Path2
        Else
            pJoinPathCoreWithSep = Path1 & Path2
        End If
    Else
        If StartsWith(Path2, PathSep) Then
            pJoinPathCoreWithSep = Path1 & Path2
        Else
            pJoinPathCoreWithSep = Path1 & PathSep & Path2
        End If
    End If
End Function

Private Function pNormalizeUrlRelativePathForJoin(ByVal PathWithSuffix As String) As String
    Dim path_part As String
    Dim path_suffix As String
    Call pSplitPathSuffix(path_part, path_suffix, PathWithSuffix)
    pNormalizeUrlRelativePathForJoin = pNormalizePathSeparators(path_part) & path_suffix
End Function

Private Sub pSplitPathSuffix(ByRef PathPart As String, ByRef PathSuffix As String, ByVal PathWithSuffix As String)
    Dim query_pos As Long
    query_pos = InStr(1, PathWithSuffix, "?", vbBinaryCompare)

    Dim fragment_pos As Long
    fragment_pos = InStr(1, PathWithSuffix, "#", vbBinaryCompare)

    Dim suffix_pos As Long
    If query_pos > 0 And fragment_pos > 0 Then
        suffix_pos = IIf(query_pos < fragment_pos, query_pos, fragment_pos)
    ElseIf query_pos > 0 Then
        suffix_pos = query_pos
    Else
        suffix_pos = fragment_pos
    End If

    If suffix_pos = 0 Then
        PathPart = PathWithSuffix
        PathSuffix = ""
    Else
        PathPart = Left$(PathWithSuffix, suffix_pos - 1)
        PathSuffix = Mid$(PathWithSuffix, suffix_pos)
    End If
End Sub

Private Function pNormalizePathSeparators(ByVal PathLikeString As String) As String
    pNormalizePathSeparators = Replace(PathLikeString, G_FS_PATH_SEP, "/")
End Function

Private Function pNormalizeWindowsPathSeparators(ByVal PathLikeString As String) As String
    pNormalizeWindowsPathSeparators = Replace(PathLikeString, "/", G_FS_PATH_SEP)
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 絶対パス化
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 基準ディレクトリを使ってパスを絶対パス化します。
'*
'* @param ParentAbsolutePath 基準にする絶対ディレクトリ パス。
'* @param PathLikeString 絶対パス化するパス文字列。
'* @return 絶対パス。
'*
'* @details
'* PathLikeString が相対パスの場合は ParentAbsolutePath を基準にして絶対パス化します。
'* PathLikeString が絶対パスの場合は ParentAbsolutePath を使わず、正規化した絶対パスを返します。
'* `\foo` のようなルート相対パスは ParentAbsolutePath のルートを基準にします。
'* `C:foo` のようなドライブ相対パスはカレントディレクトリ依存になるためエラーにします。
'* URL は FSO に渡さず、パス部分だけを `/` 区切りで正規化します。
'*
'* @note
'* PathLikeString が絶対パスまたは URL の場合は ParentAbsolutePath を使わず、そのパスを正規化して返します。
'* PathLikeString が相対パスの場合は JoinPath と同じ結合規則に従います。
'* そのため ParentAbsolutePath が URL サフィックスを含む場合、後続相対パスはサフィックスの後ろへ結合されます。
Public Function GetAbsolutePathFromParent(ByVal ParentAbsolutePath As String, ByVal PathLikeString As String) As String
    If IsDriveRelativePath(ParentAbsolutePath) Or Not IsAbsolutePath(ParentAbsolutePath) Then Err.Raise vbObjectError + 1, "Function GetAbsolutePathFromParent", "ParentAbsolutePath には絶対パスを指定してください。(" & ParentAbsolutePath & ")"
    If IsDriveRelativePath(PathLikeString) Then Err.Raise vbObjectError + 1, "Function GetAbsolutePathFromParent", "ドライブ相対パスは指定できません。(" & PathLikeString & ")"

    Dim parent_path As String
    parent_path = pNormalizeAbsolutePath(ParentAbsolutePath)

    Dim target_path As String
    If PathLikeString = "" Or PathLikeString = "." Or PathLikeString = ".\" Or PathLikeString = "./" Then
        target_path = parent_path
    ElseIf IsAbsolutePath(PathLikeString) Then
        target_path = PathLikeString
    ElseIf pIsRootRelativePath(PathLikeString) Then
        target_path = JoinPath(GetPathRoot(parent_path), Mid$(pNormalizePathPartSeparators(PathLikeString), 2) & pGetPathSuffix(PathLikeString))
    Else
        target_path = JoinPath(parent_path, PathLikeString)
    End If

    GetAbsolutePathFromParent = pNormalizeAbsolutePath(target_path)
End Function

Private Function pNormalizeAbsolutePath(ByVal PathLikeString As String) As String
    If IsUrlPath(PathLikeString) Then
        pNormalizeAbsolutePath = pNormalizeUrlPath(PathLikeString)
    Else
        pNormalizeAbsolutePath = pNormalizeLocalAbsolutePath(PathLikeString)
    End If
End Function

Private Function pNormalizeUrlPath(ByVal UrlPath As String) As String
    Dim url_prefix As String
    Dim url_path As String
    Dim url_suffix As String
    Call pSplitUrl(url_prefix, url_path, url_suffix, UrlPath)
    pNormalizeUrlPath = url_prefix & pNormalizeSlashPath(url_path, True) & url_suffix
End Function

Private Sub pSplitUrl(ByRef UrlPrefix As String, ByRef UrlPathPart As String, ByRef UrlSuffix As String, ByVal UrlPath As String)
    Dim scheme_sep_pos As Long
    scheme_sep_pos = InStr(1, UrlPath, "://", vbBinaryCompare)
    If scheme_sep_pos = 0 Then Err.Raise vbObjectError + 1, "Function GetPathRoot", "URL 形式ではありません。(" & UrlPath & ")"

    Dim scan_start As Long
    scan_start = scheme_sep_pos + Len("://")

    Dim path_start As Long
    Dim suffix_start As Long
    Dim char_idx As Long
    For char_idx = scan_start To Len(UrlPath)
        Dim current_char As String
        current_char = Mid$(UrlPath, char_idx, 1)
        If current_char = "/" Or current_char = G_FS_PATH_SEP Then
            path_start = char_idx
            Exit For
        ElseIf current_char = "?" Or current_char = "#" Then
            suffix_start = char_idx
            Exit For
        End If
    Next char_idx

    If path_start = 0 Then
        If suffix_start = 0 Then
            UrlPrefix = UrlPath
            UrlSuffix = ""
        Else
            UrlPrefix = Left$(UrlPath, suffix_start - 1)
            UrlSuffix = Mid$(UrlPath, suffix_start)
        End If
        UrlPathPart = "/"
        Exit Sub
    End If

    UrlPrefix = Left$(UrlPath, path_start - 1)
    Call pSplitPathSuffix(UrlPathPart, UrlSuffix, Mid$(UrlPath, path_start))
    UrlPathPart = pNormalizePathSeparators(UrlPathPart)
End Sub

Private Function pNormalizeLocalAbsolutePath(ByVal PathLikeString As String) As String
    Dim slash_path As String
    slash_path = pNormalizePathSeparators(PathLikeString)

    If IsDriveRelativePath(slash_path) Then Err.Raise vbObjectError + 1, "Function GetAbsolutePathFromParent", "ドライブ相対パスは指定できません。(" & PathLikeString & ")"

    If IsDriveAbsolutePath(slash_path) Then
        Dim drive_root As String
        drive_root = UCase$(Left$(slash_path, 1)) & ":/"

        Dim drive_rest As String
        drive_rest = Mid$(slash_path, Len(drive_root) + 1)

        Dim drive_normalized_path As String
        drive_normalized_path = pNormalizeSlashPath("/" & drive_rest, True)
        pNormalizeLocalAbsolutePath = pNormalizeWindowsPathSeparators(drive_root & Mid$(drive_normalized_path, 2))
    ElseIf IsUncPath(slash_path) Then
        Dim unc_root As String
        unc_root = pGetUncPathRootSlash(slash_path)

        Dim unc_rest As String
        unc_rest = Mid$(slash_path, Len(unc_root) + 1)

        Dim unc_normalized_path As String
        unc_normalized_path = pNormalizeSlashPath("/" & unc_rest, True)
        pNormalizeLocalAbsolutePath = pNormalizeWindowsPathSeparators(unc_root & Mid$(unc_normalized_path, 2))
    Else
        Err.Raise vbObjectError + 1, "Function GetAbsolutePathFromParent", "絶対パスを指定してください。(" & PathLikeString & ")"
    End If
End Function

Private Function pNormalizeSlashPath(ByVal PathPart As String, ByVal IsAbsolute As Boolean) As String
    Dim slash_path As String
    slash_path = pNormalizePathSeparators(PathPart)

    Dim path_parts As Variant
    path_parts = Split(slash_path, "/")

    Dim stack_items() As String
    Dim stack_count As Long
    stack_count = 0

    Dim path_part As Variant
    For Each path_part In path_parts
        Dim part_text As String
        part_text = CStr(path_part)

        If part_text = "" Or part_text = "." Then
            ' no operation
        ElseIf part_text = ".." Then
            If stack_count > 0 Then
                stack_count = stack_count - 1
            ElseIf Not IsAbsolute Then
                Call pPushPathSegment(stack_items, stack_count, part_text)
            End If
        Else
            Call pPushPathSegment(stack_items, stack_count, part_text)
        End If
    Next path_part

    Dim result_path As String
    If IsAbsolute Then result_path = "/"

    Dim item_idx As Long
    For item_idx = 0 To stack_count - 1
        If result_path <> "" And Not EndsWith(result_path, "/") Then result_path = result_path & "/"
        result_path = result_path & stack_items(item_idx)
    Next item_idx

    If result_path = "" Then
        If IsAbsolute Then
            result_path = "/"
        Else
            result_path = "."
        End If
    End If

    pNormalizeSlashPath = result_path
End Function

Private Sub pPushPathSegment(ByRef StackItems() As String, ByRef StackCount As Long, ByVal PathSegment As String)
    If StackCount = 0 Then
        ReDim StackItems(0 To 0)
    Else
        ReDim Preserve StackItems(0 To StackCount)
    End If

    StackItems(StackCount) = PathSegment
    StackCount = StackCount + 1
End Sub

Private Function pIsRootRelativePath(ByVal PathWithSuffix As String) As Boolean
    pIsRootRelativePath = StartsWith(pNormalizePathPartSeparators(PathWithSuffix), "/")
End Function

Private Function pNormalizePathPartSeparators(ByVal PathWithSuffix As String) As String
    Dim path_part As String
    Dim path_suffix As String
    Call pSplitPathSuffix(path_part, path_suffix, PathWithSuffix)
    pNormalizePathPartSeparators = pNormalizePathSeparators(path_part)
End Function

Private Function pGetPathSuffix(ByVal PathWithSuffix As String) As String
    Dim path_part As String
    Dim path_suffix As String
    Call pSplitPathSuffix(path_part, path_suffix, PathWithSuffix)
    pGetPathSuffix = path_suffix
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' パス ルート取得
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* パスのルート部分を取得します。
'*
'* @param PathLikeString ルートを取得するパス文字列。
'* @return ドライブパスでは `C:\`、UNC パスでは `\\server\share\`、URL では `scheme://authority/` 形式のルート。
'*
'* @details
'* 先頭がアルファベット 1 文字、2 文字目が `:` の場合はドライブパスとして扱います。
'* 先頭 2 文字が `\\` または `//` の場合は UNC パスとして扱います。
'* `scheme://` 形式の URL も絶対パスとして扱います。
'* それ以外のパス文字列ではエラーを発生させます。
Public Function GetPathRoot(ByVal PathLikeString As String) As String
    If IsUrlPath(PathLikeString) Then
        Dim url_prefix As String
        Dim url_path As String
        Dim url_suffix As String
        Call pSplitUrl(url_prefix, url_path, url_suffix, PathLikeString)
        GetPathRoot = url_prefix & "/"
    ElseIf IsDriveRelativePath(PathLikeString) Then
        Err.Raise vbObjectError + 1, "Function GetPathRoot", "ドライブ相対パスのルートは解決できません。(" & PathLikeString & ")"
    ElseIf IsDriveAbsolutePath(PathLikeString) Then
        GetPathRoot = UCase$(Left$(PathLikeString, 1)) & ":" & G_FS_PATH_SEP
    ElseIf IsUncPath(PathLikeString) Then
        GetPathRoot = pNormalizeWindowsPathSeparators(pGetUncPathRootSlash(PathLikeString))
    Else
        Err.Raise vbObjectError + 1, "Function GetPathRoot", "ドライブパス、UNC パス、または URL を指定してください。(" & PathLikeString & ")"
    End If
End Function

Private Function pGetUncPathRootSlash(ByVal UncPath As String) As String
    Dim slash_path As String
    slash_path = pNormalizePathSeparators(UncPath)

    Dim server_sep As Long
    server_sep = InStr(3, slash_path, "/")
    If server_sep <= 3 Then Err.Raise vbObjectError + 1, "Function GetPathRoot", "UNC パスのルートを解決できません。(" & UncPath & ")"

    Dim share_sep As Long
    share_sep = InStr(server_sep + 1, slash_path, "/")

    Dim share_name As String
    If share_sep = 0 Then
        share_name = Mid$(slash_path, server_sep + 1)
        If share_name = "" Then Err.Raise vbObjectError + 1, "Function GetPathRoot", "UNC パスのルートを解決できません。(" & UncPath & ")"

        pGetUncPathRootSlash = slash_path
        If Not EndsWith(pGetUncPathRootSlash, "/") Then pGetUncPathRootSlash = pGetUncPathRootSlash & "/"
    Else
        share_name = Mid$(slash_path, server_sep + 1, share_sep - server_sep - 1)
        If share_name = "" Then Err.Raise vbObjectError + 1, "Function GetPathRoot", "UNC パスのルートを解決できません。(" & UncPath & ")"

        pGetUncPathRootSlash = Left$(slash_path, share_sep)
    End If
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' パス種別判定
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* ドライブ指定を持つパスか判定します。
'*
'* @param TestPath 判定対象のパス文字列。
'* @return ドライブ指定を持つパスの場合は True、それ以外は False。
'*
'* @details
'* `C:\foo`、`C:/foo`、`C:foo`、`C:` のように先頭がドライブ指定の場合は True とします。
Public Function IsDrivePath(ByVal TestPath As String) As Boolean
    If Len(TestPath) < 2 Then Exit Function
    If Mid$(TestPath, 2, 1) <> ":" Then Exit Function
    IsDrivePath = pIsAsciiAlphabet(Left$(TestPath, 1))
End Function

Private Function pIsAsciiAlphabet(ByVal TestChar As String) As Boolean
    If Len(TestChar) <> 1 Then Exit Function
    pIsAsciiAlphabet = ("A" <= UCase$(TestChar) And UCase$(TestChar) <= "Z")
End Function

'* ドライブ絶対パスか判定します。
'*
'* @param TestPath 判定対象のパス文字列。
'* @return ドライブ絶対パスの場合は True、それ以外は False。
'*
'* @details
'* `C:\foo` または `C:/foo` のようなドライブ絶対パスを True とします。`C:foo` のようなドライブ相対パスは False です。
Public Function IsDriveAbsolutePath(ByVal TestPath As String) As Boolean
    If Not IsDrivePath(TestPath) Then Exit Function
    If Len(TestPath) < 3 Then Exit Function
    IsDriveAbsolutePath = (Mid$(TestPath, 3, 1) = G_FS_PATH_SEP Or Mid$(TestPath, 3, 1) = "/")
End Function

'* ドライブ相対パスか判定します。
'*
'* @param TestPath 判定対象のパス文字列。
'* @return ドライブ相対パスの場合は True、それ以外は False。
'*
'* @details
'* `C:foo` または `C:` のようにドライブ指定を持ち、区切り文字で始まらないパスを True とします。
Public Function IsDriveRelativePath(ByVal TestPath As String) As Boolean
    If Not IsDrivePath(TestPath) Then Exit Function
    IsDriveRelativePath = Not IsDriveAbsolutePath(TestPath)
End Function

'* UNC パスか判定します。
'*
'* @param TestPath 判定対象のパス文字列。
'* @return UNC パスの場合は True、それ以外は False。
'*
'* @details
'* `\\server\share` または `//server/share` のような UNC パスを True とします。
Public Function IsUncPath(ByVal TestPath As String) As Boolean
    If Not (Left$(TestPath, 2) = "\\" Or Left$(TestPath, 2) = "//") Then Exit Function

    On Error Resume Next
    Err.Clear
    Dim root_path As String
    root_path = pGetUncPathRootSlash(TestPath)
    IsUncPath = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'* URL 形式のパスか判定します。
'*
'* @param TestPath 判定対象のパス文字列。
'* @return URL 形式の場合は True、それ以外は False。
'*
'* @details
'* `scheme://` の形式を URL として扱います。スキーム名は英字で始まり、英数字、`+`、`.`、`-` を含めることができます。
Public Function IsUrlPath(ByVal TestPath As String) As Boolean
    Dim scheme_sep_pos As Long
    scheme_sep_pos = InStr(1, TestPath, "://", vbBinaryCompare)
    If scheme_sep_pos <= 1 Then Exit Function

    Dim scheme_name As String
    scheme_name = Left$(TestPath, scheme_sep_pos - 1)
    If Not pIsAsciiAlphabet(Left$(scheme_name, 1)) Then Exit Function

    Dim char_idx As Long
    For char_idx = 2 To Len(scheme_name)
        Dim scheme_char As String
        scheme_char = Mid$(scheme_name, char_idx, 1)
        If Not (pIsAsciiAlphabet(scheme_char) Or pIsAsciiDigit(scheme_char) Or scheme_char = "+" Or scheme_char = "." Or scheme_char = "-") Then Exit Function
    Next char_idx

    IsUrlPath = True
End Function

Private Function pIsAsciiDigit(ByVal TestChar As String) As Boolean
    If Len(TestChar) <> 1 Then Exit Function
    pIsAsciiDigit = ("0" <= TestChar And TestChar <= "9")
End Function

'* あるパス文字列が絶対パスかを判定します。
'*
'* @param TestPath 判定対象のパス文字列
'* @return 絶対パスの場合は True、それ以外は False
'*
'* @details
'* 入力されたパス文字列が絶対パスかどうかを判定します。
'* URL 形式の文字列も絶対パスとして扱います。
Public Function IsAbsolutePath(ByVal TestPath As String) As Boolean
    If IsDriveRelativePath(TestPath) Then Exit Function

    Err.Clear
    On Error Resume Next
    Call GetPathRoot(TestPath)
    IsAbsolutePath = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' パス分割
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* パスを親パスと末端パスに分割します。
'*
'* @param ParentPath [出力] 親パス。
'* @param LeafPath [出力] 末端パス。
'* @param Path 入力パス文字列。
'* @param IgnoreEndSep [省略可] 入力パス文字列の末尾の区切り文字を無視するか否か。
'*
'* @details
'* Path を最後のパス区切り文字で親パスと末端パスに分割します。
'* URL では query / fragment 内の区切り文字を分割対象から除外し、PathSuffix は LeafPath に含めて返します。
Public Sub SplitPath( _
        ByRef ParentPath As String, _
        ByRef LeafPath As String, _
        ByVal Path As String, _
        Optional ByVal IgnoreEndSep As Boolean = False)

    If IsUrlPath(Path) Then
        Call pSplitUrlPath(ParentPath, LeafPath, Path, IgnoreEndSep)
    Else
        Call pSplitWindowsPath(ParentPath, LeafPath, Path, IgnoreEndSep)
    End If
End Sub

Private Sub pSplitUrlPath( _
        ByRef ParentPath As String, _
        ByRef LeafPath As String, _
        ByVal UrlPath As String, _
        ByVal IgnoreEndSep As Boolean)

    Dim url_prefix As String
    Dim url_path As String
    Dim url_suffix As String
    Call pSplitUrl(url_prefix, url_path, url_suffix, UrlPath)

    Dim normalized_path As String
    normalized_path = pNormalizeSlashPath(url_path, True)
    If EndsWith(url_path, "/") And normalized_path <> "/" Then normalized_path = normalized_path & "/"
    url_path = normalized_path

    If EndsWith(url_path, "/") And url_path <> "/" Then
        url_path = Left$(url_path, Len(url_path) - 1)
        If Not IgnoreEndSep Then
            ParentPath = url_prefix & url_path
            LeafPath = url_suffix
            Exit Sub
        End If
    End If

    Dim last_sep As Long
    last_sep = InStrRev(url_path, "/")
    If last_sep <= 1 Then
        ParentPath = url_prefix & "/"
        LeafPath = Mid$(url_path, 2) & url_suffix
    Else
        ParentPath = url_prefix & Left$(url_path, last_sep - 1)
        LeafPath = Mid$(url_path, last_sep + 1) & url_suffix
    End If
End Sub

Private Sub pSplitWindowsPath( _
        ByRef ParentPath As String, _
        ByRef LeafPath As String, _
        ByVal WindowsPath As String, _
        ByVal IgnoreEndSep As Boolean)

    WindowsPath = pNormalizeWindowsPathSeparators(WindowsPath)

    Dim root_path As String
    root_path = pGetWindowsPathRootForSplit(WindowsPath)

    If root_path <> "" And StrComp(WindowsPath, root_path, vbTextCompare) = 0 Then
        If IgnoreEndSep Then
            ParentPath = ""
            LeafPath = WindowsPath
        Else
            ParentPath = WindowsPath
            LeafPath = ""
        End If
        Exit Sub
    End If

    If EndsWith(WindowsPath, G_FS_PATH_SEP) Then
        WindowsPath = Left$(WindowsPath, Len(WindowsPath) - 1)
        If Not IgnoreEndSep Then
            ParentPath = WindowsPath
            LeafPath = ""
            Exit Sub
        End If
    End If

    Dim last_sep As Long
    last_sep = InStrRev(WindowsPath, G_FS_PATH_SEP)
    If last_sep = 0 Then
        ParentPath = ""
        LeafPath = WindowsPath
    ElseIf root_path <> "" And last_sep = Len(root_path) Then
        ParentPath = root_path
        LeafPath = Mid$(WindowsPath, last_sep + 1)
    Else
        ParentPath = Left$(WindowsPath, last_sep - 1)
        LeafPath = Mid$(WindowsPath, last_sep + 1)
    End If
End Sub

Private Function pGetWindowsPathRootForSplit(ByVal WindowsPath As String) As String
    If IsDriveAbsolutePath(WindowsPath) Then
        pGetWindowsPathRootForSplit = Left$(WindowsPath, 3)
    ElseIf IsUncPath(WindowsPath) Then
        pGetWindowsPathRootForSplit = pNormalizeWindowsPathSeparators(pGetUncPathRootSlash(WindowsPath))
    End If
End Function

'* 末端パスをベース名、拡張子、パスサフィックスに分解します。
'*
'* @param BaseName [出力] 拡張子とパスサフィックスを除いた末端名。
'* @param Extension [出力] `.` を含む拡張子。拡張子がない場合は空文字列。
'* @param PathSuffix [出力] query / fragment を含むパスサフィックス。
'* @param LeafPath 入力末端パス文字列。
'* @param AsUrl [省略可] LeafPath を URL の末端パスとして扱い、query / fragment を分離するか。既定値は False。
'*
'* @details
'* 既定では `?` や `#` を末端名の一部として扱い、最後の `.` で分割します。
'* AsUrl が True の場合だけ、`?` または `#` 以降を PathSuffix とし、それより前の文字列を最後の `.` で分割します。
Public Sub ParseLeafPath( _
        ByRef BaseName As String, _
        ByRef Extension As String, _
        ByRef PathSuffix As String, _
        ByVal LeafPath As String, _
        Optional ByVal AsUrl As Boolean = False)

    Dim leaf_body As String
    If AsUrl Then
        Call pSplitPathSuffix(leaf_body, PathSuffix, LeafPath)
    Else
        leaf_body = LeafPath
        PathSuffix = ""
    End If

    Dim last_period As Long
    last_period = InStrRev(leaf_body, ".")

    If 1 < last_period Then
        BaseName = Left$(leaf_body, last_period - 1)
        Extension = Mid$(leaf_body, last_period)
    Else
        BaseName = leaf_body
        Extension = ""
    End If
End Sub

'* パスの最後の部分を除いたパスを取得します。
'*
'* @param Path 入力パス文字列
'* @param IgnoreEndSep [省略可] 入力パス文字列の末尾の区切り文字を無視するか否か
'* @return 最後の部分を除いたパス文字列
'*
'* @details
'* 入力されたパス文字列の最後のパス区切り文字より前の部分を返します。
'* 例えば、`Path\to\File` の場合は `Path\to` を返します。
Public Function GetParentPath(ByVal Path As String, Optional ByVal IgnoreEndSep As Boolean = False) As String
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, Path, IgnoreEndSep:=IgnoreEndSep)

    GetParentPath = parent_path
End Function

'* パスの最後の部分を取得します。
'*
'* @param Path 入力パス文字列
'* @param BaseName [省略可] ベース名を含めるか。規定値は True
'* @param Extension [省略可] 拡張子を含めるか。規定値は True
'* @param IgnoreEndSep [省略可] 入力パス文字列の末尾の区切り文字を無視するか否か
'* @return 最後の部分 (ファイル名またはフォルダ名)
'*
'* @details
'* 入力されたパス文字列の最後のパス区切り文字より後の部分を返します。
'* オプション引数 `Extension` を False にすると、拡張子を除去した結果を返します。
'* URL の PathSuffix は返却に含めません。PathSuffix が必要な場合は SplitPath と ParseLeafPath を使用します。
'* 例えば、`Path\to\File.txt` の場合は `File` を返します。
Public Function GetLeafFromPath(ByVal Path As String, Optional ByVal BaseName As Boolean = True, Optional ByVal Extension As Boolean = True, Optional ByVal IgnoreEndSep As Boolean = False) As String
    If Not BaseName And Not Extension Then Exit Function

    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, Path, IgnoreEndSep:=IgnoreEndSep)

    Dim base_name As String
    Dim file_ext As String
    Dim path_suffix As String
    Call ParseLeafPath(base_name, file_ext, path_suffix, leaf_path, AsUrl:=IsUrlPath(Path))

    If BaseName Then GetLeafFromPath = GetLeafFromPath & base_name
    If Extension Then GetLeafFromPath = GetLeafFromPath & file_ext
End Function


' #############################################################################
'
' 日付操作関連
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 日付判定
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* ある日付が、未設定かどうかを判定します。
'*
'* @param TestDate 判定対象の日付
'* @return 未設定の日付の場合は True、それ以外は False
'*
'* @details
'* 指定された日付が `G_DATE_MINIMUM` 未満、または `G_DATE_NULL` と等しい場合に未設定とみなして True を返します。
'* `G_DATE_NULL` は未設定の日付を表し、`G_DATE_MINIMUM` は有効な日付の最小値を表します。
Public Function IsNullDate(ByVal TestDate As Date) As Boolean
    If TestDate < G_DATE_MINIMUM Then
        IsNullDate = True
    Else
        IsNullDate = False
    End If
End Function


' #############################################################################
'
' 各種変換
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 文字列・配列変換
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 文字列を文字の配列に変換します。
'*
'* @param Expression 文字列。
'* @return 文字の配列。
'*
'* @details
'* 文字列を文字の配列に変換します。
Public Function ConvertStringToCharArray(ByVal Expression As String) As String()
    Dim result_ubound As Long
    result_ubound = Len(Expression) - 1

    If result_ubound < 0 Then
        ConvertStringToCharArray = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    ReDim result_value(0 To result_ubound)

    Dim char_idx As Long
    For char_idx = 0 To result_ubound
        result_value(char_idx) = Mid(Expression, char_idx + 1, 1)
    Next char_idx

    ConvertStringToCharArray = result_value
End Function

'* String() 型の配列を Variant() 型の配列に変換します。
'*
'* @param StringArray 変換対象の String 型の配列
'* @return Variant 型の配列
'*
'* @details
'* String 型の配列を Variant 型に変換します。各要素の内容はそのまま維持されます。
'* 0 要素配列は 0 要素の Variant 配列として返します。
Public Function ConvertArrayStringToVariant(ByRef StringArray() As String) As Variant()
    If IsEmptyArray(StringArray) Then
        ConvertArrayStringToVariant = EmptyVariantArray()
        Exit Function
    End If

    Dim result_value() As Variant
    Dim item_idx As Long

    ReDim result_value(LBound(StringArray) To UBound(StringArray))

    For item_idx = LBound(StringArray) To UBound(StringArray)
        result_value(item_idx) = StringArray(item_idx)
    Next item_idx

    ConvertArrayStringToVariant = result_value
End Function

'* Variant() 型の配列を String() 型の配列に変換します。
'*
'* @param VariantArray 変換対象の Variant 型の配列
'* @return String 型の配列
'*
'* @details
'* Variant 型の配列を String 型に変換します。各要素の内容はそのまま維持されます。
'* 0 要素配列は 0 要素の String 配列として返します。
Public Function ConvertArrayVariantToString(ByRef VariantArray() As Variant) As String()
    If IsEmptyArray(VariantArray) Then
        ConvertArrayVariantToString = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    Dim item_idx As Long

    ReDim result_value(LBound(VariantArray) To UBound(VariantArray))

    For item_idx = LBound(VariantArray) To UBound(VariantArray)
        result_value(item_idx) = VariantArray(item_idx)
    Next item_idx

    ConvertArrayVariantToString = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Boolean 変換
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Boolean 型の値を文字列に変換します。
'*
'* @param BooleanValue 変換対象の Boolean 値
'* @param FlagOnString True の場合に返す文字列 (既定値は "■")
'* @param FlagOffString False の場合に返す文字列 (既定値は空文字列)
'* @return 変換後の文字列
'*
'* @details
'* Boolean 値を指定された文字列に変換します。
'* True の場合は `FlagOnString`、False の場合は `FlagOffString` を返します。
Public Function ConvertBooleanToString( _
        ByVal BooleanValue As Boolean, _
        Optional ByVal FlagOnString As String = "■", _
        Optional ByVal FlagOffString As String = "") As String

    If BooleanValue Then
        ConvertBooleanToString = FlagOnString
    Else
        ConvertBooleanToString = FlagOffString
    End If
End Function

'* 文字列を Boolean 型の値に変換します。
'*
'* @param FlagValue 変換対象の文字列
'* @param FlagOnString True とみなす文字列 (既定値は "■")
'* @param FlagOffString False とみなす文字列 (既定値は空文字列)
'* @return 変換後の Boolean 値
'*
'* @details
'* 指定された文字列を基に Boolean 値を判定します。
'* `FlagOnString` の場合は True、`FlagOffString` の場合は False を返します。
'* それ以外の値の場合はエラーを発生させます。
Public Function ConvertStringToBoolean( _
        ByVal FlagValue As String, _
        Optional ByVal FlagOnString As String = "■", _
        Optional ByVal FlagOffString As String = "") As Boolean

    If FlagValue = FlagOnString Then
        ConvertStringToBoolean = True
    ElseIf FlagValue = FlagOffString Then
        ConvertStringToBoolean = False
    Else
        Err.Raise Number:=vbObjectError + 1, Source:="Sub CreateBackupFile", Description:="許容される FlagValue は「" & FlagOnString & "」か「" & FlagOffString & "」です。(" & FlagValue & ")"
    End If
End Function


' #############################################################################
'
' 配列関連
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 配列次元変換
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 2 次元配列を 1 次元配列に変換します。
'*
'* @param OriginalArray 2 次元配列。
'* @param ColumnDirection [省略可] 列方向 (縦方向、1 次元目) に読み取っていくか。デフォルトは False で行列方向 (横方向、2 次元目) に読み取っていきます。
'* @return 1 次元配列。
'*
'* @details
'* 2 次元配列を 0 ベースの 1 次元配列に変換します。
Public Function ConvertArray2dTo1d( _
        ByVal OriginalArray As Variant, _
        Optional ByVal ColumnDirection As Boolean = False) As Variant()

    If Not IsArray(OriginalArray) Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray2dTo1d", Description:="引数が配列ではありません。(" & TypeName(OriginalArray) & ")"
    End If

    Dim row_low_bnd As Long
    Dim row_up_bnd As Long
    Dim col_low_bnd As Long
    Dim col_up_bnd As Long
    Dim total_cnt As Long
    Dim res_arr() As Variant
    Dim arr_idx As Long
    Dim row_index As Long
    Dim col_index As Long

    row_low_bnd = LBound(OriginalArray, 1)
    row_up_bnd = UBound(OriginalArray, 1)
    col_low_bnd = LBound(OriginalArray, 2)
    col_up_bnd = UBound(OriginalArray, 2)

    total_cnt = (row_up_bnd - row_low_bnd + 1) * (col_up_bnd - col_low_bnd + 1)
    ReDim res_arr(0 To total_cnt - 1)

    arr_idx = 0
    If ColumnDirection Then
        For col_index = col_low_bnd To col_up_bnd
            For row_index = row_low_bnd To row_up_bnd
                Call pAssignVariant(res_arr(arr_idx), OriginalArray(row_index, col_index))
                arr_idx = arr_idx + 1
            Next row_index
        Next col_index
    Else
        For row_index = row_low_bnd To row_up_bnd
            For col_index = col_low_bnd To col_up_bnd
                Call pAssignVariant(res_arr(arr_idx), OriginalArray(row_index, col_index))
                arr_idx = arr_idx + 1
            Next col_index
        Next row_index
    End If

    ConvertArray2dTo1d = res_arr
End Function

'* 1 次元配列を 2 次元配列に変換します。
'*
'* @param OriginalArray 1 次元配列。
'* @param RowLBound [省略可] 行インデックス (1 次元目) の下限。デフォルトは 1 です。
'* @param ColumnLBound [省略可] 列インデックス (2 次元目) の下限。デフォルトは 1 です。
'* @param RowCount [省略可] 行 (1 次元目) の要素数。デフォルトは無制限です。
'* @param ColCount [省略可] 列 (2 次元目) の要素数。デフォルトは 1 です。
'* @param ColumnDirection [省略可] 列方向 (縦方向、1 次元目) に埋めていくか。デフォルトは False で行列方向 (横方向、2 次元目) に埋めていきます。
'* @return 2 次元配列。
'*
'* @details
'* 1 次元配列を 2 次元配列に変換します。
'* 0 要素配列は対象外です。呼び出し側で空配列を分岐してください。
Public Function ConvertArray1dTo2d( _
        ByVal OriginalArray As Variant, _
        Optional ByVal RowLBound As Long = 1, _
        Optional ByVal ColLBound As Long = 1, _
        Optional ByVal RowCount As Long = -1, _
        Optional ByVal ColCount As Long = 1, _
        Optional ByVal ColumnDirection As Boolean = False) As Variant()

    If Not IsArray(OriginalArray) Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray1dTo2d", Description:="引数が配列ではありません。(" & TypeName(OriginalArray) & ")"
    End If

    If RowCount <= 0 And ColCount <= 0 Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray1dTo2d", Description:="行数と列数の両方が無制限として指定 (0 以下を指定) されました。"
    End If

    If 0 < RowCount And 0 < ColCount Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray1dTo2d", Description:="行数と列数の両方の要素数が指定されました。"
    End If

    Dim orig_count As Long
    orig_count = UBound(OriginalArray) - LBound(OriginalArray) + 1

    Dim count_1 As Long, count_2 As Long
    If 0 < RowCount Then
        count_1 = RowCount
        count_2 = (orig_count + RowCount - 1) \ RowCount
    Else
        count_1 = (orig_count + ColCount - 1) \ ColCount
        count_2 = ColCount
    End If

    Dim max_idx_1 As Long, max_idx_2 As Long
    max_idx_1 = RowLBound + count_1 - 1
    max_idx_2 = ColLBound + count_2 - 1

    Dim result_arr() As Variant
    ReDim result_arr(RowLBound To max_idx_1, ColLBound To max_idx_2)

    Dim idx_1 As Long, idx_2 As Long
    idx_1 = RowLBound
    idx_2 = ColLBound

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(OriginalArray)
    Do While enum_obj.MoveNext()
        If IsObject(enum_obj.Current) Then
            Set result_arr(idx_1, idx_2) = enum_obj.Current
        Else
            result_arr(idx_1, idx_2) = enum_obj.Current
        End If

        If ColumnDirection Then
            idx_1 = idx_1 + 1
            If max_idx_1 < idx_1 Then
                idx_1 = RowLBound
                idx_2 = idx_2 + 1
            End If
        Else
            idx_2 = idx_2 + 1
            If max_idx_2 < idx_2 Then
                idx_1 = idx_1 + 1
                idx_2 = ColLBound
            End If
        End If
    Loop

    ConvertArray1dTo2d = result_arr
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 配列境界・状態
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 要素を持たない Variant 配列を生成します。
'*
'* @return 0 ベースで要素数 0 の Variant 配列。
'*
'* @details
'* 未初期化配列ではなく、`LBound = 0` / `UBound = -1` の空配列を返します。
Public Function EmptyVariantArray() As Variant()
    EmptyVariantArray = Array()
End Function

'* 要素を持たない String 配列を生成します。
'*
'* @return 0 ベースで要素数 0 の String 配列。
'*
'* @details
'* 未初期化配列ではなく、`LBound = 0` / `UBound = -1` の空配列を返します。
Public Function EmptyStringArray() As String()
    EmptyStringArray = Filter(Split(vbNullString, vbNullChar), vbNullChar)
End Function

'* 配列の添字の最大および最小を得ます。
'*
'* @param LBoundArray [出力] 各次元の最小の添字
'* @param UBoundArray [出力] 各次元の最大の添字
'* @param TargetArray 調査対象の配列
'*
'* @details
'* 配列の添字の最大および最小を得ます。未初期化配列の場合はエラーを発生させます。
Public Sub GetArrayBounds(ByRef LBoundArray() As Long, ByRef UBoundArray() As Long, ByVal TargetArray As Variant)
    If Not IsArray(TargetArray) Then
        Err.Raise vbObjectError + 1, "Sub GetArrayBounds", "配列ではありません。(" & TypeName(TargetArray) & ")"
    End If

    Dim result_l() As Long
    Dim result_u() As Long

    Dim dim_count As Long
    dim_count = pGetArrayBoundsCore(result_l, result_u, TargetArray)
    If dim_count = 0 Then
        Err.Raise vbObjectError + 1, "Sub GetArrayBounds", "未初期化配列です。"
    End If

    LBoundArray = result_l
    UBoundArray = result_u
End Sub

'* 引数が空の配列かを判定します。
'*
'* @param TargetArray 判定対象の配列
'* @return 空の配列の場合は True、それ以外は False
'*
'* @details
'* 指定された引数が配列であり、要素が存在しない場合に True を返します。
'* 配列でない場合は False を返します。
'* 配列の境界値取得 (`UBound` / `LBound`) が失敗する場合は空の配列とみなします。
Public Function IsEmptyArray(ByVal TargetArray As Variant) As Boolean
    If Not IsArray(TargetArray) Then
        IsEmptyArray = False
        Exit Function
    End If

    Dim result_value As Boolean
    result_value = False

    Err.Clear
    On Error Resume Next
    result_value = UBound(TargetArray) < LBound(TargetArray)
    If Err.Number <> 0 Then
        result_value = True
        Err.Clear
    End If
    On Error GoTo 0

    IsEmptyArray = result_value
End Function

'* 配列の列挙子を取得します。
'*
'* @param TargetArray 列挙対象の配列
'* @param Descending [省略可] 降順で列挙するか否か。
'* @param IsReadOnly [省略可] 読み取り専用にするか否か。
'* @return 配列の列挙子 (IEnumerator オブジェクト)
'*
'* @details
'* 指定された 1 次元配列を列挙するための `IEnumerator` オブジェクトを返します。
'* 空の 1 次元配列は有効入力として扱い、MoveNext は False を返します。
'* 未初期化配列および多次元配列はエラーを発生させます。
'* 配列列挙子では Update / Remove は使用できません。
'* 配列がオブジェクトの場合はそのまま参照を設定し、それ以外の場合は値を設定します。
'*
'* 使用例:
'* @code
'* Dim enum_obj As IEnumerator
'* Set enum_obj = GetArrayEnumerator(some_arr)
'* Do While enum_obj.MoveNext()
'*     Dim item_var as Variable
'*     item_var = enum_obj.Current
'*     'Debug.Print item_var
'* Loop
'* @endcode
Public Function GetArrayEnumerator(ByVal TargetArray As Variant, Optional ByVal Descending As Boolean = False, Optional ByVal IsReadOnly As Boolean = False) As IEnumerator
    Dim result_value As Enumerator
    Set result_value = New Enumerator
    Call result_value.Initialize(TargetArray, Descending:=Descending, IsReadOnly:=IsReadOnly)
    Set GetArrayEnumerator = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 配列検索
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 配列に指定された値が含まれるかを判定します。
'*
'* @param CheckItem 判定対象の値
'* @param TargetArray 判定対象の配列
'* @return 配列に値が含まれる場合は True、それ以外は False
'*
'* @details
'* 配列に指定された値が含まれるかを判定します。
Public Function ContainsInArray(ByVal CheckItem As Variant, ByRef TargetArray As Variant) As Boolean
    Dim arr_item As Variant

    For Each arr_item In TargetArray
        If pArrayItemsEqual(arr_item, CheckItem) Then
            ContainsInArray = True
            Exit Function
        End If
    Next arr_item

    ContainsInArray = False
End Function

Private Function pArrayItemsEqual(ByVal Item1 As Variant, ByVal Item2 As Variant) As Boolean
    If IsObject(Item1) Or IsObject(Item2) Then
        If IsObject(Item1) And IsObject(Item2) Then
            Dim obj_1 As Object
            Dim obj_2 As Object
            Set obj_1 = Item1
            Set obj_2 = Item2
            pArrayItemsEqual = (obj_1 Is obj_2)
        Else
            pArrayItemsEqual = False
        End If
        Exit Function
    End If

    If IsArray(Item1) Or IsArray(Item2) Then
        pArrayItemsEqual = False
        Exit Function
    End If

    If IsNull(Item1) Or IsNull(Item2) Then
        pArrayItemsEqual = (IsNull(Item1) And IsNull(Item2))
        Exit Function
    End If

    If IsError(Item1) Or IsError(Item2) Then
        If IsError(Item1) And IsError(Item2) Then
            pArrayItemsEqual = (CLng(Item1) = CLng(Item2))
        Else
            pArrayItemsEqual = False
        End If
        Exit Function
    End If

    If IsEmpty(Item1) Or IsEmpty(Item2) Then
        pArrayItemsEqual = (IsEmpty(Item1) And IsEmpty(Item2))
        Exit Function
    End If

    pArrayItemsEqual = (Item1 = Item2)
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 配列ソート
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 配列をソートします。
'*
'* @param TargetArray ソート対象の配列
'* @param Descending 降順ソートを行う場合は True。既定値は False (昇順ソート)
'*
'* @details
'* 配列をソートします。
'*
'* @note
'* 配列の各要素が不等号演算で比較可能であることが前提です。
Public Sub SortArray(ByRef TargetArray As Variant, Optional ByVal Descending As Boolean = False)
    If Not IsArray(TargetArray) Then
        Err.Raise vbObjectError + 1, Source:="Sub SortArray", Description:="Argument is not an array. (" & TypeName(TargetArray) & ")"
    End If

    If IsEmptyArray(TargetArray) Then Exit Sub
    Call pValidateSortArrayItems(TargetArray)
    Call pSortArrayCore(TargetArray, LBound(TargetArray), UBound(TargetArray), Descending)
End Sub

Private Sub pValidateSortArrayItems(ByRef TargetArray As Variant)
    Dim arr_item As Variant
    For Each arr_item In TargetArray
        If IsObject(arr_item) Then
            Err.Raise vbObjectError + 1, Source:="Sub SortArray", Description:="Object items are not supported."
        End If
    Next arr_item
End Sub

Private Sub pSortArrayCore(ByRef TargetArray As Variant, ByVal MinIndex As Long, ByVal MaxIndex As Long, ByVal Descending As Boolean)
    If MinIndex >= MaxIndex Then Exit Sub
    Dim item_idx As Long
    Dim pos_lng As Long: pos_lng = MinIndex

    Call pSortSwap(TargetArray(MinIndex), TargetArray(Int((MinIndex + MaxIndex) \ 2)))
    For item_idx = MinIndex + 1 To MaxIndex
        If pSortIsLessThan(TargetArray(item_idx), TargetArray(MinIndex), Descending) Then
            pos_lng = pos_lng + 1
            Call pSortSwap(TargetArray(item_idx), TargetArray(pos_lng))
        End If
    Next
    Call pSortSwap(TargetArray(MinIndex), TargetArray(pos_lng))

    Call pSortArrayCore(TargetArray, MinIndex, pos_lng - 1, Descending)
    Call pSortArrayCore(TargetArray, pos_lng + 1, MaxIndex, Descending)
End Sub

Private Sub pSortSwap(ByRef Item1 As Variant, ByRef Item2 As Variant)
    Dim tmp_var As Variant
    Call pAssignVariant(tmp_var, Item1)
    Call pAssignVariant(Item1, Item2)
    Call pAssignVariant(Item2, tmp_var)
End Sub

Private Function pSortIsLessThan(ByVal Item1 As Variant, ByVal Item2 As Variant, ByVal Descending As Boolean) As Boolean
    If Descending Then
        pSortIsLessThan = (Item2 < Item1)
    Else
        pSortIsLessThan = (Item1 < Item2)
    End If
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 配列結合
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 配列を連結します。
'*
'* @param Array1 最初の配列
'* @param Array2 2 番目の配列
'* @param OtherArrays 可変長の追加配列
'* @return 連結された配列
'*
'* @details
'* 複数の配列を結合して 1 つの配列を返します。
'* 引数が配列でない場合は長さ 1 の配列として扱われます。
Public Function ConcatArray(ByVal Array1 As Variant, ByVal Array2 As Variant, ParamArray OtherArrays() As Variant) As Variant()
    Dim result_value() As Variant
    Dim arr_length As Long

    arr_length = pGetArrayLengthCore(Array1)
    arr_length = arr_length + pGetArrayLengthCore(Array2)

    Dim param_idx As Long
    Dim param_item As Variant
    If UBound(OtherArrays) <> -1 Then
        For param_idx = LBound(OtherArrays) To UBound(OtherArrays)
            Call pAssignVariant(param_item, OtherArrays(param_idx))
            arr_length = arr_length + pGetArrayLengthCore(param_item)
        Next param_idx
    End If

    If arr_length = 0 Then
        ConcatArray = Array()
        Exit Function
    End If

    'If IsArray(Array1) And Not IsEmptyArray(Array1) Then
    '    ReDim result_value(LBound(Array1) To LBound(Array1) + arr_length - 1)
    'Else
        ReDim result_value(0 To arr_length - 1)
    'End If

    Dim result_idx As Long
    result_idx = LBound(result_value)

    Call pConcatArrayCore(result_value, result_idx, Array1)
    Call pConcatArrayCore(result_value, result_idx, Array2)

    If UBound(OtherArrays) <> -1 Then
        For param_idx = LBound(OtherArrays) To UBound(OtherArrays)
            Call pAssignVariant(param_item, OtherArrays(param_idx))
            Call pConcatArrayCore(result_value, result_idx, param_item)
        Next param_idx
    End If

    ConcatArray = result_value
End Function

Private Function pGetArrayLengthCore(ByVal TargetArray As Variant) As Long
    Dim result_value As Long

    If IsEmptyArray(TargetArray) Then
        result_value = 0
    ElseIf IsArray(TargetArray) Then
        result_value = UBound(TargetArray) - LBound(TargetArray) + 1
    Else
        result_value = 1
    End If

    pGetArrayLengthCore = result_value
End Function

Private Sub pConcatArrayCore(ByRef ResultArray() As Variant, ByRef CurrentIndex As Long, ByVal TargetArray As Variant)
    If IsEmptyArray(TargetArray) Then
        Exit Sub
    End If

    Dim item_idx As Long
    If IsArray(TargetArray) Then
        For item_idx = LBound(TargetArray) To UBound(TargetArray)
            Call pAssignVariant(ResultArray(CurrentIndex), TargetArray(item_idx))
            CurrentIndex = CurrentIndex + 1
        Next item_idx
    Else
        Call pAssignVariant(ResultArray(CurrentIndex), TargetArray)
        CurrentIndex = CurrentIndex + 1
    End If
End Sub

Private Sub pAssignVariant(ByRef TargetValue As Variant, ByVal SourceValue As Variant)
    If IsObject(SourceValue) Then
        Set TargetValue = SourceValue
    Else
        TargetValue = SourceValue
    End If
End Sub


' #############################################################################
'
' String 関連
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' ID 表記関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* ID 番号と名前を連結した表示名を返します。
'*
'* @param IDNumber ID 番号。
'* @param Name 名前。
'* @param NumFormat ID 番号に適用する書式。省略時は 2 桁ゼロ埋めです。
'* @param Separator ID 番号と名前の区切り文字。
'* @return 整形した ID 表示名。
'*
'* @details
'* Format 関数で IDNumber を文字列化し、Separator と Name を連結します。
Public Function FormatIDName( _
        ByVal IDNumber As Integer, _
        ByVal Name As String, _
        Optional ByVal NumFormat As String = "00", _
        Optional ByVal Separator As String = ".") As Variant

    FormatIDName = Format(IDNumber, NumFormat) & Separator & Name
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 型文字列関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 要素型契約名として使用できる文字列かどうかを返します。
'*
'* @param ElementTypeKey 確認対象の要素型契約名。
'* @return 要素型契約名として使用できる場合は True。
'*
'* @details
'* クラス モジュール名として扱えるように、1 文字目は英字、2 文字目以降は英字、数字、アンダースコア、全体は 31 文字以内に制限します。
Public Function IsValidElementTypeKey(ByVal ElementTypeKey As String) As Boolean
    Dim result_value As Boolean
    result_value = False

    Dim key_length As Long
    key_length = Len(ElementTypeKey)
    If key_length < 1 Or 31 < key_length Then
        IsValidElementTypeKey = result_value
        Exit Function
    End If

    If Not pIsElementTypeKeyFirstChar(Mid$(ElementTypeKey, 1, 1)) Then
        IsValidElementTypeKey = result_value
        Exit Function
    End If

    Dim char_idx As Long
    For char_idx = 2 To key_length
        If Not pIsElementTypeKeyFollowingChar(Mid$(ElementTypeKey, char_idx, 1)) Then
            IsValidElementTypeKey = result_value
            Exit Function
        End If
    Next char_idx

    result_value = True
    IsValidElementTypeKey = result_value
End Function

'* 値の型文字列を返します。
'*
'* @param Value 型文字列を取得する値。
'* @param ObjectKeyMode オブジェクトのキー化方法。
'* @param IncludeArrayBounds 配列の境界情報を含めるか否か。True の場合は IncludeArrayRank より優先します。
'* @param IncludeArrayRank 配列の次元数を含めるか否か。
'* @return 値の型文字列。
'*
'* @details
'* スカラー値は VBA の TypeName と同じ型名を返します。
'* オブジェクトは ObjectKeyMode に応じて Object@ClassName、IEquatable@ClassName、IDuplicateCheckable@ClassName の形式で返します。
'* 配列は末尾に []、[,]、[1:2,1:3] のいずれかを付けます。
Public Function GetTypeString( _
        ByVal Value As Variant, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE, _
        Optional ByVal IncludeArrayBounds As Boolean = False, _
        Optional ByVal IncludeArrayRank As Boolean = False, _
        Optional ByVal UseElementTypeKey As Boolean = False) As String

    Dim result_value As String
    result_value = pGetTypeStringCore(Value, ObjectKeyMode, IncludeArrayBounds, IncludeArrayRank, False, UseElementTypeKey, "Function GetTypeString")

    GetTypeString = result_value
End Function

Private Function pGetTypeStringCore( _
        ByVal Value As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal IncludeArrayBounds As Boolean, _
        ByVal IncludeArrayRank As Boolean, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If IsArray(Value) Then
        result_value = pGetArrayTypeName(Value, ObjectKeyMode, IncludeArrayBounds, IncludeArrayRank, UsePrimitiveTypeName, ErrorSource)
    ElseIf IsObject(Value) Then
        If Value Is Nothing Then
            result_value = "Nothing"
        Else
            result_value = pGetObjectKeyModeName(ObjectKeyMode, ErrorSource) & "@" & pGetElementTypeName(Value, UseElementTypeKey, ErrorSource)
        End If
    Else
        result_value = TypeName(Value)
        If UsePrimitiveTypeName And pIsPrimitiveValueKeyTypeName(result_value) Then result_value = "Primitive"
    End If

    pGetTypeStringCore = result_value
End Function

Private Function pGetArrayTypeName( _
        ByVal TargetArray As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal IncludeArrayBounds As Boolean, _
        ByVal IncludeArrayRank As Boolean, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal ErrorSource As String) As String

    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Dim dim_count As Long
    dim_count = pGetArrayBoundsCore(lbound_arr, ubound_arr, TargetArray)

    Dim item_type_name As String
    item_type_name = pGetArrayItemTypeName(TypeName(TargetArray))

    Dim result_value As String
    If UsePrimitiveTypeName And pIsPrimitiveValueKeyTypeName(item_type_name) Then
        result_value = "Primitive"
    ElseIf pIsObjectArrayType(item_type_name) Then
        Dim object_array_type_name As String
        object_array_type_name = item_type_name
        If object_array_type_name = "Object" Then object_array_type_name = pGetObjectKeyModeName(ObjectKeyMode, ErrorSource)
        result_value = pGetObjectKeyModeName(ObjectKeyMode, ErrorSource) & "@" & object_array_type_name
    Else
        result_value = item_type_name
    End If

    If IncludeArrayBounds And 0 < dim_count Then
        result_value = result_value & "[" & pGetArrayBoundsString(lbound_arr, ubound_arr, dim_count) & "]"
    ElseIf IncludeArrayRank And 0 < dim_count Then
        result_value = result_value & "[" & pGetArrayRankString(dim_count) & "]"
    Else
        result_value = result_value & "[]"
    End If

    pGetArrayTypeName = result_value
End Function

Private Function pGetArrayRankString(ByVal DimCount As Long) As String
    Dim result_value As String

    Dim dim_idx As Long
    For dim_idx = 2 To DimCount
        result_value = result_value & ","
    Next dim_idx

    pGetArrayRankString = result_value
End Function

Private Function pGetArrayBoundsCore( _
        ByRef LBoundArray() As Long, _
        ByRef UBoundArray() As Long, _
        ByVal TargetArray As Variant) As Long

    Dim dim_count As Long
    dim_count = 0

    Do
        Dim lower_bound As Long
        Dim upper_bound As Long

        Err.Clear
        On Error Resume Next
            lower_bound = LBound(TargetArray, dim_count + 1)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit Do
            End If

            upper_bound = UBound(TargetArray, dim_count + 1)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit Do
            End If
        On Error GoTo 0

        ReDim Preserve LBoundArray(0 To dim_count)
        ReDim Preserve UBoundArray(0 To dim_count)
        LBoundArray(dim_count) = lower_bound
        UBoundArray(dim_count) = upper_bound

        dim_count = dim_count + 1
    Loop While True

    pGetArrayBoundsCore = dim_count
End Function

Private Function pGetArrayItemTypeName(ByVal ArrayTypeName As String) As String
    Dim result_value As String
    If Right$(ArrayTypeName, 2) = "()" Then
        result_value = Left$(ArrayTypeName, Len(ArrayTypeName) - 2)
    Else
        result_value = ArrayTypeName
    End If

    pGetArrayItemTypeName = result_value
End Function

Private Function pIsObjectArrayType(ByVal ItemTypeName As String) As Boolean
    Dim result_value As Boolean
    Select Case ItemTypeName
        Case "Boolean", "Byte", "Currency", "Date", "Decimal", "Double", "Integer", "Long", "LongLong", "LongPtr", "Short", "Single", "String", "Variant"
            result_value = False
        Case Else
            result_value = True
    End Select

    pIsObjectArrayType = result_value
End Function

Private Function pGetArrayBoundsString(ByRef LBoundArray() As Long, ByRef UBoundArray() As Long, ByVal DimCount As Long) As String
    Dim result_value As String

    Dim dim_idx As Long
    For dim_idx = 0 To DimCount - 1
        If result_value <> "" Then result_value = result_value & ","
        result_value = result_value & CStr(LBoundArray(dim_idx)) & ":" & CStr(UBoundArray(dim_idx))
    Next dim_idx

    pGetArrayBoundsString = result_value
End Function

Private Function pGetElementTypeName( _
        ByVal Value As Object, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If UseElementTypeKey And TypeOf Value Is IElementTypeProvider Then
        Dim element_type_provider As IElementTypeProvider
        Set element_type_provider = Value

        result_value = element_type_provider.ElementTypeKey
        If Not IsValidElementTypeKey(result_value) Then
            Err.Raise vbObjectError + 1, ErrorSource, "ElementTypeKey はクラス モジュール名に使用可能な文字で指定してください。(" & result_value & ")"
        End If
    Else
        result_value = TypeName(Value)
    End If

    pGetElementTypeName = result_value
End Function

Private Function pIsElementTypeKeyFirstChar(ByVal TargetChar As String) As Boolean
    Dim char_code As Long
    char_code = AscW(TargetChar)

    Dim result_value As Boolean
    result_value = ((AscW("A") <= char_code And char_code <= AscW("Z")) _
            Or (AscW("a") <= char_code And char_code <= AscW("z")))

    pIsElementTypeKeyFirstChar = result_value
End Function

Private Function pIsElementTypeKeyFollowingChar(ByVal TargetChar As String) As Boolean
    Dim char_code As Long
    char_code = AscW(TargetChar)

    Dim result_value As Boolean
    result_value = (pIsElementTypeKeyFirstChar(TargetChar) _
            Or (AscW("0") <= char_code And char_code <= AscW("9")) _
            Or TargetChar = "_")

    pIsElementTypeKeyFollowingChar = result_value
End Function

Private Function pGetObjectKeyModeName(ByVal ObjectKeyMode As Long, ByVal ErrorSource As String) As String
    Dim result_value As String
    Select Case ObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE
            result_value = "Object"
        Case G_OBJECT_KEY_MODE_I_EQUATABLE
            result_value = "IEquatable"
        Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            result_value = "IDuplicateCheckable"
        Case Else
            Err.Raise vbObjectError + 1, ErrorSource, "未対応のオブジェクト キー化モードです。(" & CStr(ObjectKeyMode) & ")"
    End Select

    pGetObjectKeyModeName = result_value
End Function

Private Function pIsPrimitiveValueKeyTypeName(ByVal TypeNameString As String) As Boolean
    Dim result_value As Boolean
    Select Case TypeNameString
        Case "Boolean", "Byte", "Date", "Decimal", "Double", "Integer", "Long", "LongLong", "LongPtr", "Short", "Single", "String"
            result_value = True
        Case Else
            result_value = False
    End Select

    pIsPrimitiveValueKeyTypeName = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 値キー関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 値を型情報付きのキー文字列へ変換します。
'*
'* @param Value キー化する値。
'* @param ObjectKeyMode オブジェクトのキー化方法。
'* @return 型情報付きのキー文字列。
'*
'* @details
'* Dictionary などのキーとして使用するための内部表現を返します。
'* スカラー値、特殊値、配列の次元/境界/要素、オブジェクトの識別方法を含めて文字列化します。
'* スカラー値も型名で区別します。
Public Function GetTypedValueKey( _
        ByVal Value As Variant, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE, _
        Optional ByVal UseElementTypeKey As Boolean = False) As String

    Dim result_value As String
    result_value = pGetValueKeyCore(Value, ObjectKeyMode, False, UseElementTypeKey, "Function GetTypedValueKey")

    GetTypedValueKey = result_value
End Function

'* 値をキー文字列へ変換します。
'*
'* @param Value キー化する値。
'* @param ObjectKeyMode オブジェクトのキー化方法。
'* @return キー文字列。
'*
'* @details
'* `GetTypedValueKey` と同じ規則でキー化しますが、Error / Currency / Variant / Null / Empty / オブジェクト型を除くプリミティブ値は Primitive として同一視します。
Public Function GetValueKey( _
        ByVal Value As Variant, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE, _
        Optional ByVal UseElementTypeKey As Boolean = False) As String

    Dim result_value As String
    result_value = pGetValueKeyCore(Value, ObjectKeyMode, True, UseElementTypeKey, "Function GetValueKey")

    GetValueKey = result_value
End Function

Private Function pGetValueKeyCore( _
        ByVal Value As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If IsArray(Value) Then
        result_value = pGetArrayValueKeyString(Value, ObjectKeyMode, UsePrimitiveTypeName, ErrorSource)
    ElseIf IsObject(Value) Then
        result_value = pGetObjectValueKeyString(Value, ObjectKeyMode, UseElementTypeKey, ErrorSource)
    Else
        result_value = pGetTypeStringCore(Value, ObjectKeyMode, False, False, UsePrimitiveTypeName, False, ErrorSource) _
                & "(" & pGetScalarValueKeyString(Value) & ")"
    End If

    pGetValueKeyCore = result_value
End Function

Private Function pGetArrayValueKeyString( _
        ByVal TargetArray As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal ErrorSource As String) As String

    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Dim dim_count As Long
    dim_count = pGetArrayBoundsCore(lbound_arr, ubound_arr, TargetArray)

    Dim array_prefix As String
    array_prefix = pGetTypeStringCore(TargetArray, ObjectKeyMode, True, False, UsePrimitiveTypeName, False, ErrorSource)

    Dim result_value As String
    If dim_count = 0 Then
        result_value = array_prefix & "()"
    Else
        Dim total_count As Long
        total_count = pGetArrayItemCount(lbound_arr, ubound_arr, dim_count)

        If total_count = 0 Then
            result_value = array_prefix & "()"
        Else
            Dim item_type_name As String
            item_type_name = pGetArrayItemTypeName(TypeName(TargetArray))

            Dim flat_items() As Variant
            Call pFlattenArray(flat_items, TargetArray, total_count)

            Dim index_arr() As Long
            ReDim index_arr(0 To dim_count - 1)

            result_value = array_prefix & "(" _
                    & pGetArrayItemsString(flat_items, lbound_arr, ubound_arr, index_arr, 0, item_type_name, ObjectKeyMode, UsePrimitiveTypeName, ErrorSource) & ")"
        End If
    End If

    pGetArrayValueKeyString = result_value
End Function

Private Function pIsEmptyArrayBounds(ByRef LBoundArray() As Long, ByRef UBoundArray() As Long, ByVal DimCount As Long) As Boolean
    Dim result_value As Boolean
    result_value = False

    Dim dim_idx As Long
    For dim_idx = 0 To DimCount - 1
        If UBoundArray(dim_idx) < LBoundArray(dim_idx) Then
            result_value = True
            Exit For
        End If
    Next dim_idx

    pIsEmptyArrayBounds = result_value
End Function

Private Function pGetArrayItemCount(ByRef LBoundArray() As Long, ByRef UBoundArray() As Long, ByVal DimCount As Long) As Long
    Dim result_value As Long
    result_value = 1

    Dim dim_idx As Long
    For dim_idx = 0 To DimCount - 1
        result_value = result_value * (UBoundArray(dim_idx) - LBoundArray(dim_idx) + 1)
    Next dim_idx

    pGetArrayItemCount = result_value
End Function

Private Sub pFlattenArray(ByRef FlatItems() As Variant, ByVal TargetArray As Variant, ByVal ItemCount As Long)
    ReDim FlatItems(0 To ItemCount - 1)

    Dim item_idx As Long
    item_idx = 0

    Dim item_value As Variant
    For Each item_value In TargetArray
        If IsObject(item_value) Then
            Set FlatItems(item_idx) = item_value
        Else
            FlatItems(item_idx) = item_value
        End If
        item_idx = item_idx + 1
    Next item_value
End Sub

Private Function pGetArrayItemsString( _
        ByRef FlatItems() As Variant, _
        ByRef LBoundArray() As Long, _
        ByRef UBoundArray() As Long, _
        ByRef IndexArray() As Long, _
        ByVal CurrentDim As Long, _
        ByVal ItemTypeName As String, _
        ByVal ObjectKeyMode As Long, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    Dim item_idx As Long

    If CurrentDim = UBound(IndexArray) Then
        For item_idx = LBoundArray(CurrentDim) To UBoundArray(CurrentDim)
            If result_value <> "" Then result_value = result_value & ","
            IndexArray(CurrentDim) = item_idx
            result_value = result_value & pGetArrayItemString(FlatItems(pGetForEachFlatIndex(IndexArray, LBoundArray, UBoundArray)), ItemTypeName, ObjectKeyMode, UsePrimitiveTypeName, ErrorSource)
        Next item_idx
    Else
        For item_idx = LBoundArray(CurrentDim) To UBoundArray(CurrentDim)
            If result_value <> "" Then result_value = result_value & ","
            IndexArray(CurrentDim) = item_idx
            result_value = result_value & "(" & pGetArrayItemsString(FlatItems, LBoundArray, UBoundArray, IndexArray, CurrentDim + 1, ItemTypeName, ObjectKeyMode, UsePrimitiveTypeName, ErrorSource) & ")"
        Next item_idx
    End If

    pGetArrayItemsString = result_value
End Function

Private Function pGetForEachFlatIndex( _
        ByRef IndexArray() As Long, _
        ByRef LBoundArray() As Long, _
        ByRef UBoundArray() As Long) As Long

    Dim result_value As Long
    Dim multiplier As Long
    result_value = 0
    multiplier = 1

    Dim dim_idx As Long
    For dim_idx = LBound(IndexArray) To UBound(IndexArray)
        result_value = result_value + (IndexArray(dim_idx) - LBoundArray(dim_idx)) * multiplier
        multiplier = multiplier * (UBoundArray(dim_idx) - LBoundArray(dim_idx) + 1)
    Next dim_idx

    pGetForEachFlatIndex = result_value
End Function

Private Function pGetArrayItemString( _
        ByVal ItemValue As Variant, _
        ByVal ItemTypeName As String, _
        ByVal ObjectKeyMode As Long, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If ItemTypeName = "Variant" Or pIsObjectArrayType(ItemTypeName) Then
        result_value = pGetValueKeyCore(ItemValue, ObjectKeyMode, UsePrimitiveTypeName, False, ErrorSource)
    Else
        result_value = pGetScalarValueKeyString(ItemValue)
    End If

    pGetArrayItemString = result_value
End Function

Private Function pGetObjectValueKeyString( _
        ByVal Value As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If Value Is Nothing Then
        result_value = "Nothing()"
    Else
        Dim object_type_name As String
        object_type_name = pGetTypeStringCore(Value, ObjectKeyMode, False, False, False, UseElementTypeKey, ErrorSource)

        Select Case ObjectKeyMode
            Case G_OBJECT_KEY_MODE_REFERENCE
                result_value = object_type_name & "(" & Hex$(ObjPtr(Value)) & ")"

            Case G_OBJECT_KEY_MODE_I_EQUATABLE
                If Not TypeOf Value Is IEquatable Then
                    Err.Raise vbObjectError + 1, ErrorSource, "IEquatable ではないオブジェクトを IEquatable モードでキー化しようとしました。(" & object_type_name & ")"
                End If

                Dim eq_item As IEquatable
                Set eq_item = Value
                result_value = object_type_name & "(" & pEscapeValueKeyString(eq_item.GetIdentityString()) & ")"

            Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
                If Not TypeOf Value Is IDuplicateCheckable Then
                    Err.Raise vbObjectError + 1, ErrorSource, "IDuplicateCheckable ではないオブジェクトを IDuplicateCheckable モードでキー化しようとしました。(" & object_type_name & ")"
                End If

                Dim dup_item As IDuplicateCheckable
                Set dup_item = Value
                result_value = object_type_name & "(" & pEscapeValueKeyString(dup_item.GetKey()) & ")"

            Case Else
                Err.Raise vbObjectError + 1, ErrorSource, "未対応のオブジェクト キー化モードです。(" & CStr(ObjectKeyMode) & ")"
        End Select
    End If

    pGetObjectValueKeyString = result_value
End Function

Private Function pGetScalarValueKeyString(ByVal Value As Variant) As String
    Dim result_value As String
    If IsError(Value) Then
        result_value = CStr(CLng(Value))
    ElseIf IsNull(Value) Or IsEmpty(Value) Then
        result_value = ""
    ElseIf VarType(Value) = vbDate Then
        result_value = pEscapeValueKeyString(Format$(Value, "yyyy-mm-dd\Thh:nn:ss"))
    ElseIf VarType(Value) = vbString Then
        result_value = pEscapeValueKeyString(CStr(Value))
    Else
        result_value = pEscapeValueKeyString(CStr(Value))
    End If

    pGetScalarValueKeyString = result_value
End Function

Private Function pEscapeValueKeyString(ByVal Expression As String) As String
    Dim result_value As String
    result_value = Replace(Expression, "\", "\\")
    result_value = Replace(result_value, vbTab, "\t")
    result_value = Replace(result_value, "(", "\(")
    result_value = Replace(result_value, ")", "\)")
    result_value = Replace(result_value, "[", "\[")
    result_value = Replace(result_value, "]", "\]")
    result_value = Replace(result_value, ",", "\,")
    result_value = Replace(result_value, ":", "\:")
    result_value = Replace(result_value, "@", "\@")
    result_value = Replace(result_value, "=", "\=")

    pEscapeValueKeyString = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 複合キー関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 複数キーを型情報付きの文字列として連結して返します。
'*
'* @param DictionaryKeys キー (可変長引数)。0 件を指定した場合は空文字列を返します。
'* @return 連結されたキー文字列。
'*
'* @details
'* 指定された複数のキーを `GetTypedValueKey` と同じ基本規則でキー化し、1 つの文字列として返します。
'* キーに使用される区切り文字はタブ文字 (`vbTab`) です。
'* 各キー内のタブ文字は `GetTypedValueKey` でエスケープされます。
Public Function GetTypedMultiKey(ParamArray DictionaryKeys() As Variant) As String
    Dim key_arr As Variant
    key_arr = DictionaryKeys

    GetTypedMultiKey = GetTypedMultiKeyByModeFromArray(G_OBJECT_KEY_MODE_REFERENCE, key_arr)
End Function

'* オブジェクト キー化モードを指定して、複数キーを型情報付きの文字列として連結して返します。
'*
'* @param ObjectKeyMode トップレベルのオブジェクト キー化方法。
'* @param DictionaryKeys キー (可変長引数)。0 件を指定した場合は空文字列を返します。
'* @return 連結されたキー文字列。
'*
'* @details
'* トップレベルのオブジェクト値は ObjectKeyMode を優先してキー化し、指定モードに対応していないオブジェクトは参照キーへフォールバックします。
'* 配列引数は `G_OBJECT_KEY_MODE_REFERENCE` でキー化するため、配列要素内のオブジェクトには ObjectKeyMode を適用しません。
Public Function GetTypedMultiKeyByMode(ByVal ObjectKeyMode As Long, ParamArray DictionaryKeys() As Variant) As String
    Dim key_arr As Variant
    key_arr = DictionaryKeys

    GetTypedMultiKeyByMode = GetTypedMultiKeyByModeFromArray(ObjectKeyMode, key_arr)
End Function

'* オブジェクト キー化モードを指定して、配列から複数キーを型情報付きの文字列として連結して返します。
'*
'* @param ObjectKeyMode トップレベルのオブジェクト キー化方法。
'* @param DictionaryKeys ParamArray と同じ形式の 1 次元配列。
'* @return 連結されたキー文字列。
'*
'* @details
'* VBA では配列を ParamArray へ展開できないため、その回避用に用意した関数です。
'* 通常の呼び出しでは `GetTypedMultiKeyByMode` を使用してください。
'* トップレベルのオブジェクト値は ObjectKeyMode を優先してキー化し、指定モードに対応していないオブジェクトは参照キーへフォールバックします。
'* 配列引数は `G_OBJECT_KEY_MODE_REFERENCE` でキー化するため、配列要素内のオブジェクトには ObjectKeyMode を適用しません。
Public Function GetTypedMultiKeyByModeFromArray( _
        ByVal ObjectKeyMode As Long, _
        ByVal DictionaryKeys As Variant, _
        Optional ByVal UseElementTypeKey As Boolean = False) As String

    GetTypedMultiKeyByModeFromArray = pGetMultiKeyByModeFromArrayCore( _
            ObjectKeyMode, _
            DictionaryKeys, _
            False, _
            UseElementTypeKey, _
            "Function GetTypedMultiKeyByModeFromArray")
End Function

'* 複数キーを連結して返します。
'*
'* @param DictionaryKeys キー (可変長引数)。0 件を指定した場合は空文字列を返します。
'* @return 連結されたキー文字列。
'*
'* @details
'* 指定された複数のキーを `GetValueKey` と同じ基本規則でキー化し、1 つの文字列として返します。
'* Error / Currency / Variant / Null / Empty / オブジェクト型を除くプリミティブ値は Primitive として同一視します。
Public Function GetMultiKey(ParamArray DictionaryKeys() As Variant) As String
    Dim key_arr As Variant
    key_arr = DictionaryKeys

    GetMultiKey = GetMultiKeyByModeFromArray(G_OBJECT_KEY_MODE_REFERENCE, key_arr)
End Function

'* オブジェクト キー化モードを指定して、複数キーを連結して返します。
'*
'* @param ObjectKeyMode トップレベルのオブジェクト キー化方法。
'* @param DictionaryKeys キー (可変長引数)。0 件を指定した場合は空文字列を返します。
'* @return 連結されたキー文字列。
'*
'* @details
'* トップレベルのオブジェクト値は ObjectKeyMode を優先してキー化し、指定モードに対応していないオブジェクトは参照キーへフォールバックします。
'* 配列引数は `G_OBJECT_KEY_MODE_REFERENCE` でキー化するため、配列要素内のオブジェクトには ObjectKeyMode を適用しません。
'* Error / Currency / Variant / Null / Empty / オブジェクト型を除くプリミティブ値は Primitive として同一視します。
Public Function GetMultiKeyByMode(ByVal ObjectKeyMode As Long, ParamArray DictionaryKeys() As Variant) As String
    Dim key_arr As Variant
    key_arr = DictionaryKeys

    GetMultiKeyByMode = GetMultiKeyByModeFromArray(ObjectKeyMode, key_arr)
End Function

'* オブジェクト キー化モードを指定して、配列から複数キーを連結して返します。
'*
'* @param ObjectKeyMode トップレベルのオブジェクト キー化方法。
'* @param DictionaryKeys ParamArray と同じ形式の 1 次元配列。
'* @return 連結されたキー文字列。
'*
'* @details
'* VBA では配列を ParamArray へ展開できないため、その回避用に用意した関数です。
'* 通常の呼び出しでは `GetMultiKeyByMode` を使用してください。
'* トップレベルのオブジェクト値は ObjectKeyMode を優先してキー化し、指定モードに対応していないオブジェクトは参照キーへフォールバックします。
'* 配列引数は `G_OBJECT_KEY_MODE_REFERENCE` でキー化するため、配列要素内のオブジェクトには ObjectKeyMode を適用しません。
'* Error / Currency / Variant / Null / Empty / オブジェクト型を除くプリミティブ値は Primitive として同一視します。
Public Function GetMultiKeyByModeFromArray( _
        ByVal ObjectKeyMode As Long, _
        ByVal DictionaryKeys As Variant, _
        Optional ByVal UseElementTypeKey As Boolean = False) As String

    GetMultiKeyByModeFromArray = pGetMultiKeyByModeFromArrayCore( _
            ObjectKeyMode, _
            DictionaryKeys, _
            True, _
            UseElementTypeKey, _
            "Function GetMultiKeyByModeFromArray")
End Function

Private Function pGetMultiKeyByModeFromArrayCore( _
        ByVal ObjectKeyMode As Long, _
        ByVal DictionaryKeys As Variant, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Call pGetObjectKeyModeName(ObjectKeyMode, ErrorSource)

    Dim lower_bound As Long
    Dim upper_bound As Long
    Call pGetArgumentArrayBounds(lower_bound, upper_bound, DictionaryKeys, ErrorSource)

    If upper_bound < lower_bound Then
        pGetMultiKeyByModeFromArrayCore = ""
        Exit Function
    End If

    Dim result_value As String
    Dim item_idx As Long
    For item_idx = lower_bound To upper_bound
        If result_value <> "" Then result_value = result_value & vbTab
        result_value = result_value & pGetFallbackValueKeyCore(DictionaryKeys(item_idx), ObjectKeyMode, UsePrimitiveTypeName, UseElementTypeKey, ErrorSource)
    Next item_idx

    pGetMultiKeyByModeFromArrayCore = result_value
End Function

Private Sub pGetArgumentArrayBounds( _
        ByRef LowerBoundValue As Long, _
        ByRef UpperBoundValue As Long, _
        ByVal TargetArray As Variant, _
        ByVal ErrorSource As String)

    If Not IsArray(TargetArray) Then
        Err.Raise vbObjectError + 1, ErrorSource, "ParamArray と同じ形式の 1 次元配列を指定してください。"
    End If

    On Error GoTo INVALID_ARRAY
    LowerBoundValue = LBound(TargetArray, 1)
    UpperBoundValue = UBound(TargetArray, 1)
    On Error GoTo 0

    Dim second_dim_lower_bound As Long
    On Error Resume Next
    second_dim_lower_bound = LBound(TargetArray, 2)
    If Err.Number = 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError + 1, ErrorSource, "ParamArray と同じ形式の 1 次元配列を指定してください。"
    End If
    Err.Clear
    On Error GoTo 0
    Exit Sub

INVALID_ARRAY:
    On Error GoTo 0
    Err.Raise vbObjectError + 1, ErrorSource, "初期化済みの 1 次元配列を指定してください。"
End Sub

Private Function pGetFallbackValueKeyCore( _
        ByVal Value As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If IsArray(Value) Then
        result_value = pGetValueKeyCore(Value, G_OBJECT_KEY_MODE_REFERENCE, UsePrimitiveTypeName, False, ErrorSource)
    ElseIf IsObject(Value) Then
        result_value = pGetFallbackObjectValueKeyString(Value, ObjectKeyMode, UseElementTypeKey, ErrorSource)
    Else
        result_value = pGetValueKeyCore(Value, G_OBJECT_KEY_MODE_REFERENCE, UsePrimitiveTypeName, False, ErrorSource)
    End If

    pGetFallbackValueKeyCore = result_value
End Function

Private Function pGetFallbackObjectValueKeyString( _
        ByVal Value As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim effective_object_key_mode As Long
    Select Case ObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE
            effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE

        Case G_OBJECT_KEY_MODE_I_EQUATABLE
            If Value Is Nothing Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            ElseIf TypeOf Value Is IEquatable Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_I_EQUATABLE
            Else
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            End If

        Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            If Value Is Nothing Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            ElseIf TypeOf Value Is IDuplicateCheckable Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            Else
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            End If

        Case Else
            Err.Raise vbObjectError + 1, ErrorSource, "未対応のオブジェクト キー化モードです。(" & CStr(ObjectKeyMode) & ")"
    End Select

    pGetFallbackObjectValueKeyString = pGetObjectValueKeyString(Value, effective_object_key_mode, UseElementTypeKey, ErrorSource)
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 文字列配列関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 文字列の配列の diff を取ります。
'*
'* @param OldArray [入出力] 古い文字列の配列。ReDim 可能である必要があります。
'* @param NewArray [入出力] 新しい文字列の配列。ReDim 可能である必要があります。
'* @param ChangeTypeArray [出力] 変更種別の配列。ReDim 可能である必要があります。渡される配列は、未初期化状態でも構いません。結果が 1 件以上の場合、下限は常に 0 です。結果が 0 件の場合は空配列になります。
'* @param EnableReplaceType [省略可] 削除と追加が連続する部分を置換として扱うか否か。デフォルトは True で、置換として扱います。
'* @param ReplaceCost [省略可] 削除と追加のコストを 1 としたとき、置換のコストをどうするか (0～2)。デフォルトは 1.5。
'*
'* @details
'* 配列 OldArray, NewArray の最小編集距離を算出し、
'* (Old行, New行, 変更種別) の並びをそれぞれの配列に上書きします。
'* 結果配列は入力配列の下限に関わらず、0 ベースに正規化されます。
'* OldArray / NewArray は未初期化状態または空配列でも入力できます。両方が空の場合、出力配列も空配列になります。
'*
'* 変更種別は "" (一致), "DEL" (削除), "ADD" (追加), "MOD" (置換) のいずれかとなります。
'* 種別が一致の場合は OldArray、NewArray の両方の当該インデックスは元の文字列になります。
'* 削除の場合は、OldArray の当該インデックスは元の文字列が入り、NewArray の当該インデックスは空文字列となります。
'* 追加の場合は、NewArray の当該インデックスは新しい文字列が入り、OldArray の当該インデックスは空文字列となります。
'* 置換の場合は、OldArray の当該インデックスは元の文字列が、NewArray の当該インデックスは新しい文字列が入ります。
Public Sub DiffStringArray( _
        ByRef OldArray() As String, _
        ByRef NewArray() As String, _
        ByRef ChangeTypeArray() As String, _
        Optional ByVal EnableReplaceType As Boolean = True, _
        Optional ByVal ReplaceCost As Double = 1.5)

    Dim diff_old_coll As New Collection
    Dim diff_new_coll As New Collection
    Dim diff_op_coll  As New Collection

    If Not EnableReplaceType Then ReplaceCost = 3

    Dim old_start_idx As Long
    Dim old_end_idx As Long
    Dim new_start_idx As Long
    Dim new_end_idx As Long

    If IsEmptyArray(OldArray) Then
        old_start_idx = 0
        old_end_idx = -1
    Else
        old_start_idx = LBound(OldArray)
        old_end_idx = UBound(OldArray)
    End If

    If IsEmptyArray(NewArray) Then
        new_start_idx = 0
        new_end_idx = -1
    Else
        new_start_idx = LBound(NewArray)
        new_end_idx = UBound(NewArray)
    End If

    ' 再帰的に差分計算（Hirschberg アルゴリズム）
    Call pHirschbergDiffRecursive(OldArray, old_start_idx, old_end_idx, _
                              NewArray, new_start_idx, new_end_idx, _
                              diff_old_coll, diff_new_coll, diff_op_coll, ReplaceCost)

    Dim num_items_val As Long, loop_idx_val As Long
    num_items_val = diff_old_coll.Count
    If num_items_val = 0 Then
        Erase OldArray
        Erase NewArray
        Erase ChangeTypeArray
        Exit Sub
    End If

    ReDim OldArray(0 To num_items_val - 1)
    ReDim NewArray(0 To num_items_val - 1)
    ReDim ChangeTypeArray(0 To num_items_val - 1)

    For loop_idx_val = 1 To num_items_val
        OldArray(loop_idx_val - 1) = diff_old_coll(loop_idx_val)
        NewArray(loop_idx_val - 1) = diff_new_coll(loop_idx_val)
        ChangeTypeArray(loop_idx_val - 1) = diff_op_coll(loop_idx_val)
    Next loop_idx_val
End Sub

Private Sub pHirschbergDiffRecursive(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, _
        ByRef DiffOldColl As Collection, ByRef DiffNewColl As Collection, ByRef DiffOpColl As Collection, ByVal ReplaceCost As Double)

    Dim SMALL_THRESHOLD As Long: SMALL_THRESHOLD = 1

    Dim old_count_val As Long, new_count_val As Long
    old_count_val = EndIdxOld - StartIdxOld + 1
    new_count_val = EndIdxNew - StartIdxNew + 1

    If old_count_val = 0 Then
        Dim loop_idx_new As Long
        For loop_idx_new = StartIdxNew To EndIdxNew
            DiffOldColl.Add ""
            DiffNewColl.Add NewArr(loop_idx_new)
            DiffOpColl.Add "ADD"
        Next loop_idx_new
        Exit Sub
    ElseIf new_count_val = 0 Then
        Dim loop_idx_old As Long
        For loop_idx_old = StartIdxOld To EndIdxOld
            DiffOldColl.Add OldArr(loop_idx_old)
            DiffNewColl.Add ""
            DiffOpColl.Add "DEL"
        Next loop_idx_old
        Exit Sub
    ElseIf old_count_val <= SMALL_THRESHOLD Or new_count_val <= SMALL_THRESHOLD Then
        Call pDiffSmall(OldArr, StartIdxOld, EndIdxOld, NewArr, StartIdxNew, EndIdxNew, DiffOldColl, DiffNewColl, DiffOpColl, ReplaceCost)
        Exit Sub
    End If

    Dim mid_idx_val As Long
    mid_idx_val = (StartIdxOld + EndIdxOld) \ 2

    Dim forward_costs_arr As Variant
    forward_costs_arr = pComputeForwardCost(OldArr, StartIdxOld, mid_idx_val, NewArr, StartIdxNew, EndIdxNew, ReplaceCost)

    Dim backward_costs_arr As Variant
    backward_costs_arr = pComputeBackwardCost(OldArr, mid_idx_val + 1, EndIdxOld, NewArr, StartIdxNew, EndIdxNew, ReplaceCost)

    Dim split_j_val As Long, loop_idx_new_val As Long
    Dim min_cost_val As Double
    min_cost_val = 1E+99
    Dim new_size_val As Long
    new_size_val = new_count_val
    Dim cur_cost_val As Double
    For loop_idx_new_val = 0 To new_size_val
        cur_cost_val = forward_costs_arr(loop_idx_new_val) + backward_costs_arr(new_size_val - loop_idx_new_val)
        If cur_cost_val < min_cost_val Then
            min_cost_val = cur_cost_val
            split_j_val = loop_idx_new_val
        End If
    Next loop_idx_new_val

    Call pHirschbergDiffRecursive(OldArr, StartIdxOld, mid_idx_val, NewArr, StartIdxNew, StartIdxNew + split_j_val - 1, _
                              DiffOldColl, DiffNewColl, DiffOpColl, ReplaceCost)
    Call pHirschbergDiffRecursive(OldArr, mid_idx_val + 1, EndIdxOld, NewArr, StartIdxNew + split_j_val, EndIdxNew, _
                              DiffOldColl, DiffNewColl, DiffOpColl, ReplaceCost)
End Sub

Private Function pComputeForwardCost(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, ByVal ReplaceCost As Double) As Variant

    Dim COST_ADD As Double: COST_ADD = 1
    Dim COST_DEL As Double: COST_DEL = 1
    Dim COST_EQUAL As Double: COST_EQUAL = 0

    Dim new_count_val As Long, loop_new_val As Long, loop_old_val As Long, inner_idx_val As Long
    new_count_val = EndIdxNew - StartIdxNew + 1
    Dim cost_arr() As Double
    ReDim cost_arr(0 To new_count_val)

    cost_arr(0) = 0
    For loop_new_val = 1 To new_count_val
        cost_arr(loop_new_val) = cost_arr(loop_new_val - 1) + COST_ADD
    Next loop_new_val

    Dim prev_cost_val As Double, temp_cost_val As Double, sub_cost_val As Double
    For loop_old_val = StartIdxOld To EndIdxOld
        prev_cost_val = cost_arr(0)
        cost_arr(0) = cost_arr(0) + COST_DEL
        For inner_idx_val = 1 To new_count_val
            temp_cost_val = cost_arr(inner_idx_val)
            If OldArr(loop_old_val) = NewArr(StartIdxNew + inner_idx_val - 1) Then
                sub_cost_val = COST_EQUAL
            Else
                sub_cost_val = ReplaceCost
            End If
            cost_arr(inner_idx_val) = Application.Min(cost_arr(inner_idx_val) + COST_DEL, _
                                                       cost_arr(inner_idx_val - 1) + COST_ADD, _
                                                       prev_cost_val + sub_cost_val)
            prev_cost_val = temp_cost_val
        Next inner_idx_val
    Next loop_old_val

    pComputeForwardCost = cost_arr
End Function

Private Function pComputeBackwardCost(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, ByVal ReplaceCost As Double) As Variant

    Dim COST_ADD As Double: COST_ADD = 1
    Dim COST_DEL As Double: COST_DEL = 1
    Dim COST_EQUAL As Double: COST_EQUAL = 0

    Dim new_count_val As Long, loop_new_val As Long, loop_old_val As Long, inner_idx_val As Long
    new_count_val = EndIdxNew - StartIdxNew + 1
    Dim cost_arr() As Double
    ReDim cost_arr(0 To new_count_val)

    cost_arr(0) = 0
    For loop_new_val = 1 To new_count_val
        cost_arr(loop_new_val) = cost_arr(loop_new_val - 1) + COST_ADD
    Next loop_new_val

    Dim prev_cost_val As Double, temp_cost_val As Double, sub_cost_val As Double
    For loop_old_val = EndIdxOld To StartIdxOld Step -1
        prev_cost_val = cost_arr(0)
        cost_arr(0) = cost_arr(0) + COST_DEL
        For inner_idx_val = 1 To new_count_val
            temp_cost_val = cost_arr(inner_idx_val)
            If OldArr(loop_old_val) = NewArr(EndIdxNew - inner_idx_val + 1) Then
                sub_cost_val = COST_EQUAL
            Else
                sub_cost_val = ReplaceCost
            End If
            cost_arr(inner_idx_val) = Application.Min(cost_arr(inner_idx_val) + COST_DEL, _
                                                       cost_arr(inner_idx_val - 1) + COST_ADD, _
                                                       prev_cost_val + sub_cost_val)
            prev_cost_val = temp_cost_val
        Next inner_idx_val
    Next loop_old_val

    pComputeBackwardCost = cost_arr
End Function

Private Sub pDiffSmall(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, _
        ByRef DiffOldColl As Collection, ByRef DiffNewColl As Collection, ByRef DiffOpColl As Collection, ByVal ReplaceCost As Double)

    Dim COST_ADD As Double: COST_ADD = 1
    Dim COST_DEL As Double: COST_DEL = 1
    Dim COST_EQUAL As Double: COST_EQUAL = 0

    Dim old_count_val As Long, new_count_val As Long
    old_count_val = EndIdxOld - StartIdxOld + 1
    new_count_val = EndIdxNew - StartIdxNew + 1
    Dim dp_arr() As Double, op_arr() As String
    ReDim dp_arr(0 To old_count_val, 0 To new_count_val)
    ReDim op_arr(0 To old_count_val, 0 To new_count_val)

    Dim loop_old_val As Long, loop_new_val As Long
    dp_arr(0, 0) = 0
    op_arr(0, 0) = "DONE"
    For loop_old_val = 1 To old_count_val
        dp_arr(loop_old_val, 0) = dp_arr(loop_old_val - 1, 0) + COST_DEL
        op_arr(loop_old_val, 0) = "DEL"
    Next loop_old_val
    For loop_new_val = 1 To new_count_val
        dp_arr(0, loop_new_val) = dp_arr(0, loop_new_val - 1) + COST_ADD
        op_arr(0, loop_new_val) = "ADD"
    Next loop_new_val

    Dim inner_old_val As Long, inner_new_val As Long
    Dim sub_cost_val As Double, del_cost_val As Double, add_cost_val As Double, mod_cost_val As Double
    For inner_old_val = 1 To old_count_val
        For inner_new_val = 1 To new_count_val
            If OldArr(StartIdxOld + inner_old_val - 1) = NewArr(StartIdxNew + inner_new_val - 1) Then
                sub_cost_val = COST_EQUAL
            Else
                sub_cost_val = ReplaceCost
            End If
            del_cost_val = dp_arr(inner_old_val - 1, inner_new_val) + COST_DEL
            add_cost_val = dp_arr(inner_old_val, inner_new_val - 1) + COST_ADD
            mod_cost_val = dp_arr(inner_old_val - 1, inner_new_val - 1) + sub_cost_val
            dp_arr(inner_old_val, inner_new_val) = Application.Min(del_cost_val, add_cost_val, mod_cost_val)
            If dp_arr(inner_old_val, inner_new_val) = mod_cost_val Then
                op_arr(inner_old_val, inner_new_val) = IIf(sub_cost_val = COST_EQUAL, "", "MOD")
            ElseIf dp_arr(inner_old_val, inner_new_val) = del_cost_val Then
                op_arr(inner_old_val, inner_new_val) = "DEL"
            Else
                op_arr(inner_old_val, inner_new_val) = "ADD"
            End If
        Next inner_new_val
    Next inner_old_val

    Dim pos_old_val As Long, pos_new_val As Long
    pos_old_val = old_count_val: pos_new_val = new_count_val
    Dim temp_old_arr() As String, temp_new_arr() As String, temp_op_arr() As String
    ReDim temp_old_arr(0 To old_count_val + new_count_val - 1)
    ReDim temp_new_arr(0 To old_count_val + new_count_val - 1)
    ReDim temp_op_arr(0 To old_count_val + new_count_val - 1)
    Dim step_idx_val As Long: step_idx_val = 0
    Do While pos_old_val > 0 Or pos_new_val > 0
        Dim cur_op_val As String
        cur_op_val = op_arr(pos_old_val, pos_new_val)
        Select Case cur_op_val
            Case ""
                temp_old_arr(step_idx_val) = OldArr(StartIdxOld + pos_old_val - 1)
                temp_new_arr(step_idx_val) = NewArr(StartIdxNew + pos_new_val - 1)
                temp_op_arr(step_idx_val) = ""
                pos_old_val = pos_old_val - 1: pos_new_val = pos_new_val - 1
            Case "MOD"
                temp_old_arr(step_idx_val) = OldArr(StartIdxOld + pos_old_val - 1)
                temp_new_arr(step_idx_val) = NewArr(StartIdxNew + pos_new_val - 1)
                temp_op_arr(step_idx_val) = "MOD"
                pos_old_val = pos_old_val - 1: pos_new_val = pos_new_val - 1
            Case "DEL"
                temp_old_arr(step_idx_val) = OldArr(StartIdxOld + pos_old_val - 1)
                temp_new_arr(step_idx_val) = ""
                temp_op_arr(step_idx_val) = "DEL"
                pos_old_val = pos_old_val - 1
            Case "ADD"
                temp_old_arr(step_idx_val) = ""
                temp_new_arr(step_idx_val) = NewArr(StartIdxNew + pos_new_val - 1)
                temp_op_arr(step_idx_val) = "ADD"
                pos_new_val = pos_new_val - 1
        End Select
        step_idx_val = step_idx_val + 1
    Loop

    Dim num_steps_val As Long: num_steps_val = step_idx_val
    Dim rev_old_arr() As String, rev_new_arr() As String, rev_op_arr() As String
    ReDim rev_old_arr(0 To num_steps_val - 1)
    ReDim rev_new_arr(0 To num_steps_val - 1)
    ReDim rev_op_arr(0 To num_steps_val - 1)
    For loop_old_val = 0 To num_steps_val - 1
        rev_old_arr(loop_old_val) = temp_old_arr(num_steps_val - 1 - loop_old_val)
        rev_new_arr(loop_old_val) = temp_new_arr(num_steps_val - 1 - loop_old_val)
        rev_op_arr(loop_old_val) = temp_op_arr(num_steps_val - 1 - loop_old_val)
    Next loop_old_val

    Dim loop_idx_final_val As Long
    For loop_idx_final_val = 0 To num_steps_val - 1
        DiffOldColl.Add rev_old_arr(loop_idx_final_val)
        DiffNewColl.Add rev_new_arr(loop_idx_final_val)
        DiffOpColl.Add rev_op_arr(loop_idx_final_val)
    Next loop_idx_final_val
End Sub

'* String の配列が空かどうかを判定します。
'*
'* @param StringArray 判定対象の String 型配列
'* @param BlankAsEmpty [省略可] 要素数 1 で中身が空文字列の配列を空扱いするか否か。既定値は True。
'* @return 空の場合は True、それ以外は False
'*
'* @details
'* 配列が初期化されていない場合、または要素が 1 未満の場合に True を返します。
'* BlankAsEmpty が True の場合は、要素数 1 で唯一の要素が空文字列の場合も True を返します。
Public Function IsEmptyStringArray( _
        ByRef StringArray() As String, _
        Optional ByVal BlankAsEmpty As Boolean = True) As Boolean
    If (Not StringArray) = -1 Then
        ' 未初期化
        IsEmptyStringArray = True
    ElseIf UBound(StringArray) < LBound(StringArray) Then
        ' 長さが 1 未満
        IsEmptyStringArray = True
    ElseIf LBound(StringArray) < UBound(StringArray) Then
        ' 長さが 2 以上
        IsEmptyStringArray = False
    ElseIf BlankAsEmpty And StringArray(LBound(StringArray)) = "" Then
        ' 長さが 1 で、内容が空文字列
        IsEmptyStringArray = True
    Else
        ' それ以外
        IsEmptyStringArray = False
    End If
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 文字列置換関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 文字列に対して複数条件で置換を実行します。
'*
'* @param Expression 対象の文字列
'* @param Find1 最初の置換対象文字列
'* @param Replace1 最初の置換後文字列 (文字列または配列)
'* @param FindReplacePairs 可変長の置換対象文字列と置換後文字列のペア
'* @return 置換後の文字列配列
'*
'* @details
'* 指定された条件で文字列を置換し、すべての結果を配列で返します。
'* 置換後文字列が配列の場合、それぞれの組み合わせを結果に含めます。
'* 置換対象文字列に空文字列は指定できません。
'* 置換後文字列が空配列の場合はエラーを発生させます。
Public Function ReplaceMulti(ByVal Expression As String, ByVal Find1 As String, ByVal Replace1 As Variant, ParamArray FindReplacePairs() As Variant) As String()
    Dim result_list As ObjectList
    Set result_list = New ObjectList

    If LBound(FindReplacePairs) <= UBound(FindReplacePairs) Then
        If (UBound(FindReplacePairs) - LBound(FindReplacePairs) + 1) Mod 2 <> 0 Then
            Err.Raise vbObjectError + 1, "Function ReplaceMulti", "可変長文字列が奇数個です。置換対象と置換後文字列の対応が取れていません。"
            Exit Function
        End If
    End If

    Call result_list.Add(Expression)
    Call pReplaceMultiCore(result_list, Find1, Replace1)

    If LBound(FindReplacePairs) <= UBound(FindReplacePairs) Then
        Dim item_idx As Long
        For item_idx = LBound(FindReplacePairs) To UBound(FindReplacePairs) - 1 Step 2
            Call pReplaceMultiCore(result_list, FindReplacePairs(item_idx), FindReplacePairs(item_idx + 1))
        Next item_idx
    End If

    ReplaceMulti = result_list.ConvertToStringArray()
End Function

Private Sub pReplaceMultiCore(ByRef ResultList As ObjectList, ByVal FindString As String, ByVal Replaces As Variant)
    If FindString = "" Then
        Err.Raise vbObjectError + 1, "Function ReplaceMulti", "置換対象文字列に空文字列は指定できません。"
    End If

    If IsArray(Replaces) Then
        If IsEmptyArray(Replaces) Then
            Err.Raise vbObjectError + 1, "Function ReplaceMulti", "置換後文字列の候補配列が空です。"
        End If
    End If

    Dim result_value As ObjectList
    Set result_value = New ObjectList

    Dim enum_obj As IEnumerator
    Set enum_obj = ResultList.GetEnumerator()
    Do While enum_obj.MoveNext()
        Dim current_item As String
        current_item = enum_obj.Current
        If IsArray(Replaces) Then
            Dim replace_item As Variant
            For Each replace_item In Replaces
                Call result_value.Add(Replace(current_item, FindString, replace_item))
            Next replace_item
        Else
            Call result_value.Add(Replace(current_item, FindString, Replaces))
        End If
    Loop

    Set ResultList = result_value
End Sub


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 改行文字関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Sub pValidateLineSeparatorEscSeqChar(ByVal EscSeqChar As String, ByVal ErrorSource As String)
    If Len(EscSeqChar) <> 1 Then
        Err.Raise vbObjectError + 1, ErrorSource, "EscSeqChar は 1 文字を指定してください。"
    End If
End Sub

'* 改行記号をエスケープします。
'*
'* @param Expression 対象の文字列
'* @param EscSeqChar エスケープ記号 (既定値は `\`)
'* @return 改行記号をエスケープした文字列
'*
'* @details
'* 改行記号 (`vbCr`, `vbLf`) を指定されたエスケープ記号で置換します。
'* EscSeqChar が空文字列の場合は省略扱いとし、`\` を使用します。
'* EscSeqChar は空文字列を除き 1 文字だけ指定できます。
Public Function EscapeLineSeparator(ByVal Expression As String, Optional ByVal EscSeqChar As String = "") As String
    If EscSeqChar = "" Then EscSeqChar = "\"
    Call pValidateLineSeparatorEscSeqChar(EscSeqChar, "Function EscapeLineSeparator")

    Dim result_value As String
    result_value = Replace(Expression, EscSeqChar, EscSeqChar & EscSeqChar)
    result_value = Replace(result_value, vbCr, EscSeqChar & "r")
    result_value = Replace(result_value, vbLf, EscSeqChar & "n")
    EscapeLineSeparator = result_value
End Function

'* 改行記号のエスケープを解除します。
'*
'* @param Expression 対象の文字列
'* @param EscSeqChar エスケープ記号 (既定値は `\`)
'* @return エスケープ解除後の文字列
'*
'* @details
'* エスケープされた改行記号 (`\n`, `\r`) を元の改行記号 (`vbCr`, `vbLf`) に戻します。
'* EscSeqChar が空文字列の場合は省略扱いとし、`\` を使用します。
'* EscSeqChar は空文字列を除き 1 文字だけ指定できます。
'* 未知のエスケープ列はエスケープ文字を捨て、後続文字だけを残します。
'* 文字列末尾のエスケープ文字はエラーとします。
Public Function UnescapeLineSeparator(ByVal Expression As String, Optional ByVal EscSeqChar As String = "\") As String
    If EscSeqChar = "" Then EscSeqChar = "\"
    Call pValidateLineSeparatorEscSeqChar(EscSeqChar, "Function UnescapeLineSeparator")

    Dim result_value As String
    Dim item_idx As Long
    Dim str_len As Long
    Dim cur_char As String

    str_len = Len(Expression)
    item_idx = 1
    result_value = ""

    Do While item_idx <= str_len
        cur_char = Mid$(Expression, item_idx, 1)

        If cur_char = EscSeqChar Then
            ' エスケープシーケンスの開始
            If item_idx < str_len Then
                Dim next_char As String
                next_char = Mid$(Expression, item_idx + 1, 1)

                Select Case next_char
                    Case "n"
                        result_value = result_value & vbLf
                        item_idx = item_idx + 2
                    Case "r"
                        result_value = result_value & vbCr
                        item_idx = item_idx + 2
                    Case Else
                        result_value = result_value & next_char
                        item_idx = item_idx + 2
                End Select
            Else
                ' 文字列の最後にエスケープ文字がある場合
                Err.Raise vbObjectError + 512, "Function UnescapeLineSeparator", "文字列の最後にエスケープ文字があります。"
            End If
        Else
            ' エスケープ文字でない場合はそのまま追加
            result_value = result_value & cur_char
            item_idx = item_idx + 1
        End If
    Loop

    UnescapeLineSeparator = result_value
End Function

'* 改行記号で文字列を分割して配列を返します。'* 改行記号で文字列を分割して配列を返します。
'*
'* @param StringList 対象の文字列
'* @return 改行記号で分割された文字列配列
'*
'* @details
'* 改行記号に基づいて文字列を分割し、配列で返します。
Public Function SplitByLineSeparator(ByVal StringList As String) As String()
    SplitByLineSeparator = Split(pUnifyLineSeparatorCore(StringList), vbLf)
End Function

'* 改行記号を統一します。
'*
'* @param Expression 対象の文字列
'* @param LineSep 統一後の改行記号 (既定値は `vbLf`)
'* @return 統一後の文字列
'*
'* @details
'* 改行記号を指定された記号 (`vbLf`, `vbCr`, `vbCrLf`) に置換します。
Public Function UnifyLineSeparator(ByVal Expression As String, Optional ByVal LineSep As String = vbLf) As String
    If LineSep <> vbLf And LineSep <> vbCr And LineSep <> vbCrLf Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function UnifyLineSeparator", Description:="改行記号ではありません。(" & LineSep & ")"
    End If

    UnifyLineSeparator = Replace(pUnifyLineSeparatorCore(Expression), vbLf, LineSep)
End Function

Private Function pUnifyLineSeparatorCore(ByVal Expression As String) As String
    pUnifyLineSeparatorCore = Replace(Replace(Expression, vbCrLf, vbLf), vbCr, vbLf)
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 文字列結合関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* String が格納された ObjectList を、Delimiter で区切って一つの String にします。
'*
'* @param SourceList 文字列を格納した ObjectList
'* @param Delimiter 区切り文字 (既定値は半角スペース)
'* @return 区切り文字で連結された文字列
'*
'* @details
'* 指定された文字列リストを区切り文字で連結し、単一の文字列として返します。
Public Function JoinStringList(ByVal SourceList As ObjectList, Optional ByVal Delimiter As String = " ") As String
    Dim result_value As String

    If SourceList.Count = 0 Then Exit Function

    result_value = SourceList.Item(0)

    Dim item_idx As Long
    For item_idx = 1 To SourceList.Count - 1
        result_value = result_value & Delimiter & SourceList.Item(item_idx)
    Next item_idx

    JoinStringList = result_value
End Function

'* String が格納された ObjectSet を、Delimiter で区切って一つの String にします。
'*
'* @param SourceSet 文字列を格納した ObjectSet
'* @param Delimiter 区切り文字 (既定値は半角スペース)
'* @return 区切り文字で連結された文字列
'*
'* @details
'* 指定された文字列セットを区切り文字で連結し、単一の文字列として返します。
Public Function JoinStringSet(ByVal SourceSet As ObjectSet, Optional ByVal Delimiter As String = " ") As String
    Dim result_value As String

    If SourceSet.Count = 0 Then Exit Function

    result_value = SourceSet.Item(0)

    Dim item_idx As Long
    For item_idx = 1 To SourceSet.Count - 1
        result_value = result_value & Delimiter & SourceSet.Item(item_idx)
    Next item_idx

    JoinStringSet = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' メッセージ分割関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* メッセージ文字列をページに分割して表示する MsgBox です。
'*
'* @param MessageString 表示するメッセージ文字列
'* @param Title MsgBox のタイトル (オプション)
'*
'* @details
'* 指定されたメッセージをページに分割し、ページごとに MsgBox で表示します。
Public Sub MsgBoxPage(ByVal MessageString As String, Optional ByVal Title As String = "")
    Dim msg_pages() As String
    Dim msg_page As Variant 'String
    Dim total_num As Long
    Dim page_num As Long

    msg_pages = SplitMessage(MessageString)

    total_num = UBound(msg_pages) - LBound(msg_pages) + 1

    page_num = 1
    For Each msg_page In msg_pages
        MsgBox msg_page, Title:=Title & "(" & page_num & "/" & total_num & ")"
        page_num = page_num + 1
    Next msg_page
End Sub

'* メッセージ文字列をページに分割します。
'*
'* @param MessageString 対象の文字列
'* @param PageSize ページの最大サイズ (バイト数) (既定値は 1023)
'* @return ページに分割された文字列の配列
'*
'* @details
'* 指定されたメッセージ文字列を改行単位で分割し、指定されたバイト数の範囲でページ化します。
Public Function SplitMessage(ByVal MessageString As String, Optional ByVal PageSize As Long = 1023) As Variant
    If PageSize < 4 Then Err.Raise vbObjectError + 1, "Function SplitMessage", "PageSize は 4 以上を指定してください。(" & PageSize & ")"

    Dim msgs_list As ObjectList
    Dim msg_str As String
    Dim has_msg_str As Boolean
    Dim lines_arr() As String
    Dim line_str As Variant 'String
    Dim taken_str As String
    Dim rem_str As String
    Dim result_value() As String
    Dim item_idx As Long

    Set msgs_list = New ObjectList

    If MessageString = "" Then
        ReDim result_value(0 To 0)
        result_value(0) = ""
        SplitMessage = result_value
        Exit Function
    End If

    lines_arr() = SplitByLineSeparator(MessageString)

    For Each line_str In lines_arr
        If (Not has_msg_str And LenB(StrConv(line_str, vbFromUnicode)) <= PageSize) _
                Or (has_msg_str And LenB(StrConv(msg_str, vbFromUnicode)) + 2 + LenB(StrConv(line_str, vbFromUnicode)) <= PageSize) Then
            If Not has_msg_str Then
                msg_str = line_str
                has_msg_str = True
            Else
                msg_str = msg_str & vbCrLf & line_str
            End If
        Else
            If has_msg_str Then
                Call msgs_list.Add(msg_str)
            End If

            Do While LenB(StrConv(line_str, vbFromUnicode)) > PageSize
                Call pTakeString(taken_str, rem_str, line_str, PageSize)
                Call msgs_list.Add(taken_str)
                line_str = rem_str
            Loop

            msg_str = line_str
            has_msg_str = True
        End If
    Next line_str

    If has_msg_str Then Call msgs_list.Add(msg_str)

    ReDim result_value(0 To msgs_list.Count - 1)
    For item_idx = LBound(result_value) To UBound(result_value)
        result_value(item_idx) = msgs_list.Item(item_idx)
    Next item_idx

    SplitMessage = result_value
End Function

Private Sub pTakeString( _
        ByRef TakenString As String, _
        ByRef RemainingString As String, _
        ByVal Expression As String, _
        ByVal LengthByte As Long)

    Dim char_text As String
    Dim item_idx As Long
    Dim result_value As String
    Dim total_bytes As Long
    Dim current_bytes As Long

    For item_idx = 1 To Len(Expression)
        char_text = Mid(Expression, item_idx, 1)
        current_bytes = LenB(StrConv(char_text, vbFromUnicode))

        If total_bytes + current_bytes > LengthByte Then Exit For

        result_value = result_value + char_text
        total_bytes = total_bytes + current_bytes
    Next item_idx

    TakenString = result_value
    RemainingString = Mid(Expression, item_idx)
End Sub


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' トリム関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 文字列の前後の半角空白、タブ、改行記号を削除します。
'*
'* @param Expression 対象の文字列
'* @param IgnoreHead 先頭の空白を無視する場合は True (既定値は False)
'* @param IgnoreTail 末尾の空白を無視する場合は True (既定値は False)
'* @param RemoveFullWidthSpace 全角空白も削除する場合は True (既定値は False)
'* @return 前後の空白を削除した文字列
'*
'* @details
'* 指定された条件に基づいて、文字列の前後から空白文字を削除します。
Public Function Strip( _
        ByVal Expression As String, _
        Optional ByVal IgnoreHead As Boolean = False, _
        Optional ByVal IgnoreTail As Boolean = False, _
        Optional ByVal RemoveFullWidthSpace As Boolean = False) As String

    Dim head_pos As Long
    Dim tail_pos As Long
    Dim test_char As String

    head_pos = 1
    tail_pos = Len(Expression)

    If tail_pos = 0 Then
        Strip = Expression
        Exit Function
    End If

    If Not IgnoreHead Then
        ' 先頭の空白、タブ、改行を削除
        Do While head_pos <= tail_pos
            test_char = Mid(Expression, head_pos, 1)
            If Not pIsWhitespace(test_char, RemoveFullWidthSpace) Then
                Exit Do
            End If

            head_pos = head_pos + 1
        Loop
    End If

    If Not IgnoreTail Then
        ' 末尾の空白、タブ、改行を削除
        Do While tail_pos >= head_pos
            test_char = Mid(Expression, tail_pos, 1)
            If Not pIsWhitespace(test_char, RemoveFullWidthSpace) Then
                Exit Do
            End If

            tail_pos = tail_pos - 1
        Loop
    End If

    ' 最終的な文字列を返す
    Strip = Mid(Expression, head_pos, tail_pos - head_pos + 1)
End Function

Private Function pIsWhitespace(ByVal Character As String, ByVal IncludeFullWidthSpace As Boolean) As Boolean
    Dim result_value As Boolean

    If Character = " " Then
        result_value = True
    ElseIf Character = vbTab Then
        result_value = True
    ElseIf Character = vbCrLf Then
        result_value = True
    ElseIf Character = vbCr Then
        result_value = True
    ElseIf Character = vbLf Then
        result_value = True
    ElseIf (Character = "　" And IncludeFullWidthSpace) Then
        result_value = True
    Else
        result_value = False
    End If

    pIsWhitespace = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 文字列判定関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 文字列が指定の文字列で始まるかを判定します。
'*
'* @param Expression 判定対象の文字列
'* @param SearchString 検索対象の文字列
'* @return 指定の文字列で始まる場合は True、それ以外は False
'*
'* @details
'* 判定対象の文字列が指定の文字列で始まるかを判定します。SearchString が空文字列の場合は True を返します。
Public Function StartsWith(ByVal Expression As String, ByVal SearchString As String) As Boolean
    If SearchString = "" Then
        StartsWith = True
        Exit Function
    End If

    If Len(Expression) < Len(SearchString) Then
        StartsWith = False
        Exit Function
    End If

    If Left(Expression, Len(SearchString)) = SearchString Then
        StartsWith = True
    Else
        StartsWith = False
    End If
End Function

'* 文字列が指定の文字列で終わるかを判定します。
'*
'* @param Expression 判定対象の文字列
'* @param SearchString 検索対象の文字列
'* @return 指定の文字列で終わる場合は True、それ以外は False
'*
'* @details
'* 判定対象の文字列が指定の文字列で終わるかを判定します。SearchString が空文字列の場合は True を返します。
Public Function EndsWith(ByVal Expression As String, ByVal SearchString As String) As Boolean
    If SearchString = "" Then
        EndsWith = True
        Exit Function
    End If

    If Len(Expression) < Len(SearchString) Then
        EndsWith = False
        Exit Function
    End If

    If Right(Expression, Len(SearchString)) = SearchString Then
        EndsWith = True
    Else
        EndsWith = False
    End If
End Function

'* 文字列が指定の文字列で括られているかを判定します。
'*
'* @param Expression 判定対象の文字列
'* @param QuoteString 開始文字列
'* @param EndString 終了文字列 (既定値は `QuoteString`) 終了文字列が省略された場合、開始文字列と同じものを使用します。
'* @return 指定の文字列で括られている場合は True、それ以外は False
'*
'* @details
'* 判定対象の文字列が指定の文字列で括られているかを判定します。
Public Function IsQuotedWith(ByVal Expression As String, ByVal QuoteString As String, Optional ByVal EndString As String = "") As Boolean
    If QuoteString = "" Then
        IsQuotedWith = False
        Exit Function
    End If

    If EndString = "" Then
        EndString = QuoteString
    End If

    If Len(Expression) < 2 Or Len(Expression) < (Len(QuoteString) + Len(EndString)) Then
        IsQuotedWith = False
        Exit Function
    End If

    If Left(Expression, Len(QuoteString)) = QuoteString And Right(Expression, Len(EndString)) = EndString Then
        IsQuotedWith = True
    Else
        IsQuotedWith = False
    End If
End Function


' #############################################################################
'
' 数値関連
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 最大値・最小値
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 引数として渡された Long 型の値のうち、最大のものを返します。
'*
'* @param Number1 最初の数値
'* @param Number2 2 番目の数値
'* @param Numbers その他の数値 (可変長引数)
'* @return 最大の数値
'*
'* @details
'* 最低 2 つの引数を指定し、それに加えて可変長引数として任意の個数の数値を指定できます。
Public Function MaxLng(ByVal Number1 As Long, ByVal Number2 As Long, ParamArray Numbers() As Variant) As Long
    Dim result_value As Long

    result_value = Number1
    If result_value < Number2 Then result_value = Number2

    If UBound(Numbers) <> -1 Then
        Dim item_idx As Long
        For item_idx = LBound(Numbers) To UBound(Numbers)
            If result_value < Numbers(item_idx) Then result_value = Numbers(item_idx)
        Next item_idx
    End If

    MaxLng = result_value
End Function

'* 引数として渡された Double 型の値のうち、最大のものを返します。
'*
'* @param Number1 最初の数値
'* @param Number2 2 番目の数値
'* @param Numbers その他の数値 (可変長引数)
'* @return 最大の数値
'*
'* @details
'* 最低 2 つの引数を指定し、それに加えて可変長引数として任意の個数の数値を指定できます。
Public Function MaxDbl(ByVal Number1 As Double, ByVal Number2 As Double, ParamArray Numbers() As Variant) As Double
    Dim result_value As Double

    result_value = Number1
    If result_value < Number2 Then result_value = Number2

    If UBound(Numbers) <> -1 Then
        Dim item_idx As Long
        For item_idx = LBound(Numbers) To UBound(Numbers)
            If result_value < Numbers(item_idx) Then result_value = Numbers(item_idx)
        Next item_idx
    End If

    MaxDbl = result_value
End Function

'* 引数として渡された Long 型の値のうち、最小のものを返します。
'*
'* @param Number1 最初の数値
'* @param Number2 2 番目の数値
'* @param Numbers その他の数値 (可変長引数)
'* @return 最小の数値
'*
'* @details
'* 最低 2 つの引数を指定し、それに加えて可変長引数として任意の個数の数値を指定できます。
Public Function MinLng(ByVal Number1 As Long, ByVal Number2 As Long, ParamArray Numbers() As Variant) As Long
    Dim result_value As Long

    result_value = Number1
    If Number2 < result_value Then result_value = Number2

    If UBound(Numbers) <> -1 Then
        Dim item_idx As Long
        For item_idx = LBound(Numbers) To UBound(Numbers)
            If Numbers(item_idx) < result_value Then result_value = Numbers(item_idx)
        Next item_idx
    End If

    MinLng = result_value
End Function

'* 引数として渡された Double 型の値のうち、最小のものを返します。
'*
'* @param Number1 最初の数値
'* @param Number2 2 番目の数値
'* @param Numbers その他の数値 (可変長引数)
'* @return 最小の数値
'*
'* @details
'* 最低 2 つの引数を指定し、それに加えて可変長引数として任意の個数の数値を指定できます。
Public Function MinDbl(ByVal Number1 As Double, ByVal Number2 As Double, ParamArray Numbers() As Variant) As Double
    Dim result_value As Double

    result_value = Number1
    If Number2 < result_value Then result_value = Number2

    If UBound(Numbers) <> -1 Then
        Dim item_idx As Long
        For item_idx = LBound(Numbers) To UBound(Numbers)
            If Numbers(item_idx) < result_value Then result_value = Numbers(item_idx)
        Next item_idx
    End If

    MinDbl = result_value
End Function


' #############################################################################
'
' Integer, Long 関連
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Long ビット文字列関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Long 型の値を 2 進数表記の文字列に変換します。
'*
'* @param LongValue 変換対象の Long 型の値
'* @return 2 進数表記の文字列
'*
'* @details
'* 指定された Long 型の値を 2 進数表記の文字列に変換します。
Public Function LongToBin(ByVal LongValue As Long) As String
    Dim high_bit As String
    Dim long_value As Long
    Dim item_idx As Long
    Dim result_value As String

    If 0 <= LongValue Then
        high_bit = "0"
        long_value = LongValue
    Else
        high_bit = "1"
        long_value = LongValue And &H7FFFFFFF
    End If

    For item_idx = 1 To 31
        result_value = (long_value Mod 2) & result_value
        long_value = long_value \ 2
    Next item_idx

    result_value = high_bit & result_value

    LongToBin = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' ビット シフト関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'Sub Test_BitShift()
'    Dim test_values(0 To 7) As Long
'    Dim arr_idx As Integer
'    'For arr_idx = 0 To 3
'    '    test_values(arr_idx) = 32767& * Rnd
'    '    test_values(arr_idx) = test_values(arr_idx) * (2 ^ 16) + 32767& * Rnd
'    '    If Rnd < 0.5 Then test_values(arr_idx) = -test_values(arr_idx) - 1
'    'next
'
'    test_values(0) = &H80010003
'    test_values(1) = &H80010002
'    test_values(2) = &HE0010003
'    test_values(3) = &HE0010002
'    test_values(4) = &H40010003
'    test_values(5) = &H40010002
'    test_values(6) = &H70010003
'    test_values(7) = &H70010002
'
'    arr_idx = 0 ' 0 ～ 7 の範囲で変更
'
'    Dim test_value As Long
'    test_value = test_values(arr_idx)
'
'    'Debug.Print "----------------------"
'    'Debug.Print "left"
'    'Debug.Print test_value & "(" & LongToBin(test_value) & ")"
'
'    Dim item_idx As Long
'    For item_idx = 33 To -33 Step -1
'        'Debug.Print LongToBin(BitLeft(test_value, item_idx)) & " | " & item_idx
'    Next
'    'Debug.Print "right (arithmetic)"
'    For item_idx = 0 To 33
'        'Debug.Print LongToBin(BitRight(test_value, item_idx, Arithmetic:=True)) & " | " & item_idx
'    Next
'End Sub

'* ビット左シフト演算を行う関数。
'*
'* @param TargetValue 左シフトする対象の値。
'* @param ShiftCount シフトするビット数。
'* @return 左シフトされた結果の値。
'*
'* @details
'* 引数として指定された Long 値に対して、左ビットシフト演算を行います。
'* 負のシフト数が指定された場合は、論理右シフトとして処理します。
'* ShiftCount が 32 以上、または -32 以下の場合は 0 を返します。
Public Function BitLeft(ByVal TargetValue As Long, ByVal ShiftCount As Long) As Long
    ' シフト数 0 の場合は、そのまま返す
    If ShiftCount = 0 Then
        BitLeft = TargetValue
        Exit Function
    End If

    ' シフト数が 32 以上の場合は、すべて消える
    If 31 < ShiftCount Then
        BitLeft = &H0&
        Exit Function
    End If

    ' シフト数が -32 以下の場合は、逆方向の論理右シフトですべて消える
    If ShiftCount < -31 Then
        BitLeft = &H0&
        Exit Function
    End If

    ' シフト数が負の場合は、論理右シフトとして処理
    If ShiftCount < 0 Then
        BitLeft = BitRight(TargetValue, -ShiftCount)
        Exit Function
    End If

    ' ShiftCount が 31 のときは桁あふれするので先に処理
    If ShiftCount = 31 Then
        ' ShiftCount が 31 のときは、TargetValue の最も右のビットが符合ビット (最も左) になる。
        ' その他は 0 になる。
        If TargetValue Mod 2 = 0 Then
            ' 偶数のときは最も右のビットは 0 なので、符号ビットは 0
            BitLeft = &H0&
        Else
            ' 奇数のときは最も右のビットは 1 なので、符号ビットは 1
            BitLeft = &H80000000
        End If
        Exit Function
    End If

    ' 以降、シフト数が 1 ～ 30 の計算

    Dim result_value As Long

   ' 桁あふれ対策のマスクを準備
   ' 他言語でいう「&HFFFFFFFF >> (ShiftCount + 1)」と同じ値を得るための処理
   ' 消えるはずの上位ビットと、処理後に符合ビットになるビットを 0 にするために使用する。
    Dim mask_value As Long
    mask_value = (2& ^ (31 - ShiftCount)) - 1

    ' シフト処理
    ' 他言語でいう「TargetValue << ShiftCount」を 2 の累乗を掛けることで模擬
    result_value = (TargetValue And mask_value) * (2 ^ ShiftCount)

    ' 最上位ビット取得用マスクを準備
    ' 他言語でいう「&H1& << (32 - ShiftCount)」と同じ値を得るための処理
    ' 処理後に符号ビットになるはずのビットを得るために使用する。
    Dim high_bit_mask As Long
    high_bit_mask = &H1& * (2& ^ (31 - ShiftCount))

    ' 符号ビット (最上位ビット) の処理
    ' 0 でないなら符号ビットを立てる。
    If (TargetValue And high_bit_mask) <> 0 Then
        result_value = result_value Or &H80000000
    End If

    BitLeft = result_value
End Function

'* ビット右シフト演算を行う関数。
'*
'* @param TargetValue 右シフトする対象の値。
'* @param ShiftCount シフトするビット数。
'* @param Arithmetic 算術シフトを行う場合はTrue。デフォルトはFalse。
'* @return 右シフトされた結果の値。
'*
'* @details
'* 引数として指定された Long 値に対して、右ビットシフト演算を行います。
'* 負のシフト数が指定された場合は、左シフトとして処理します。
'* 算術シフトが要求された場合、左側が符号ビットで埋められます。
'* ShiftCount が -32 以下の場合は 0 を返します。
'* ShiftCount が 32 以上の場合、論理シフトでは 0、算術シフトでは符号ビットで埋めた値を返します。
Public Function BitRight(ByVal TargetValue As Long, ByVal ShiftCount As Long, Optional ByVal Arithmetic As Boolean = False) As Long
    ' シフト数 0 の場合は、そのまま返す
    If ShiftCount = 0 Then
        BitRight = TargetValue
        Exit Function
    End If

    ' シフト数が 32 以上の場合
    If 31 < ShiftCount Then
        If Arithmetic And TargetValue < 0 Then
            ' すべて符号ビットで埋められる
            BitRight = &HFFFFFFFF
        Else
            ' すべて消える
            BitRight = 0
        End If
        Exit Function
    End If

    ' シフト数が -32 以下の場合は、逆方向の左シフトですべて消える
    If ShiftCount < -31 Then
        BitRight = &H0&
        Exit Function
    End If

    ' シフト数が負の場合は、左シフトとして処理
    If ShiftCount < 0 Then
        BitRight = BitLeft(TargetValue, -ShiftCount)
        Exit Function
    End If

    ' ShiftCount が 31 のときは桁あふれするので先に処理
    If ShiftCount = 31 Then
        If TargetValue < 0 Then
            If Arithmetic Then
                ' 31 桁は符号ビット (1) で埋められ、32 桁目は符号ビットそのもの (1)
                BitRight = &HFFFFFFFF
            Else
                ' 31 桁は 0 で埋められ、32 桁目は符号ビットそのもの (1)
                BitRight = &H1&
            End If
        Else
            ' 31 桁は符号ビット (0) で埋められ、32 桁目は符号ビットそのもの (0)
            ' 31 桁は 0 で埋められ、32 桁目は符号ビットそのもの (0)
            ' いずれにせよ、すべて 0
            BitRight = &H0&
        End If
        Exit Function
    End If

    ' 以降、シフト数が 1 ～ 30 の計算

    Dim result_value As Long
    ' まずは符号ビットを倒す
    result_value = TargetValue And &H7FFFFFFF

    ' 1 個シフト
    ' 整数の範囲内で 2 で 1 回割ることで、右シフトを模擬
    result_value = result_value \ 2

    ' 符号ビットを 1 個シフト
    ' 符号ビットが 1 (つまり負) のときに、第 2 ビット目を立てることで模擬
    If TargetValue < 0 Then
        result_value = result_value Or &H40000000
    End If

    ' 残りのシフト
    ' 整数の範囲内で、2 の (ShiftCount - 1) 乗 で割ることで、右シフトを模擬
    ' (シフト数が 1 のときは無駄な処理になるが、1 以外の時に不要な比較をするのとどちらが良いか…)
    result_value = result_value \ (2 ^ (ShiftCount - 1))

    If Arithmetic And TargetValue < 0 Then
        ' 他言語でいう「&HFFFFFFFF << (32 - ShiftCount - 1)」と同じ値を得るための処理
        Dim sign_mask As Long
        sign_mask = Not ((2& ^ (32 - ShiftCount)) - 1)

        ' 論理シフトのせいで 0 になっている部分を 1 で埋める
        result_value = result_value Or sign_mask
    End If

    BitRight = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 符号なし Long 比較関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Long 型の値を符号なし整数として比較する関数。
'*
'* @param ValueA 比較する最初の値。
'* @param ValueB 比較する2番目の値。
'* @return 値の比較結果。A = B の場合は 0、A < B の場合は -1、A > B の場合は 1。
'*
'* @details
'* 値 A と値 B を符号なし整数として比較します。
Public Function CompareAsUnsignedLong(ByVal ValueA As Long, ByVal ValueB As Long) As Integer
    If ValueA = ValueB Then
        CompareAsUnsignedLong = 0
    Else
        If 0 <= ValueA Then
            If 0 <= ValueB Then
                ' 両方正の時は普通に比較する。
                If ValueA < ValueB Then
                    CompareAsUnsignedLong = -1
                Else
                    CompareAsUnsignedLong = 1
                End If
            Else
                ' A が正、B が負のときは A < B
                CompareAsUnsignedLong = -1
            End If
        Else
            If 0 <= ValueB Then
                ' A が負、B が正のときは A > B
                CompareAsUnsignedLong = 1
            Else
                ' 両方負の時は普通に比較する。
                If ValueA < ValueB Then
                    CompareAsUnsignedLong = -1
                Else
                    CompareAsUnsignedLong = 1
                End If
            End If
        End If
    End If
End Function

'* 符号なし整数として A < B の関係を判定します。
'*
'* @param ValueA 比較する最初の値
'* @param ValueB 比較する 2 番目の値
'* @return A が B より小さい場合は True、それ以外は False
'*
'* @details
'* 2 つの Long 型の値を符号なし整数として比較し、A < B の場合に True を返します。
'* この関数の判定結果は、`CompareAsUnsignedLong` 関数が負の値を返す場合と一致します。
Public Function IsLessThanUnsignedLong(ByVal ValueA As Long, ByVal ValueB As Long) As Boolean
    If CompareAsUnsignedLong(ValueA, ValueB) < 0 Then
        IsLessThanUnsignedLong = True
    Else
        IsLessThanUnsignedLong = False
    End If
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 符号なし Long 加減算関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 2 個の Long 型の値を符号なし整数として加算します。
'*
'* @param ValueA 加算対象の最初の値
'* @param ValueB 加算対象の 2 番目の値
'* @return 符号なしとして扱った加算結果
'*
'* @details
'* 2 個の Long 型の値を符号なし整数として扱い、加算を行います。
'* 加算結果が `FFFFFFFF` を超える場合はエラーとなります。
Public Function AddUnsignedLong(ByVal ValueA As Long, ByVal ValueB As Long) As Long
    Dim a_high As Long
    Dim a_low As Long
    Call pSeparateUnsignedLong(a_high, a_low, ValueA)

    Dim b_high As Long
    Dim b_low As Long
    Call pSeparateUnsignedLong(b_high, b_low, ValueB)

    Dim carry_bit_l As Long
    Dim add_result_l As Long
    Call pAddUnsignedLongCore(add_result_l, carry_bit_l, a_low, b_low)

    Dim carry_bit_ha As Long
    Dim add_result_ha As Long
    Call pAddUnsignedLongCore(add_result_ha, carry_bit_ha, a_high, carry_bit_l)

    If 0 < carry_bit_ha Then Err.Raise vbObjectError + 1, "Function AddUnsignedLong", "加算の結果が FFFFFFFF を超えます(" & Hex(ValueA) & " + " & Hex(ValueB) & ")"

    Dim carry_bit_h As Long
    Dim add_result_h As Long
    Call pAddUnsignedLongCore(add_result_h, carry_bit_h, add_result_ha, b_high)
    If 0 < carry_bit_h Then Err.Raise vbObjectError + 1, "Function AddUnsignedLong", "加算の結果が FFFFFFFF を超えます(" & Hex(ValueA) & " + " & Hex(ValueB) & ")"

    Dim high_bit As Boolean
    If 32767 < add_result_h Then
        high_bit = True
        add_result_h = add_result_h And &H7FFF&
    Else
        high_bit = False
    End If

    add_result_h = add_result_h * 65536

    Dim result_value As Long
    result_value = add_result_h + add_result_l

    If high_bit Then
        result_value = result_value Or &H80000000
    End If

    AddUnsignedLong = result_value
End Function

Private Function pCombineUnsignedLong(ByVal HighPart As Long, ByVal LowPart As Long) As Long
    Dim high_bit As Boolean
    If 32767 < HighPart Then
        high_bit = True
        HighPart = HighPart And &H7FFF&
    Else
        high_bit = False
    End If

    Dim result_value As Long
    result_value = HighPart * 65536 + LowPart
    If high_bit Then
        result_value = result_value Or &H80000000
    End If

    pCombineUnsignedLong = result_value
End Function

Private Sub pSeparateUnsignedLong(ByRef HighPart As Long, ByRef LowPart As Long, ByVal TargetValue As Long)
    Dim high_bit As Boolean
    If TargetValue < 0 Then
        high_bit = True
        TargetValue = TargetValue And &H7FFFFFFF
    Else
        high_bit = False
    End If

    LowPart = TargetValue And &HFFFF&

    TargetValue = TargetValue \ 65536

    If high_bit Then
        HighPart = TargetValue Or &H8000&
    Else
        HighPart = TargetValue
    End If
End Sub

Private Sub pAddUnsignedLongCore(ByRef ResultValue As Long, ByRef CarryBit As Long, ByVal ValueA As Long, ByVal ValueB As Long)
    Dim result_value As Long
    result_value = ValueA + ValueB
    If 65535 < result_value Then
        ResultValue = result_value - 65536
        CarryBit = 1
    Else
        ResultValue = result_value
        CarryBit = 0
    End If
End Sub

'* 2 個の Long 型の値を符号なし整数として減算します。
'*
'* @param ValueA 減算される値
'* @param ValueB 減算する値
'* @return 符号なしとして扱った減算結果
'*
'* @details
'* 2 個の Long 型の値を符号なし整数として扱い、減算を行います。
'* 減算結果が負になる場合 (`A < B`) はエラーとなります。
Public Function SubtractUnsignedLong(ByVal ValueA As Long, ByVal ValueB As Long) As Long
    If ValueA = ValueB Then
        SubtractUnsignedLong = 0
        Exit Function
    End If

    If IsLessThanUnsignedLong(ValueA, ValueB) Then
        Err.Raise vbObjectError + 1, "Function SubtractUnsignedLong", "第 1 引数より第 2 引数のほうが大きく、計算結果が負になります。(a: " & Hex(ValueA) & ", b: " & Hex(ValueB) & ")"
        Exit Function
    End If

    Dim a_high As Long
    Dim a_low As Long
    Call pSeparateUnsignedLong(a_high, a_low, ValueA)

    Dim b_high As Long
    Dim b_low As Long
    Call pSeparateUnsignedLong(b_high, b_low, ValueB)

    Dim borrow_bit_l As Long
    Dim subtract_result_l As Long
    If a_low < b_low Then
        subtract_result_l = a_low + 65536 - b_low
        borrow_bit_l = 1
    Else
        subtract_result_l = a_low - b_low
        borrow_bit_l = 0
    End If

    Dim subtract_result_h As Long
    subtract_result_h = a_high - b_high - borrow_bit_l
    If subtract_result_h < 0 Then
        Err.Raise vbObjectError + 1, "Function SubtractUnsignedLong", "第 1 引数より第 2 引数のほうが大きく、計算結果が負になります。(a: " & Hex(ValueA) & ", b: " & Hex(ValueB) & ")"
    End If

    SubtractUnsignedLong = pCombineUnsignedLong(subtract_result_h, subtract_result_l)
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' 整数型判定
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* 指定された値が整数 (Integer 型) に変換可能かを判定します。
'*
'* @param Value 判定対象の値
'* @return 整数に変換可能な場合は True、それ以外は False
'*
'* @details
'* 指定された値が数値であり、`Integer` 型に変換可能な範囲内である場合に True を返します。
'* 値が数値でない場合や範囲外の場合は False を返します。
Public Function IsInteger(ByVal Value As Variant) As Boolean
    Dim numeric_value As Double
    If Not pTryConvertToDouble(numeric_value, Value) Then
        IsInteger = False
        Exit Function
    End If

    If numeric_value < C_INTEGER_MIN Or C_INTEGER_MAX < numeric_value Then
        IsInteger = False
        Exit Function
    End If

    IsInteger = (numeric_value = Fix(numeric_value))
End Function

'* 指定された値が長整数 (Long 型) に変換可能かを判定します。
'*
'* @param Value 判定対象の値
'* @return 長整数に変換可能な場合は True、それ以外は False
'*
'* @details
'* 指定された値が数値であり、`Long` 型に変換可能な範囲内である場合に True を返します。
'* 値が数値でない場合や範囲外の場合は False を返します。
Public Function IsLong(ByVal Value As Variant) As Boolean
    Dim numeric_value As Double
    If Not pTryConvertToDouble(numeric_value, Value) Then
        IsLong = False
        Exit Function
    End If

    If numeric_value < C_LONG_MIN Or C_LONG_MAX < numeric_value Then
        IsLong = False
        Exit Function
    End If

    IsLong = (numeric_value = Fix(numeric_value))
End Function

Private Function pTryConvertToDouble(ByRef ResultValue As Double, ByVal Value As Variant) As Boolean
    If Not IsNumeric(Value) Then
        pTryConvertToDouble = False
        Exit Function
    End If

    Err.Clear
    On Error Resume Next
    ResultValue = CDbl(Value)
    If Err.Number <> 0 Then
        Err.Clear
        pTryConvertToDouble = False
    Else
        pTryConvertToDouble = True
    End If
    On Error GoTo 0
End Function


' #############################################################################
'
' Excel 関連
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel ファイル フォーマット判定
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* ファイル パスの拡張子から、ファイルフォーマットを得ます。
'*
'* @param FileNameOrPath ファイルのパス文字列。
'* @return ファイル フォーマット。判定できなかった場合は、xlOpenXMLWorkbook となります。
'*
'* @details
'* ファイル パスの拡張子から、ファイルフォーマットを得ます。
'* 拡張子の大小文字は区別しません。
Public Function GetExcelFileFormat(ByVal FileNameOrPath As String) As Long
    Select Case LCase$(GetLeafFromPath(FileNameOrPath, BaseName:=False, Extension:=True))
     Case ".xlsm"
        GetExcelFileFormat = xlOpenXMLWorkbookMacroEnabled
     Case ".xltm"
        GetExcelFileFormat = xlOpenXMLTemplateMacroEnabled
     Case ".xls"
        GetExcelFileFormat = xlExcel8
     Case ".xla"
        GetExcelFileFormat = xlAddIn8
     Case ".xlam"
        GetExcelFileFormat = xlOpenXMLAddIn
     Case ".xlsb"
        GetExcelFileFormat = xlExcel12
     Case ".xlt"
        GetExcelFileFormat = xlTemplate8
     Case ".xltx"
        GetExcelFileFormat = xlOpenXMLTemplate
     Case ".xlw"
        GetExcelFileFormat = xlExcel4Workbook
     Case ".csv"
        GetExcelFileFormat = xlCSV
     Case ".txt"
        GetExcelFileFormat = xlCurrentPlatformText
     Case ".dbf"
        GetExcelFileFormat = xlDBF4
     Case ".dif"
        GetExcelFileFormat = xlDIF
     Case ".htm"
        GetExcelFileFormat = xlHtml
     Case ".html"
        GetExcelFileFormat = xlHtml
     Case ".ods"
        GetExcelFileFormat = xlOpenDocumentSpreadsheet
     Case ".slk"
        GetExcelFileFormat = xlSYLK
     Case Else
        GetExcelFileFormat = xlOpenXMLWorkbook
    End Select
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel アドレス種別判定
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Range アドレスが複数選択 (例: A1, B2:C3, D4) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 複数選択されている場合は True、それ以外は False
'*
'* @details
'* Range アドレスが複数選択 (例: A1, B2:C3, D4) かどうかをチェックします。
Public Function IsMultiRange(ByVal AddressString As String) As Boolean
    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, AddressString)

    IsMultiRange = 0 < InStr(cell_address, ",")
End Function

'* Range アドレスが Area (単一セルを除く連続した複数セル範囲) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return Area の場合は True、それ以外は False
'*
'* @details
'* この共通モジュールでは Area を「単一セルを除く、連続した複数セル範囲」として扱います。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsArea(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsArea = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsArea = range_bounds.IsArea
End Function

'* Range アドレスが Cell (単一セル範囲) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 単一セル範囲の場合は True、それ以外は False
'*
'* @details
'* Range アドレスが Cell (例: A1) かどうかをチェックします。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsCell(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsCell = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsCell = range_bounds.IsCell
End Function

'* Range アドレスが行全体 (例: 1:2) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 行全体である場合は True、それ以外は False
'*
'* @details
'* Range アドレスが行全体 (例: 1:2) かどうかをチェックします。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsEntireRow(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsEntireRow = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsEntireRow = range_bounds.IsEntireRow
End Function

'* Range アドレスが列全体 (例: A:B) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 列全体である場合は True、それ以外は False
'*
'* @details
'* Range アドレスが列全体 (例: A:B) かどうかをチェックします。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsEntireColumn(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsEntireColumn = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsEntireColumn = range_bounds.IsEntireColumn
End Function

'* Range アドレスが 1 行形状 (例: A1 や A1:B1 や 1:1) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 1 行形状の場合は True、それ以外は False
'*
'* @details
'* 単一セルも 1 行形状として True になります。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsOneRow(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneRow = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneRow = range_bounds.IsOneRow
End Function

'* Range アドレスが 1 列形状 (例: A1 や A1:A2 や A:A) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 1 列形状の場合は True、それ以外は False
'*
'* @details
'* 単一セルも 1 列形状として True になります。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsOneColumn(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneColumn = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneColumn = range_bounds.IsOneColumn
End Function

'* Range アドレスが 1 行だけで構成される Area かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 1 行だけで構成される Area の場合は True、それ以外は False
'*
'* @details
'* A1:B1 や 1:1 は True、A1 は Cell のため False になります。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsOneRowArea(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneRowArea = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneRowArea = range_bounds.IsOneRowArea
End Function

'* Range アドレスが 1 列だけで構成される Area かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 1 列だけで構成される Area の場合は True、それ以外は False
'*
'* @details
'* A1:A2 や A:A は True、A1 は Cell のため False になります。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsOneColumnArea(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneColumnArea = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneColumnArea = range_bounds.IsOneColumnArea
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel アドレス生成
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* パラメータを指定して範囲のアドレス表記文字列を得ます。
'*
'* @param StartRow 開始行番号 (省略時は未指定)
'* @param StartColumn 開始列番号 (省略時は未指定)
'* @param FinishRow 終了行番号 (省略時は未指定)
'* @param FinishColumn 終了列番号 (省略時は未指定)
'* @param IsAbsoluteStartRow 開始行を絶対参照で指定する場合は True (既定値は False)
'* @param IsAbsoluteStartColumn 開始列を絶対参照で指定する場合は True (既定値は False)
'* @param IsAbsoluteFinishRow 終了行を絶対参照で指定する場合は True (既定値は False)
'* @param IsAbsoluteFinishColumn 終了列を絶対参照で指定する場合は True (既定値は False)
'* @param ReferenceRow 相対参照の基準行番号 (既定値は 1)
'* @param ReferenceColumn 相対参照の基準列番号 (既定値は 1)
'* @param AddressType アドレス形式 ("A1" または "R1C1") (既定値は "A1")
'* @param SheetName シート名 (省略時は未指定)
'* @param BookName ブック名 (省略時は未指定)
'* @return 指定された条件に基づいて生成された範囲のアドレス文字列
'*
'* @details
'* 行、列、絶対参照・相対参照などのパラメータを指定して、Excel の範囲アドレスを文字列で生成します。
'*
'* 行指定が省略された場合、列範囲として扱われます。列指定が省略された場合、行範囲として扱われます。
'* 行指定と列指定の両方を省略することはできません。
'* BookName を指定する場合、SheetName も指定する必要があります。
Public Function RangeAddress( _
        Optional ByVal StartRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal StartColumn As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishColumn As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal IsAbsoluteStartRow As Boolean = False, _
        Optional ByVal IsAbsoluteStartColumn As Boolean = False, _
        Optional ByVal IsAbsoluteFinishRow As Boolean = False, _
        Optional ByVal IsAbsoluteFinishColumn As Boolean = False, _
        Optional ByVal ReferenceRow As Long = 0, _
        Optional ByVal ReferenceColumn As Long = 0, _
        Optional ByVal AddressType As String = "A1", _
        Optional ByVal SheetName As String = "", _
        Optional ByVal BookName As String = "") As String

    Call pValidateBookAndSheetAddress(BookName, SheetName, "Function RangeAddress")

    ' 行指定および列指定の確認
    If StartRow = G_OMIT_CELL_INDEX And StartColumn = G_OMIT_CELL_INDEX Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="開始行と開始列の両方を省略することはできません。"
        Exit Function
    End If

    If StartRow = G_OMIT_CELL_INDEX And FinishRow <> G_OMIT_CELL_INDEX Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="開始行のみを省略することはできません。(finish_row: " & FinishRow & ")"
        Exit Function
    ElseIf StartRow <> G_OMIT_CELL_INDEX And FinishRow = G_OMIT_CELL_INDEX Then
        If IsAbsoluteStartRow = IsAbsoluteFinishRow Then
            ' 絶対アドレス指定が一致していたら、FinishRow を補完する。
            FinishRow = StartRow
        ElseIf Not IsAbsoluteFinishRow Then
            ' IsAbsoluteFinishRow が初期値 (False) のままなら、FinishRow と IsAbsoluteFinishRow を補完する。
            FinishRow = StartRow
            IsAbsoluteFinishRow = IsAbsoluteStartRow
        Else
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="終了行を省略することはできません。(start_row: " & StartRow & ")"
            Exit Function
        End If
    End If

    If StartColumn = G_OMIT_CELL_INDEX And FinishColumn <> G_OMIT_CELL_INDEX Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="開始列のみを省略することはできません。(finish_col: " & FinishColumn & ")"
        Exit Function
    ElseIf StartColumn <> G_OMIT_CELL_INDEX And FinishColumn = G_OMIT_CELL_INDEX Then
        If IsAbsoluteStartColumn = IsAbsoluteFinishColumn Then
            ' 絶対アドレス指定が一致していたら、FinishColumn を補完する。
            FinishColumn = StartColumn
        ElseIf Not IsAbsoluteFinishColumn Then
            ' IsAbsoluteFinishColumn が初期値 (False) のままなら、FinishColumn と IsAbsoluteFinishColumn を補完する。
            FinishColumn = StartColumn
            IsAbsoluteFinishColumn = IsAbsoluteStartColumn
        Else
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="終了列を省略することはできません。(start_col: " & StartColumn & ")"
            Exit Function
        End If
    End If

    Dim is_row_range As Boolean
    If StartColumn = G_OMIT_CELL_INDEX Then
        is_row_range = True
    Else
        is_row_range = False
    End If

    Dim is_col_range As Boolean
    If StartRow = G_OMIT_CELL_INDEX Then
        is_col_range = True
    Else
        is_col_range = False
    End If

    Dim is_cell_range As Boolean
    If StartRow <> G_OMIT_CELL_INDEX And FinishRow = StartRow And IsAbsoluteStartRow = IsAbsoluteFinishRow _
            And StartColumn <> G_OMIT_CELL_INDEX And FinishColumn = StartColumn And IsAbsoluteStartColumn = IsAbsoluteFinishColumn Then
        is_cell_range = True
    Else
        is_cell_range = False
    End If

    Dim result_value As String
    result_value = ExcelBookAndSheetAddress(BookName, SheetName)

    Select Case AddressType
     Case "A1"
        If is_cell_range Then
            result_value = result_value & pA1columnAddressCore(StartColumn, IsAbsoluteStartColumn, ReferenceColumn) & pA1RowAddressCore(StartRow, IsAbsoluteStartRow, ReferenceRow)
        Else
            If is_row_range Then
                result_value = result_value & pA1RowAddressCore(StartRow, IsAbsoluteStartRow, ReferenceRow) & ":" & pA1RowAddressCore(FinishRow, IsAbsoluteFinishRow, ReferenceRow)
            ElseIf is_col_range Then
                result_value = result_value & pA1columnAddressCore(StartColumn, IsAbsoluteStartColumn, ReferenceColumn) & ":" & pA1columnAddressCore(FinishColumn, IsAbsoluteFinishColumn, ReferenceColumn)
            Else
                result_value = result_value & pA1columnAddressCore(StartColumn, IsAbsoluteStartColumn, ReferenceColumn) & pA1RowAddressCore(StartRow, IsAbsoluteStartRow, ReferenceRow) & _
                        ":" & pA1columnAddressCore(FinishColumn, IsAbsoluteFinishColumn, ReferenceColumn) & pA1RowAddressCore(FinishRow, IsAbsoluteFinishRow, ReferenceRow)
            End If
        End If
     Case "R1C1"
        If is_cell_range Then
            result_value = result_value & "R" & pR1C1RowAddressCore(StartRow, IsAbsoluteStartRow) & "C" & pR1C1ColumnAddressCore(StartColumn, IsAbsoluteStartColumn)
        Else
            If is_row_range Then
                result_value = result_value & "R" & pR1C1RowAddressCore(StartRow, IsAbsoluteStartRow) & ":R" & pR1C1RowAddressCore(FinishRow, IsAbsoluteFinishRow)
            ElseIf is_col_range Then
                result_value = result_value & "C" & pR1C1ColumnAddressCore(StartColumn, IsAbsoluteStartColumn) & ":C" & pR1C1ColumnAddressCore(FinishColumn, IsAbsoluteFinishColumn)
            Else
                result_value = result_value & "R" & pR1C1RowAddressCore(StartRow, IsAbsoluteStartRow) & "C" & pR1C1ColumnAddressCore(StartColumn, IsAbsoluteStartColumn) & _
                        ":" & "R" & pR1C1RowAddressCore(FinishRow, IsAbsoluteFinishRow) & "C" & pR1C1ColumnAddressCore(FinishColumn, IsAbsoluteFinishColumn)
            End If
        End If
     Case Else
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="AddressType は A1 か R1C1 です。(" & AddressType & ")"
        Exit Function
    End Select

    RangeAddress = result_value
End Function

Private Function pA1columnAddressCore(ByVal ColumnIndex As Long, ByVal IsAbsolute As Boolean, ByVal ReferenceColumn As Long) As String
    Dim actual_column As Double
    If IsAbsolute Then
        actual_column = ColumnIndex
        If actual_column < 1 Or G_COL_MAX < actual_column Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="列インデックスが範囲外です。(" & ColumnIndex & ")"
        End If

        pA1columnAddressCore = "$" & ExcelA1ColumnAddress(CLng(actual_column))
    Else
        actual_column = CDbl(ReferenceColumn) + CDbl(ColumnIndex)
        If actual_column < 1 Or G_COL_MAX < actual_column Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="列インデックスが範囲外です。(" & ReferenceColumn & " + " & ColumnIndex & " = " & CStr(actual_column) & ")"
        End If

        pA1columnAddressCore = ExcelA1ColumnAddress(CLng(actual_column))
    End If
End Function

Private Function pA1RowAddressCore(ByVal RowIndex As Long, ByVal IsAbsolute As Boolean, ByVal ReferenceRow As Long) As String
    Dim actual_row As Double
    If IsAbsolute Then
        actual_row = RowIndex
        If actual_row < 1 Or G_ROW_MAX < actual_row Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="行インデックスが範囲外です。(" & RowIndex & ")"
        End If

        pA1RowAddressCore = "$" & CStr(CLng(actual_row))
    Else
        actual_row = CDbl(ReferenceRow) + CDbl(RowIndex)
        If actual_row < 1 Or G_ROW_MAX < actual_row Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="行インデックスが範囲外です。(" & ReferenceRow & " + " & RowIndex & " = " & CStr(actual_row) & ")"
        End If

        pA1RowAddressCore = CStr(CLng(actual_row))
    End If
End Function

Private Function pR1C1RowAddressCore(ByVal RowIndex As Long, ByVal IsAbsolute As Boolean) As String
    If IsAbsolute Then
        If RowIndex < 1 Or G_ROW_MAX < RowIndex Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="行インデックスが範囲外です。(" & RowIndex & ")"
        End If

        pR1C1RowAddressCore = CStr(RowIndex)
    Else
        If RowIndex <> 0 Then
            pR1C1RowAddressCore = "[" & CStr(RowIndex) & "]"
        Else
            pR1C1RowAddressCore = ""
        End If
    End If
End Function

Private Function pR1C1ColumnAddressCore(ByVal ColumnIndex As Long, ByVal IsAbsolute As Boolean) As String
    If IsAbsolute Then
        If ColumnIndex < 1 Or G_COL_MAX < ColumnIndex Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="列インデックスが範囲外です。(" & ColumnIndex & ")"
        End If

        pR1C1ColumnAddressCore = CStr(ColumnIndex)
    Else
        If ColumnIndex <> 0 Then
            pR1C1ColumnAddressCore = "[" & CStr(ColumnIndex) & "]"
        Else
            pR1C1ColumnAddressCore = ""
        End If
    End If
End Function

'* ブック名とシート名を Excel アドレス形式にした文字列を得ます。
'*
'* @param BookName ブック名 (省略時は空文字列)
'* @param SheetName シート名 (省略時は空文字列)
'* @return 指定されたブック名およびシート名を含む Excel アドレス形式の文字列
'*
'* @details
'* ブック名とシート名を組み合わせて、Excel で使用可能なアドレス形式の文字列を生成します。
'* BookName を指定する場合、SheetName も指定する必要があります。
Public Function ExcelBookAndSheetAddress(Optional ByVal BookName As String = "", Optional ByVal SheetName As String = "") As String
    Call pValidateBookAndSheetAddress(BookName, SheetName, "Function ExcelBookAndSheetAddress")

    If pIsNeedQuoteBookAndSheetAddressCore(BookName, SheetName) Then
        If BookName <> "" Then
            ExcelBookAndSheetAddress = "'[" & pEscapeBookAndSheetAddressCore(BookName) & "]" & pEscapeBookAndSheetAddressCore(SheetName) & "'!"
        ElseIf SheetName <> "" Then
            ExcelBookAndSheetAddress = "'" & pEscapeBookAndSheetAddressCore(SheetName) & "'!"
        Else
            'Debug.Print "Function ExcelBookAndSheetAddress: 想定外"
        End If
    Else
        If BookName <> "" Then
            ExcelBookAndSheetAddress = "[" & BookName & "]" & SheetName & "!"
        ElseIf SheetName <> "" Then
            ExcelBookAndSheetAddress = SheetName & "!"
        Else
            ExcelBookAndSheetAddress = ""
        End If
    End If
End Function

Private Sub pValidateBookAndSheetAddress(ByVal BookName As String, ByVal SheetName As String, ByVal ErrorSource As String)
    If BookName <> "" And SheetName = "" Then
        Err.Raise vbObjectError + 1, ErrorSource, "ブック名を指定する場合はシート名も指定してください。(" & BookName & ")"
    End If
End Sub

Private Function pIsNeedQuoteBookAndSheetAddressCore(ByRef BookName As String, ByRef SheetName As String) As Boolean
    If BookName = "" And SheetName = "" Then
        pIsNeedQuoteBookAndSheetAddressCore = False
    ElseIf pBookNameNeedsQuote(BookName) Or pSheetNameNeedsQuote(SheetName) Then
        pIsNeedQuoteBookAndSheetAddressCore = True
    Else
        pIsNeedQuoteBookAndSheetAddressCore = False
    End If
End Function

Private Function pBookNameNeedsQuote(ByVal BookName As String) As Boolean
    If BookName = "" Then Exit Function

    pBookNameNeedsQuote = (0 < InStr(BookName, " ") Or 0 < InStr(BookName, "'") Or 0 < InStr(BookName, "!") _
            Or 0 < InStr(BookName, "[") Or 0 < InStr(BookName, "]") Or 0 < InStr(BookName, "(") Or 0 < InStr(BookName, ")"))
End Function

Private Function pSheetNameNeedsQuote(ByVal SheetName As String) As Boolean
    If SheetName = "" Then Exit Function

    If Not pIsUnquotedSheetName(SheetName) Then
        pSheetNameNeedsQuote = True
        Exit Function
    End If

    If pLooksLikeA1Reference(SheetName) Or pLooksLikeR1C1Reference(SheetName) Then
        pSheetNameNeedsQuote = True
    Else
        pSheetNameNeedsQuote = False
    End If
End Function

Private Function pIsUnquotedSheetName(ByVal SheetName As String) As Boolean
    If SheetName = "" Then Exit Function
    If Not pIsUnquotedSheetNameStartChar(Left$(SheetName, 1)) Then Exit Function

    Dim char_idx As Long
    For char_idx = 2 To Len(SheetName)
        If Not pIsUnquotedSheetNameChar(Mid$(SheetName, char_idx, 1)) Then Exit Function
    Next char_idx

    pIsUnquotedSheetName = True
End Function

Private Function pIsUnquotedSheetNameStartChar(ByVal TestChar As String) As Boolean
    pIsUnquotedSheetNameStartChar = (pIsAsciiAlphabet(TestChar) Or TestChar = "_" Or pIsNonAsciiChar(TestChar))
End Function

Private Function pIsUnquotedSheetNameChar(ByVal TestChar As String) As Boolean
    pIsUnquotedSheetNameChar = (pIsUnquotedSheetNameStartChar(TestChar) Or pIsAsciiDigit(TestChar) Or TestChar = ".")
End Function

Private Function pIsNonAsciiChar(ByVal TestChar As String) As Boolean
    Dim char_code As Long
    char_code = AscW(TestChar)
    pIsNonAsciiChar = (char_code < 0 Or 127 < char_code)
End Function

Private Function pLooksLikeA1Reference(ByVal SheetName As String) As Boolean
    Err.Clear
    On Error Resume Next
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, SheetName)
    pLooksLikeA1Reference = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function pLooksLikeR1C1Reference(ByVal SheetName As String) As Boolean
    Dim upper_name As String
    upper_name = UCase$(SheetName)
    If Left$(upper_name, 1) <> "R" Then Exit Function

    Dim c_pos As Long
    c_pos = InStr(2, upper_name, "C")
    If c_pos <= 2 Or c_pos = Len(upper_name) Then Exit Function
    If 0 < InStr(c_pos + 1, upper_name, "C") Then Exit Function

    Dim row_text As String
    Dim col_text As String
    row_text = Mid$(upper_name, 2, c_pos - 2)
    col_text = Mid$(upper_name, c_pos + 1)

    If Not pIsAsciiDigitString(row_text) Or Not pIsAsciiDigitString(col_text) Then Exit Function

    Dim row_num As Double
    Dim col_num As Double
    row_num = CDbl(row_text)
    col_num = CDbl(col_text)

    pLooksLikeR1C1Reference = (1 <= row_num And row_num <= G_ROW_MAX And 1 <= col_num And col_num <= G_COL_MAX)
End Function

Private Function pIsAsciiDigitString(ByVal TestString As String) As Boolean
    If TestString = "" Then Exit Function

    Dim char_idx As Long
    For char_idx = 1 To Len(TestString)
        If Not pIsAsciiDigit(Mid$(TestString, char_idx, 1)) Then Exit Function
    Next char_idx

    pIsAsciiDigitString = True
End Function

Private Function pEscapeBookAndSheetAddressCore(ByVal BookOrSheetName As String) As String
    pEscapeBookAndSheetAddressCore = Replace(Replace(Replace(BookOrSheetName, "'", "''"), "[", "("), "]", ")")
End Function

'* 列番号を A1 形式の列名 (例: A, B, ... Z, AA, AB...) に変換します。
'*
'* @param ColumnIndex 列番号 (1 以上 G_COL_MAX 以下の値)
'* @return 列番号に対応する A1 形式の列名
'*
'* @details
'* 指定された列番号を基に、Excel の A1 形式の列名を生成します。
'* 列番号が 1 未満または G_COL_MAX を超える場合、エラーになります。
Public Function ExcelA1ColumnAddress(ByVal ColumnIndex As Long) As String
    If ColumnIndex < 1 Or G_COL_MAX < ColumnIndex Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function ExcelA1ColumnAddress", Description:="列インデックスが範囲外です。(" & ColumnIndex & ")"
    End If

    Dim result_value As String
    Dim temp_num As Long

    Do While ColumnIndex > 0
        temp_num = (ColumnIndex - 1) Mod 26
        result_value = Chr(temp_num + 65) & result_value
        ColumnIndex = (ColumnIndex - 1) \ 26
    Loop

    ExcelA1ColumnAddress = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel アドレス分割
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Excel のアドレス表記から各情報を取り出します。
'*
'* @param FolderPath [出力] 結果として取得されるフォルダパス
'* @param BookName [出力] 結果として取得されるブック名
'* @param SheetName [出力] 結果として取得されるシート名
'* @param CellAddress [出力] 結果として取得されるセルアドレス
'* @param AddressString 分解対象となる Excel アドレス文字列
'*
'* @details
'* 指定された Excel アドレス文字列を分解し、フォルダパス、ブック名、シート名、セルアドレスに分けて出力します。
'* 形式が正しくない場合はエラーにします。
Public Sub SplitExcelAddress(ByRef FolderPath As String, ByRef BookName As String, ByRef SheetName As String, ByRef CellAddress As String, ByVal AddressString As String)
    FolderPath = ""
    BookName = ""
    SheetName = ""
    CellAddress = ""

    If AddressString = "" Then
        Call pRaiseInvalidExcelAddress(AddressString)
    End If

    Dim delimiter_idx As Long
    delimiter_idx = pFindExcelAddressDelimiter(AddressString)
    If delimiter_idx < 0 Then
        Call pRaiseInvalidExcelAddress(AddressString)
    End If

    Dim location_part As String
    If delimiter_idx = 0 Then
        CellAddress = AddressString
        If 0 < InStr(CellAddress, "[") Or 0 < InStr(CellAddress, "]") Or 0 < InStr(CellAddress, "'") Then
            Call pRaiseInvalidExcelAddress(AddressString)
        End If
    Else
        location_part = Left(AddressString, delimiter_idx - 1)
        CellAddress = Mid(AddressString, delimiter_idx + 1)

        If location_part = "" Or CellAddress = "" Then
            Call pRaiseInvalidExcelAddress(AddressString)
        End If

        If StartsWith(location_part, "'") Then
            If Not EndsWith(location_part, "'") Or Len(location_part) < 2 Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
            location_part = Mid(location_part, 2, Len(location_part) - 2)
            location_part = Replace(location_part, "''", "'")
        ElseIf 0 < InStr(location_part, "'") Then
            Call pRaiseInvalidExcelAddress(AddressString)
        End If

        Dim close_book_idx As Long
        close_book_idx = InStrRev(location_part, "]")
        If 0 < close_book_idx Then
            Dim open_book_idx As Long
            open_book_idx = InStrRev(Left(location_part, close_book_idx - 1), "[")
            If open_book_idx < 1 Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If

            FolderPath = Left(location_part, open_book_idx - 1)
            BookName = Mid(location_part, open_book_idx + 1, close_book_idx - open_book_idx - 1)
            SheetName = Mid(location_part, close_book_idx + 1)
            If BookName = "" Or SheetName = "" Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
        Else
            If 0 < InStr(location_part, "[") Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
            SheetName = location_part
            If SheetName = "" Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
        End If
    End If

    If CellAddress = "" Then
        Call pRaiseInvalidExcelAddress(AddressString)
    End If
End Sub

Private Function pFindExcelAddressDelimiter(ByVal AddressString As String) As Long
    Dim result_value As Long
    Dim is_in_quote As Boolean
    Dim char_idx As Long
    For char_idx = 1 To Len(AddressString)
        Dim current_char As String
        current_char = Mid$(AddressString, char_idx, 1)
        If current_char = "'" Then
            If is_in_quote And char_idx < Len(AddressString) And Mid$(AddressString, char_idx + 1, 1) = "'" Then
                char_idx = char_idx + 1
            Else
                is_in_quote = Not is_in_quote
            End If
        ElseIf current_char = "!" And Not is_in_quote Then
            If result_value <> 0 Then
                pFindExcelAddressDelimiter = -1
                Exit Function
            End If
            result_value = char_idx
        End If
    Next

    If is_in_quote Then
        pFindExcelAddressDelimiter = -1
    Else
        pFindExcelAddressDelimiter = result_value
    End If
End Function

Private Sub pRaiseInvalidExcelAddress(ByVal AddressString As String)
    Err.Raise Number:=vbObjectError + 1, Source:="Sub SplitExcelAddress", Description:="Excel アドレス文字列の形式が正しくありません。(" & AddressString & ")"
End Sub

'* A1 形式の単一矩形範囲アドレスを、開始・終了インデックスへ分解します。
'*
'* @param StartRow [出力] 開始行番号。列範囲の場合は G_OMIT_CELL_INDEX。
'* @param StartColumn [出力] 開始列番号。行範囲の場合は G_OMIT_CELL_INDEX。
'* @param FinishRow [出力] 終了行番号。列範囲の場合は G_OMIT_CELL_INDEX。
'* @param FinishColumn [出力] 終了列番号。行範囲の場合は G_OMIT_CELL_INDEX。
'* @param AddressString 分解対象の A1 形式アドレス。ブック名・シート名は含めない。
'*
'* @details
'* A1、A1:B2、1:3、A:C、$A$1:$B$2 を扱います。
'* 複数範囲、R1C1 形式、ブック名・シート名付きアドレス、不完全なアドレスはエラーにします。
Public Sub SplitA1RangeAddress( _
        ByRef StartRow As Long, _
        ByRef StartColumn As Long, _
        ByRef FinishRow As Long, _
        ByRef FinishColumn As Long, _
        ByVal AddressString As String)

    StartRow = G_OMIT_CELL_INDEX
    StartColumn = G_OMIT_CELL_INDEX
    FinishRow = G_OMIT_CELL_INDEX
    FinishColumn = G_OMIT_CELL_INDEX

    Dim normalized_address As String
    normalized_address = Trim(AddressString)
    If normalized_address = "" Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If

    If 0 < InStr(normalized_address, ",") Or 0 < InStr(normalized_address, "!") _
            Or 0 < InStr(normalized_address, "[") Or 0 < InStr(normalized_address, "]") _
            Or 0 < InStr(normalized_address, "'") Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If

    Dim address_parts() As String
    address_parts = Split(normalized_address, ":")

    Dim start_type As Long
    Dim finish_type As Long
    If UBound(address_parts) = 0 Then
        Call pSplitA1AddressToken(StartRow, StartColumn, start_type, address_parts(0), AddressString)
        If start_type <> C_A1_TOKEN_CELL Then
            Call pRaiseInvalidA1RangeAddress(AddressString)
        End If
        FinishRow = StartRow
        FinishColumn = StartColumn
    ElseIf UBound(address_parts) = 1 Then
        Call pSplitA1AddressToken(StartRow, StartColumn, start_type, address_parts(0), AddressString)
        Call pSplitA1AddressToken(FinishRow, FinishColumn, finish_type, address_parts(1), AddressString)
        If start_type <> finish_type Then
            Call pRaiseInvalidA1RangeAddress(AddressString)
        End If

        If start_type = C_A1_TOKEN_ROW Then
            StartColumn = G_OMIT_CELL_INDEX
            FinishColumn = G_OMIT_CELL_INDEX
        ElseIf start_type = C_A1_TOKEN_COLUMN Then
            StartRow = G_OMIT_CELL_INDEX
            FinishRow = G_OMIT_CELL_INDEX
        End If
    Else
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If

    If StartRow <> G_OMIT_CELL_INDEX And FinishRow <> G_OMIT_CELL_INDEX And FinishRow < StartRow Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If
    If StartColumn <> G_OMIT_CELL_INDEX And FinishColumn <> G_OMIT_CELL_INDEX And FinishColumn < StartColumn Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If
End Sub

Private Sub pSplitA1AddressToken( _
        ByRef RowIndex As Long, _
        ByRef ColumnIndex As Long, _
        ByRef TokenType As Long, _
        ByVal AddressToken As String, _
        ByVal OriginalAddressString As String)

    RowIndex = G_OMIT_CELL_INDEX
    ColumnIndex = G_OMIT_CELL_INDEX
    TokenType = 0

    Dim normalized_token As String
    normalized_token = UCase(AddressToken)
    If normalized_token = "" Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If

    Dim char_idx As Long
    char_idx = 1

    Dim has_leading_dollar As Boolean
    If Mid$(normalized_token, char_idx, 1) = "$" Then
        has_leading_dollar = True
        char_idx = char_idx + 1
        If Len(normalized_token) < char_idx Then
            Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
        End If
    End If

    Dim col_text As String
    Do While char_idx <= Len(normalized_token)
        Dim char_code As Long
        char_code = Asc(Mid$(normalized_token, char_idx, 1))
        If char_code < Asc("A") Or Asc("Z") < char_code Then Exit Do
        col_text = col_text & Chr(char_code)
        char_idx = char_idx + 1
    Loop

    If col_text = "" And has_leading_dollar Then
        If char_idx <= Len(normalized_token) And Mid$(normalized_token, char_idx, 1) = "$" Then
            Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
        End If
    ElseIf col_text <> "" Then
        If char_idx <= Len(normalized_token) And Mid$(normalized_token, char_idx, 1) = "$" Then
            char_idx = char_idx + 1
            If Len(normalized_token) < char_idx Then
                Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
            End If
        End If
    End If

    Dim row_text As String
    Do While char_idx <= Len(normalized_token)
        char_code = Asc(Mid$(normalized_token, char_idx, 1))
        If char_code < Asc("0") Or Asc("9") < char_code Then Exit Do
        row_text = row_text & Chr(char_code)
        char_idx = char_idx + 1
    Loop

    If char_idx <= Len(normalized_token) Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If

    If col_text <> "" Then
        ColumnIndex = pA1ColumnIndex(col_text, OriginalAddressString)
    End If
    If row_text <> "" Then
        RowIndex = pA1RowIndex(row_text, OriginalAddressString)
    End If

    If col_text <> "" And row_text <> "" Then
        TokenType = C_A1_TOKEN_CELL
    ElseIf row_text <> "" Then
        TokenType = C_A1_TOKEN_ROW
    ElseIf col_text <> "" Then
        TokenType = C_A1_TOKEN_COLUMN
    Else
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If
End Sub

Private Function pA1ColumnIndex(ByVal ColumnAddress As String, ByVal OriginalAddressString As String) As Long
    Dim result_value As Long
    Dim char_idx As Long
    For char_idx = 1 To Len(ColumnAddress)
        Dim char_code As Long
        char_code = Asc(Mid(ColumnAddress, char_idx, 1))
        result_value = result_value * 26 + char_code - Asc("A") + 1
        If G_COL_MAX < result_value Then
            Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
        End If
    Next

    If result_value < 1 Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If
    pA1ColumnIndex = result_value
End Function

Private Function pA1RowIndex(ByVal RowAddress As String, ByVal OriginalAddressString As String) As Long
    If Len(CStr(G_ROW_MAX)) < Len(RowAddress) Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If

    Dim result_value As Long
    result_value = CLng(RowAddress)
    If result_value < 1 Or G_ROW_MAX < result_value Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If
    pA1RowIndex = result_value
End Function

Private Sub pRaiseInvalidA1RangeAddress(ByVal AddressString As String)
    Err.Raise Number:=vbObjectError + 1, Source:="Sub SplitA1RangeAddress", Description:="A1 形式の単一矩形範囲アドレスではありません。(" & AddressString & ")"
End Sub


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Range 値変換
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Range の値を格納した ObjectList を返します。
'*
'* @param TargetRange 対象となる範囲
'* @param IgnoreEmpty 空のセルを無視する場合は True (既定値は False)
'* @param GetText Text プロパティを取得する場合は True (既定値は False)
'* @return 対象範囲内のセルを文字列化して格納した ObjectList
'*
'* @details
'* 指定された範囲の各セルを Value または Text として文字列化し、ObjectList に格納して返します。
'* IgnoreEmpty が True の場合、空のセルはリストに含まれません。
Public Function ConvertRangeToStringList(ByVal TargetRange As WorksheetRangeBounds, Optional ByVal IgnoreEmpty As Boolean = False, Optional ByVal GetText As Boolean = False) As ObjectList
    Dim result_value As ObjectList
    Set result_value = New ObjectList

    Dim enum_obj As IEnumerator
    Set enum_obj = TargetRange.GetEnumerator( _
            EnumerationMode:=G_RANGE_ENUM_MODE_CELLS_HORIZONTAL, _
            Descending:=False)
    Do While enum_obj.MoveNext()
        Dim cell_item As WorksheetRangeBounds
        Set cell_item = enum_obj.Current

        Dim cell_text As String
        Call WsSrv.ReadCell(cell_item, cell_text, GetText:=GetText)

        If Not IgnoreEmpty Or cell_text <> "" Then
            Call result_value.Add(cell_text)
        End If
    Loop

    Set ConvertRangeToStringList = result_value
End Function

