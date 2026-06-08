Attribute VB_Name = "Fx_Common"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! CommonModules が Excel ワークシート関数として公開する UDF をまとめた標準モジュールです。
'!
' #############################################################################

'* 文字列同士を比較し、差分の文字列を取得します。
'*
'* @param Expression1 元の文字列。
'* @param Expression2 変更後の文字列。
'* @param ExtractType 抽出タイプ。0 を指定すると、元の文字列から削除された文字列を取得します。1 を指定すると、元の文字列に追加された文字列を取得します。
'* @return 差分の文字列。
'*
'* @details
'* 文字列同士を比較し、差分の文字を前から順に結合した文字列を取得します。
'* たとえば、Expression1 が「aaa [bbb ccc] ddd」、Expression2 が「aaa bbb ddd eee」で、ExtractType が 0 の場合は、「[ccc] 」が返ります。
Public Function DIFFSTR( _
    ByVal Expression1 As String, _
    ByVal Expression2 As String, _
    Optional ByVal ExtractType As Integer = 0) As Variant  ' String

    Call InitializeUdfCommonService

    '=== 1) 空文字列が混在するケースの簡易処理 ==================
    If Expression1 = "" Or Expression2 = "" Then
        If Expression1 = "" Then
            If ExtractType = 0 Then
                DIFFSTR = ""
            Else
                DIFFSTR = Expression2
            End If
        Else
            ' Expression2 が空文字列
            If ExtractType = 0 Then
                DIFFSTR = Expression1
            Else
                DIFFSTR = ""
            End If
        End If
        Exit Function
    End If

    '=== 2) 文字列を文字配列に変換 ==============================
    Dim expr_1_arr() As String
    Dim expr_2_arr() As String
    expr_1_arr = ConvertStringToCharArray(Expression1)
    expr_2_arr = ConvertStringToCharArray(Expression2)

    '=== 3) DiffStringArray を使用して差分を算出 ===============
    Dim diff_type_arr() As String
    Call DiffStringArray(expr_1_arr, expr_2_arr, diff_type_arr, EnableReplaceType:=False)

    '=== 4) 抽出タイプに応じて、差分文字列を再構築 ==============
    Dim result_str As String
    result_str = ""

    Dim i_idx As Long
    Dim last_idx As Long
    last_idx = UBound(expr_1_arr)  ' = UBound(expr_2_arr) = UBound(diff_type_arr)

    For i_idx = LBound(expr_1_arr) To last_idx
        Dim diff_tag As String
        diff_tag = diff_type_arr(i_idx)

        Select Case ExtractType
            Case 0
                ' 削除で旧文字列側が取り除かれた文字を集める
                If (diff_tag = "DEL") Then
                    result_str = result_str & expr_1_arr(i_idx)
                End If
            Case 1
                ' 追加で新文字列側に増えた文字を集める
                If (diff_tag = "ADD") Then
                    result_str = result_str & expr_2_arr(i_idx)
                End If
        End Select
    Next i_idx

    DIFFSTR = result_str
End Function

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
