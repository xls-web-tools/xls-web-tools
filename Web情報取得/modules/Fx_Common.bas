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
