Attribute VB_Name = "Test_ColumnExpression"
Option Explicit
Option Base 0

Public Sub Test_ColumnExpression_ANDOR条件を評価できる(ByVal Assert As UnitTestAssert)
    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("([申請者] == ""山田"" OR [代理申請者] == ""山田"") AND [処理済み日] != ""済""", G_COLUMN_EXPRESSION_KIND_CONDITION, "OutputConditionExpression")

    Dim column_values As ObjectDictionary
    Set column_values = pColumnValues( _
            "申請者", "", _
            "代理申請者", "山田", _
            "処理済み日", "")

    Dim actual As Boolean

    On Error Resume Next
    actual = column_expression.EvaluateCondition(column_values)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    On Error GoTo 0

    Call Assert.IsTrue(actual, "列値が条件に一致する場合は True")
End Sub

Public Sub Test_ColumnExpression_条件不一致はFalseを返す(ByVal Assert As UnitTestAssert)
    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("[状態] == ""完了""", G_COLUMN_EXPRESSION_KIND_CONDITION, "OutputConditionExpression")

    Dim actual As Boolean

    On Error Resume Next
    actual = column_expression.EvaluateCondition(pColumnValues("状態", "未完了"))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    On Error GoTo 0

    Call Assert.IsFalse(actual, "列値が条件に一致しない場合は False")
End Sub

Public Sub Test_ColumnExpression_参照列名は登場順で重複除去する(ByVal Assert As UnitTestAssert)
    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("([A] == ""x"" OR [B] == ""y"") AND [A] != ""z""", G_COLUMN_EXPRESSION_KIND_CONDITION, "OutputConditionExpression")

    Dim actual_names As ObjectList

    On Error Resume Next
    Set actual_names = column_expression.ReferencedColumnNames
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    On Error GoTo 0

    Assert.EqualsNumeric 2, actual_names.Count
    Assert.Equals "A", CStr(actual_names.Item(0))
    Assert.Equals "B", CStr(actual_names.Item(1))
End Sub

Public Sub Test_ColumnExpression_IF値式は選択した枝だけ評価する(ByVal Assert As UnitTestAssert)
    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("IF([判定] == ""Y"", [存在], [未入力])", G_COLUMN_EXPRESSION_KIND_VALUE, "ValueExpression(採用列)")

    Dim actual As String

    On Error Resume Next
    actual = column_expression.EvaluateValue(pColumnValues("判定", "Y", "存在", "採用値"))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    On Error GoTo 0

    Assert.Equals "採用値", actual
End Sub

Public Sub Test_ColumnExpression_OR条件は左がTrueなら右を評価しない(ByVal Assert As UnitTestAssert)
    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("[A] == ""x"" OR [未入力] == ""y""", G_COLUMN_EXPRESSION_KIND_CONDITION, "OutputConditionExpression")

    Dim actual As Boolean

    On Error Resume Next
    actual = column_expression.EvaluateCondition(pColumnValues("A", "x"))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    On Error GoTo 0

    Assert.IsTrue actual
End Sub

Public Sub Test_ColumnExpression_文字列リテラルのエスケープを解釈する(ByVal Assert As UnitTestAssert)
    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("[件名] == ""A\""B\nC""", G_COLUMN_EXPRESSION_KIND_CONDITION, "OutputConditionExpression")

    Dim actual As Boolean

    On Error Resume Next
    actual = column_expression.EvaluateCondition(pColumnValues("件名", "A" & Chr$(34) & "B" & vbLf & "C"))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    On Error GoTo 0

    Assert.IsTrue actual
End Sub

Public Sub Test_ColumnExpression_比較時だけ前後空白を除去する(ByVal Assert As UnitTestAssert)
    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("[件名] == "" 案件A """, G_COLUMN_EXPRESSION_KIND_CONDITION, "OutputConditionExpression")

    Dim actual As Boolean

    On Error Resume Next
    actual = column_expression.EvaluateCondition(pColumnValues("件名", "  案件A  "))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    On Error GoTo 0

    Assert.IsTrue actual
End Sub

Public Sub Test_ColumnExpression_ValueExpression結果に追加Trimしない(ByVal Assert As UnitTestAssert)
    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("[件名]", G_COLUMN_EXPRESSION_KIND_VALUE, "ValueExpression(件名別名)")

    Dim actual As String

    On Error Resume Next
    actual = column_expression.EvaluateValue(pColumnValues("件名", "  案件A  "))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    On Error GoTo 0

    Assert.Equals "  案件A  ", actual
End Sub

Public Sub Test_ColumnExpression_未知escapeは初期化時にエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("[件名] == ""A\x""", G_COLUMN_EXPRESSION_KIND_CONDITION, "OutputConditionExpression")

    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue 0 < InStr(1, Err.Description, "未知", vbTextCompare)
End Sub

Public Sub Test_ColumnExpression_未定義列参照は評価時にエラー(ByVal Assert As UnitTestAssert)
    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("[未定義] == """"", G_COLUMN_EXPRESSION_KIND_CONDITION, "OutputConditionExpression")

    Dim column_values As ObjectDictionary
    Set column_values = New_ObjectDictionary()

    On Error Resume Next
    Call column_expression.EvaluateCondition(column_values)

    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue 0 < InStr(1, Err.Description, "未定義", vbTextCompare)
End Sub

Public Sub Test_ColumnExpression_OutputColumnName大小違いは別列として参照できる(ByVal Assert As UnitTestAssert)
    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("[Name] == ""UPPER"" AND [name] == ""LOWER""", G_COLUMN_EXPRESSION_KIND_CONDITION, "OutputConditionExpression")

    Dim actual As Boolean

    On Error Resume Next
    actual = column_expression.EvaluateCondition(pColumnValues("Name", "UPPER", "name", "LOWER"))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    On Error GoTo 0

    Assert.IsTrue actual
End Sub

Public Sub Test_ColumnExpression_LFを含む列名を実LFで参照できる(ByVal Assert As UnitTestAssert)
    Dim column_name As String
    column_name = "件名" & vbLf & "詳細"

    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("[" & column_name & "] == ""案件A""", G_COLUMN_EXPRESSION_KIND_CONDITION, "OutputConditionExpression")

    Dim actual As Boolean

    On Error Resume Next
    actual = column_expression.EvaluateCondition(pColumnValues(column_name, "案件A"))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    On Error GoTo 0

    Assert.IsTrue actual
End Sub

Public Sub Test_ColumnExpression_列参照内backslashnはLFとして扱わない(ByVal Assert As UnitTestAssert)
    Dim column_name As String
    column_name = "件名" & vbLf & "詳細"

    Dim column_expression As ColumnExpression
    Set column_expression = New_ColumnExpression("[件名\n詳細] == ""案件A""", G_COLUMN_EXPRESSION_KIND_CONDITION, "OutputConditionExpression")

    On Error Resume Next
    Call column_expression.EvaluateCondition(pColumnValues(column_name, "案件A"))

    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue 0 < InStr(1, Err.Description, "未定義", vbTextCompare)
End Sub

Private Function pColumnValues(ParamArray Pairs() As Variant) As ObjectDictionary
    Dim result_values As ObjectDictionary
    Set result_values = New_ObjectDictionary()

    Dim pair_idx As Long
    If LBound(Pairs) <= UBound(Pairs) Then
        For pair_idx = LBound(Pairs) To UBound(Pairs) Step 2
            Call result_values.Add(CStr(Pairs(pair_idx)), CStr(Pairs(pair_idx + 1)))
        Next pair_idx
    End If

    Set pColumnValues = result_values
End Function
