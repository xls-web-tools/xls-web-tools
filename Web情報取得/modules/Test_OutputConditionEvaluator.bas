Attribute VB_Name = "Test_OutputConditionEvaluator"
Option Explicit
Option Base 0

Public Sub Test_OutputConditionEvaluator_ANDOR条件を評価できる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))
    Call detail_defs.Add(New_DetailColumnDefinition("代理申請者", "#proxy"))
    Call detail_defs.Add(New_DetailColumnDefinition("処理済み日", "#processed"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("([申請者] == ""山田"" OR [代理申請者] == ""山田"") AND [処理済み日] != ""済""", detail_defs)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 2)
    Call detail_values.Update(0, "")
    Call detail_values.Update(1, "山田")
    Call detail_values.Update(2, "")

    Dim actual As Boolean
    actual = evaluator.ShouldOutput(detail_values)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(actual, "条件に一致する詳細値は出力対象になる")
End Sub

Public Sub Test_OutputConditionEvaluator_条件不一致はFalseを返す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[申請者] == ""山田""", detail_defs)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 0)
    Call detail_values.Update(0, "佐藤")

    Dim actual As Boolean
    actual = evaluator.ShouldOutput(detail_values)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsFalse(actual, "条件に一致しない詳細値は出力対象外になる")
End Sub

Public Sub Test_OutputConditionEvaluator_output未指定の抽出列も条件参照できる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("判定", "#decision"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[判定] == ""対象""", detail_defs)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 0)
    Call detail_values.Update(0, "対象")

    Dim actual As Boolean
    actual = evaluator.ShouldOutput(detail_values)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(actual, "output ヘッダーにない抽出列の条件一致で出力対象になる")
End Sub

Public Sub Test_OutputConditionEvaluator_未定義列参照はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[未定義] == ""山田""", detail_defs)

    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(0 < InStr(1, Err.Description, "未定義", vbTextCompare), "未定義の OutputColumnName がエラー説明に含まれる")
End Sub

Public Sub Test_OutputConditionEvaluator_OutputColumnName重複はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester2"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("", detail_defs)

    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(0 < InStr(1, Err.Description, "重複", vbTextCompare), "重複した OutputColumnName がエラーになる")
End Sub

Public Sub Test_OutputConditionEvaluator_OutputColumnName大小違いは別列として参照できる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("Name", "#upper-name"))
    Call detail_defs.Add(New_DetailColumnDefinition("name", "#lower-name"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[Name] == ""UPPER"" AND [name] == ""LOWER""", detail_defs)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 1)
    Call detail_values.Update(0, "UPPER")
    Call detail_values.Update(1, "LOWER")

    Dim actual As Boolean
    actual = evaluator.ShouldOutput(detail_values)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(actual, "大文字小文字だけ異なる OutputColumnName を別の抽出列として条件参照できる")
End Sub

Public Sub Test_OutputConditionEvaluator_LFを含むOutputColumnNameを実LFで参照できる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim column_name As String
    column_name = "件名" & vbLf & "詳細"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition(column_name, "#subject"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[" & column_name & "] == ""対象""", detail_defs)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 0)
    Call detail_values.Update(0, "対象")

    Dim actual As Boolean
    actual = evaluator.ShouldOutput(detail_values)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(actual, "LF を含む OutputColumnName を実 LF の列参照で条件評価できる")
End Sub

Public Sub Test_OutputConditionEvaluator_列参照内backslashnはLFとして扱わない(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名" & vbLf & "詳細", "#subject"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[件名\n詳細] == ""対象""", detail_defs)

    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(0 < InStr(1, Err.Description, "未定義", vbTextCompare), "列参照内の \n は実 LF に変換されない")
End Sub

Public Sub Test_OutputConditionEvaluator_文字列リテラル内backslashnはLFとして扱う(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[件名] == ""A\nB""", detail_defs)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 0)
    Call detail_values.Update(0, "A" & vbLf & "B")

    Dim actual As Boolean
    actual = evaluator.ShouldOutput(detail_values)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(actual, "文字列リテラル内の \n は LF として比較される")
End Sub

Public Sub Test_OutputConditionEvaluator_LF違いは別列として参照できる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim first_column_name As String
    first_column_name = "判定" & vbLf & "A"

    Dim second_column_name As String
    second_column_name = "判定" & vbLf & "B"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition(first_column_name, "#decision-a"))
    Call detail_defs.Add(New_DetailColumnDefinition(second_column_name, "#decision-b"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[" & second_column_name & "] == ""対象""", detail_defs)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 1)
    Call detail_values.Update(0, "対象外")
    Call detail_values.Update(1, "対象")

    Dim actual As Boolean
    actual = evaluator.ShouldOutput(detail_values)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(actual, "LF で異なる OutputColumnName は別列として条件参照できる")
End Sub

Public Sub Test_OutputConditionEvaluator_固定管理列名と同名の抽出列を条件参照できる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition(G_WEB_OUTPUT_COL_TARGET_ID, "#detail-target-id"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[" & G_WEB_OUTPUT_COL_TARGET_ID & "] == ""DETAIL-001""", detail_defs)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 0)
    Call detail_values.Update(0, "DETAIL-001")

    Dim actual As Boolean
    actual = evaluator.ShouldOutput(detail_values)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(actual, "固定管理列名と同名の抽出列を条件式から参照する")
End Sub

Public Sub Test_OutputConditionEvaluator_文字列リテラルのエスケープを解釈する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator

    Dim expression_text As String
    expression_text = "[件名] == " & Chr$(34) & "A\" & Chr$(34) & "B\nC" & Chr$(34)
    Call evaluator.Initialize(expression_text, detail_defs)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 0)
    Call detail_values.Update(0, "A" & Chr$(34) & "B" & vbLf & "C")

    Dim actual As Boolean
    actual = evaluator.ShouldOutput(detail_values)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(actual, "エスケープ文字列が評価される")
End Sub

Public Sub Test_OutputConditionEvaluator_リテラルの前後空白を除去して比較する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[申請者] == "" 山田 """, detail_defs)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 0)
    Call detail_values.Update(0, "山田")

    Dim actual As Boolean
    actual = evaluator.ShouldOutput(detail_values)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(actual, "リテラルの前後空白は比較前に除去される")
End Sub

Public Sub Test_OutputConditionEvaluator_未知escapeはエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator

    Dim expression_text As String
    expression_text = "[件名] == " & Chr$(34) & "A\q" & Chr$(34)
    Call evaluator.Initialize(expression_text, detail_defs)

    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(0 < InStr(1, Err.Description, "未知", vbTextCompare), "未知 escape がエラー説明に含まれる")
End Sub

Public Sub Test_OutputConditionEvaluator_条件参照列は任意列として読み取る(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("判定", "#decision", IsRequired:=True, BlankMode:="ErrorIfBlank"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[判定] == """"", detail_defs)

    Dim referenced_defs As ObjectList
    Set referenced_defs = evaluator.ReferencedColumnDefinitions

    Dim referenced_def As DetailColumnDefinition
    Set referenced_def = referenced_defs.Item(0)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, referenced_defs.Count
    Assert.Equals "判定", referenced_def.OutputColumnName
    Call Assert.IsFalse(referenced_def.IsRequired, "条件参照列は selector 欠落を空文字列扱いできるよう任意列として読む")
    Assert.Equals "AllowBlank", referenced_def.BlankMode
End Sub

Public Sub Test_OutputConditionEvaluator_派生列を条件参照列として返す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("元列", "#source"))
    Call detail_defs.Add(New_DetailColumnDefinition("派生", "", ValueExpression:="[元列]"))

    Dim evaluator As OutputConditionEvaluator
    Set evaluator = New OutputConditionEvaluator
    Call evaluator.Initialize("[派生] == ""対象""", detail_defs)

    Dim referenced_defs As ObjectList
    Set referenced_defs = evaluator.ReferencedColumnDefinitions

    Dim referenced_def As DetailColumnDefinition
    Set referenced_def = referenced_defs.Item(0)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, referenced_defs.Count
    Assert.Equals "派生", referenced_def.OutputColumnName
    Assert.IsTrue referenced_def.IsDerived
    Assert.Equals "[元列]", referenced_def.ValueExpression
    Assert.Equals "AllowBlank", referenced_def.BlankMode
End Sub
