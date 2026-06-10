Attribute VB_Name = "Test_DetailValueResolver"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! DetailValueResolver のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_DetailValueResolver_通常抽出列だけをWebDriver抽出対象にする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject"))
    Call detail_defs.Add(New_DetailColumnDefinition("件名別名", "", ValueExpression:="[件名]"))

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver

    ' --- Act ---
    Call resolver.Initialize(detail_defs)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, resolver.ExtractionColumnDefinitions.Count
    Assert.Equals "件名", resolver.ExtractionColumnDefinitions.Item(0).OutputColumnName
End Sub

Public Sub Test_DetailValueResolver_単純列参照の派生値を解決する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject"))
    Call detail_defs.Add(New_DetailColumnDefinition("件名別名", "", ValueExpression:="[件名]"))

    Dim extraction_values As ArrayObject
    Set extraction_values = New ArrayObject
    Call extraction_values.ReDimArray(0, 0)
    Call extraction_values.Update(0, "案件A")

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver
    Call resolver.Initialize(detail_defs)

    Dim actual_values As ArrayObject

    ' --- Act ---
    Set actual_values = resolver.ResolveValues(extraction_values)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "案件A", CStr(actual_values.Item(1))
End Sub

Public Sub Test_DetailValueResolver_ValueExpression結果に追加Trimしない(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject"))
    Call detail_defs.Add(New_DetailColumnDefinition("件名別名", "", ValueExpression:="[件名]"))

    Dim extraction_values As ArrayObject
    Set extraction_values = New ArrayObject
    Call extraction_values.ReDimArray(0, 0)
    Call extraction_values.Update(0, "  案件A  ")

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver
    Call resolver.Initialize(detail_defs)

    Dim actual_values As ArrayObject

    ' --- Act ---
    Set actual_values = resolver.ResolveValues(extraction_values)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "  案件A  ", CStr(actual_values.Item(1))
End Sub

Public Sub Test_DetailValueResolver_IF値式で空欄時の代替値を返す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("列1", "#col1"))
    Call detail_defs.Add(New_DetailColumnDefinition("列2", "#col2"))
    Call detail_defs.Add(New_DetailColumnDefinition("採用列", "", ValueExpression:="IF([列2] == """", [列1], [列2])"))

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver
    Call resolver.Initialize(detail_defs)

    Dim actual_values As ArrayObject

    ' --- Act ---
    Set actual_values = resolver.ResolveValues(pValues("A", ""))

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "A", CStr(actual_values.Item(2))
End Sub

Public Sub Test_DetailValueResolver_IF条件で論理演算子とエスケープを使える(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))
    Call detail_defs.Add(New_DetailColumnDefinition("代理申請者", "#proxy"))
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#subject"))
    Call detail_defs.Add(New_DetailColumnDefinition("判定", "", ValueExpression:="IF(([申請者] == ""山田"" OR [代理申請者] == ""山田"") AND NOT [件名] == ""A\""B\nC"", ""対象\n行"", ""対象外"")"))

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver
    Call resolver.Initialize(detail_defs)

    Dim actual_values As ArrayObject

    ' --- Act ---
    Set actual_values = resolver.ResolveValues(pValues("", "山田", "その他"))

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "対象" & vbLf & "行", CStr(actual_values.Item(3))
End Sub

Public Sub Test_DetailValueResolver_派生列が他の派生列を参照できる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("元列", "#source"))
    Call detail_defs.Add(New_DetailColumnDefinition("別名", "", ValueExpression:="[元列]"))
    Call detail_defs.Add(New_DetailColumnDefinition("再別名", "", ValueExpression:="[別名]"))

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver
    Call resolver.Initialize(detail_defs)

    Dim actual_values As ArrayObject

    ' --- Act ---
    Set actual_values = resolver.ResolveValues(pValues("案件A"))

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "案件A", CStr(actual_values.Item(2))
End Sub

Public Sub Test_DetailValueResolver_派生列評価順は設定行順に依存しない(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("再別名", "", ValueExpression:="[別名]"))
    Call detail_defs.Add(New_DetailColumnDefinition("元列", "#source"))
    Call detail_defs.Add(New_DetailColumnDefinition("別名", "", ValueExpression:="[元列]"))

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver
    Call resolver.Initialize(detail_defs)

    Dim actual_values As ArrayObject

    ' --- Act ---
    Set actual_values = resolver.ResolveValues(pValues("案件A"))

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "案件A", CStr(actual_values.Item(0))
End Sub

Public Sub Test_DetailValueResolver_未対応関数はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("元列", "#source"))
    Call detail_defs.Add(New_DetailColumnDefinition("派生", "", ValueExpression:="COALESCE([元列], """")"))

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver

    ' --- Act ---
    Call resolver.Initialize(detail_defs)

    ' --- Assert ---
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue 0 < InStr(1, Err.Description, "ValueExpression", vbTextCompare)
End Sub

Public Sub Test_DetailValueResolver_文字列結合はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("元列", "#source"))
    Call detail_defs.Add(New_DetailColumnDefinition("派生", "", ValueExpression:="""A"" & ""B"""))

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver

    ' --- Act ---
    Call resolver.Initialize(detail_defs)

    ' --- Assert ---
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue 0 < InStr(1, Err.Description, "ValueExpression", vbTextCompare)
End Sub

Public Sub Test_DetailValueResolver_未定義列参照はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("元列", "#source"))
    Call detail_defs.Add(New_DetailColumnDefinition("派生", "", ValueExpression:="[未定義]"))

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver

    ' --- Act ---
    Call resolver.Initialize(detail_defs)

    ' --- Assert ---
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue 0 < InStr(1, Err.Description, "未定義", vbTextCompare)
End Sub

Public Sub Test_DetailValueResolver_自己参照はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("自己", "", ValueExpression:="[自己]"))

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver

    ' --- Act ---
    Call resolver.Initialize(detail_defs)

    ' --- Assert ---
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue 0 < InStr(1, Err.Description, "自己参照", vbTextCompare)
End Sub

Public Sub Test_DetailValueResolver_派生列同士の循環参照はエラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("A", "", ValueExpression:="[B]"))
    Call detail_defs.Add(New_DetailColumnDefinition("B", "", ValueExpression:="[A]"))

    Dim resolver As DetailValueResolver
    Set resolver = New DetailValueResolver

    ' --- Act ---
    Call resolver.Initialize(detail_defs)

    ' --- Assert ---
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue 0 < InStr(1, Err.Description, "循環", vbTextCompare)
End Sub

Private Function pValues(ParamArray Values() As Variant) As ArrayObject
    Dim result_values As ArrayObject
    Set result_values = New ArrayObject

    If LBound(Values) <= UBound(Values) Then
        Call result_values.ReDimArray(0, UBound(Values) - LBound(Values))
    End If

    Dim value_idx As Long
    For value_idx = LBound(Values) To UBound(Values)
        Call result_values.Update(value_idx - LBound(Values), Values(value_idx))
    Next value_idx

    Set pValues = result_values
End Function
