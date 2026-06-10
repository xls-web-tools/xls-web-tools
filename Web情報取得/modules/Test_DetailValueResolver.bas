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
