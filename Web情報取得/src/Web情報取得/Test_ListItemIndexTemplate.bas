Attribute VB_Name = "Test_ListItemIndexTemplate"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 一覧項目 index テンプレートのユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_ListItemIndexTemplate_IndexとRowNumberを置換する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim index_template As ListItemIndexTemplate
    Set index_template = New_ListItemIndexTemplate()

    ' --- Act ---
    Dim actual_value As String
    actual_value = index_template.Apply("#list tr:nth-child({{rowNumber}}) a[data-index='{{index}}']", 2)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "#list tr:nth-child(3) a[data-index='2']", actual_value
End Sub

Public Sub Test_ListItemIndexTemplate_置換トークンがなければ元の文字列を返す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim index_template As ListItemIndexTemplate
    Set index_template = New_ListItemIndexTemplate()

    ' --- Act ---
    Dim actual_value As String
    actual_value = index_template.Apply("#target-id", 4)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "#target-id", actual_value
End Sub
