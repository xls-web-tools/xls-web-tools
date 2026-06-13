Attribute VB_Name = "Test_TransitionOperationResolve"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 画面遷移操作 resolver のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_TransitionOperationResolver_一覧遷移操作名を大文字小文字を区別せずに解決する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.ListTransitionOperationName = "openlist"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "#open-list", WaitSelector:="#list-ready"))
    Set tool_settings.TransitionOperations = operations

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Dim actual_operation As TransitionOperation
    Set actual_operation = resolver.ResolveListTransitionOperation()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "OpenList", actual_operation.OperationName
    Assert.Equals "#open-list", actual_operation.ActionSelector
    Assert.Equals "#list-ready", actual_operation.WaitSelector
End Sub

Public Sub Test_TransitionOperationResolver_詳細遷移操作ではIndexテンプレートを置換する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.DetailTransitionOperationName = "OpenDetail"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation( _
            "OpenDetail", _
            "#list tr:nth-child({{rowNumber}}) a", _
            ActionScript:="openDetail({{index}})", _
            WaitSelector:="#detail-{{rowNumber}}"))
    Set tool_settings.TransitionOperations = operations

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Dim actual_operation As TransitionOperation
    Set actual_operation = resolver.ResolveDetailTransitionOperation(2)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "OpenDetail", actual_operation.OperationName
    Assert.Equals "#list tr:nth-child(3) a", actual_operation.ActionSelector
    Assert.Equals "openDetail(2)", actual_operation.ActionScript
    Assert.Equals "#detail-3", actual_operation.WaitSelector
End Sub

Public Sub Test_TransitionOperationResolver_未定義の操作名ならエラーにする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.NextPageOperationName = "NextPage"
    Set tool_settings.TransitionOperations = New_ObjectList("TransitionOperation")

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Call resolver.ResolveNextPageOperation

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "TransitionOperations", vbTextCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "NextPage", vbTextCompare)
End Sub
