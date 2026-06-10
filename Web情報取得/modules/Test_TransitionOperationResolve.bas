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

Public Sub Test_TransitionOperationResolver_一覧遷移操作名を大文字小文字非区別で解決する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.ListTransitionOperationName = "openlist"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "css selector", "#open-list", WaitConditionName:="ListReady"))
    Set tool_settings.TransitionOperations = operations

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Dim actual_operation As TransitionOperation
    Set actual_operation = resolver.ResolveListTransitionOperation()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "OpenList", actual_operation.OperationName
    Assert.Equals "#open-list", actual_operation.LocatorValue
    Assert.Equals "ListReady", actual_operation.WaitConditionName
End Sub

Public Sub Test_TransitionOperationResolver_詳細遷移操作のIndexテンプレートを解決する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.DetailTransitionOperationName = "OpenDetail"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation( _
            "OpenDetail", _
            "css selector", _
            "#list tr:nth-child({{rowNumber}}) a", _
            Script:="openDetail({{index}})", _
            WaitConditionName:="DetailReady"))
    Set tool_settings.TransitionOperations = operations

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Dim actual_operation As TransitionOperation
    Set actual_operation = resolver.ResolveDetailTransitionOperation(2)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "OpenDetail", actual_operation.OperationName
    Assert.Equals "css selector", actual_operation.LocatorType
    Assert.Equals "#list tr:nth-child(3) a", actual_operation.LocatorValue
    Assert.Equals "openDetail(2)", actual_operation.Script
    Assert.Equals "DetailReady", actual_operation.WaitConditionName
End Sub

Public Sub Test_TransitionOperationResolver_復帰と詳細の有無を返す(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.DetailTransitionOperationName = ""
    tool_settings.ReturnToListOperationName = "ReturnToList"

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Dim actual_has_detail As Boolean
    actual_has_detail = resolver.HasDetailTransitionOperation

    Dim actual_has_return As Boolean
    actual_has_return = resolver.HasReturnToListOperation

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_has_detail
    Assert.IsTrue actual_has_return
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
    Assert.IsTrue 0 < InStr(1, Err.Description, "画面遷移操作が見つかりません", vbTextCompare)
End Sub
