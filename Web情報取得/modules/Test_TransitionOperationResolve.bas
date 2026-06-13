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
    Dim actual_chain As TransitionOperationChain
    Set actual_chain = resolver.ResolveListTransitionOperation()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, actual_chain.Count

    Dim actual_operation As TransitionOperation
    Set actual_operation = actual_chain.Item(0)
    Assert.Equals "OpenList", actual_operation.OperationName
    Assert.Equals "#open-list", actual_operation.ActionSelector
    Assert.Equals "#list-ready", actual_operation.WaitSelector
End Sub

Public Sub Test_TransitionOperationResolver_一覧遷移操作チェーンを順に解決する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.ListTransitionOperationName = "OpenOther >> OpenList"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenOther", "#open-other", WaitSelector:="#other-ready"))
    Call operations.Add(New_TransitionOperation("OpenList", "#open-list", WaitSelector:="#list-ready"))
    Set tool_settings.TransitionOperations = operations

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Dim actual_chain As TransitionOperationChain
    Set actual_chain = resolver.ResolveListTransitionOperation()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_chain.Count
    Assert.Equals "OpenOther", actual_chain.Item(0).OperationName
    Assert.Equals "#other-ready", actual_chain.Item(0).WaitSelector
    Assert.Equals "OpenList", actual_chain.Item(1).OperationName
    Assert.Equals "#list-ready", actual_chain.Item(1).WaitSelector
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
    Dim actual_chain As TransitionOperationChain
    Set actual_chain = resolver.ResolveDetailTransitionOperation(2)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, actual_chain.Count

    Dim actual_operation As TransitionOperation
    Set actual_operation = actual_chain.Item(0)
    Assert.Equals "OpenDetail", actual_operation.OperationName
    Assert.Equals "#list tr:nth-child(3) a", actual_operation.ActionSelector
    Assert.Equals "openDetail(2)", actual_operation.ActionScript
    Assert.Equals "#detail-3", actual_operation.WaitSelector
End Sub

Public Sub Test_TransitionOperationResolver_詳細遷移操作チェーンの全要素にIndexテンプレートを適用する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.DetailTransitionOperationName = "OpenMenu >> OpenDetail"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation( _
            "OpenMenu", _
            "#list tr:nth-child({{rowNumber}}) .menu", _
            WaitSelector:="#menu-{{index}}"))
    Call operations.Add(New_TransitionOperation( _
            "OpenDetail", _
            "#list tr:nth-child({{rowNumber}}) a", _
            ActionScript:="openDetail({{index}})", _
            WaitSelector:="#detail-{{rowNumber}}"))
    Set tool_settings.TransitionOperations = operations

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Dim actual_chain As TransitionOperationChain
    Set actual_chain = resolver.ResolveDetailTransitionOperation(2)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_chain.Count
    Assert.Equals "#list tr:nth-child(3) .menu", actual_chain.Item(0).ActionSelector
    Assert.Equals "#menu-2", actual_chain.Item(0).WaitSelector
    Assert.Equals "#list tr:nth-child(3) a", actual_chain.Item(1).ActionSelector
    Assert.Equals "openDetail(2)", actual_chain.Item(1).ActionScript
    Assert.Equals "#detail-3", actual_chain.Item(1).WaitSelector
End Sub

Public Sub Test_TransitionOperationResolver_復帰と次ページの遷移操作チェーンを解決する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.ReturnToListOperationName = "CloseDetail >> BackToList"
    tool_settings.NextPageOperationName = "OpenPager >> NextPage"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("CloseDetail", "#close-detail", WaitSelector:="#closed"))
    Call operations.Add(New_TransitionOperation("BackToList", "#back-list", WaitSelector:="#list-ready"))
    Call operations.Add(New_TransitionOperation("OpenPager", "#open-pager", WaitSelector:="#pager-ready"))
    Call operations.Add(New_TransitionOperation("NextPage", "#next-page", WaitSelector:="#next-ready"))
    Set tool_settings.TransitionOperations = operations

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Dim return_chain As TransitionOperationChain
    Set return_chain = resolver.ResolveReturnToListOperation()

    Dim next_chain As TransitionOperationChain
    Set next_chain = resolver.ResolveNextPageOperation()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, return_chain.Count
    Assert.Equals "CloseDetail", return_chain.Item(0).OperationName
    Assert.Equals "BackToList", return_chain.Item(1).OperationName
    Assert.EqualsNumeric 2, next_chain.Count
    Assert.Equals "OpenPager", next_chain.Item(0).OperationName
    Assert.Equals "NextPage", next_chain.Item(1).OperationName
End Sub

Public Sub Test_TransitionOperationResolver_同じ操作名をチェーン内で複数回使える(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.ListTransitionOperationName = "OpenList >> OpenList"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenList", "#open-list", WaitSelector:="#list-ready"))
    Set tool_settings.TransitionOperations = operations

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Dim actual_chain As TransitionOperationChain
    Set actual_chain = resolver.ResolveListTransitionOperation()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_chain.Count
    Assert.Equals "OpenList", actual_chain.Item(0).OperationName
    Assert.Equals "OpenList", actual_chain.Item(1).OperationName
End Sub

Public Sub Test_TransitionOperationResolver_空要素なら設定項目名とチェーン文字列と要素名を含めてエラーにする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.ListTransitionOperationName = "OpenOther >>  >> OpenList"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("OpenOther", "#open-other", WaitSelector:="#other-ready"))
    Call operations.Add(New_TransitionOperation("OpenList", "#open-list", WaitSelector:="#list-ready"))
    Set tool_settings.TransitionOperations = operations

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Call resolver.ResolveListTransitionOperation

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "ListTransitionOperationName", vbTextCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "OpenOther >>  >> OpenList", vbTextCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "ElementName=", vbTextCompare)
End Sub

Public Sub Test_TransitionOperationResolver_区切り文字を含む操作名なら設定項目名とチェーン文字列と要素名を含めてエラーにする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.ListTransitionOperationName = "Open >> List"

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")
    Call operations.Add(New_TransitionOperation("Open >> List", "#bad-name", WaitSelector:="#bad-ready"))
    Set tool_settings.TransitionOperations = operations

    Dim resolver As TransitionOperationResolver
    Set resolver = New_TransitionOperationResolver(tool_settings)

    ' --- Act ---
    Call resolver.ResolveListTransitionOperation

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.IsTrue 0 < InStr(1, Err.Description, "ListTransitionOperationName", vbTextCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "Open >> List", vbTextCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "ElementName=Open >> List", vbTextCompare)
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
    Assert.IsTrue 0 < InStr(1, Err.Description, "NextPageOperationName", vbTextCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "OperationChain=NextPage", vbTextCompare)
    Assert.IsTrue 0 < InStr(1, Err.Description, "ElementName=NextPage", vbTextCompare)
End Sub
