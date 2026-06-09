Attribute VB_Name = "Test_OutputSheetWriter"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! OutputSheetWriter のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_OutputSheetWriter_固定管理列と詳細列のヘッダーを用意する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub
    Call pSetEmptyOutput(ws_stub)

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title", OutputEnabled:=True))
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.EnsureHeaders

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub

    Dim actual_values As Variant
    actual_values = pGetLatestWrittenRangeValues(ws_stub)

    Assert.Equals G_WEB_OUTPUT_COL_TARGET_ID, CStr(actual_values(1, 1))
    Assert.Equals G_WEB_OUTPUT_COL_STATUS, CStr(actual_values(1, 2))
    Assert.Equals G_WEB_OUTPUT_COL_ERROR, CStr(actual_values(1, 3))
    Assert.Equals G_WEB_OUTPUT_COL_DOWNLOAD_STATUS, CStr(actual_values(1, 4))
    Assert.Equals "件名", CStr(actual_values(1, 5))
    Assert.EqualsNumeric 5, UBound(actual_values, 2)
End Sub

Public Sub Test_OutputSheetWriter_既存outputを現行ヘッダーへ正規化する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title", OutputEnabled:=True))
    Call detail_defs.Add(New_DetailColumnDefinition("判定", "#decision"))
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester", OutputEnabled:=True))
    Call detail_defs.Add(New_DetailColumnDefinition("部署", "#department", OutputEnabled:=True))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim existing_values(1 To 2, 1 To 8) As Variant
    existing_values(1, 1) = "件名"
    existing_values(1, 2) = G_WEB_OUTPUT_COL_TARGET_ID
    existing_values(1, 3) = "詳細ページID"
    existing_values(1, 4) = G_WEB_OUTPUT_COL_STATUS
    existing_values(1, 5) = "旧列"
    existing_values(1, 6) = "旧列"
    existing_values(1, 7) = "申請者"
    existing_values(1, 8) = G_WEB_OUTPUT_COL_ERROR
    existing_values(2, 1) = "案件A"
    existing_values(2, 2) = "T-001"
    existing_values(2, 3) = "D-001"
    existing_values(2, 4) = G_WEB_STATUS_OK
    existing_values(2, 5) = "old1"
    existing_values(2, 6) = "old2"
    existing_values(2, 7) = "山田"
    existing_values(2, 8) = ""

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = pSetExistingOutputValues(ws_stub, existing_values)

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.EnsureHeaders

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, ws_stub.Store.GetCallCount("ClearRange", used_bounds)

    Dim actual_values As Variant
    actual_values = pGetLatestWrittenRangeValues(ws_stub)

    Assert.Equals G_WEB_OUTPUT_COL_TARGET_ID, CStr(actual_values(1, 1))
    Assert.Equals G_WEB_OUTPUT_COL_STATUS, CStr(actual_values(1, 2))
    Assert.Equals G_WEB_OUTPUT_COL_ERROR, CStr(actual_values(1, 3))
    Assert.Equals G_WEB_OUTPUT_COL_DOWNLOAD_STATUS, CStr(actual_values(1, 4))
    Assert.Equals "件名", CStr(actual_values(1, 5))
    Assert.Equals "申請者", CStr(actual_values(1, 6))
    Assert.Equals "部署", CStr(actual_values(1, 7))

    Assert.Equals "T-001", CStr(actual_values(2, 1))
    Assert.Equals G_WEB_STATUS_OK, CStr(actual_values(2, 2))
    Assert.Equals "", CStr(actual_values(2, 3))
    Assert.Equals "", CStr(actual_values(2, 4))
    Assert.Equals "案件A", CStr(actual_values(2, 5))
    Assert.Equals "山田", CStr(actual_values(2, 6))
    Assert.Equals "", CStr(actual_values(2, 7))
    Assert.EqualsNumeric 7, UBound(actual_values, 2)
End Sub

Public Sub Test_OutputSheetWriter_現行ヘッダー重複は出力シート不整合エラー(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title", OutputEnabled:=True))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim existing_values(1 To 1, 1 To 4) As Variant
    existing_values(1, 1) = G_WEB_OUTPUT_COL_TARGET_ID
    existing_values(1, 2) = "旧列"
    existing_values(1, 3) = "旧列"
    existing_values(1, 4) = G_WEB_OUTPUT_COL_TARGET_ID
    Call pSetExistingOutputValues(ws_stub, existing_values)

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.EnsureHeaders

    ' --- Assert ---
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call Assert.IsTrue(0 < InStr(1, Err.Description, "出力シート不整合", vbTextCompare), "現行ヘッダー重複が出力シート不整合として扱われる")
End Sub

Public Sub Test_OutputSheetWriter_診断出力は対象ID一致行を上書きする(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title", OutputEnabled:=True))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=1, Sheet:="result")

    Dim found_rows As ObjectList
    Set found_rows = New_ObjectList("WorksheetRangeBounds")
    Call found_rows.Add(New_RangeBounds(Row:=5, Column:=1, Sheet:="result"))
    Call ws_stub.Store.SetReturn("Find", found_rows, "T-001", search_bounds, True, True, True, True)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 0)
    Call detail_values.Update(0, "案件A")

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.WriteDiagnosticRow("T-001", G_WEB_STATUS_OK, "", detail_values, G_WEB_DOWNLOAD_STATUS_NO_FILE)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 5, 1, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 2, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 4, G_WEB_DOWNLOAD_STATUS_NO_FILE)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 5, "案件A")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=2, Column:=1, Sheet:="result"))
End Sub

Public Sub Test_OutputSheetWriter_既存行扱いを判定する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act & Assert ---
    tool_settings.ExistingRowMode = G_WEB_ROW_MODE_SKIP_EXISTING
    Assert.IsFalse writer.ShouldWriteExistingRow(G_WEB_STATUS_OK)
    Assert.IsTrue writer.ShouldWriteExistingRow(G_WEB_STATUS_ERROR)
    Assert.IsTrue writer.ShouldWriteExistingRow(G_WEB_STATUS_OK, IsDiagnosticMode:=True)

    Call writer.ShouldWriteExistingRow("UNKNOWN")
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.Equals "既存行の取得状態が不正です。(UNKNOWN)", Err.Description
    Err.Clear

    tool_settings.ExistingRowMode = G_WEB_ROW_MODE_OVERWRITE
    Assert.IsTrue writer.ShouldWriteExistingRow(G_WEB_STATUS_OK)
    Assert.IsTrue writer.ShouldWriteExistingRow(G_WEB_STATUS_ERROR)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_OutputSheetWriter_本番収集対象判定は対象ID主キーとExistingRowModeに従う(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=1, Sheet:="result")

    Dim ok_rows As ObjectList
    Set ok_rows = New_ObjectList("WorksheetRangeBounds")
    Call ok_rows.Add(New_RangeBounds(Row:=5, Column:=1, Sheet:="result"))
    Call ws_stub.Store.SetReturn("Find", ok_rows, "T-OK", search_bounds, True, True, True, True)
    Call ws_stub.Store.SetReturn("ReadCell", G_WEB_STATUS_OK, New_RangeBounds(Row:=5, Column:=2, Sheet:="result"), False)

    Dim error_rows As ObjectList
    Set error_rows = New_ObjectList("WorksheetRangeBounds")
    Call error_rows.Add(New_RangeBounds(Row:=6, Column:=1, Sheet:="result"))
    Call ws_stub.Store.SetReturn("Find", error_rows, "T-ERROR", search_bounds, True, True, True, True)
    Call ws_stub.Store.SetReturn("ReadCell", G_WEB_STATUS_ERROR, New_RangeBounds(Row:=6, Column:=2, Sheet:="result"), False)

    Dim missing_rows As ObjectList
    Set missing_rows = New_ObjectList("WorksheetRangeBounds")
    Call ws_stub.Store.SetReturn("Find", missing_rows, "T-NEW", search_bounds, True, True, True, True)

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act & Assert ---
    tool_settings.ExistingRowMode = G_WEB_ROW_MODE_SKIP_EXISTING
    Assert.IsFalse writer.ShouldCollectTarget("T-OK")
    Assert.IsTrue writer.ShouldCollectTarget("T-ERROR")
    Assert.IsTrue writer.ShouldCollectTarget("T-NEW")

    tool_settings.ExistingRowMode = G_WEB_ROW_MODE_OVERWRITE
    Assert.IsTrue writer.ShouldCollectTarget("T-OK")
    Assert.IsTrue writer.ShouldCollectTarget("T-ERROR")
    Assert.IsTrue writer.ShouldCollectTarget("T-NEW")

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_OutputSheetWriter_本番出力は対象ID一致行を更新し未登録なら末尾へ追加する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title", OutputEnabled:=True))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=1, Sheet:="result")

    Dim existing_rows As ObjectList
    Set existing_rows = New_ObjectList("WorksheetRangeBounds")
    Call existing_rows.Add(New_RangeBounds(Row:=5, Column:=1, Sheet:="result"))
    Call ws_stub.Store.SetReturn("Find", existing_rows, "T-001", search_bounds, True, True, True, True)

    Dim missing_rows As ObjectList
    Set missing_rows = New_ObjectList("WorksheetRangeBounds")
    Call ws_stub.Store.SetReturn("Find", missing_rows, "T-002", search_bounds, True, True, True, True)

    Dim used_search_bounds As WorksheetRangeBounds
    Set used_search_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=5, Sheet:="result")

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=6, FinishColumn:=5, Sheet:="result")
    Call ws_stub.Store.SetReturn("GetUsedRangeBounds", used_bounds, used_search_bounds, True, True, True, False)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 0)
    Call detail_values.Update(0, "案件A")

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.WriteCollectionRow("T-001", G_WEB_STATUS_OK, "", detail_values, G_WEB_DOWNLOAD_STATUS_DOWNLOADED)
    Call writer.WriteCollectionRow("T-002", G_WEB_STATUS_ERROR, "missing", detail_values, G_WEB_DOWNLOAD_STATUS_ERROR)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 5, 1, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 2, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 4, G_WEB_DOWNLOAD_STATUS_DOWNLOADED)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 5, "案件A")
    Call pAssertWrittenCell(Assert, ws_stub, 7, 1, "T-002")
    Call pAssertWrittenCell(Assert, ws_stub, 7, 2, G_WEB_STATUS_ERROR)
    Call pAssertWrittenCell(Assert, ws_stub, 7, 3, "missing")
    Call pAssertWrittenCell(Assert, ws_stub, 7, 4, G_WEB_DOWNLOAD_STATUS_ERROR)
    Call pAssertWrittenCell(Assert, ws_stub, 7, 5, "案件A")
End Sub

Public Sub Test_OutputSheetWriter_全詳細列非出力なら固定管理列だけを書き込む(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title"))
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=1, Sheet:="result")

    Dim missing_rows As ObjectList
    Set missing_rows = New_ObjectList("WorksheetRangeBounds")
    Call ws_stub.Store.SetReturn("Find", missing_rows, "T-001", search_bounds, True, True, True, True)

    Dim used_search_bounds As WorksheetRangeBounds
    Set used_search_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=4, Sheet:="result")

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=1, FinishColumn:=4, Sheet:="result")
    Call ws_stub.Store.SetReturn("GetUsedRangeBounds", used_bounds, used_search_bounds, True, True, True, False)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 1)
    Call detail_values.Update(0, "案件A")
    Call detail_values.Update(1, "山田太郎")

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.WriteCollectionRow("T-001", G_WEB_STATUS_OK, "", detail_values, G_WEB_DOWNLOAD_STATUS_NO_FILE)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 2, 1, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 2, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 2, 3, "")
    Call pAssertWrittenCell(Assert, ws_stub, 2, 4, G_WEB_DOWNLOAD_STATUS_NO_FILE)
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=2, Column:=5, Sheet:="result"))
End Sub

Private Sub pSetEmptyOutput(ByVal WsStub As WorksheetServiceTestDouble)
    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = pOutputUsedSearchBounds()

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=0, FinishColumn:=0, Sheet:="result")
    Call WsStub.Store.SetReturn("GetUsedRangeBounds", used_bounds, search_bounds, True, True, True, False)
End Sub

Private Function pSetExistingOutputValues(ByVal WsStub As WorksheetServiceTestDouble, ByRef ExistingValues As Variant) As WorksheetRangeBounds
    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = pOutputUsedSearchBounds()

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = New_RangeBounds( _
            Row:=1, _
            Column:=1, _
            FinishRow:=UBound(ExistingValues, 1), _
            FinishColumn:=UBound(ExistingValues, 2), _
            Sheet:="result")
    Call WsStub.Store.SetReturn("GetUsedRangeBounds", used_bounds, search_bounds, True, True, True, False)
    Call WsStub.Store.SetReturn("ReadRange", ExistingValues, used_bounds)

    Set pSetExistingOutputValues = used_bounds
End Function

Private Function pOutputUsedSearchBounds() As WorksheetRangeBounds
    Set pOutputUsedSearchBounds = New_RangeBounds( _
            Row:=1, _
            Column:=1, _
            FinishRow:=G_ROW_MAX, _
            FinishColumn:=G_COL_MAX, _
            Sheet:="result")
End Function

Private Function pGetLatestWrittenRangeValues(ByVal WsStub As WorksheetServiceTestDouble) As Variant
    Dim target_bounds As WorksheetRangeBounds
    Set target_bounds = New_RangeBounds(Row:=1, Column:=1, Sheet:="result")

    Dim call_record As TestDoubleCallRecord
    Set call_record = WsStub.Store.GetLatestCall("WriteRange", target_bounds)

    pGetLatestWrittenRangeValues = call_record.GetArgument(1)
End Function

Private Sub pAssertHeader(ByVal Assert As UnitTestAssert, ByVal WsStub As WorksheetServiceTestDouble, ByVal ColumnIndex As Long, ByVal ExpectedHeader As String)
    Dim target_bounds As WorksheetRangeBounds
    Set target_bounds = New_RangeBounds(Row:=1, Column:=ColumnIndex, Sheet:="result")

    Dim call_record As TestDoubleCallRecord
    Set call_record = WsStub.Store.GetLatestCall("WriteCell", target_bounds)

    Assert.Equals ExpectedHeader, CStr(call_record.GetArgument(1)), CaseName:="Column " & CStr(ColumnIndex)
End Sub

Private Sub pAssertWrittenCell( _
        ByVal Assert As UnitTestAssert, _
        ByVal WsStub As WorksheetServiceTestDouble, _
        ByVal RowIndex As Long, _
        ByVal ColumnIndex As Long, _
        ByVal ExpectedValue As String)

    Dim target_bounds As WorksheetRangeBounds
    Set target_bounds = New_RangeBounds(Row:=RowIndex, Column:=ColumnIndex, Sheet:="result")

    Assert.EqualsNumeric 1, WsStub.Store.GetCallCount("WriteCell", target_bounds), CaseName:="Row " & CStr(RowIndex) & ", Column " & CStr(ColumnIndex) & " call count"

    Dim call_record As TestDoubleCallRecord
    Set call_record = WsStub.Store.GetLatestCall("WriteCell", target_bounds)

    Assert.Equals ExpectedValue, CStr(call_record.GetArgument(1)), CaseName:="Row " & CStr(RowIndex) & ", Column " & CStr(ColumnIndex)
End Sub
