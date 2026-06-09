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

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title"))
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.EnsureHeaders

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertHeader(Assert, ws_stub, 1, G_WEB_OUTPUT_COL_DETAIL_PAGE_ID)
    Call pAssertHeader(Assert, ws_stub, 2, G_WEB_OUTPUT_COL_TARGET_ID)
    Call pAssertHeader(Assert, ws_stub, 3, G_WEB_OUTPUT_COL_STATUS)
    Call pAssertHeader(Assert, ws_stub, 4, G_WEB_OUTPUT_COL_ERROR)
    Call pAssertHeader(Assert, ws_stub, 5, G_WEB_OUTPUT_COL_DOWNLOAD_STATUS)
    Call pAssertHeader(Assert, ws_stub, 6, "件名")
    Call pAssertHeader(Assert, ws_stub, 7, "申請者")
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

    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=G_ROW_MAX, FinishColumn:=2, Sheet:="result")

    Dim found_rows As ObjectList
    Set found_rows = New_ObjectList("WorksheetRangeBounds")
    Call found_rows.Add(New_RangeBounds(Row:=5, Column:=2, Sheet:="result"))
    Call ws_stub.Store.SetReturn("Find", found_rows, "T-001", search_bounds, True, True, True, True)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 0)
    Call detail_values.Update(0, "案件A")

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.WriteDiagnosticRow("", "T-001", G_WEB_STATUS_OK, "", detail_values, G_WEB_DOWNLOAD_STATUS_NO_FILE)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 5, 2, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 3, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 5, G_WEB_DOWNLOAD_STATUS_NO_FILE)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 6, "案件A")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=2, Column:=2, Sheet:="result"))
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
    Set search_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=G_ROW_MAX, FinishColumn:=2, Sheet:="result")

    Dim ok_rows As ObjectList
    Set ok_rows = New_ObjectList("WorksheetRangeBounds")
    Call ok_rows.Add(New_RangeBounds(Row:=5, Column:=2, Sheet:="result"))
    Call ws_stub.Store.SetReturn("Find", ok_rows, "T-OK", search_bounds, True, True, True, True)
    Call ws_stub.Store.SetReturn("ReadCell", G_WEB_STATUS_OK, New_RangeBounds(Row:=5, Column:=3, Sheet:="result"), False)

    Dim error_rows As ObjectList
    Set error_rows = New_ObjectList("WorksheetRangeBounds")
    Call error_rows.Add(New_RangeBounds(Row:=6, Column:=2, Sheet:="result"))
    Call ws_stub.Store.SetReturn("Find", error_rows, "T-ERROR", search_bounds, True, True, True, True)
    Call ws_stub.Store.SetReturn("ReadCell", G_WEB_STATUS_ERROR, New_RangeBounds(Row:=6, Column:=3, Sheet:="result"), False)

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
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=G_ROW_MAX, FinishColumn:=2, Sheet:="result")

    Dim existing_rows As ObjectList
    Set existing_rows = New_ObjectList("WorksheetRangeBounds")
    Call existing_rows.Add(New_RangeBounds(Row:=5, Column:=2, Sheet:="result"))
    Call ws_stub.Store.SetReturn("Find", existing_rows, "T-001", search_bounds, True, True, True, True)

    Dim missing_rows As ObjectList
    Set missing_rows = New_ObjectList("WorksheetRangeBounds")
    Call ws_stub.Store.SetReturn("Find", missing_rows, "T-002", search_bounds, True, True, True, True)

    Dim used_search_bounds As WorksheetRangeBounds
    Set used_search_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=6, Sheet:="result")

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=6, FinishColumn:=6, Sheet:="result")
    Call ws_stub.Store.SetReturn("GetUsedRangeBounds", used_bounds, used_search_bounds, True, True, True, False)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 0)
    Call detail_values.Update(0, "案件A")

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.WriteCollectionRow("", "T-001", G_WEB_STATUS_OK, "", detail_values, G_WEB_DOWNLOAD_STATUS_DOWNLOADED)
    Call writer.WriteCollectionRow("", "T-002", G_WEB_STATUS_ERROR, "missing", detail_values, G_WEB_DOWNLOAD_STATUS_ERROR)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 5, 2, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 3, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 5, G_WEB_DOWNLOAD_STATUS_DOWNLOADED)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 6, "案件A")
    Call pAssertWrittenCell(Assert, ws_stub, 7, 2, "T-002")
    Call pAssertWrittenCell(Assert, ws_stub, 7, 3, G_WEB_STATUS_ERROR)
    Call pAssertWrittenCell(Assert, ws_stub, 7, 4, "missing")
    Call pAssertWrittenCell(Assert, ws_stub, 7, 5, G_WEB_DOWNLOAD_STATUS_ERROR)
    Call pAssertWrittenCell(Assert, ws_stub, 7, 6, "案件A")
End Sub
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
