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

Public Sub Test_OutputSheetWriter_固定管理列ヘッダーだけを用意する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseOutputSheetStubs()

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.EnsureHeaders

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 1, 1, G_WEB_OUTPUT_COL_TARGET_ID)
    Call pAssertWrittenCell(Assert, ws_stub, 1, 2, G_WEB_OUTPUT_COL_STATUS)
    Call pAssertWrittenCell(Assert, ws_stub, 1, 3, G_WEB_OUTPUT_COL_ERROR)
    Call pAssertWrittenCell(Assert, ws_stub, 1, 4, G_WEB_OUTPUT_COL_DOWNLOAD_STATUS)
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCountAll("WriteRange")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCountAll("ClearRange")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=1, Column:=5, Sheet:="result"))
End Sub

Public Sub Test_OutputSheetWriter_診断出力はoutputヘッダーに一致する詳細列へ書き込む(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseOutputSheetStubs()
    Call pSetOutputHeaders(ws_stub, "件名")

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title"))
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Call pSetFindRows(ws_stub, "T-001", 5)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 1)
    Call detail_values.Update(0, "案件A")
    Call detail_values.Update(1, "山田太郎")

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.WriteDiagnosticRow("T-001", G_WEB_STATUS_OK, "", detail_values, G_WEB_DOWNLOAD_STATUS_NO_FILE)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 5, 1, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 2, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 3, "")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 4, G_WEB_DOWNLOAD_STATUS_NO_FILE)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 5, "案件A")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=5, Column:=6, Sheet:="result"))
End Sub

Public Sub Test_OutputSheetWriter_OK行は空白ヘッダー列と未定義ヘッダー列を保持する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseOutputSheetStubs()
    Call pSetOutputHeaders(ws_stub, "件名", "", "未定義", "申請者")

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title"))
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Call pSetFindRows(ws_stub, "T-001", 5)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 1)
    Call detail_values.Update(0, "案件A")
    Call detail_values.Update(1, "山田太郎")

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.WriteDiagnosticRow("T-001", G_WEB_STATUS_OK, "", detail_values, G_WEB_DOWNLOAD_STATUS_NO_FILE)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 5, 5, "案件A")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 8, "山田太郎")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=5, Column:=6, Sheet:="result"))
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=5, Column:=7, Sheet:="result"))
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCountAll("ClearRange")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCountAll("WriteRange")
End Sub

Public Sub Test_OutputSheetWriter_ERROR行は定義済み出力列だけを空欄化する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseOutputSheetStubs()
    Call pSetOutputHeaders(ws_stub, "件名", "", "未定義", "申請者")

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title"))
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Call pSetFindRows(ws_stub, "T-001", 5)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 1)
    Call detail_values.Update(0, "古い案件")
    Call detail_values.Update(1, "古い申請者")

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.WriteDiagnosticRow("T-001", G_WEB_STATUS_ERROR, "missing", detail_values, G_WEB_DOWNLOAD_STATUS_ERROR)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 5, 5, "")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 8, "")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=5, Column:=6, Sheet:="result"))
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=5, Column:=7, Sheet:="result"))
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCountAll("ClearRange")
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCountAll("WriteRange")
End Sub

Public Sub Test_OutputSheetWriter_重複ヘッダーは同じ抽出値を両方へ出力する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseOutputSheetStubs()
    Call pSetOutputHeaders(ws_stub, "件名", "件名")

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Call pSetFindRows(ws_stub, "T-001", 5)

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
    Call pAssertWrittenCell(Assert, ws_stub, 5, 5, "案件A")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 6, "案件A")
End Sub

Public Sub Test_OutputSheetWriter_固定管理列名ヘッダーは同名抽出値を出力する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseOutputSheetStubs()
    Call pSetOutputHeaders(ws_stub, G_WEB_OUTPUT_COL_TARGET_ID, G_WEB_OUTPUT_COL_STATUS, G_WEB_OUTPUT_COL_ERROR, G_WEB_OUTPUT_COL_DOWNLOAD_STATUS)

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition(G_WEB_OUTPUT_COL_TARGET_ID, "#detail-target-id"))
    Call detail_defs.Add(New_DetailColumnDefinition(G_WEB_OUTPUT_COL_STATUS, "#detail-status"))
    Call detail_defs.Add(New_DetailColumnDefinition(G_WEB_OUTPUT_COL_ERROR, "#detail-error"))
    Call detail_defs.Add(New_DetailColumnDefinition(G_WEB_OUTPUT_COL_DOWNLOAD_STATUS, "#detail-download-status"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Call pSetFindRows(ws_stub, "T-001", 5)

    Dim detail_values As ArrayObject
    Set detail_values = New ArrayObject
    Call detail_values.ReDimArray(0, 3)
    Call detail_values.Update(0, "DETAIL-ID")
    Call detail_values.Update(1, "DETAIL-STATUS")
    Call detail_values.Update(2, "DETAIL-ERROR")
    Call detail_values.Update(3, "DETAIL-DOWNLOAD")

    Dim writer As OutputSheetWriter
    Set writer = New_OutputSheetWriter(tool_settings)

    ' --- Act ---
    Call writer.WriteDiagnosticRow("T-001", G_WEB_STATUS_OK, "", detail_values, G_WEB_DOWNLOAD_STATUS_NO_FILE)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Call pAssertWrittenCell(Assert, ws_stub, 5, 1, "T-001")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 2, G_WEB_STATUS_OK)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 3, "")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 4, G_WEB_DOWNLOAD_STATUS_NO_FILE)
    Call pAssertWrittenCell(Assert, ws_stub, 5, 5, "DETAIL-ID")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 6, "DETAIL-STATUS")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 7, "DETAIL-ERROR")
    Call pAssertWrittenCell(Assert, ws_stub, 5, 8, "DETAIL-DOWNLOAD")
End Sub

Public Sub Test_OutputSheetWriter_固定管理列名ヘッダーも未定義なら保持する(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseOutputSheetStubs()
    Call pSetOutputHeaders(ws_stub, G_WEB_OUTPUT_COL_TARGET_ID)

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Call pSetFindRows(ws_stub, "T-001", 5)

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
    Assert.EqualsNumeric 0, ws_stub.Store.GetCallCount("WriteCell", New_RangeBounds(Row:=5, Column:=5, Sheet:="result"))
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
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseOutputSheetStubs()

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Call pSetFindRows(ws_stub, "T-OK", 5)
    Call ws_stub.Store.SetReturn("ReadCell", G_WEB_STATUS_OK, New_RangeBounds(Row:=5, Column:=2, Sheet:="result"), False)

    Call pSetFindRows(ws_stub, "T-ERROR", 6)
    Call ws_stub.Store.SetReturn("ReadCell", G_WEB_STATUS_ERROR, New_RangeBounds(Row:=6, Column:=2, Sheet:="result"), False)

    Call pSetFindRows(ws_stub, "T-NEW", 0)

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
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseOutputSheetStubs()
    Call pSetOutputHeaders(ws_stub, "件名")
    Call pSetManagedUsedBounds(ws_stub, 6)

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Call pSetFindRows(ws_stub, "T-001", 5)
    Call pSetFindRows(ws_stub, "T-002", 0)

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
    Call pAssertWrittenCell(Assert, ws_stub, 7, 5, "")
End Sub

Public Sub Test_OutputSheetWriter_outputヘッダーなしなら固定管理列だけを書き込む(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = pUseOutputSheetStubs()
    Call pSetEmptyOutputHeaders(ws_stub)
    Call pSetManagedUsedBounds(ws_stub, 1)

    Dim tool_settings As ToolSettingsTestDouble
    Set tool_settings = New ToolSettingsTestDouble
    tool_settings.OutputSheetName = "result"

    Dim detail_defs As ObjectList
    Set detail_defs = New_ObjectList("DetailColumnDefinition")
    Call detail_defs.Add(New_DetailColumnDefinition("件名", "#title"))
    Call detail_defs.Add(New_DetailColumnDefinition("申請者", "#requester"))
    Set tool_settings.DetailColumnDefinitions = detail_defs

    Call pSetFindRows(ws_stub, "T-001", 0)

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

Private Function pUseOutputSheetStubs() As WorksheetServiceTestDouble
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Set WbSrv = wb_stub

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Set pUseOutputSheetStubs = ws_stub
End Function

Private Sub pSetFindRows(ByVal WsStub As WorksheetServiceTestDouble, ByVal TargetId As String, ByVal RowIndex As Long)
    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=1, Sheet:="result")

    Dim found_rows As ObjectList
    Set found_rows = New_ObjectList("WorksheetRangeBounds")
    If RowIndex <> 0 Then
        Call found_rows.Add(New_RangeBounds(Row:=RowIndex, Column:=1, Sheet:="result"))
    End If

    Call WsStub.Store.SetReturn("Find", found_rows, TargetId, search_bounds, True, True, True, True)
End Sub

Private Sub pSetManagedUsedBounds(ByVal WsStub As WorksheetServiceTestDouble, ByVal FinishRow As Long)
    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=4, Sheet:="result")

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=FinishRow, FinishColumn:=4, Sheet:="result")
    Call WsStub.Store.SetReturn("GetUsedRangeBounds", used_bounds, search_bounds, True, True, True, False)
End Sub

Private Sub pSetOutputHeaders(ByVal WsStub As WorksheetServiceTestDouble, ParamArray HeaderNames() As Variant)
    Dim header_count As Long
    header_count = UBound(HeaderNames) - LBound(HeaderNames) + 1

    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = pOutputHeaderSearchBounds()

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = New_RangeBounds(Row:=1, Column:=5, FinishRow:=1, FinishColumn:=4 + header_count, Sheet:="result")

    Dim header_values() As Variant
    ReDim header_values(1 To 1, 1 To header_count)

    Dim header_idx As Long
    For header_idx = 1 To header_count
        header_values(1, header_idx) = CStr(HeaderNames(LBound(HeaderNames) + header_idx - 1))
    Next header_idx

    Call WsStub.Store.SetReturn("GetUsedRangeBounds", used_bounds, search_bounds, True, True, True, False)
    Call WsStub.Store.SetReturn("ReadRange", header_values, used_bounds)
End Sub

Private Sub pSetEmptyOutputHeaders(ByVal WsStub As WorksheetServiceTestDouble)
    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = pOutputHeaderSearchBounds()

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = New_RangeBounds(Row:=1, Column:=5, FinishRow:=0, FinishColumn:=0, Sheet:="result")
    Call WsStub.Store.SetReturn("GetUsedRangeBounds", used_bounds, search_bounds, True, True, True, False)
End Sub

Private Function pOutputHeaderSearchBounds() As WorksheetRangeBounds
    Set pOutputHeaderSearchBounds = New_RangeBounds( _
            Row:=1, _
            Column:=5, _
            FinishRow:=1, _
            FinishColumn:=G_COL_MAX, _
            Sheet:="result")
End Function

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
