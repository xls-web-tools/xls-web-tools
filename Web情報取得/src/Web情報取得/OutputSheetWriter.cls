VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutputSheetWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Web情報取得の出力先シートへの書き込みを管理します。
'!
' #############################################################################

Private Const C_HEADER_ROW As Long = 1
Private Const C_DIAGNOSTIC_OUTPUT_ROW As Long = 2
Private Const C_COL_TARGET_ID As Long = 1
Private Const C_COL_STATUS As Long = 2
Private Const C_COL_ERROR As Long = 3
Private Const C_COL_DOWNLOAD_STATUS As Long = 4
Private Const C_COL_DETAIL_START As Long = 5
Private Const C_NOT_FOUND As Long = -1

Private pIsInitialized As Boolean
Private pSettings As IToolSettings

'* 設定を初期化します。
'*
'* @param Settings Web情報取得の設定。
'*
'* @details
'* 出力先シート名、詳細ページ列定義、既存行扱いを Settings から参照できる状態にします。
'* Settings が Nothing の場合はエラーにします。
Public Sub Initialize(ByVal Settings As IToolSettings)
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class OutputSheetWriter.Initialize", Description:="Settings は必須です。"
    End If

    Set pSettings = Settings
    pIsInitialized = True
End Sub

'* 固定管理列の既定ヘッダーを用意します。
'*
'* @details
'* A:D は列位置で固定管理列として扱います。1 行目は利用者が変更できる表示名のため、
'* 既に値、数式、または空文字列が入っているセルは変更しません。本当に空白のセルだけ既定値を書き込みます。
Public Sub EnsureHeaders()
    Call pCheckInit

    Call pWriteManagedHeaderIfBlank(C_COL_TARGET_ID, G_WEB_OUTPUT_COL_TARGET_ID)
    Call pWriteManagedHeaderIfBlank(C_COL_STATUS, G_WEB_OUTPUT_COL_STATUS)
    Call pWriteManagedHeaderIfBlank(C_COL_ERROR, G_WEB_OUTPUT_COL_ERROR)
    Call pWriteManagedHeaderIfBlank(C_COL_DOWNLOAD_STATUS, G_WEB_OUTPUT_COL_DOWNLOAD_STATUS)
End Sub

Private Sub pWriteManagedHeaderIfBlank(ByVal ColumnIndex As Long, ByVal HeaderValue As String)
    Dim target_bounds As WorksheetRangeBounds
    Set target_bounds = New_RangeBounds(Row:=C_HEADER_ROW, Column:=ColumnIndex, Sheet:=pSettings.OutputSheetName)

    If WsSrv.IsEmptyCell(target_bounds, IgnoreEmptyString:=False) Then
        Call pWriteOutputCell(C_HEADER_ROW, ColumnIndex, HeaderValue)
    End If
End Sub

'* 診断用の出力行を書き込みます。
'*
'* @param TargetId 対象ID。
'* @param Status 取得状態。
'* @param ErrorDescription エラー内容。
'* @param DetailValues 詳細ページ列定義に対応する抽出値。
'* @param DownloadStatus ダウンロード状態。
'*
'* @details
'* TargetId の既存行があればその行へ上書きし、見つからない場合は 2 行目へ書き込みます。
Public Sub WriteDiagnosticRow( _
        ByVal TargetId As String, _
        ByVal Status As String, _
        ByVal ErrorDescription As String, _
        ByVal DetailValues As ArrayObject, _
        Optional ByVal DownloadStatus As String = "")
    Call pCheckInit

    If DetailValues Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class OutputSheetWriter.WriteDiagnosticRow", Description:="DetailValues は必須です。"
    End If

    Dim output_row As Long
    output_row = pResolveDiagnosticOutputRow(TargetId)

    Call pWriteOutputRow(output_row, TargetId, Status, ErrorDescription, DownloadStatus, DetailValues)
End Sub

Private Function pResolveDiagnosticOutputRow(ByVal TargetId As String) As Long
    Dim target_row As Long
    target_row = pFindExistingOutputRow(C_COL_TARGET_ID, TargetId)
    If target_row <> 0 Then
        pResolveDiagnosticOutputRow = target_row
        Exit Function
    End If

    pResolveDiagnosticOutputRow = C_DIAGNOSTIC_OUTPUT_ROW
End Function

'* 本番収集用の出力行を書き込みます。
'*
'* @param TargetId 対象ID。
'* @param Status 取得状態。
'* @param ErrorDescription エラー内容。
'* @param DetailValues 詳細ページ列定義に対応する抽出値。
'* @param DownloadStatus ダウンロード状態。
'*
'* @details
'* TargetId の既存行があればその行へ上書きし、見つからない場合は出力先シートの末尾へ追加します。
Public Sub WriteCollectionRow( _
        ByVal TargetId As String, _
        ByVal Status As String, _
        ByVal ErrorDescription As String, _
        ByVal DetailValues As ArrayObject, _
        Optional ByVal DownloadStatus As String = "")
    Call pCheckInit

    If DetailValues Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class OutputSheetWriter.WriteCollectionRow", Description:="DetailValues は必須です。"
    End If

    Dim target_row As Long
    target_row = pFindExistingOutputRow(C_COL_TARGET_ID, TargetId)

    Dim output_row As Long
    If target_row = 0 Then
        output_row = pResolveNextOutputRow()
    Else
        output_row = target_row
    End If

    Call pWriteOutputRow( _
            output_row, _
            TargetId, _
            Status, _
            ErrorDescription, _
            DownloadStatus, _
            DetailValues, _
            FillBlankCells:=(target_row <> 0 And pSettings.ExistingRowMode = G_WEB_ROW_MODE_FILL_BLANK_CELLS))
End Sub

Private Function pResolveNextOutputRow() As Long
    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds( _
            Row:=C_HEADER_ROW, _
            Column:=C_COL_TARGET_ID, _
            FinishRow:=G_ROW_MAX, _
            FinishColumn:=C_COL_DOWNLOAD_STATUS, _
            Sheet:=pSettings.OutputSheetName)

    Dim used_bounds As WorksheetRangeBounds
    Set used_bounds = WsSrv.GetUsedRangeBounds(search_bounds)

    Dim result_row As Long
    result_row = used_bounds.FinishRow + 1
    If result_row < C_DIAGNOSTIC_OUTPUT_ROW Then result_row = C_DIAGNOSTIC_OUTPUT_ROW

    pResolveNextOutputRow = result_row
End Function

'* 対象IDを今回の本番収集で取得すべきか返します。
'*
'* @param TargetId 対象ID。
'* @return 未登録、または ExistingRowMode 上再取得すべき場合は True。
'*
'* @details
'* TargetId を主キーとして出力先シートを検索します。既存行が見つからない場合は取得対象とし、
'* 既存行が見つかった場合は取得状態と ExistingRowMode で判定します。
Public Function ShouldCollectTarget(ByVal TargetId As String) As Boolean
    Call pCheckInit

    If Trim$(TargetId) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class OutputSheetWriter.ShouldCollectTarget", Description:="TargetId は必須です。"
    End If

    Dim target_row As Long
    target_row = pFindExistingOutputRow(C_COL_TARGET_ID, TargetId)
    If target_row = 0 Then
        ShouldCollectTarget = True
        Exit Function
    End If

    If pSettings.ExistingRowMode = G_WEB_ROW_MODE_FILL_BLANK_CELLS Then
        ShouldCollectTarget = pHasBlankOutputCell(target_row)
        Exit Function
    End If

    Dim existing_status As String
    existing_status = CStr(WsSrv.ReadCell(New_RangeBounds(Row:=target_row, Column:=C_COL_STATUS, Sheet:=pSettings.OutputSheetName)))

    ShouldCollectTarget = ShouldWriteExistingRow(existing_status)
End Function

Private Function pHasBlankOutputCell(ByVal RowIndex As Long) As Boolean
    Dim detail_defs As ObjectList
    Set detail_defs = pSettings.DetailColumnDefinitions
    If detail_defs.Count = 0 Then Exit Function

    Dim output_columns As ObjectDictionary
    Set output_columns = pResolveOutputColumnSelections(detail_defs)
    If output_columns.Count = 0 Then Exit Function

    Dim output_col_keys() As Variant
    output_col_keys = output_columns.Keys

    Dim key_idx As Long
    For key_idx = LBound(output_col_keys) To UBound(output_col_keys)
        Dim output_col_idx As Long
        output_col_idx = CLng(output_col_keys(key_idx))

        If WsSrv.IsEmptyCell(New_RangeBounds(Row:=RowIndex, Column:=output_col_idx, Sheet:=pSettings.OutputSheetName), IgnoreEmptyString:=False) Then
            pHasBlankOutputCell = True
            Exit Function
        End If
    Next key_idx
End Function

'* 既存行を今回の実行で書き込むべきか返します。
'*
'* @param ExistingStatus 既存行の取得状態。
'* @param IsDiagnosticMode True の場合は診断モードとして常に上書き対象にします。
'* @return 今回の実行で既存行を書き込むべき場合は True。
'*
'* @details
'* 診断モードでは既存行を上書きします。本番収集では ExistingRowMode に従い、
'* SkipExisting または Overwrite のいずれかで判定します。
Public Function ShouldWriteExistingRow(ByVal ExistingStatus As String, Optional ByVal IsDiagnosticMode As Boolean = False) As Boolean
    Call pCheckInit

    If IsDiagnosticMode Then
        ShouldWriteExistingRow = True
        Exit Function
    End If

    Call pValidateExistingStatus(ExistingStatus)

    Select Case pSettings.ExistingRowMode
        Case G_WEB_ROW_MODE_SKIP_EXISTING
            ShouldWriteExistingRow = (ExistingStatus = G_WEB_STATUS_ERROR)
        Case G_WEB_ROW_MODE_OVERWRITE
            ShouldWriteExistingRow = True
        Case G_WEB_ROW_MODE_FILL_BLANK_CELLS
            ShouldWriteExistingRow = True
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:="Class OutputSheetWriter.ShouldWriteExistingRow", Description:="ExistingRowMode が不正です。(" & pSettings.ExistingRowMode & ")"
    End Select
End Function

Private Sub pValidateExistingStatus(ByVal ExistingStatus As String)
    Select Case ExistingStatus
        Case G_WEB_STATUS_OK, G_WEB_STATUS_ERROR
            Exit Sub
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:="Class OutputSheetWriter.pValidateExistingStatus", Description:="既存行の取得状態が不正です。(" & ExistingStatus & ")"
    End Select
End Sub

Private Sub pWriteOutputCell( _
        ByVal RowIndex As Long, _
        ByVal ColumnIndex As Long, _
        ByVal CellValue As String, _
        Optional ByVal TypeConvert As Boolean = False, _
        Optional ByVal OnlyIfBlank As Boolean = False)

    Dim target_bounds As WorksheetRangeBounds
    Set target_bounds = New_RangeBounds(Row:=RowIndex, Column:=ColumnIndex, Sheet:=pSettings.OutputSheetName)

    If OnlyIfBlank Then
        If CellValue = "" Then Exit Sub
        If Not WsSrv.IsEmptyCell(target_bounds, IgnoreEmptyString:=False) Then Exit Sub
    End If

    Call WsSrv.WriteCell( _
            target_bounds, _
            CellValue, _
            TypeConvert:=TypeConvert)
End Sub

Private Function pFindExistingOutputRow(ByVal ColumnIndex As Long, ByVal LookupValue As String) As Long
    If Trim$(LookupValue) = "" Then Exit Function

    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds( _
            Row:=C_DIAGNOSTIC_OUTPUT_ROW, _
            Column:=ColumnIndex, _
            FinishRow:=G_ROW_MAX, _
            FinishColumn:=ColumnIndex, _
            Sheet:=pSettings.OutputSheetName)

    Dim found_cells As ObjectList
    Set found_cells = WsSrv.Find(LookupValue, search_bounds)
    If found_cells.Count = 0 Then Exit Function

    Dim found_bounds As WorksheetRangeBounds
    Set found_bounds = found_cells.Item(0)
    pFindExistingOutputRow = found_bounds.Row
End Function

Private Sub pWriteOutputRow( _
        ByVal RowIndex As Long, _
        ByVal TargetId As String, _
        ByVal Status As String, _
        ByVal ErrorDescription As String, _
        ByVal DownloadStatus As String, _
        ByVal DetailValues As ArrayObject, _
        Optional ByVal FillBlankCells As Boolean = False)

    Call pWriteOutputCell(RowIndex, C_COL_TARGET_ID, TargetId, OnlyIfBlank:=FillBlankCells)
    Call pWriteOutputCell(RowIndex, C_COL_STATUS, Status, OnlyIfBlank:=FillBlankCells)
    Call pWriteOutputCell(RowIndex, C_COL_ERROR, ErrorDescription, OnlyIfBlank:=FillBlankCells)
    Call pWriteOutputCell(RowIndex, C_COL_DOWNLOAD_STATUS, DownloadStatus, OnlyIfBlank:=FillBlankCells)
    Call pWriteDetailOutputCells(RowIndex, DetailValues, BlankValues:=(Status <> G_WEB_STATUS_OK), FillBlankCells:=FillBlankCells)
End Sub

Private Sub pWriteDetailOutputCells( _
        ByVal RowIndex As Long, _
        ByVal DetailValues As ArrayObject, _
        Optional ByVal BlankValues As Boolean = False, _
        Optional ByVal FillBlankCells As Boolean = False)
    Dim detail_defs As ObjectList
    Set detail_defs = pSettings.DetailColumnDefinitions
    If detail_defs.Count = 0 Then Exit Sub

    Dim output_columns As ObjectDictionary
    Set output_columns = pResolveOutputColumnSelections(detail_defs)
    If output_columns.Count = 0 Then Exit Sub

    Dim output_col_keys() As Variant
    output_col_keys = output_columns.Keys

    Dim key_idx As Long
    For key_idx = LBound(output_col_keys) To UBound(output_col_keys)
        Dim output_col_idx As Long
        output_col_idx = CLng(output_col_keys(key_idx))

        Dim detail_idx As Long
        detail_idx = CLng(output_columns.Item(output_col_idx))

        Dim output_value As String
        If BlankValues Then
            output_value = ""
        Else
            output_value = CStr(DetailValues.Item(detail_idx))
        End If

        Dim detail_def As DetailColumnDefinition
        Set detail_def = detail_defs.Item(detail_idx)

        Call pWriteOutputCell( _
                RowIndex, _
                output_col_idx, _
                output_value, _
                pShouldTypeConvertOutputValue(detail_def, output_value), _
                OnlyIfBlank:=FillBlankCells)
    Next key_idx
End Sub

Private Function pResolveOutputColumnSelections(ByVal DetailDefinitions As ObjectList) As ObjectDictionary
    Dim result_values As ObjectDictionary
    Set result_values = New_ObjectDictionary()

    Dim header_bounds As WorksheetRangeBounds
    Set header_bounds = pOutputHeaderUsedBounds()
    If header_bounds.IsEmpty Then
        Set pResolveOutputColumnSelections = result_values
        Exit Function
    End If

    Dim header_values As Variant
    header_values = WsSrv.ReadRange(header_bounds)

    Dim header_col_idx As Long
    For header_col_idx = LBound(header_values, 2) To UBound(header_values, 2)
        Dim header_text As String
        header_text = Trim$(pCellText(header_values(LBound(header_values, 1), header_col_idx)))

        If header_text <> "" Then
            Dim detail_idx As Long
            detail_idx = pFindDetailDefinitionIndex(DetailDefinitions, header_text)
            If detail_idx <> C_NOT_FOUND Then
                Dim output_col_idx As Long
                output_col_idx = header_bounds.Column + header_col_idx - LBound(header_values, 2)
                Call result_values.Add(output_col_idx, detail_idx)
            End If
        End If
    Next header_col_idx

    Set pResolveOutputColumnSelections = result_values
End Function

Private Function pShouldTypeConvertOutputValue(ByVal DetailDefinition As DetailColumnDefinition, ByVal OutputValue As String) As Boolean
    pShouldTypeConvertOutputValue = ( _
            StrComp(DetailDefinition.OutputValueType, "Auto", vbTextCompare) = 0 _
            And OutputValue <> "")
End Function

Private Function pOutputHeaderUsedBounds() As WorksheetRangeBounds
    Dim search_bounds As WorksheetRangeBounds
    Set search_bounds = New_RangeBounds( _
            Row:=C_HEADER_ROW, _
            Column:=C_COL_DETAIL_START, _
            FinishRow:=C_HEADER_ROW, _
            FinishColumn:=G_COL_MAX, _
            Sheet:=pSettings.OutputSheetName)

    Set pOutputHeaderUsedBounds = WsSrv.GetUsedRangeBounds(search_bounds)
End Function

Private Function pFindDetailDefinitionIndex(ByVal DetailDefinitions As ObjectList, ByVal HeaderText As String) As Long
    Dim def_idx As Long
    For def_idx = 0 To DetailDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = DetailDefinitions.Item(def_idx)
        If StrComp(detail_def.OutputColumnName, HeaderText, vbBinaryCompare) = 0 Then
            pFindDetailDefinitionIndex = def_idx
            Exit Function
        End If
    Next def_idx

    pFindDetailDefinitionIndex = C_NOT_FOUND
End Function

Private Function pCellText(ByVal CellValue As Variant) As String
    If IsError(CellValue) Or IsNull(CellValue) Or IsEmpty(CellValue) Then
        pCellText = ""
    Else
        pCellText = CStr(CellValue)
    End If
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class OutputSheetWriter.pCheckInit", Description:="初期化されていません。"
    End If
End Sub
