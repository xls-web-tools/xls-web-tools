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
Private Const C_COL_DETAIL_PAGE_ID As Long = 1
Private Const C_COL_TARGET_ID As Long = 2
Private Const C_COL_STATUS As Long = 3
Private Const C_COL_ERROR As Long = 4
Private Const C_COL_DETAIL_START As Long = 5

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
        Err.Raise Number:=vbObjectError + 1, Source:="Class OutputSheetWriter", Description:="Settings は必須です。"
    End If

    Set pSettings = Settings
    pIsInitialized = True
End Sub

'* 固定管理列と詳細ページ列定義のヘッダーを用意します。
'*
'* @details
'* 1 行目に固定管理列を出力し、その右側へ settings の詳細ページ列定義を表示順に出力します。
'* 実際のワークシート書き込みは WsSrv.WriteCell 経由で行います。
Public Sub EnsureHeaders()
    Call pCheckInit

    Call pWriteHeader(C_COL_DETAIL_PAGE_ID, G_WEB_OUTPUT_COL_DETAIL_PAGE_ID)
    Call pWriteHeader(C_COL_TARGET_ID, G_WEB_OUTPUT_COL_TARGET_ID)
    Call pWriteHeader(C_COL_STATUS, G_WEB_OUTPUT_COL_STATUS)
    Call pWriteHeader(C_COL_ERROR, G_WEB_OUTPUT_COL_ERROR)

    Dim detail_defs As ObjectList
    Set detail_defs = pSettings.DetailColumnDefinitions

    Dim def_idx As Long
    For def_idx = 0 To detail_defs.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = detail_defs.Item(def_idx)
        Call pWriteHeader(C_COL_DETAIL_START + def_idx, detail_def.OutputColumnName)
    Next def_idx
End Sub

Private Sub pWriteHeader(ByVal ColumnIndex As Long, ByVal HeaderText As String)
    Call WsSrv.WriteCell( _
            New_RangeBounds(Row:=C_HEADER_ROW, Column:=ColumnIndex, Sheet:=pSettings.OutputSheetName), _
            HeaderText, _
            TypeConvert:=False)
End Sub

'* 診断用の出力行を書き込みます。
'*
'* @param DetailPageId 詳細ページID。
'* @param TargetId 対象ID。
'* @param Status 取得状態。
'* @param ErrorDescription エラー内容。
'* @param DetailValues 詳細ページ列定義に対応する抽出値。
'*
'* @details
'* TargetId または DetailPageId の既存行があればその行へ上書きし、見つからない場合は 2 行目へ書き込みます。
Public Sub WriteDiagnosticRow( _
        ByVal DetailPageId As String, _
        ByVal TargetId As String, _
        ByVal Status As String, _
        ByVal ErrorDescription As String, _
        ByVal DetailValues As ArrayObject)
    Call pCheckInit

    If DetailValues Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class OutputSheetWriter", Description:="DetailValues は必須です。"
    End If

    Dim output_row As Long
    output_row = pResolveDiagnosticOutputRow(DetailPageId, TargetId)

    Call pWriteOutputCell(output_row, C_COL_DETAIL_PAGE_ID, DetailPageId)
    Call pWriteOutputCell(output_row, C_COL_TARGET_ID, TargetId)
    Call pWriteOutputCell(output_row, C_COL_STATUS, Status)
    Call pWriteOutputCell(output_row, C_COL_ERROR, ErrorDescription)

    Dim detail_idx As Long
    For detail_idx = DetailValues.LowerBound To DetailValues.UpperBound
        Call pWriteOutputCell(output_row, C_COL_DETAIL_START + detail_idx, CStr(DetailValues.Item(detail_idx)))
    Next detail_idx
End Sub

Private Function pResolveDiagnosticOutputRow(ByVal DetailPageId As String, ByVal TargetId As String) As Long
    Dim target_row As Long
    target_row = pFindExistingOutputRow(C_COL_TARGET_ID, TargetId)
    If target_row <> 0 Then
        pResolveDiagnosticOutputRow = target_row
        Exit Function
    End If

    target_row = pFindExistingOutputRow(C_COL_DETAIL_PAGE_ID, DetailPageId)
    If target_row <> 0 Then
        pResolveDiagnosticOutputRow = target_row
        Exit Function
    End If

    pResolveDiagnosticOutputRow = C_DIAGNOSTIC_OUTPUT_ROW
End Function

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

Private Sub pWriteOutputCell(ByVal RowIndex As Long, ByVal ColumnIndex As Long, ByVal CellValue As String)
    Call WsSrv.WriteCell( _
            New_RangeBounds(Row:=RowIndex, Column:=ColumnIndex, Sheet:=pSettings.OutputSheetName), _
            CellValue, _
            TypeConvert:=False)
End Sub

'* 既存行を今回の実行で書き込むべきか返します。
'*
'* @param ExistingStatus 既存行の取得状態。
'* @param IsDiagnosticMode True の場合は診断モードとして常に上書き対象にします。
'* @return 今回の実行で既存行を書き込むべき場合は True。
'*
'* @details
'* 診断モードでは既存行を上書きします。本番収集では ExistingRowMode に従い、
'* SkipExisting、RetryError、Overwrite のいずれかで判定します。
Public Function ShouldWriteExistingRow(ByVal ExistingStatus As String, Optional ByVal IsDiagnosticMode As Boolean = False) As Boolean
    Call pCheckInit

    If IsDiagnosticMode Then
        ShouldWriteExistingRow = True
        Exit Function
    End If

    Select Case pSettings.ExistingRowMode
        Case G_WEB_ROW_MODE_SKIP_EXISTING
            ShouldWriteExistingRow = False
        Case G_WEB_ROW_MODE_RETRY_ERROR
            ShouldWriteExistingRow = (ExistingStatus = G_WEB_STATUS_ERROR)
        Case G_WEB_ROW_MODE_OVERWRITE
            ShouldWriteExistingRow = True
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:="Class OutputSheetWriter", Description:="ExistingRowMode が不正です。(" & pSettings.ExistingRowMode & ")"
    End Select
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class OutputSheetWriter", Description:="初期化されていません。"
    End If
End Sub
