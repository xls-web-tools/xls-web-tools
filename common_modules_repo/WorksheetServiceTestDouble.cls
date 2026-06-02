VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetServiceTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWorksheetService

' #############################################################################
'!
'! @brief
'! ワークシートへのアクセスを管理するサービスのテストダブルです。
'!
'! @details
'! 戻り値が必要で Excel 状態や親リソースに依存するものは strict stub とします。戻り値なしの操作は spy only として正常呼び出しだけを記録します。
'! strict stub の未登録エラーおよび SetError による明示エラーは呼び出し履歴に残しません。
'!
' #############################################################################

'* スタブ設定と正常に成立した呼び出し履歴。
'* strict stub の未登録エラーおよび SetError による明示エラーは呼び出し履歴に残しません。
'* spy only は、元メソッド定義の順序で呼び出し引数を記録します。
Public Store As TestDoubleBehaviorStore

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
    Store.DefaultObjectKeyMode = G_OBJECT_KEY_MODE_I_EQUATABLE
End Sub

'* Find の strict stub です。検索結果なしを表す空 ObjectList も明示登録してください。
Public Function Find( _
        ByVal What As String, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList

    Call Store.RaiseIfError("Find", What, RangeBounds, LookInValues, LookAtWhole, MatchCase, MatchByte)

    Dim result_value As ObjectList
    Set result_value = Store.GetReturn("Find", What, RangeBounds, LookInValues, LookAtWhole, MatchCase, MatchByte)
    Call Store.RecordCall("Find", What, RangeBounds, LookInValues, LookAtWhole, MatchCase, MatchByte)

    Set Find = result_value
End Function
Private Function IWorksheetService_Find( _
        ByVal What As String, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList

    Set IWorksheetService_Find = Find( _
            What, RangeBounds, _
            LookInValues:=LookInValues, _
            LookAtWhole:=LookAtWhole, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte)
End Function

'* Sort の spy only です。SortKeyAndOrder を記録値として残します。
Public Sub Sort(ByVal RangeBounds As WorksheetRangeBounds, ParamArray SortKeyAndOrder() As Variant)
    Call pSortCore(RangeBounds, SortKeyAndOrder)
End Sub
Private Sub IWorksheetService_Sort(ByVal RangeBounds As WorksheetRangeBounds, ParamArray SortKeyAndOrder() As Variant)
    Call pSortCore(RangeBounds, SortKeyAndOrder)
End Sub

Private Sub pSortCore(ByVal RangeBounds As WorksheetRangeBounds, ByVal SortKeyAndOrder As Variant)
    Call Store.RaiseIfError("Sort", RangeBounds)
    Call Store.RecordCallWithKeyArgs("Sort", Array(RangeBounds), RangeBounds, SortKeyAndOrder)
End Sub

'* ActivateRange の spy only です。呼び出しを記録し、失敗は Store.SetError で注入します。
Public Sub ActivateRange(ByVal RangeBounds As WorksheetRangeBounds)
    Call Store.RaiseIfError("ActivateRange", RangeBounds)
    Call Store.RecordCall("ActivateRange", RangeBounds)
End Sub
Private Sub IWorksheetService_ActivateRange(ByVal RangeBounds As WorksheetRangeBounds)
    Call ActivateRange(RangeBounds)
End Sub

'* WriteCell の spy only です。書き込み値とオプションを記録値として残します。
Public Sub WriteCell( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal Expression As String, _
        Optional ByVal NumberFormat As String = "", _
        Optional ByVal AsFormula As Boolean = False, _
        Optional ByVal ClearWhenEmpty As Boolean = True, _
        Optional ByVal IgnoreEmpty As Boolean = False, _
        Optional ByVal TypeConvert As Boolean = True)

    Call Store.RaiseIfError("WriteCell", RangeBounds)
    Call Store.RecordCallWithKeyArgs("WriteCell", Array(RangeBounds), RangeBounds, Expression, NumberFormat, AsFormula, ClearWhenEmpty, IgnoreEmpty, TypeConvert)
End Sub
Private Sub IWorksheetService_WriteCell( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal Expression As String, _
        Optional ByVal NumberFormat As String = "", _
        Optional ByVal AsFormula As Boolean = False, _
        Optional ByVal ClearWhenEmpty As Boolean = True, _
        Optional ByVal IgnoreEmpty As Boolean = False, _
        Optional ByVal TypeConvert As Boolean = True)

    Call WriteCell( _
            RangeBounds, _
            Expression, _
            NumberFormat:=NumberFormat, _
            AsFormula:=AsFormula, _
            ClearWhenEmpty:=ClearWhenEmpty, _
            IgnoreEmpty:=IgnoreEmpty, _
            TypeConvert:=TypeConvert)
End Sub

'* WriteRange の spy only です。書き込み配列を記録値として残します。
Public Sub WriteRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal ValuesArray As Variant)
    Call Store.RaiseIfError("WriteRange", RangeBounds)
    Call Store.RecordCallWithKeyArgs("WriteRange", Array(RangeBounds), RangeBounds, ValuesArray)
End Sub
Private Sub IWorksheetService_WriteRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal ValuesArray As Variant)
    Call WriteRange(RangeBounds, ValuesArray)
End Sub

'* WriteArrayFormula の spy only です。配列数式を記録値として残します。
Public Sub WriteArrayFormula(ByVal RangeBounds As WorksheetRangeBounds, ByVal FormulaArray As String)
    Call Store.RaiseIfError("WriteArrayFormula", RangeBounds)
    Call Store.RecordCallWithKeyArgs("WriteArrayFormula", Array(RangeBounds), RangeBounds, FormulaArray)
End Sub
Private Sub IWorksheetService_WriteArrayFormula(ByVal RangeBounds As WorksheetRangeBounds, ByVal FormulaArray As String)
    Call WriteArrayFormula(RangeBounds, FormulaArray)
End Sub

'* ReadCell の strict stub です。戻り値は Store に明示登録し、ByRef 出力値は Store の Output 設定または既定値から返します。
Public Function ReadCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByRef Expression As String = "", Optional ByRef NumberFormat As String = "", Optional ByVal GetText As Boolean = False) As String
    Call Store.RaiseIfError("ReadCell", RangeBounds, GetText)

    Dim result_value As String
    result_value = Store.GetReturn("ReadCell", RangeBounds, GetText)
    Expression = Store.GetOutputOrDefault("ReadCell", "Expression", result_value, RangeBounds, GetText)
    NumberFormat = Store.GetOutputOrDefault("ReadCell", "NumberFormat", "", RangeBounds, GetText)
    Call Store.RecordCall("ReadCell", RangeBounds, GetText)

    ReadCell = result_value
End Function
Private Function IWorksheetService_ReadCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByRef Expression As String = "", Optional ByRef NumberFormat As String = "", Optional ByVal GetText As Boolean = False) As String
    IWorksheetService_ReadCell = ReadCell(RangeBounds, Expression, NumberFormat:=NumberFormat, GetText:=GetText)
End Function

'* ReadRange の strict stub です。戻り値は Store に明示登録してください。
Public Function ReadRange(ByVal RangeBounds As WorksheetRangeBounds) As Variant()
    Call Store.RaiseIfError("ReadRange", RangeBounds)

    Dim result_value() As Variant
    result_value = Store.GetReturn("ReadRange", RangeBounds)
    Call Store.RecordCall("ReadRange", RangeBounds)

    ReadRange = result_value
End Function
Private Function IWorksheetService_ReadRange(ByVal RangeBounds As WorksheetRangeBounds) As Variant()
    IWorksheetService_ReadRange = ReadRange(RangeBounds)
End Function

'* EvaluateFormula の strict stub です。戻り値は Store に明示登録してください。
Public Function EvaluateFormula(ByVal ContextRangeBounds As WorksheetRangeBounds, ByVal Formula As String) As Variant
    Call Store.RaiseIfError("EvaluateFormula", ContextRangeBounds, Formula)

    If IsObject(Store.GetReturn("EvaluateFormula", ContextRangeBounds, Formula)) Then
        Set EvaluateFormula = Store.GetReturn("EvaluateFormula", ContextRangeBounds, Formula)
        Call Store.RecordCall("EvaluateFormula", ContextRangeBounds, Formula)
    Else
        EvaluateFormula = Store.GetReturn("EvaluateFormula", ContextRangeBounds, Formula)
        Call Store.RecordCall("EvaluateFormula", ContextRangeBounds, Formula)
    End If
End Function
Private Function IWorksheetService_EvaluateFormula(ByVal ContextRangeBounds As WorksheetRangeBounds, ByVal Formula As String) As Variant
    IWorksheetService_EvaluateFormula = EvaluateFormula(ContextRangeBounds, Formula)
End Function

'* XLookup の strict stub です。戻り値は Store に明示登録してください。
Public Function XLookup( _
        ByVal LookupValue As Variant, _
        ByVal LookupRangeBounds As WorksheetRangeBounds, _
        ByVal ReturnRangeBounds As WorksheetRangeBounds, _
        Optional ByVal IfNotFound As Variant = "", _
        Optional ByVal MatchMode As Long = 0, _
        Optional ByVal SearchMode As Long = 1) As Variant

    Call Store.RaiseIfError("XLookup", LookupValue, LookupRangeBounds, ReturnRangeBounds, IfNotFound, MatchMode, SearchMode)

    If IsObject(Store.GetReturn("XLookup", LookupValue, LookupRangeBounds, ReturnRangeBounds, IfNotFound, MatchMode, SearchMode)) Then
        Set XLookup = Store.GetReturn("XLookup", LookupValue, LookupRangeBounds, ReturnRangeBounds, IfNotFound, MatchMode, SearchMode)
        Call Store.RecordCall("XLookup", LookupValue, LookupRangeBounds, ReturnRangeBounds, IfNotFound, MatchMode, SearchMode)
    Else
        XLookup = Store.GetReturn("XLookup", LookupValue, LookupRangeBounds, ReturnRangeBounds, IfNotFound, MatchMode, SearchMode)
        Call Store.RecordCall("XLookup", LookupValue, LookupRangeBounds, ReturnRangeBounds, IfNotFound, MatchMode, SearchMode)
    End If
End Function
Private Function IWorksheetService_XLookup( _
        ByVal LookupValue As Variant, _
        ByVal LookupRangeBounds As WorksheetRangeBounds, _
        ByVal ReturnRangeBounds As WorksheetRangeBounds, _
        Optional ByVal IfNotFound As Variant = "", _
        Optional ByVal MatchMode As Long = 0, _
        Optional ByVal SearchMode As Long = 1) As Variant

    IWorksheetService_XLookup = XLookup( _
            LookupValue, _
            LookupRangeBounds, _
            ReturnRangeBounds, _
            IfNotFound:=IfNotFound, _
            MatchMode:=MatchMode, _
            SearchMode:=SearchMode)
End Function

'* CopyCell の spy only です。コピー元とオプションを記録値として残します。
Public Sub CopyCell( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)

    Call Store.RaiseIfError("CopyCell", DestinationRangeBounds)
    Call Store.RecordCallWithKeyArgs("CopyCell", Array(DestinationRangeBounds), SourceRangeBounds, DestinationRangeBounds, AsValue, CopyNumberFormat)
End Sub
Private Sub IWorksheetService_CopyCell( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)

    Call CopyCell( _
            SourceRangeBounds, _
            DestinationRangeBounds, _
            AsValue:=AsValue, _
            CopyNumberFormat:=CopyNumberFormat)
End Sub

'* CopyRange の spy only です。コピー元とオプションを記録値として残します。
Public Sub CopyRange( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)

    Call Store.RaiseIfError("CopyRange", DestinationRangeBounds)
    Call Store.RecordCallWithKeyArgs("CopyRange", Array(DestinationRangeBounds), SourceRangeBounds, DestinationRangeBounds, AsValue, CopyNumberFormat)
End Sub
Private Sub IWorksheetService_CopyRange( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)

    Call CopyRange( _
            SourceRangeBounds, _
            DestinationRangeBounds, _
            AsValue:=AsValue, _
            CopyNumberFormat:=CopyNumberFormat)
End Sub

'* IsEmptyCell の strict stub です。セル状態に依存する戻り値は Store に明示登録してください。
Public Function IsEmptyCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
    Call Store.RaiseIfError("IsEmptyCell", RangeBounds, IgnoreEmptyString)

    Dim result_value As Boolean
    result_value = Store.GetReturn("IsEmptyCell", RangeBounds, IgnoreEmptyString)
    Call Store.RecordCall("IsEmptyCell", RangeBounds, IgnoreEmptyString)

    IsEmptyCell = result_value
End Function
Private Function IWorksheetService_IsEmptyCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
    IWorksheetService_IsEmptyCell = IsEmptyCell(RangeBounds, IgnoreEmptyString:=IgnoreEmptyString)
End Function

'* HasFormula の strict stub です。セル状態に依存する戻り値は Store に明示登録してください。
Public Function HasFormula(ByVal RangeBounds As WorksheetRangeBounds) As Boolean
    Call Store.RaiseIfError("HasFormula", RangeBounds)

    Dim result_value As Boolean
    result_value = Store.GetReturn("HasFormula", RangeBounds)
    Call Store.RecordCall("HasFormula", RangeBounds)

    HasFormula = result_value
End Function
Private Function IWorksheetService_HasFormula(ByVal RangeBounds As WorksheetRangeBounds) As Boolean
    IWorksheetService_HasFormula = HasFormula(RangeBounds)
End Function

'* GetUsedRangeBounds の strict stub です。戻り値は Store に明示登録してください。
Public Function GetUsedRangeBounds( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal IgnoreEmpty As Boolean = True, _
        Optional ByVal IsUsedStartRow As Boolean = True, _
        Optional ByVal IsUsedStartColumn As Boolean = True, _
        Optional ByVal GetRawRange As Boolean = False) As WorksheetRangeBounds

    Call Store.RaiseIfError("GetUsedRangeBounds", RangeBounds, IgnoreEmpty, IsUsedStartRow, IsUsedStartColumn, GetRawRange)

    Dim result_value As WorksheetRangeBounds
    Set result_value = Store.GetReturn("GetUsedRangeBounds", RangeBounds, IgnoreEmpty, IsUsedStartRow, IsUsedStartColumn, GetRawRange)
    Call Store.RecordCall("GetUsedRangeBounds", RangeBounds, IgnoreEmpty, IsUsedStartRow, IsUsedStartColumn, GetRawRange)

    Set GetUsedRangeBounds = result_value
End Function
Private Function IWorksheetService_GetUsedRangeBounds( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal IgnoreEmpty As Boolean = True, _
        Optional ByVal IsUsedStartRow As Boolean = True, _
        Optional ByVal IsUsedStartColumn As Boolean = True, _
        Optional ByVal GetRawRange As Boolean = False) As WorksheetRangeBounds

    Set IWorksheetService_GetUsedRangeBounds = GetUsedRangeBounds( _
            RangeBounds, _
            IgnoreEmpty:=IgnoreEmpty, _
            IsUsedStartRow:=IsUsedStartRow, _
            IsUsedStartColumn:=IsUsedStartColumn, _
            GetRawRange:=GetRawRange)
End Function

'* ClearRange の spy only です。クリア設定を記録値として残します。
Public Sub ClearRange( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal ClearAll As Boolean = False, _
        Optional ByVal ClearContents As Boolean = True, _
        Optional ByVal ClearNumberFormats As Boolean = False, _
        Optional ByVal ClearColors As Boolean = False, _
        Optional ByVal ClearComments As Boolean = False, _
        Optional ByVal ClearHyperlinks As Boolean = False)

    Call Store.RaiseIfError("ClearRange", RangeBounds)
    Call Store.RecordCallWithKeyArgs("ClearRange", Array(RangeBounds), RangeBounds, ClearAll, ClearContents, ClearNumberFormats, ClearColors, ClearComments, ClearHyperlinks)
End Sub
Private Sub IWorksheetService_ClearRange( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal ClearAll As Boolean = False, _
        Optional ByVal ClearContents As Boolean = True, _
        Optional ByVal ClearNumberFormats As Boolean = False, _
        Optional ByVal ClearColors As Boolean = False, _
        Optional ByVal ClearComments As Boolean = False, _
        Optional ByVal ClearHyperlinks As Boolean = False)

    Call ClearRange( _
            RangeBounds, _
            ClearAll:=ClearAll, _
            ClearContents:=ClearContents, _
            ClearNumberFormats:=ClearNumberFormats, _
            ClearColors:=ClearColors, _
            ClearComments:=ClearComments, _
            ClearHyperlinks:=ClearHyperlinks)
End Sub

'* RemoveDuplicates の spy only です。重複列指定を記録値として残します。
Public Sub RemoveDuplicates(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant)
    Call Store.RaiseIfError("RemoveDuplicates", RangeBounds)
    Call Store.RecordCallWithKeyArgs("RemoveDuplicates", Array(RangeBounds), RangeBounds, DuplicateColumns)
End Sub
Private Sub IWorksheetService_RemoveDuplicates(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant)
    Call RemoveDuplicates(RangeBounds, DuplicateColumns)
End Sub

'* InsertRows の spy only です。呼び出しを記録し、失敗は Store.SetError で注入します。
Public Sub InsertRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call Store.RaiseIfError("InsertRows", RangeBounds)
    Call Store.RecordCall("InsertRows", RangeBounds)
End Sub
Private Sub IWorksheetService_InsertRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call InsertRows(RangeBounds)
End Sub

'* DeleteRows の spy only です。呼び出しを記録し、失敗は Store.SetError で注入します。
Public Sub DeleteRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call Store.RaiseIfError("DeleteRows", RangeBounds)
    Call Store.RecordCall("DeleteRows", RangeBounds)
End Sub
Private Sub IWorksheetService_DeleteRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call DeleteRows(RangeBounds)
End Sub

'* SetAllDataVisible の spy only です。呼び出しを記録し、失敗は Store.SetError で注入します。
Public Sub SetAllDataVisible(ByVal RangeBounds As WorksheetRangeBounds)
    Call Store.RaiseIfError("SetAllDataVisible", RangeBounds)
    Call Store.RecordCall("SetAllDataVisible", RangeBounds)
End Sub
Private Sub IWorksheetService_SetAllDataVisible(ByVal RangeBounds As WorksheetRangeBounds)
    Call SetAllDataVisible(RangeBounds)
End Sub

'* SetSheetOutlineLevel の spy only です。アウトライン設定を記録値として残します。
Public Sub SetSheetOutlineLevel(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal RowLevels As Long = 0, Optional ByVal ColumnLevels As Long = 0)
    Call Store.RaiseIfError("SetSheetOutlineLevel", RangeBounds)
    Call Store.RecordCallWithKeyArgs("SetSheetOutlineLevel", Array(RangeBounds), RangeBounds, RowLevels, ColumnLevels)
End Sub
Private Sub IWorksheetService_SetSheetOutlineLevel(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal RowLevels As Long = 0, Optional ByVal ColumnLevels As Long = 0)
    Call SetSheetOutlineLevel(RangeBounds, RowLevels:=RowLevels, ColumnLevels:=ColumnLevels)
End Sub

'* SetSheetTabColor の spy only です。タブ色設定を記録値として残します。
Public Sub SetSheetTabColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal TabColorIndex As Long = xlColorIndexNone, _
        Optional ByVal TabColor As Long = -1)

    Call Store.RaiseIfError("SetSheetTabColor", RangeBounds)
    Call Store.RecordCallWithKeyArgs("SetSheetTabColor", Array(RangeBounds), RangeBounds, TabColorIndex, TabColor)
End Sub
Private Sub IWorksheetService_SetSheetTabColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal TabColorIndex As Long = xlColorIndexNone, _
        Optional ByVal TabColor As Long = -1)

    Call SetSheetTabColor( _
            RangeBounds, _
            TabColorIndex:=TabColorIndex, _
            TabColor:=TabColor)
End Sub

'* SetRangeColor の spy only です。範囲色設定を記録値として残します。
Public Sub SetRangeColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal FontColorIndex As Long = 0, _
        Optional ByVal FontColor As Long = -1, _
        Optional ByVal InteriorColorIndex As Long = xlColorIndexNone, _
        Optional ByVal InteriorColor As Long = -1)

    Call Store.RaiseIfError("SetRangeColor", RangeBounds)
    Call Store.RecordCallWithKeyArgs("SetRangeColor", Array(RangeBounds), RangeBounds, FontColorIndex, FontColor, InteriorColorIndex, InteriorColor)
End Sub
Private Sub IWorksheetService_SetRangeColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal FontColorIndex As Long = 0, _
        Optional ByVal FontColor As Long = -1, _
        Optional ByVal InteriorColorIndex As Long = xlColorIndexNone, _
        Optional ByVal InteriorColor As Long = -1)

    Call SetRangeColor( _
        RangeBounds, _
        FontColorIndex:=FontColorIndex, _
        FontColor:=FontColor, _
        InteriorColorIndex:=InteriorColorIndex, _
        InteriorColor:=InteriorColor)
End Sub

'* SetWrapText の spy only です。折り返し設定を記録値として残します。
Public Sub SetWrapText(ByVal RangeBounds As WorksheetRangeBounds, ByVal WrapText As Boolean)
    Call Store.RaiseIfError("SetWrapText", RangeBounds)
    Call Store.RecordCallWithKeyArgs("SetWrapText", Array(RangeBounds), RangeBounds, WrapText)
End Sub
Private Sub IWorksheetService_SetWrapText(ByVal RangeBounds As WorksheetRangeBounds, ByVal WrapText As Boolean)
    Call SetWrapText(RangeBounds, WrapText)
End Sub

'* SetShrinkToFit の spy only です。縮小表示設定を記録値として残します。
Public Sub SetShrinkToFit(ByVal RangeBounds As WorksheetRangeBounds, ByVal ShrinkToFit As Boolean)
    Call Store.RaiseIfError("SetShrinkToFit", RangeBounds)
    Call Store.RecordCallWithKeyArgs("SetShrinkToFit", Array(RangeBounds), RangeBounds, ShrinkToFit)
End Sub
Private Sub IWorksheetService_SetShrinkToFit(ByVal RangeBounds As WorksheetRangeBounds, ByVal ShrinkToFit As Boolean)
    Call SetShrinkToFit(RangeBounds, ShrinkToFit)
End Sub

'* SetAlignment の spy only です。配置設定を記録値として残します。
Public Sub SetAlignment( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal HorizontalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal VerticalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal Orientation As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal IndentLevel As Long = G_ALIGNMENT_NO_CHANGE)

    Call Store.RaiseIfError("SetAlignment", RangeBounds)
    Call Store.RecordCallWithKeyArgs("SetAlignment", Array(RangeBounds), RangeBounds, HorizontalAlignment, VerticalAlignment, Orientation, IndentLevel)
End Sub
Private Sub IWorksheetService_SetAlignment( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal HorizontalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal VerticalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal Orientation As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal IndentLevel As Long = G_ALIGNMENT_NO_CHANGE)

    Call SetAlignment( _
            RangeBounds, _
            HorizontalAlignment:=HorizontalAlignment, _
            VerticalAlignment:=VerticalAlignment, _
            Orientation:=Orientation, _
            IndentLevel:=IndentLevel)
End Sub
