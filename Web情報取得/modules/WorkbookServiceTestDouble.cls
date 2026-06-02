VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookServiceTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWorkbookService

' #############################################################################
'!
'! @brief
'! Excel ファイルのオープン、クローズ、書き込みなどを管理するサービスのテストダブルです。
'!
' #############################################################################

'* スタブ設定と正常に成立した呼び出し履歴。
'* 戻り値が必要で外部状態や親リソースに依存するものは strict stub、存在確認など保守的な既定値が自然なものは safe default とします。
'* 戻り値なしの操作は spy only とし、strict stub の未登録エラーおよび SetError による明示エラーは呼び出し履歴に残しません。
'* spy only は、元メソッド定義の順序で呼び出し引数を記録します。
Public Store As TestDoubleBehaviorStore

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* GetThisWorkbookName の safe default です。未登録時は "ThisWorkbook.xlsm" を返します。
Public Function GetThisWorkbookName() As String
    Call Store.RaiseIfError("GetThisWorkbookName")

    Dim result_value As String
    result_value = Store.GetReturnOrDefault("GetThisWorkbookName", "ThisWorkbook.xlsm")
    Call Store.RecordCall("GetThisWorkbookName")

    GetThisWorkbookName = result_value
End Function
Private Function IWorkbookService_GetThisWorkbookName() As String
    IWorkbookService_GetThisWorkbookName = GetThisWorkbookName()
End Function

'* GetThisWorkbookDirectoryPath の safe default です。未登録時は未保存 ThisWorkbook 相当として空文字列を返します。
Public Function GetThisWorkbookDirectoryPath() As String
    Call Store.RaiseIfError("GetThisWorkbookDirectoryPath")

    Dim result_value As String
    result_value = Store.GetReturnOrDefault("GetThisWorkbookDirectoryPath", "")
    Call Store.RecordCall("GetThisWorkbookDirectoryPath")

    GetThisWorkbookDirectoryPath = result_value
End Function
Private Function IWorkbookService_GetThisWorkbookDirectoryPath() As String
    IWorkbookService_GetThisWorkbookDirectoryPath = GetThisWorkbookDirectoryPath()
End Function

'* WorkbookExists の safe default です。未登録時は False を返します。
Public Function WorkbookExists(ByVal Book As String) As Boolean
    Call Store.RaiseIfError("WorkbookExists", Book)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("WorkbookExists", False, Book)
    Call Store.RecordCall("WorkbookExists", Book)

    WorkbookExists = result_value
End Function
Private Function IWorkbookService_WorkbookExists(ByVal Book As String) As Boolean
    IWorkbookService_WorkbookExists = WorkbookExists(Book)
End Function

'* GetAllWorkbooks の safe default です。未登録時は GetThisWorkbookName と連動した 1 件配列を返します。
Public Function GetAllWorkbooks() As String()
    Call Store.RaiseIfError("GetAllWorkbooks")

    Dim result_value() As String
    If Store.HasReturn("GetAllWorkbooks") Then
        result_value = Store.GetReturn("GetAllWorkbooks")
    Else
        ReDim result_value(0 To 0) As String
        result_value(0) = pGetThisWorkbookNameForDefault()
    End If

    Call Store.RecordCall("GetAllWorkbooks")
    GetAllWorkbooks = result_value
End Function
Private Function IWorkbookService_GetAllWorkbooks() As String()
    IWorkbookService_GetAllWorkbooks = GetAllWorkbooks()
End Function

Private Function pGetThisWorkbookNameForDefault() As String
    Call Store.RaiseIfError("GetThisWorkbookName")
    pGetThisWorkbookNameForDefault = Store.GetReturnOrDefault("GetThisWorkbookName", "ThisWorkbook.xlsm")
End Function

'* GetOtherWorkbooks の safe default です。未登録時は他ブックなしを表す空配列を返します。
Public Function GetOtherWorkbooks() As String()
    Call Store.RaiseIfError("GetOtherWorkbooks")

    Dim result_value() As String
    If Store.HasReturn("GetOtherWorkbooks") Then
        result_value = Store.GetReturn("GetOtherWorkbooks")
    Else
        result_value = EmptyStringArray()
    End If

    Call Store.RecordCall("GetOtherWorkbooks")
    GetOtherWorkbooks = result_value
End Function
Private Function IWorkbookService_GetOtherWorkbooks() As String()
    IWorkbookService_GetOtherWorkbooks = GetOtherWorkbooks()
End Function

'* WorksheetExists の safe default です。未登録時は False を返します。
Public Function WorksheetExists(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
    Call Store.RaiseIfError("WorksheetExists", Sheet, Book)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("WorksheetExists", False, Sheet, Book)
    Call Store.RecordCall("WorksheetExists", Sheet, Book)

    WorksheetExists = result_value
End Function
Private Function IWorkbookService_WorksheetExists(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
    IWorkbookService_WorksheetExists = WorksheetExists(Sheet, Book:=Book)
End Function

'* GetAllWorksheets の strict stub です。戻り値は Store に明示登録してください。
Public Function GetAllWorksheets(Optional ByVal Book As String = "") As String()
    Call Store.RaiseIfError("GetAllWorksheets", Book)

    Dim result_value() As String
    result_value = Store.GetReturn("GetAllWorksheets", Book)
    Call Store.RecordCall("GetAllWorksheets", Book)

    GetAllWorksheets = result_value
End Function
Private Function IWorkbookService_GetAllWorksheets(Optional ByVal Book As String = "") As String()
    IWorkbookService_GetAllWorksheets = GetAllWorksheets(Book:=Book)
End Function

'* GetOtherWorksheets の strict stub です。戻り値は Store に明示登録してください。
Public Function GetOtherWorksheets(Optional ByVal Book As String = "") As String()
    Call Store.RaiseIfError("GetOtherWorksheets", Book)

    Dim result_value() As String
    result_value = Store.GetReturn("GetOtherWorksheets", Book)
    Call Store.RecordCall("GetOtherWorksheets", Book)

    GetOtherWorksheets = result_value
End Function
Private Function IWorkbookService_GetOtherWorksheets(Optional ByVal Book As String = "") As String()
    IWorkbookService_GetOtherWorksheets = GetOtherWorksheets(Book:=Book)
End Function

'* OpenWorkbook の strict stub です。戻り値は Store に明示登録してください。
Public Function OpenWorkbook(Optional ByVal FilePath As String = "") As String
    Call Store.RaiseIfError("OpenWorkbook", FilePath)

    Dim result_value As String
    result_value = Store.GetReturn("OpenWorkbook", FilePath)
    Call Store.RecordCall("OpenWorkbook", FilePath)

    OpenWorkbook = result_value
End Function
Private Function IWorkbookService_OpenWorkbook(Optional ByVal FilePath As String = "") As String
    IWorkbookService_OpenWorkbook = OpenWorkbook(FilePath:=FilePath)
End Function

'* SaveWorkbook の strict stub です。戻り値は Store に明示登録してください。
Public Function SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
    Call Store.RaiseIfError("SaveWorkbook", Book, FilePath, Force)

    Dim result_value As String
    result_value = Store.GetReturn("SaveWorkbook", Book, FilePath, Force)
    Call Store.RecordCall("SaveWorkbook", Book, FilePath, Force)

    SaveWorkbook = result_value
End Function
Private Function IWorkbookService_SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
    IWorkbookService_SaveWorkbook = SaveWorkbook(Book, FilePath, Force:=Force)
End Function

'* IsSaved の safe default です。未登録時は False を返します。
Public Function IsSaved(ByVal Book As String) As Boolean
    Call Store.RaiseIfError("IsSaved", Book)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("IsSaved", False, Book)
    Call Store.RecordCall("IsSaved", Book)

    IsSaved = result_value
End Function
Private Function IWorkbookService_IsSaved(ByVal Book As String) As Boolean
    IWorkbookService_IsSaved = IsSaved(Book)
End Function

'* HasPath の safe default です。未登録時は False を返します。
Public Function HasPath(ByVal Book As String) As Boolean
    Call Store.RaiseIfError("HasPath", Book)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("HasPath", False, Book)
    Call Store.RecordCall("HasPath", Book)

    HasPath = result_value
End Function
Private Function IWorkbookService_HasPath(ByVal Book As String) As Boolean
    IWorkbookService_HasPath = HasPath(Book)
End Function

'* CloseWorkbook の spy only です。呼び出しを記録し、失敗は Store.SetError で注入します。
Public Sub CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
    Call Store.RaiseIfError("CloseWorkbook", Book, Force)
    Call Store.RecordCall("CloseWorkbook", Book, Force)
End Sub
Private Sub IWorkbookService_CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
    Call CloseWorkbook(Book, Force:=Force)
End Sub

'* Find の strict stub です。検索結果なしを表す空 ObjectList も明示登録してください。
Public Function Find( _
        ByVal What As String, _
        Optional ByVal Book As String = "", _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList

    Call Store.RaiseIfError("Find", What, Book, LookInValues, LookAtWhole, MatchCase, MatchByte)

    Dim result_value As ObjectList
    Set result_value = Store.GetReturn("Find", What, Book, LookInValues, LookAtWhole, MatchCase, MatchByte)
    Call Store.RecordCall("Find", What, Book, LookInValues, LookAtWhole, MatchCase, MatchByte)

    Set Find = result_value
End Function
Private Function IWorkbookService_Find( _
        ByVal What As String, _
        Optional ByVal Book As String = "", _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList

    Set IWorkbookService_Find = Find( _
            What, Book:=Book, _
            LookInValues:=LookInValues, _
            LookAtWhole:=LookAtWhole, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte)
End Function

'* AddWorksheet の strict stub です。戻り値は Store に明示登録してください。
Public Function AddWorksheet( _
        Optional ByVal Sheet As String = "", _
        Optional ByVal Book As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    Call Store.RaiseIfError("AddWorksheet", Sheet, Book, SheetIndex, Before)

    Dim result_value As String
    result_value = Store.GetReturn("AddWorksheet", Sheet, Book, SheetIndex, Before)
    Call Store.RecordCall("AddWorksheet", Sheet, Book, SheetIndex, Before)

    AddWorksheet = result_value
End Function
Private Function IWorkbookService_AddWorksheet( _
        Optional ByVal Sheet As String = "", _
        Optional ByVal Book As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    IWorkbookService_AddWorksheet = AddWorksheet( _
            Sheet:=Sheet, Book:=Book, _
            SheetIndex:=SheetIndex, Before:=Before)
End Function

'* RemoveWorksheet の spy only です。呼び出しを記録し、失敗は Store.SetError で注入します。
Public Sub RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
    Call Store.RaiseIfError("RemoveWorksheet", Sheet, Book)
    Call Store.RecordCall("RemoveWorksheet", Sheet, Book)
End Sub
Private Sub IWorkbookService_RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
    Call RemoveWorksheet(Sheet:=Sheet, Book:=Book)
End Sub

'* CopyWorksheet の strict stub です。戻り値は Store に明示登録してください。
Public Function CopyWorksheet( _
        Optional ByVal SourceWorksheetName As String = "", _
        Optional ByVal SourceWorkbookName As String = "", _
        Optional ByVal DestinationWorksheetName As String = "", _
        Optional ByVal DestinationWorkbookName As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    Call Store.RaiseIfError("CopyWorksheet", SourceWorksheetName, SourceWorkbookName, DestinationWorksheetName, DestinationWorkbookName, SheetIndex, Before)

    Dim result_value As String
    result_value = Store.GetReturn("CopyWorksheet", SourceWorksheetName, SourceWorkbookName, DestinationWorksheetName, DestinationWorkbookName, SheetIndex, Before)
    Call Store.RecordCall("CopyWorksheet", SourceWorksheetName, SourceWorkbookName, DestinationWorksheetName, DestinationWorkbookName, SheetIndex, Before)

    CopyWorksheet = result_value
End Function
Private Function IWorkbookService_CopyWorksheet( _
        Optional ByVal SourceWorksheetName As String = "", _
        Optional ByVal SourceWorkbookName As String = "", _
        Optional ByVal DestinationWorksheetName As String = "", _
        Optional ByVal DestinationWorkbookName As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    IWorkbookService_CopyWorksheet = CopyWorksheet( _
            SourceWorksheetName:=SourceWorksheetName, _
            SourceWorkbookName:=SourceWorkbookName, _
            DestinationWorksheetName:=DestinationWorksheetName, _
            DestinationWorkbookName:=DestinationWorkbookName, _
            SheetIndex:=SheetIndex, _
            Before:=Before)
End Function

'* ActivateWorksheet の spy only です。呼び出しを記録し、失敗は Store.SetError で注入します。
Public Sub ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
    Call Store.RaiseIfError("ActivateWorksheet", Sheet, Book)
    Call Store.RecordCall("ActivateWorksheet", Sheet, Book)
End Sub
Private Sub IWorkbookService_ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
    Call ActivateWorksheet(Sheet, Book:=Book)
End Sub

'* RemoveVBComponents の spy only です。呼び出しを記録し、失敗は Store.SetError で注入します。
Public Sub RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
    Call Store.RaiseIfError("RemoveVBComponents", ComponentNames, ComponentType, Book)
    Call Store.RecordCall("RemoveVBComponents", ComponentNames, ComponentType, Book)
End Sub
Private Sub IWorkbookService_RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
    Call RemoveVBComponents(ComponentNames, ComponentType:=ComponentType, Book:=Book)
End Sub
