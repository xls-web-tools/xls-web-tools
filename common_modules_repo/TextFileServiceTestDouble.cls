VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFileServiceTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITextFileService

' #############################################################################
'!
'! @brief
'! TextFileEntity を生成するサービスのテストダブルです。
'!
'! @details
'! GetTextFileEntity は lightweight fake の入口です。未登録時は新しい初期化済み TextFileEntityTestDouble を返します。
'! ファイル内容や open 失敗は自動再現しません。任意失敗は Store.SetError で GetTextFileEntity に設定してください。
'!
' #############################################################################

'* スタブ設定と正常に成立した呼び出し履歴。
'* strict stub の未登録エラーおよび SetError による明示エラーは呼び出し履歴に残しません。
Public Store As TestDoubleBehaviorStore

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* GetTextFileEntity の lightweight fake です。未登録時は新しい初期化済み TextFileEntityTestDouble を生成します。
Public Function GetTextFileEntity(ByVal FilePath As String) As ITextFileEntity
    Call Store.RaiseIfError("GetTextFileEntity", FilePath)

    Dim result_value As ITextFileEntity
    If Store.HasReturn("GetTextFileEntity", FilePath) Then
        Set result_value = Store.GetReturn("GetTextFileEntity", FilePath)
        Call pValidateEntityFilePath(result_value, FilePath)
    Else
        Dim file_stub As TextFileEntityTestDouble
        Set file_stub = New TextFileEntityTestDouble
        Call file_stub.Initialize(FilePath)
        Set result_value = file_stub
    End If

    Call Store.RecordCall("GetTextFileEntity", FilePath)
    Set GetTextFileEntity = result_value
End Function
Private Function ITextFileService_GetTextFileEntity(ByVal FilePath As String) As ITextFileEntity
    Set ITextFileService_GetTextFileEntity = GetTextFileEntity(FilePath)
End Function

Private Sub pValidateEntityFilePath(ByVal TextFileEntity As ITextFileEntity, ByVal ExpectedFilePath As String)
    On Error GoTo INVALID_ENTITY

    Dim actual_path As String
    actual_path = TextFileEntity.FilePath

    On Error GoTo 0
    If actual_path <> ExpectedFilePath Then
        Err.Raise vbObjectError + 1, "Class TextFileServiceTestDouble", "登録済み TextFileEntity の FilePath が一致しません。(expected: " & ExpectedFilePath & ", actual: " & actual_path & ")"
    End If
    Exit Sub

INVALID_ENTITY:
    Err.Raise vbObjectError + 1, "Class TextFileServiceTestDouble", "登録済み TextFileEntity が指定パスで初期化されていません。(" & ExpectedFilePath & ")"
End Sub
