VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserInputSheetTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IUserInputSheet

' #############################################################################
'!
'! @brief
'! 入力シートにアクセスするためのテストダブルです。
'!
'! @details
'! GetItemRange は strict stub とし、戻り値未登録エラーおよび SetError による明示エラーは呼び出し履歴に残しません。
'!
' #############################################################################

'* スタブ設定と正常に成立した呼び出し履歴。
Public Store As TestDoubleBehaviorStore

'* 入力領域の safe default です。未登録時は InputSheet シート全体を返します。
Public Property Get InputArea() As WorksheetRangeBounds
    Dim default_area As WorksheetRangeBounds
    Set default_area = New_RangeBounds(Sheet:="InputSheet", Book:="ThisWorkbook.xlsm")
    Set InputArea = Store.GetReturnOrDefault("InputArea", default_area)
End Property
Private Property Get IUserInputSheet_InputArea() As WorksheetRangeBounds
    Set IUserInputSheet_InputArea = InputArea
End Property

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* GetItemRange の strict stub です。戻り値は Store に明示登録してください。
Public Function GetItemRange( _
        ByVal FirstItemName As String, _
        Optional ByVal SecondItemName As String = "") As WorksheetRangeBounds

    Call Store.RaiseIfError("GetItemRange", FirstItemName, SecondItemName)

    Dim result_value As WorksheetRangeBounds
    Set result_value = Store.GetReturn("GetItemRange", FirstItemName, SecondItemName)
    Call Store.RecordCall("GetItemRange", FirstItemName, SecondItemName)

    Set GetItemRange = result_value
End Function
Private Function IUserInputSheet_GetItemRange( _
        ByVal FirstItemName As String, _
        Optional ByVal SecondItemName As String = "") As WorksheetRangeBounds

    Set IUserInputSheet_GetItemRange = GetItemRange(FirstItemName, SecondItemName:=SecondItemName)
End Function
