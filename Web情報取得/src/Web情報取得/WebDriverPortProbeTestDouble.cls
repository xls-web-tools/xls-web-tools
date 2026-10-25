VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverPortProbeTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IWebDriverPortProbe

' #############################################################################
'!
'! @brief
'! WebDriver port probe のテストダブルです。
'!
' #############################################################################

'* スタブ設定と正常に成立した呼び出し履歴。
Public Store As TestDoubleBehaviorStore

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* port が使用中か確認します。
'*
'* @param PortNumber 確認対象の port 番号。
'* @return Store に登録された使用中判定。未登録の場合は False。
'*
'* @details
'* Port 使用中のケースだけ Store.SetReturn で True を設定します。
Public Function IsPortInUse(ByVal PortNumber As Long) As Boolean
    Call Store.RaiseIfError("IsPortInUse", PortNumber)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("IsPortInUse", False, PortNumber)
    Call Store.RecordCall("IsPortInUse", PortNumber)

    IsPortInUse = result_value
End Function
Private Function IWebDriverPortProbe_IsPortInUse(ByVal PortNumber As Long) As Boolean
    IWebDriverPortProbe_IsPortInUse = IsPortInUse(PortNumber)
End Function
