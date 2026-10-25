VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverPortProbe"
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
'! WebDriver port の使用状況を確認します。
'!
' #############################################################################

'* port が使用中か確認します。
'*
'* @param PortNumber 確認対象の port 番号。
'* @return 使用中の場合は True。
'*
'* @details
'* `http://127.0.0.1:{port}/status` へ短い timeout で接続を試み、応答できる相手がいる場合は使用中とみなします。
Public Function IsPortInUse(ByVal PortNumber As Long) As Boolean
    On Error GoTo PORT_NOT_IN_USE

    Dim request_obj As Object
    Set request_obj = CreateObject("WinHttp.WinHttpRequest.5.1")
    request_obj.SetTimeouts 300, 300, 300, 300
    request_obj.Open "GET", "http://127.0.0.1:" & CStr(PortNumber) & "/status", False
    request_obj.Send

    IsPortInUse = True
    Exit Function

PORT_NOT_IN_USE:
    Err.Clear
    IsPortInUse = False
End Function
Private Function IWebDriverPortProbe_IsPortInUse(ByVal PortNumber As Long) As Boolean
    IWebDriverPortProbe_IsPortInUse = IsPortInUse(PortNumber)
End Function
