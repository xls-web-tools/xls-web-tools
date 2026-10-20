VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebDriverClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'#Interface

' #############################################################################
'!
'! @brief
'! Edge WebDriver HTTP API 呼び出しを差し替えるためのインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IWebDriverClient", Description:="このクラスはインターフェイス用です。"
End Sub

'* WebDriver HTTP API へ 1 コマンドを送信します。
'*
'* @param HttpMethod HTTP method。
'* @param EndpointPath WebDriver の endpoint path。
'* @param RequestBody JSON request body。
'* @return JSON response body。
'*
'* @details
'* 呼び出し元は WebDriver endpoint path と JSON request body を組み立てて渡します。
'* 実装側は HTTP 通信、エラー応答、呼び出し履歴の扱いをこの契約の内側で処理します。
Public Function Execute( _
        ByVal HttpMethod As String, _
        ByVal EndpointPath As String, _
        Optional ByVal RequestBody As String = "") As String
End Function
