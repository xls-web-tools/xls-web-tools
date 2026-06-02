VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverClientTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IWebDriverClient

' #############################################################################
'!
'! @brief
'! WebDriver HTTP API 呼び出しのテストダブルです。
'!
' #############################################################################

'* スタブ設定と正常に成立した呼び出し履歴。
Public Store As TestDoubleBehaviorStore

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* WebDriver HTTP API へ 1 コマンドを送信します。
'*
'* @param HttpMethod HTTP method。
'* @param EndpointPath WebDriver の endpoint path。
'* @param RequestBody JSON request body。
'* @return Store に登録された JSON response body。未登録の場合は空文字列。
'*
'* @details
'* TestDoubleBehaviorStore を使って戻り値、明示エラー、正常呼び出し履歴を管理します。
Public Function Execute( _
        ByVal HttpMethod As String, _
        ByVal EndpointPath As String, _
        Optional ByVal RequestBody As String = "") As String

    Call Store.RaiseIfError("Execute", HttpMethod, EndpointPath, RequestBody)

    Dim result_value As String
    result_value = Store.GetReturnOrDefault("Execute", "", HttpMethod, EndpointPath, RequestBody)
    Call Store.RecordCall("Execute", HttpMethod, EndpointPath, RequestBody)

    Execute = result_value
End Function
Private Function IWebDriverClient_Execute( _
        ByVal HttpMethod As String, _
        ByVal EndpointPath As String, _
        Optional ByVal RequestBody As String = "") As String

    IWebDriverClient_Execute = Execute(HttpMethod, EndpointPath, RequestBody:=RequestBody)
End Function
