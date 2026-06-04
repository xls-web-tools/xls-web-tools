VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverClient"
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
'! Edge WebDriver HTTP API 呼び出しの実装です。
'!
' #############################################################################

Private pPortNumber As Long
Private pTimeoutSeconds As Long
Private pIsInitialized As Boolean

'* WebDriver HTTP client を初期化します。
'*
'* @param PortNumber Edge WebDriver の待ち受け port。
'* @param TimeoutSeconds HTTP timeout 秒数。
'*
'* @details
'* `http://127.0.0.1:{port}` を基準 URL として WebDriver HTTP API を呼び出します。
Public Sub Initialize(ByVal PortNumber As Long, ByVal TimeoutSeconds As Long)
    If PortNumber <= 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverClient", Description:="PortNumber は 1 以上である必要があります。"
    End If
    If TimeoutSeconds <= 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverClient", Description:="TimeoutSeconds は 1 以上である必要があります。"
    End If

    pPortNumber = PortNumber
    pTimeoutSeconds = TimeoutSeconds
    pIsInitialized = True
End Sub

'* WebDriver HTTP API へ 1 コマンドを送信します。
'*
'* @param HttpMethod HTTP method。
'* @param EndpointPath WebDriver の endpoint path。
'* @param RequestBody JSON request body。
'* @return JSON response body。
'*
'* @details
'* WinHTTP を late bind で使い、追加ライブラリを導入せずに WebDriver HTTP API を呼び出します。
Public Function Execute( _
        ByVal HttpMethod As String, _
        ByVal EndpointPath As String, _
        Optional ByVal RequestBody As String = "") As String

    Call pCheckInit

    Dim request_obj As Object
    Set request_obj = CreateObject("WinHttp.WinHttpRequest.5.1")

    Dim target_url As String
    target_url = pBuildUrl(EndpointPath)

    Dim timeout_ms As Long
    timeout_ms = pTimeoutSeconds * 1000

    On Error GoTo REQUEST_ERROR

    request_obj.SetTimeouts timeout_ms, timeout_ms, timeout_ms, timeout_ms
    request_obj.Open HttpMethod, target_url, False
    request_obj.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    request_obj.Send RequestBody

    Dim response_body As String
    response_body = CStr(request_obj.ResponseText)

    Dim status_code As Long
    status_code = CLng(request_obj.Status)

    On Error GoTo 0
    If status_code < 200 Or 300 <= status_code Then
        If Not pIsWebDriverErrorResponse(response_body) Then
            Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverClient", Description:="WebDriver HTTP error " & CStr(status_code) & "。" & response_body
        End If
    End If

    Execute = response_body
    Exit Function

REQUEST_ERROR:
    Dim err_num As Long
    Dim err_desc As String
    Dim err_help_file As String
    Dim err_help_context As Long
    err_num = Err.Number
    err_desc = Err.Description
    err_help_file = Err.HelpFile
    err_help_context = Err.HelpContext

    Err.Raise err_num, "Class WebDriverClient", "WebDriver HTTP API に接続できません。(" & HttpMethod & " " & target_url & ") " & err_desc, err_help_file, err_help_context
End Function
Private Function IWebDriverClient_Execute( _
        ByVal HttpMethod As String, _
        ByVal EndpointPath As String, _
        Optional ByVal RequestBody As String = "") As String

    IWebDriverClient_Execute = Execute(HttpMethod, EndpointPath, RequestBody:=RequestBody)
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverClient", Description:="初期化されていません。"
    End If
End Sub

Private Function pIsWebDriverErrorResponse(ByVal ResponseBody As String) As Boolean
    pIsWebDriverErrorResponse = (0 < InStr(1, ResponseBody, """error""", vbTextCompare))
End Function

Private Function pBuildUrl(ByVal EndpointPath As String) As String
    If Left$(EndpointPath, 1) = "/" Then
        pBuildUrl = "http://127.0.0.1:" & CStr(pPortNumber) & EndpointPath
    Else
        pBuildUrl = "http://127.0.0.1:" & CStr(pPortNumber) & "/" & EndpointPath
    End If
End Function
