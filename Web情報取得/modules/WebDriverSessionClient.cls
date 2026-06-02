VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverSessionClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Edge WebDriver session の作成と終了を管理します。
'!
' #############################################################################

Private pClient As IWebDriverClient
Private pSettings As IToolSettings
Private pSessionId As String
Private pIsInitialized As Boolean

'* 作成済み session id。
Public Property Get SessionId() As String
    Call pCheckInit
    SessionId = pSessionId
End Property

'* WebDriver session client を初期化します。
'*
'* @param Client WebDriver HTTP API 呼び出しクライアント。
'* @param Settings Web情報取得の設定。
'*
'* @details
'* Client と Settings は Nothing にできません。session 作成時の capabilities は Settings から組み立てます。
Public Sub Initialize(ByVal Client As IWebDriverClient, ByVal Settings As IToolSettings)
    If Client Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="Client は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="Settings は必須です。"
    End If

    Set pClient = Client
    Set pSettings = Settings
    pSessionId = ""
    pIsInitialized = True
End Sub

'* Edge WebDriver session を作成します。
'*
'* @return 作成された session id。
'*
'* @details
'* settings の Headless と BrowserProfilePath を capabilities に反映し、`POST /session` を送信します。
'* response から `sessionId` を抽出し、抽出できない場合は原因が分かるエラーとして再送出します。
Public Function CreateSession() As String
    Call pCheckInit

    Dim response_body As String
    response_body = pClient.Execute("POST", "/session", pBuildCreateSessionBody())
    Call pRaiseIfWebDriverError(response_body)

    pSessionId = pExtractJsonString(response_body, "sessionId")
    If pSessionId = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver sessionId を取得できません。(" & response_body & ")"
    End If

    CreateSession = pSessionId
End Function

Private Function pBuildCreateSessionBody() As String
    Dim args_text As String
    args_text = """--user-data-dir=" & pJsonEscape(pSettings.BrowserProfilePath) & """"
    If pSettings.Headless Then args_text = args_text & ",""--headless=new"""

    pBuildCreateSessionBody = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{""args"": [" & args_text & "]}}}}"
    pBuildCreateSessionBody = Replace(pBuildCreateSessionBody, """args"": [", """args"":[")
End Function

Private Function pExtractJsonString(ByVal JsonText As String, ByVal KeyName As String) As String
    Dim target_re As RegExp
    Set target_re = New RegExp
    target_re.Global = False
    target_re.IgnoreCase = False
    target_re.Pattern = """" & KeyName & """\s*:\s*""([^""]*)"""

    Dim result_value As String
    result_value = ""
    If target_re.Test(JsonText) Then
        result_value = target_re.Execute(JsonText)(0).SubMatches(0)
    End If

    pExtractJsonString = result_value
End Function

'* Edge WebDriver session を終了します。
'*
'* @param SessionId [省略可] 終了する session id。省略時は CreateSession で保持した session id を使います。
'*
'* @details
'* `DELETE /session/{id}` を送信します。session id がない場合は終了対象不足としてエラーにします。
Public Sub QuitSession(Optional ByVal SessionId As String = "")
    Call pCheckInit

    Dim target_session_id As String
    target_session_id = SessionId
    If target_session_id = "" Then target_session_id = pSessionId

    If target_session_id = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="終了対象の WebDriver session id がありません。"
    End If

    Dim response_body As String
    response_body = pClient.Execute("DELETE", "/session/" & target_session_id, "")
    Call pRaiseIfWebDriverError(response_body)

    If pSessionId = target_session_id Then pSessionId = ""
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="初期化されていません。"
    End If
End Sub

Private Function pJsonEscape(ByVal TargetText As String) As String
    Dim result_value As String
    result_value = TargetText
    result_value = Replace(result_value, "\", "\\")
    result_value = Replace(result_value, """", "\""")
    result_value = Replace(result_value, vbCrLf, "\n")
    result_value = Replace(result_value, vbCr, "\n")
    result_value = Replace(result_value, vbLf, "\n")

    pJsonEscape = result_value
End Function

Private Sub pRaiseIfWebDriverError(ByVal ResponseBody As String)
    Dim error_name As String
    error_name = pExtractJsonString(ResponseBody, "error")
    If error_name = "" Then Exit Sub

    Dim error_message As String
    error_message = pExtractJsonString(ResponseBody, "message")

    Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver error: " & error_name & "。" & error_message
End Sub
