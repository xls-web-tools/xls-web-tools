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

'* 取得開始 URL へ遷移します。
'*
'* @details
'* settings の StartUrl を WebDriver の URL command へ渡します。認証情報の保存、解析、ログ出力は行いません。
Public Sub NavigateToStartUrl()
    Call pCheckInit

    If pSessionId = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver session が作成されていません。"
    End If

    Dim response_body As String
    response_body = pClient.Execute("POST", "/session/" & pSessionId & "/url", "{""url"":""" & pJsonEscape(pSettings.StartUrl) & """}")
    Call pRaiseIfWebDriverError(response_body)
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

'* selector が出現するまで待機します。
'*
'* @param Selector 待機する CSS selector。
'* @return 見つかった element id。
'*
'* @details
'* settings の TimeoutSeconds まで WebDriver element command を繰り返します。URL ではなく selector 出現を到達判定に使います。
Public Function WaitForSelector(ByVal Selector As String) As String
    Call pCheckInit
    Call pRequireSession

    If Trim$(Selector) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="Selector は必須です。"
    End If

    Dim deadline_time As Date
    deadline_time = DateAdd("s", pSettings.TimeoutSeconds, Now)

    Do
        Dim response_body As String
        response_body = pClient.Execute("POST", "/session/" & pSessionId & "/element", pBuildFindElementBody("css selector", Selector))

        If Not pIsWebDriverError(response_body) Then
            WaitForSelector = pExtractElementId(response_body)
            If WaitForSelector = "" Then
                Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver element id を取得できません。(" & response_body & ")"
            End If
            Exit Function
        End If

        If Not pIsNoSuchElement(response_body) Then
            Call pRaiseIfWebDriverError(response_body)
        End If
        DoEvents
    Loop While Now < deadline_time

    Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="selector が見つかりません。(" & Selector & ")"
End Function

Private Function pBuildFindElementBody(ByVal LocatorType As String, ByVal LocatorValue As String) As String
    pBuildFindElementBody = "{""using"":""" & pJsonEscape(LocatorType) & """,""value"":""" & pJsonEscape(LocatorValue) & """}"
End Function

Private Function pExtractElementId(ByVal ResponseBody As String) As String
    Dim result_value As String
    result_value = pExtractJsonString(ResponseBody, "element-6066-11e4-a52e-4f735466cecf")
    If result_value = "" Then result_value = pExtractJsonString(ResponseBody, "ELEMENT")

    pExtractElementId = result_value
End Function

Private Function pIsWebDriverError(ByVal ResponseBody As String) As Boolean
    pIsWebDriverError = (pExtractJsonString(ResponseBody, "error") <> "")
End Function

Private Function pIsNoSuchElement(ByVal ResponseBody As String) As Boolean
    pIsNoSuchElement = (LCase$(pExtractJsonString(ResponseBody, "error")) = "no such element")
End Function

'* 画面遷移操作を実行します。
'*
'* @param Operation 実行する画面遷移操作定義。
'*
'* @details
'* locator がある場合は WebDriver click を第一候補にします。locator がない場合、または click が失敗して Script が指定されている場合だけ script 実行に切り替えます。
Public Sub RunTransitionOperation(ByVal Operation As TransitionOperation)
    Call pCheckInit
    Call pRequireSession

    If Operation Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="Operation は必須です。"
    End If

    If Trim$(Operation.LocatorType) <> "" Then
        On Error GoTo CLICK_ERROR
        Call pClickElement(Operation.LocatorType, Operation.LocatorValue)
        On Error GoTo 0
        Exit Sub
    End If

    Call pExecuteScript(Operation.Script)
    Exit Sub

CLICK_ERROR:
    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    Dim err_help_file As String
    Dim err_help_context As Long
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description
    err_help_file = Err.HelpFile
    err_help_context = Err.HelpContext

    If Trim$(Operation.Script) = "" Then
        Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
    End If

    Err.Clear
    On Error GoTo 0
    Call pExecuteScript(Operation.Script)
End Sub

Private Sub pClickElement(ByVal LocatorType As String, ByVal LocatorValue As String)
    Dim element_id As String
    element_id = pFindElement(LocatorType, LocatorValue)

    Dim response_body As String
    response_body = pClient.Execute("POST", "/session/" & pSessionId & "/element/" & element_id & "/click", "{}")
    Call pRaiseIfWebDriverError(response_body)
End Sub

Private Sub pExecuteScript(ByVal Script As String)
    If Trim$(Script) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="Script は必須です。"
    End If

    Dim response_body As String
    response_body = pClient.Execute("POST", "/session/" & pSessionId & "/execute/sync", "{""script"":""" & pJsonEscape(Script) & """,""args"":[]}")
    Call pRaiseIfWebDriverError(response_body)
End Sub

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

Private Sub pRequireSession()
    If pSessionId = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver session が作成されていません。"
    End If
End Sub

Private Function pFindElement(ByVal LocatorType As String, ByVal LocatorValue As String) As String
    Dim response_body As String
    response_body = pClient.Execute("POST", "/session/" & pSessionId & "/element", pBuildFindElementBody(LocatorType, LocatorValue))
    Call pRaiseIfWebDriverError(response_body)

    pFindElement = pExtractElementId(response_body)
    If pFindElement = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver element id を取得できません。(" & response_body & ")"
    End If
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="初期化されていません。"
    End If
End Sub

Private Sub pRaiseIfWebDriverError(ByVal ResponseBody As String)
    Dim error_name As String
    error_name = pExtractJsonString(ResponseBody, "error")
    If error_name = "" Then Exit Sub

    Dim error_message As String
    error_message = pExtractJsonString(ResponseBody, "message")

    Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver error: " & error_name & "。" & error_message
End Sub
