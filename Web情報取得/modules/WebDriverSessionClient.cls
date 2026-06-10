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

Private Const C_DOWNLOAD_DIRECTORY_PREFIX As String = "xls-web-tools_"

Private pClient As IWebDriverClient
Private pSettings As IToolSettings
Private pSessionId As String
Private pDownloadDirectoryPath As String
Private pFileSystem As IFileSystemService
Private pIsInitialized As Boolean

'* 作成済み session id。
Public Property Get SessionId() As String
    Call pCheckInit
    SessionId = pSessionId
End Property

'* WebDriver session の実ダウンロード先として指定した一時ディレクトリ。
Public Property Get DownloadDirectoryPath() As String
    Call pCheckInit
    DownloadDirectoryPath = pDownloadDirectoryPath
End Property


'* WebDriver session client を初期化します。
'*
'* @param Client WebDriver HTTP API 呼び出しクライアント。
'* @param Settings Web情報取得の設定。
'* @param FileSystem ファイルシステム操作サービス。
'*
'* @details
'* Client と Settings は Nothing にできません。session 作成時の capabilities は Settings から組み立てます。
Public Sub Initialize( _
        ByVal Client As IWebDriverClient, _
        ByVal Settings As IToolSettings, _
        Optional ByVal FileSystem As IFileSystemService = Nothing)

    If Client Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="Client は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="Settings は必須です。"
    End If
    If Settings.DownloadEnabled And (FileSystem Is Nothing) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="FileSystem は必須です。"
    End If

    Set pClient = Client
    Set pSettings = Settings
    Set pFileSystem = FileSystem
    pSessionId = ""
    pDownloadDirectoryPath = ""
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

    Dim edge_options_text As String
    edge_options_text = """args"": [" & args_text & "]"
    edge_options_text = Replace(edge_options_text, """args"": [", """args"":[")

    If pSettings.DownloadEnabled Then
        pDownloadDirectoryPath = pCreateTemporaryDownloadDirectory()
        edge_options_text = edge_options_text _
                & ",""prefs"":{""download.default_directory"":""" & pJsonEscape(pDownloadDirectoryPath) & """,""download.prompt_for_download"":false,""download.directory_upgrade"":true}"
    Else
        pDownloadDirectoryPath = ""
    End If

    pBuildCreateSessionBody = "{""capabilities"":{""alwaysMatch"":{""browserName"":""MicrosoftEdge"",""ms:edgeOptions"":{" & edge_options_text & "}}}}"
End Function

Private Function pCreateTemporaryDownloadDirectory() As String
    If pFileSystem Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="FileSystem は必須です。"
    End If

    pCreateTemporaryDownloadDirectory = pFileSystem.CreateTemporaryDirectory(C_DOWNLOAD_DIRECTORY_PREFIX)
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
        response_body = pFindElementResponse("css selector", Selector)

        If Not pIsWebDriverError(response_body) Then
            WaitForSelector = pExtractElementId(response_body)
            If WaitForSelector = "" Then
                Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver element id を取得できません。(" & response_body & ")"
            End If
            Call pSwitchToTopLevelContext
            Exit Function
        End If

        If Not pIsNoSuchElement(response_body) Then
            Call pRaiseIfWebDriverError(response_body)
        End If
        DoEvents
    Loop While Now < deadline_time

    On Error Resume Next
    Call pSwitchToTopLevelContext
    On Error GoTo 0

    Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="selector が見つかりません。(" & Selector & ")"
End Function

'* selector に一致する要素数を数えます。
'*
'* @param Selector CSS selector。
'* @return selector に一致する要素数。
'*
'* @details
'* frame chain selector は WaitForSelector と同じ `frame selector >> target selector` 形式を受け付けます。
'* 0 件はエラーではなく 0 を返します。戻る前に top-level context へ戻します。
Public Function CountElements(ByVal Selector As String) As Long
    Call pCheckInit
    Call pRequireSession

    If Trim$(Selector) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="Selector は必須です。"
    End If

    Dim response_body As String
    response_body = pFindElementsResponse("css selector", Selector)
    Call pRaiseIfWebDriverError(response_body)

    CountElements = pCountElementIds(response_body)
    Call pSwitchToTopLevelContext
End Function

'* ダウンロードリンクの browser-resolved href が HTTP(S) URL なら遷移します。
'*
'* @param LocatorType locator 種別。
'* @param LocatorValue locator 値。
'* @return URL 遷移を実行したか。
Public Function TryNavigateDownloadLink(ByVal LocatorType As String, ByVal LocatorValue As String) As Boolean
    Call pCheckInit
    Call pRequireSession

    TryNavigateDownloadLink = pTryNavigateDownloadLink(LocatorType, LocatorValue)
End Function

'* 要素をクリックします。
'*
'* @param LocatorType locator 種別。
'* @param LocatorValue locator 値。
Public Sub ClickElement(ByVal LocatorType As String, ByVal LocatorValue As String)
    Call pCheckInit
    Call pRequireSession

    Call pClickElement(LocatorType, LocatorValue)
End Sub

'* selector に一致する要素の表示テキストを読み取ります。
'*
'* @param Selector CSS selector。
'* @return 要素の表示テキスト。
'*
'* @details
'* frame chain selector は WaitForSelector と同じ `frame selector >> target selector` 形式を受け付けます。
'* WebDriver の element text command を使い、URL 直叩きや DOM の直接構築は行いません。
Public Function ReadElementText(ByVal Selector As String) As String
    Call pCheckInit
    Call pRequireSession

    If Trim$(Selector) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="Selector は必須です。"
    End If

    Dim element_id As String
    element_id = pFindElement("css selector", Selector)

    Dim response_body As String
    response_body = pClient.Execute("GET", "/session/" & pSessionId & "/element/" & element_id & "/text", "")
    Call pRaiseIfWebDriverError(response_body)

    Dim result_value As String
    result_value = pExtractJsonString(response_body, "value")

    Call pSwitchToTopLevelContext

    ReadElementText = result_value
End Function

'* 詳細ページ列定義に対応する値を読み取ります。
'*
'* @param DetailColumnDefinitions 詳細ページ列定義の一覧。
'* @return 列定義と同じ順序の抽出値。
'*
'* @details
'* WebDriver の execute/sync で DOM から列定義順の配列を一括抽出します。
'* frame chain selector、TextContent、InnerText、Attribute の抽出種別を受け付けます。
Public Function ReadDetailColumnValues(ByVal DetailColumnDefinitions As ObjectList) As ArrayObject
    Call pCheckInit
    Call pRequireSession

    If DetailColumnDefinitions Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="DetailColumnDefinitions は必須です。"
    End If

    Dim response_body As String
    response_body = pClient.Execute( _
            "POST", _
            "/session/" & pSessionId & "/execute/sync", _
            pBuildDetailExtractionScriptBody(DetailColumnDefinitions))
    Call pRaiseIfWebDriverError(response_body)

    If response_body <> "" Then
        Set ReadDetailColumnValues = pExtractJsonStringArray(response_body, "value")
        Exit Function
    End If

    Set ReadDetailColumnValues = pReadDetailColumnValuesByElementText(DetailColumnDefinitions)
End Function

Private Function pReadDetailColumnValuesByElementText(ByVal DetailColumnDefinitions As ObjectList) As ArrayObject
    Dim result_values As ArrayObject
    Set result_values = New ArrayObject

    If 0 < DetailColumnDefinitions.Count Then
        Call result_values.ReDimArray(0, DetailColumnDefinitions.Count - 1)
    End If

    Dim def_idx As Long
    For def_idx = 0 To DetailColumnDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = DetailColumnDefinitions.Item(def_idx)
        Call result_values.Update(def_idx, Trim$(ReadElementText(detail_def.Selector)))
    Next def_idx

    Set pReadDetailColumnValuesByElementText = result_values
End Function

Private Function pBuildDetailExtractionScriptBody(ByVal DetailColumnDefinitions As ObjectList) As String
    pBuildDetailExtractionScriptBody = "{""script"":""" & pJsonEscape(pDetailExtractionScript()) & """,""args"":[" & pBuildDetailDefinitionJson(DetailColumnDefinitions) & "]}"
End Function

Private Function pDetailExtractionScript() As String
    Dim script_text As String
    script_text = "return (function(defs) {"
    script_text = script_text & "function trimValue(value) { if (value === null || value === undefined) { return ''; } return String(value).replace(/^\s+|\s+$/g, ''); }"
    script_text = script_text & "function findElement(selector) { var parts = String(selector).split('>>'); var root = document; for (var i = 0; i < parts.length - 1; i++) { var frame = root.querySelector(trimValue(parts[i])); if (!frame || !frame.contentDocument) { return null; } root = frame.contentDocument; } return root.querySelector(trimValue(parts[parts.length - 1])); }"
    script_text = script_text & "return defs.map(function(def) { var element = findElement(def.selector); if (!element) { if (def.isRequired) { throw new Error('Required selector not found: ' + def.outputColumnName); } return ''; } var extractType = String(def.extractType || 'InnerText').toLowerCase(); var value = ''; if (extractType === 'attribute') { value = element.getAttribute(def.attributeName); } else if (extractType === 'textcontent') { value = element.textContent; } else { value = element.innerText; } value = trimValue(value); if (value === '' && def.isRequired && String(def.blankMode || '').toLowerCase() === 'errorifblank') { throw new Error('Required value is blank: ' + def.outputColumnName); } return value; });"
    script_text = script_text & "})(arguments[0]);"

    pDetailExtractionScript = script_text
End Function

Private Function pBuildDetailDefinitionJson(ByVal DetailColumnDefinitions As ObjectList) As String
    Dim result_text As String
    result_text = "["

    Dim def_idx As Long
    For def_idx = 0 To DetailColumnDefinitions.Count - 1
        If 0 < def_idx Then result_text = result_text & ","

        Dim detail_def As DetailColumnDefinition
        Set detail_def = DetailColumnDefinitions.Item(def_idx)

        result_text = result_text & "{" _
                & """outputColumnName"":""" & pJsonEscape(detail_def.OutputColumnName) & """," _
                & """selector"":""" & pJsonEscape(detail_def.Selector) & """," _
                & """extractType"":""" & pJsonEscape(detail_def.ExtractType) & """," _
                & """attributeName"":""" & pJsonEscape(detail_def.AttributeName) & """," _
                & """isRequired"":" & pJsonBoolean(detail_def.IsRequired) & "," _
                & """blankMode"":""" & pJsonEscape(detail_def.BlankMode) & """}"
    Next def_idx

    result_text = result_text & "]"
    pBuildDetailDefinitionJson = result_text
End Function

Private Function pJsonBoolean(ByVal TargetValue As Boolean) As String
    If TargetValue Then
        pJsonBoolean = "true"
    Else
        pJsonBoolean = "false"
    End If
End Function

Private Function pExtractJsonStringArray(ByVal JsonText As String, ByVal KeyName As String) As ArrayObject
    Dim key_idx As Long
    key_idx = InStr(1, JsonText, """" & KeyName & """", vbBinaryCompare)
    If key_idx = 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="JSON 配列のキーを取得できません。(" & KeyName & ")"
    End If

    Dim open_idx As Long
    open_idx = InStr(key_idx, JsonText, "[", vbBinaryCompare)

    Dim close_idx As Long
    close_idx = InStrRev(JsonText, "]", -1, vbBinaryCompare)
    If open_idx = 0 Or close_idx <= open_idx Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="JSON 配列を取得できません。(" & JsonText & ")"
    End If

    Dim array_text As String
    array_text = Mid$(JsonText, open_idx + 1, close_idx - open_idx - 1)

    Dim value_re As RegExp
    Set value_re = New RegExp
    value_re.Global = True
    value_re.IgnoreCase = False
    value_re.Pattern = Chr$(34) & "((\\.|[^" & Chr$(34) & "])*)" & Chr$(34)

    Dim matches As MatchCollection
    Set matches = value_re.Execute(array_text)

    Dim result_values As ArrayObject
    Set result_values = New ArrayObject
    If 0 < matches.Count Then
        Call result_values.ReDimArray(0, matches.Count - 1)
    End If

    Dim match_idx As Long
    For match_idx = 0 To matches.Count - 1
        Call result_values.Update(match_idx, pJsonUnescape(CStr(matches.Item(match_idx).SubMatches(0))))
    Next match_idx

    Set pExtractJsonStringArray = result_values
End Function

Private Function pJsonUnescape(ByVal TargetText As String) As String
    Dim result_value As String
    result_value = TargetText
    result_value = Replace(result_value, "\n", vbLf)
    result_value = Replace(result_value, "\r", vbCr)
    result_value = Replace(result_value, "\t", vbTab)
    result_value = Replace(result_value, "\" & Chr$(34), Chr$(34))
    result_value = Replace(result_value, "\\", "\")

    pJsonUnescape = result_value
End Function

Private Function pCountElementIds(ByVal ResponseBody As String) As Long
    Dim key_idx As Long
    key_idx = InStr(1, ResponseBody, """value""", vbBinaryCompare)
    If key_idx = 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver elements response の value を取得できません。(" & ResponseBody & ")"
    End If

    Dim open_idx As Long
    open_idx = InStr(key_idx, ResponseBody, "[", vbBinaryCompare)

    Dim close_idx As Long
    close_idx = InStrRev(ResponseBody, "]", -1, vbBinaryCompare)
    If open_idx = 0 Or close_idx < open_idx Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver elements response の配列を取得できません。(" & ResponseBody & ")"
    End If

    Dim array_text As String
    array_text = Mid$(ResponseBody, open_idx + 1, close_idx - open_idx - 1)

    Dim element_re As RegExp
    Set element_re = New RegExp
    element_re.Global = True
    element_re.IgnoreCase = False
    element_re.Pattern = """(element-6066-11e4-a52e-4f735466cecf|ELEMENT)""\s*:"

    pCountElementIds = element_re.Execute(array_text).Count
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


Private Function pFindElementResponse(ByVal LocatorType As String, ByVal LocatorValue As String) As String
    Dim target_locator_type As String
    target_locator_type = LocatorType

    Dim target_locator_value As String
    target_locator_value = LocatorValue

    Call pSwitchToTopLevelContext

    If pIsFrameChainLocator(LocatorType, LocatorValue) Then
        Dim locator_parts As Variant
        locator_parts = pSplitFrameChain(LocatorValue)

        Dim part_idx As Long
        For part_idx = LBound(locator_parts) To UBound(locator_parts) - 1
            Dim frame_selector As String
            frame_selector = Trim$(CStr(locator_parts(part_idx)))
            If frame_selector = "" Then
                Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="frame selector は空にできません。(" & LocatorValue & ")"
            End If

            Dim frame_response As String
            frame_response = pClient.Execute("POST", "/session/" & pSessionId & "/element", pBuildFindElementBody("css selector", frame_selector))
            If pIsWebDriverError(frame_response) Then
                pFindElementResponse = frame_response
                Exit Function
            End If

            Dim frame_element_id As String
            frame_element_id = pExtractElementId(frame_response)
            If frame_element_id = "" Then
                Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="frame element id を取得できません。(" & frame_response & ")"
            End If

            Call pSwitchToFrame(frame_element_id)
        Next part_idx

        target_locator_type = "css selector"
        target_locator_value = Trim$(CStr(locator_parts(UBound(locator_parts))))
        If target_locator_value = "" Then
            Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="target selector は空にできません。(" & LocatorValue & ")"
        End If
    End If

    pFindElementResponse = pClient.Execute("POST", "/session/" & pSessionId & "/element", pBuildFindElementBody(target_locator_type, target_locator_value))
End Function

Private Function pFindElementsResponse(ByVal LocatorType As String, ByVal LocatorValue As String) As String
    Dim target_locator_type As String
    target_locator_type = LocatorType

    Dim target_locator_value As String
    target_locator_value = LocatorValue

    Call pSwitchToTopLevelContext

    If pIsFrameChainLocator(LocatorType, LocatorValue) Then
        Dim locator_parts As Variant
        locator_parts = pSplitFrameChain(LocatorValue)

        Dim part_idx As Long
        For part_idx = LBound(locator_parts) To UBound(locator_parts) - 1
            Dim frame_selector As String
            frame_selector = Trim$(CStr(locator_parts(part_idx)))
            If frame_selector = "" Then
                Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="frame selector は空にできません。(" & LocatorValue & ")"
            End If

            Dim frame_response As String
            frame_response = pClient.Execute("POST", "/session/" & pSessionId & "/element", pBuildFindElementBody("css selector", frame_selector))
            If pIsWebDriverError(frame_response) Then
                pFindElementsResponse = frame_response
                Exit Function
            End If

            Dim frame_element_id As String
            frame_element_id = pExtractElementId(frame_response)
            If frame_element_id = "" Then
                Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="frame element id を取得できません。(" & frame_response & ")"
            End If

            Call pSwitchToFrame(frame_element_id)
        Next part_idx

        target_locator_type = "css selector"
        target_locator_value = Trim$(CStr(locator_parts(UBound(locator_parts))))
        If target_locator_value = "" Then
            Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="target selector は空にできません。(" & LocatorValue & ")"
        End If
    End If

    pFindElementsResponse = pClient.Execute("POST", "/session/" & pSessionId & "/elements", pBuildFindElementBody(target_locator_type, target_locator_value))
End Function

Private Sub pSwitchToTopLevelContext()
    Dim response_body As String
    response_body = pClient.Execute("POST", "/session/" & pSessionId & "/frame", "{""id"":null}")
    Call pRaiseIfWebDriverError(response_body)
End Sub

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

Private Function pTryNavigateDownloadLink(ByVal LocatorType As String, ByVal LocatorValue As String) As Boolean
    Dim download_href As String
    download_href = pReadElementHref(LocatorType, LocatorValue)

    If Not pIsNavigableDownloadHref(download_href) Then
        pTryNavigateDownloadLink = False
        Exit Function
    End If

    Dim response_body As String
    response_body = pClient.Execute("POST", "/session/" & pSessionId & "/url", "{""url"":""" & pJsonEscape(download_href) & """}")
    Call pRaiseIfWebDriverError(response_body)

    pTryNavigateDownloadLink = True
End Function

Private Function pReadElementHref(ByVal LocatorType As String, ByVal LocatorValue As String) As String
    Dim element_id As String
    element_id = pFindElement(LocatorType, LocatorValue)

    Dim response_body As String
    response_body = pClient.Execute("POST", "/session/" & pSessionId & "/execute/sync", pBuildElementHrefScriptBody(element_id))
    Call pRaiseIfWebDriverError(response_body)

    pReadElementHref = pExtractJsonString(response_body, "value")
End Function

Private Function pBuildElementHrefScriptBody(ByVal ElementId As String) As String
    pBuildElementHrefScriptBody = "{""script"":""return arguments[0].href || arguments[0].getAttribute('href') || '';"",""args"": [{""element-6066-11e4-a52e-4f735466cecf"":""" & pJsonEscape(ElementId) & """}]}"
End Function

Private Function pIsNavigableDownloadHref(ByVal DownloadHref As String) As Boolean
    Dim normalized_href As String
    normalized_href = LCase$(Trim$(DownloadHref))

    pIsNavigableDownloadHref = (Left$(normalized_href, 7) = "http://" _
            Or Left$(normalized_href, 8) = "https://")
End Function

Private Sub pClickElement(ByVal LocatorType As String, ByVal LocatorValue As String)
    Dim deadline_time As Date
    deadline_time = DateAdd("s", pSettings.TimeoutSeconds, Now)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    Dim err_help_file As String
    Dim err_help_context As Long

    Do
        If pTryClickElement(LocatorType, LocatorValue, err_num, err_source, err_desc, err_help_file, err_help_context) Then Exit Sub
        If Not pIsRetryableElementLookupError(err_desc) Then
            Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
        End If
        DoEvents
    Loop While Now < deadline_time

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
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

Private Function pBuildFindElementBody(ByVal LocatorType As String, ByVal LocatorValue As String) As String
    pBuildFindElementBody = "{""using"":""" & pJsonEscape(LocatorType) & """,""value"":""" & pJsonEscape(LocatorValue) & """}"
End Function

Private Function pTryClickElement( _
        ByVal LocatorType As String, _
        ByVal LocatorValue As String, _
        ByRef ErrorNumber As Long, _
        ByRef ErrorSource As String, _
        ByRef ErrorDescription As String, _
        ByRef ErrorHelpFile As String, _
        ByRef ErrorHelpContext As Long) As Boolean

    On Error GoTo CLICK_ERROR

    Dim element_id As String
    element_id = pFindElement(LocatorType, LocatorValue)

    Dim response_body As String
    response_body = pClient.Execute("POST", "/session/" & pSessionId & "/element/" & element_id & "/click", "{}")
    If pIsElementNotInteractable(response_body) Then
        Call pClickElementByScript(element_id)
    Else
        Call pRaiseIfWebDriverError(response_body)
    End If

    pTryClickElement = True
    Exit Function

CLICK_ERROR:
    ErrorNumber = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description
    ErrorHelpFile = Err.HelpFile
    ErrorHelpContext = Err.HelpContext
    Err.Clear
    pTryClickElement = False
End Function

Private Function pIsRetryableElementLookupError(ByVal ErrorDescription As String) As Boolean
    pIsRetryableElementLookupError = (0 < InStr(1, ErrorDescription, "no such element", vbTextCompare) _
            Or 0 < InStr(1, ErrorDescription, "stale element reference", vbTextCompare))
End Function

Private Function pIsElementNotInteractable(ByVal ResponseBody As String) As Boolean
    pIsElementNotInteractable = (LCase$(pExtractJsonString(ResponseBody, "error")) = "element not interactable")
End Function

Private Sub pClickElementByScript(ByVal ElementId As String)
    Dim response_body As String
    response_body = pClient.Execute("POST", "/session/" & pSessionId & "/execute/sync", pBuildElementClickScriptBody(ElementId))
    Call pRaiseIfWebDriverError(response_body)
End Sub

Private Function pBuildElementClickScriptBody(ByVal ElementId As String) As String
    pBuildElementClickScriptBody = "{""script"":""arguments[0].click();"",""args"":[{""element-6066-11e4-a52e-4f735466cecf"":""" & pJsonEscape(ElementId) & """}]}"
End Function

Private Sub pRequireSession()
    If pSessionId = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver session が作成されていません。"
    End If
End Sub

Private Function pFindElement(ByVal LocatorType As String, ByVal LocatorValue As String) As String
    On Error GoTo FIND_ERROR

    Dim response_body As String
    response_body = pFindElementResponse(LocatorType, LocatorValue)
    Call pRaiseIfWebDriverError(response_body)

    pFindElement = pExtractElementId(response_body)
    If pFindElement = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="WebDriver element id を取得できません。(" & response_body & ")"
    End If
    Exit Function

FIND_ERROR:
    Dim err_num As Long
    Dim err_desc As String
    Dim err_help_file As String
    Dim err_help_context As Long
    err_num = Err.Number
    err_desc = Err.Description
    err_help_file = Err.HelpFile
    err_help_context = Err.HelpContext

    Err.Raise err_num, "Class WebDriverSessionClient", err_desc & vbCrLf & "LocatorType=" & LocatorType & ", LocatorValue=" & LocatorValue, err_help_file, err_help_context
End Function

Private Function pIsFrameChainLocator(ByVal LocatorType As String, ByVal LocatorValue As String) As Boolean
    pIsFrameChainLocator = (LCase$(Trim$(LocatorType)) = "css selector" _
            And 0 < InStr(1, LocatorValue, ">>", vbBinaryCompare))
End Function

Private Function pSplitFrameChain(ByVal LocatorValue As String) As Variant
    Dim locator_parts As Variant
    locator_parts = Split(LocatorValue, ">>")
    If UBound(locator_parts) < 1 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionClient", Description:="frame chain selector が不正です。(" & LocatorValue & ")"
    End If

    pSplitFrameChain = locator_parts
End Function

Private Sub pSwitchToFrame(ByVal FrameElementId As String)
    Dim response_body As String
    response_body = pClient.Execute("POST", "/session/" & pSessionId & "/frame", pBuildSwitchToFrameBody(FrameElementId))
    Call pRaiseIfWebDriverError(response_body)
End Sub

Private Function pBuildSwitchToFrameBody(ByVal FrameElementId As String) As String
    pBuildSwitchToFrameBody = "{""id"":{""element-6066-11e4-a52e-4f735466cecf"":""" & pJsonEscape(FrameElementId) & """}}"
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
