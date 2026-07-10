VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IWebDriverProcess

' #############################################################################
'!
'! @brief
'! Edge WebDriver process の起動と終了を管理します。
'!
' #############################################################################

Private Const C_MINIMIZED_NO_FOCUS As Long = 7
Private Const C_HTTP_TIMEOUT_MS As Long = 1000

Private pFileSystem As IFileSystemService
Private pPortProbe As IWebDriverPortProbe
Private pWebDriverPort As Long
Private pIsInitialized As Boolean
Private pIsRunning As Boolean

'* 起動済みかどうか。
Public Property Get IsRunning() As Boolean
    IsRunning = pIsRunning
End Property
Private Property Get IWebDriverProcess_IsRunning() As Boolean
    IWebDriverProcess_IsRunning = IsRunning
End Property

'* WebDriver process を初期化します。
'*
'* @param PortProbe port 使用状況確認に使う probe。
'*
'* @details
'* PortProbe は Nothing にできません。
Public Sub Initialize(ByVal PortProbe As IWebDriverPortProbe)
    If PortProbe Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess.Initialize", Description:="PortProbe は必須です。"
    End If

    Set pFileSystem = FsSrv
    Set pPortProbe = PortProbe
    pWebDriverPort = 0
    pIsRunning = False
    pIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Call StopProcess
End Sub

'* Edge WebDriver process を起動します。
'*
'* @param Settings Web情報取得の設定。
'*
'* @details
'* WebDriver 実行ファイルの存在と port 使用状況を確認してから起動します。既存 port 使用中の場合、他プロセスを終了せずエラーにします。
Public Sub Start(ByVal Settings As IToolSettings)
    Call pCheckInit
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess.Start", Description:="Settings は必須です。"
    End If

    If Not pFileSystem.IsFile(Settings.WebDriverPath) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess.Start", Description:="Edge WebDriver 実行ファイルが配置されていません。(" & Settings.WebDriverPath & ")"
    End If

    If pPortProbe.IsPortInUse(Settings.WebDriverPort) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess.Start", Description:="WebDriver port が使用中です。他プロセスは終了しません。(" & CStr(Settings.WebDriverPort) & ")"
    End If

    On Error GoTo START_ERROR

    Dim shell_obj As Object
    Set shell_obj = CreateObject("WScript.Shell")
    Call shell_obj.Run(pBuildCommand(Settings), C_MINIMIZED_NO_FOCUS, False)

    pWebDriverPort = Settings.WebDriverPort
    pIsRunning = True
    Call pWaitUntilReady(Settings)

    On Error GoTo 0
    Exit Sub

START_ERROR:
    Dim err_num As Long
    Dim err_desc As String
    Dim err_help_file As String
    Dim err_help_context As Long
    err_num = Err.Number
    err_desc = Err.Description
    err_help_file = Err.HelpFile
    err_help_context = Err.HelpContext

    Call StopProcess
    Err.Raise err_num, "Class WebDriverProcess.Start", err_desc, err_help_file, err_help_context
End Sub
Private Sub IWebDriverProcess_Start(ByVal Settings As IToolSettings)
    Call Start(Settings)
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess.pCheckInit", Description:="初期化されていません。"
    End If
End Sub

Private Function pBuildCommand(ByVal Settings As IToolSettings) As String
    pBuildCommand = pQuoteArgument(Settings.WebDriverPath) & " --port=" & CStr(Settings.WebDriverPort)
End Function

Private Sub pWaitUntilReady(ByVal Settings As IToolSettings)
    Dim timeout_seconds As Long
    timeout_seconds = Settings.TimeoutSeconds
    If timeout_seconds <= 0 Then timeout_seconds = 1

    Dim deadline_time As Date
    deadline_time = DateAdd("s", timeout_seconds, Now)

    Do
        If pPortProbe.IsPortInUse(Settings.WebDriverPort) Then Exit Sub
        DoEvents
    Loop While Now < deadline_time

    Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess.pWaitUntilReady", Description:="Edge WebDriver が起動待ち時間内に応答しません。port=" & CStr(Settings.WebDriverPort) & ", timeout=" & CStr(timeout_seconds) & " 秒"
End Sub

'* 起動した Edge WebDriver process を終了します。
'*
'* @details
'* 起動時に記録した port の WebDriver service へ shutdown を要求します。他プロセスは終了しません。
Public Sub StopProcess()
    On Error Resume Next
    If pIsRunning And 0 < pWebDriverPort Then
        Call pShutdownWebDriverService(pWebDriverPort)
    End If
    pWebDriverPort = 0
    pIsRunning = False
    On Error GoTo 0
End Sub
Private Sub IWebDriverProcess_StopProcess()
    Call StopProcess
End Sub

Private Sub pShutdownWebDriverService(ByVal WebDriverPort As Long)
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call request.SetTimeouts(C_HTTP_TIMEOUT_MS, C_HTTP_TIMEOUT_MS, C_HTTP_TIMEOUT_MS, C_HTTP_TIMEOUT_MS)
    Call request.Open("GET", "http://127.0.0.1:" & CStr(WebDriverPort) & "/shutdown", False)
    Call request.Send
End Sub

Private Function pQuoteArgument(ByVal ArgumentText As String) As String
    pQuoteArgument = """" & Replace(ArgumentText, """", """""") & """"
End Function
