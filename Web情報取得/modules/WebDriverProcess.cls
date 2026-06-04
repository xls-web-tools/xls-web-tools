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

Private pFileSystem As IFileSystemService
Private pPortProbe As IWebDriverPortProbe
Private pExec As Object
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
'* @param FileSystem ファイル存在確認に使うサービス。
'* @param PortProbe port 使用状況確認に使う probe。
'*
'* @details
'* FileSystem と PortProbe は Nothing にできません。
Public Sub Initialize(ByVal FileSystem As IFileSystemService, ByVal PortProbe As IWebDriverPortProbe)
    If FileSystem Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess", Description:="FileSystem は必須です。"
    End If
    If PortProbe Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess", Description:="PortProbe は必須です。"
    End If

    Set pFileSystem = FileSystem
    Set pPortProbe = PortProbe
    Set pExec = Nothing
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
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess", Description:="Settings は必須です。"
    End If

    If Not pFileSystem.IsFile(Settings.WebDriverPath) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess", Description:="Edge WebDriver 実行ファイルが配置されていません。(" & Settings.WebDriverPath & ")"
    End If

    If pPortProbe.IsPortInUse(Settings.WebDriverPort) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess", Description:="WebDriver port が使用中です。他プロセスは終了しません。(" & CStr(Settings.WebDriverPort) & ")"
    End If

    On Error GoTo START_ERROR

    Dim shell_obj As Object
    Set shell_obj = CreateObject("WScript.Shell")
    Set pExec = shell_obj.Exec(pBuildCommand(Settings))
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
    Err.Raise err_num, "Class WebDriverProcess", err_desc, err_help_file, err_help_context
End Sub
Private Sub IWebDriverProcess_Start(ByVal Settings As IToolSettings)
    Call Start(Settings)
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess", Description:="初期化されていません。"
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
        Call pRaiseIfProcessExited
        If pPortProbe.IsPortInUse(Settings.WebDriverPort) Then Exit Sub
        DoEvents
    Loop While Now < deadline_time

    Call pRaiseIfProcessExited
    Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess", Description:="Edge WebDriver が起動待ち時間内に応答しません。port=" & CStr(Settings.WebDriverPort) & ", timeout=" & CStr(timeout_seconds) & " 秒"
End Sub

'* 起動した Edge WebDriver process を終了します。
'*
'* @details
'* この instance が起動した process だけを終了対象にします。他プロセスは終了しません。
Public Sub StopProcess()
    On Error Resume Next
    If Not pExec Is Nothing Then
        If pExec.Status = 0 Then Call pExec.Terminate
    End If
    Set pExec = Nothing
    pIsRunning = False
    On Error GoTo 0
End Sub
Private Sub IWebDriverProcess_StopProcess()
    Call StopProcess
End Sub

Private Sub pRaiseIfProcessExited()
    If pExec Is Nothing Then Exit Sub
    If pExec.Status = 0 Then Exit Sub

    pIsRunning = False
    Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverProcess", Description:="Edge WebDriver process が起動直後に終了しました。ExitCode=" & CStr(pExec.ExitCode) & pBuildProcessOutputMessage()
End Sub

Private Function pBuildProcessOutputMessage() As String
    On Error Resume Next

    Dim output_text As String
    output_text = Trim$(CStr(pExec.StdErr.ReadAll))
    If output_text = "" Then output_text = Trim$(CStr(pExec.StdOut.ReadAll))

    On Error GoTo 0

    If output_text = "" Then
        pBuildProcessOutputMessage = ""
    Else
        pBuildProcessOutputMessage = " Output=" & output_text
    End If
End Function

Private Function pQuoteArgument(ByVal ArgumentText As String) As String
    pQuoteArgument = """" & Replace(ArgumentText, """", """""") & """"
End Function
