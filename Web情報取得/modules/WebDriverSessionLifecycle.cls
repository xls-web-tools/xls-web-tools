VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverSessionLifecycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! ブラウザ取得の WebDriver process と session lifecycle を管理します。
'!
' #############################################################################

Private pProcess As IWebDriverProcess
Private pSessionClient As WebDriverSessionClient
Private pSettings As IToolSettings
Private pKeepVisibleBrowserOnError As Boolean
Private pSessionId As String
Private pIsInitialized As Boolean

'* WebDriver session client。
Public Property Get SessionClient() As WebDriverSessionClient
    Call pCheckInit
    Set SessionClient = pSessionClient
End Property

'* 作成済み session id。
Public Property Get SessionId() As String
    Call pCheckInit
    SessionId = pSessionId
End Property

'* WebDriver session lifecycle を初期化します。
'*
'* @param Process WebDriver process 操作。
'* @param SessionClient WebDriver session client。
'* @param Settings Web情報取得の設定。
'* @param KeepVisibleBrowserOnError True の場合、可視ブラウザ診断のエラー時に session と process を残します。
'*
'* @details
'* process 起動、session 作成、正常終了、エラー時 cleanup の順序をこの Module に閉じます。
Public Sub Initialize( _
        ByVal Process As IWebDriverProcess, _
        ByVal SessionClient As WebDriverSessionClient, _
        ByVal Settings As IToolSettings, _
        Optional ByVal KeepVisibleBrowserOnError As Boolean = False)

    If Process Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionLifecycle", Description:="Process は必須です。"
    End If
    If SessionClient Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionLifecycle", Description:="SessionClient は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionLifecycle", Description:="Settings は必須です。"
    End If

    Set pProcess = Process
    Set pSessionClient = SessionClient
    Set pSettings = Settings
    pKeepVisibleBrowserOnError = KeepVisibleBrowserOnError
    pSessionId = ""
    pIsInitialized = True
End Sub

'* WebDriver process を起動し、session を作成します。
'*
'* @return 作成された session id。
'*
'* @details
'* 失敗時の cleanup は CleanupAfterError に委譲します。
Public Function StartSession() As String
    Call pCheckInit

    Call pProcess.Start(pSettings)
    pSessionId = pSessionClient.CreateSession()

    StartSession = pSessionId
End Function

'* WebDriver session と process を正常終了します。
Public Sub FinishSession()
    Call pCheckInit

    If pSessionId <> "" Then
        Call pSessionClient.QuitSession(pSessionId)
        pSessionId = ""
    End If
    Call pProcess.StopProcess
End Sub

'* エラー時に WebDriver session と process の cleanup を試みます。
'*
'* @param ErrorDescription 再送出するエラー文。可視ブラウザを残す場合は診断用メッセージを追記します。
'*
'* @details
'* 可視ブラウザ診断で session 作成済みの場合は、利用者が画面確認できるよう session と process を残します。
Public Sub CleanupAfterError(ByRef ErrorDescription As String)
    Call pCheckInit

    If pShouldKeepVisibleBrowserOnError Then
        ErrorDescription = ErrorDescription & vbCrLf & "Headless=False のため、診断用にブラウザを残しました。確認後、ブラウザを閉じてください。"
        Exit Sub
    End If

    On Error Resume Next
    If pSessionId <> "" Then Call pSessionClient.QuitSession(pSessionId)
    pSessionId = ""
    Call pProcess.StopProcess
    On Error GoTo 0
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSessionLifecycle", Description:="初期化されていません。"
    End If
End Sub

Private Function pShouldKeepVisibleBrowserOnError() As Boolean
    pShouldKeepVisibleBrowserOnError = (pKeepVisibleBrowserOnError And (Not pSettings.Headless) And pSessionId <> "")
End Function
