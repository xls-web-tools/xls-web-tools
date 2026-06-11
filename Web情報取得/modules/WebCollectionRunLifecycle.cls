VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebCollectionRunLifecycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! ブラウザ取得 run の outer lifecycle を管理します。
'!
' #############################################################################

Private pLifecycle As WebDriverSessionLifecycle
Private pSettings As IToolSettings
Private pIsInitialized As Boolean

'* ブラウザ取得 run lifecycle を初期化します。
'*
'* @param Lifecycle WebDriver session lifecycle。
'* @param Settings Web情報取得の設定。
'*
'* @details
'* ブラウザ取得 run lifecycle を初期化します。
Public Sub Initialize( _
        ByVal Lifecycle As WebDriverSessionLifecycle, _
        ByVal Settings As IToolSettings)

    If Lifecycle Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunLifecycle", Description:="Lifecycle は必須です。"
    End If
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunLifecycle", Description:="Settings は必須です。"
    End If

    Set pLifecycle = Lifecycle
    Set pSettings = Settings
    pIsInitialized = True
End Sub

'* 一覧画面到達までの準備と一覧画面到達後処理を実行します。
'*
'* @param Body 一覧画面到達後処理。
'* @return 作成された session id。
'*
'* @details
'* 準備、一覧画面到達後処理、正常終了、エラー時 cleanup の順序をこの Module に閉じます。
Public Function Run(ByVal Body As IWebCollectionRunBody) As String
    Call pCheckInit
    If Body Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunLifecycle", Description:="Body は必須です。"
    End If

    Dim result_value As String
    Dim preparation As WebCollectionPreparation

    On Error GoTo ON_ERROR

    Set preparation = New_WebCollectionPreparation(pLifecycle, pSettings)
    result_value = preparation.Run()
    If result_value = "" Then
        Call preparation.ClearProgress
        On Error GoTo 0
        Run = result_value
        Exit Function
    End If
    Call Body.RunAfterPreparation
    Call pLifecycle.FinishSession

    On Error GoTo 0
    Run = result_value
    Exit Function

ON_ERROR:
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

    On Error Resume Next
    If Not preparation Is Nothing Then Call preparation.ClearProgress
    Call Body.ClearProgressAfterError
    On Error GoTo 0
    Call pLifecycle.CleanupAfterError(err_desc)

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebCollectionRunLifecycle", Description:="初期化されていません。"
    End If
End Sub
