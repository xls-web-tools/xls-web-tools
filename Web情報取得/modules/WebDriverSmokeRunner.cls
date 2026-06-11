VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverSmokeRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! WebDriver process 起動から session 終了までの smoke path を実行します。
'!
' #############################################################################

Private pLifecycle As WebDriverSessionLifecycle
Private pIsInitialized As Boolean

'* WebDriver smoke runner を初期化します。
'*
'* @param Lifecycle WebDriver session lifecycle。
'*
'* @details
'* process 起動、session 作成、quit、process 終了の依存を lifecycle に委譲します。
Public Sub Initialize(ByVal Lifecycle As WebDriverSessionLifecycle)
    If Lifecycle Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSmokeRunner", Description:="Lifecycle は必須です。"
    End If

    Set pLifecycle = Lifecycle
    pIsInitialized = True
End Sub

'* WebDriver smoke path を実行します。
'*
'* @return 作成された session id。
'*
'* @details
'* WebDriver session lifecycle の開始と正常終了を通します。
Public Function Run() As String
    Call pCheckInit

    On Error GoTo ON_ERROR

    Dim result_value As String
    result_value = pLifecycle.StartSession()
    If result_value = "" Then
        On Error GoTo 0
        Run = result_value
        Exit Function
    End If
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

    Call pLifecycle.CleanupAfterError(err_desc)

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WebDriverSmokeRunner", Description:="初期化されていません。"
    End If
End Sub
