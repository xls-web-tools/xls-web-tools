VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverProcessTestDouble"
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
'! WebDriver process 操作のテストダブルです。
'!
' #############################################################################

'* スタブ設定と正常に成立した呼び出し履歴。
Public Store As TestDoubleBehaviorStore

Private pIsRunning As Boolean

'* 起動済みかどうか。
Public Property Get IsRunning() As Boolean
    IsRunning = pIsRunning
End Property
Private Property Get IWebDriverProcess_IsRunning() As Boolean
    IWebDriverProcess_IsRunning = IsRunning
End Property

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
    pIsRunning = False
End Sub

'* Edge WebDriver process を起動します。
'*
'* @param Settings Web情報取得の設定。
'*
'* @details
'* 呼び出しを記録し、失敗は Store.SetError で注入します。
Public Sub Start(ByVal Settings As IToolSettings)
    Call Store.RaiseIfError("Start")
    pIsRunning = True
    Call Store.RecordCall("Start")
End Sub
Private Sub IWebDriverProcess_Start(ByVal Settings As IToolSettings)
    Call Start(Settings)
End Sub

'* 起動した Edge WebDriver process を終了します。
'*
'* @details
'* 呼び出しを記録し、失敗は Store.SetError で注入します。
Public Sub StopProcess()
    Call Store.RaiseIfError("StopProcess")
    pIsRunning = False
    Call Store.RecordCall("StopProcess")
End Sub
Private Sub IWebDriverProcess_StopProcess()
    Call StopProcess
End Sub
