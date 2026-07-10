VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebDriverProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'#Interface

' #############################################################################
'!
'! @brief
'! Edge WebDriver process 操作を差し替えるためのインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IWebDriverProcess", Description:="このクラスはインターフェイス用です。"
End Sub

'* 起動済みかどうか。
Public Property Get IsRunning() As Boolean
End Property

'* Edge WebDriver process を起動します。
'*
'* @param Settings Web情報取得の設定。
'*
'* @details
'* WebDriver 実行ファイルの存在と port 使用状況を確認してから起動します。既存 port 使用中の場合、他プロセスを終了せずエラーにします。
Public Sub Start(ByVal Settings As IToolSettings)
End Sub

'* 起動した Edge WebDriver process を終了します。
'*
'* @details
'* この process instance が起動した process だけを終了対象にします。他プロセスは終了しません。
Public Sub StopProcess()
End Sub
