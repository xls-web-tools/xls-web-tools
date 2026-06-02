VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebDriverPortProbe"
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
'! WebDriver port の使用状況確認を差し替えるためのインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IWebDriverPortProbe", Description:="このクラスはインターフェイス用です。"
End Sub

'* port が使用中か確認します。
'*
'* @param PortNumber 確認対象の port 番号。
'* @return 使用中の場合は True。
'*
'* @details
'* 実装側は指定 port へ接続できるかを確認します。呼び出し側は True の場合、他プロセスを終了せず中断します。
Public Function IsPortInUse(ByVal PortNumber As Long) As Boolean
End Function
