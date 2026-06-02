VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStringable"
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
'! オブジェクトが文字列表現で表現できることを示すインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IStringable", Description:="このクラスはインターフェイス用です。"
End Sub

'* このオブジェクトの文字列表現を返します。
'*
'* @details
'* このオブジェクトの文字列表現を返します。
Public Function ToString() As String
End Function
