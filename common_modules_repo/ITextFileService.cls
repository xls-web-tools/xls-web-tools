VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITextFileService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#Interface

' #############################################################################
'!
'! @brief
'! ITextFileEntity を生成するサービスのインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface ITextFileService", Description:="このクラスはインターフェイス用です。"
End Sub

'* TextFileEntity を生成します。
'*
'* @param FilePath ファイルのパス。
'*
'* @details
'* TextFileEntity を生成します。
Public Function GetTextFileEntity(ByVal FilePath As String) As ITextFileEntity
End Function
