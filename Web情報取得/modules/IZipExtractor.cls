VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IZipExtractor"
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
'! ZIP 展開処理を差し替え可能にするインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IZipExtractor", Description:="このクラスはインターフェイス用です。"
End Sub

'* ZIP ファイルを指定ディレクトリへ展開します。
Public Sub ExtractZip( _
        ByVal ZipFilePath As String, _
        ByVal DestinationDirectoryPath As String)
End Sub