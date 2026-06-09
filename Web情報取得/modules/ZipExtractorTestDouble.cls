VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZipExtractorTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IZipExtractor

' #############################################################################
'!
'! @brief
'! ZIP 展開処理のテストダブルです。
'!
' #############################################################################

Private pStore As TestDoubleBehaviorStore

'* スタブ設定と呼び出し履歴。
Public Property Get Store() As TestDoubleBehaviorStore
    Set Store = pStore
End Property

Private Sub Class_Initialize()
    Set pStore = New TestDoubleBehaviorStore
End Sub

'* ExtractZip の spy only です。失敗は Store.SetError で注入します。
Public Sub ExtractZip( _
        ByVal ZipFilePath As String, _
        ByVal DestinationDirectoryPath As String)

    Call Store.RaiseIfError("ExtractZip", ZipFilePath, DestinationDirectoryPath)
    Call Store.RecordCall("ExtractZip", ZipFilePath, DestinationDirectoryPath)
End Sub
Private Sub IZipExtractor_ExtractZip( _
        ByVal ZipFilePath As String, _
        ByVal DestinationDirectoryPath As String)

    Call ExtractZip(ZipFilePath, DestinationDirectoryPath)
End Sub