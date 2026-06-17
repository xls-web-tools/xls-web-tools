VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellZipExtractor"
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
'! Windows Shell.Application を使う ZIP 展開処理です。
'!
' #############################################################################

Private Const C_COPY_HERE_NO_UI As Long = 16

'* ZIP ファイルを指定ディレクトリへ展開します。
'*
'* @param ZipFilePath 展開対象の ZIP ファイルパス。
'* @param DestinationDirectoryPath 展開先ディレクトリパス。
'*
'* @details
'* Windows Shell.Application の CopyHere を UI なしで使い、DownloadedFileSaver から ZIP flatten 保存を実行できるようにします。
Public Sub ExtractZip( _
        ByVal ZipFilePath As String, _
        ByVal DestinationDirectoryPath As String)

    Dim shell_app As Object
    Set shell_app = CreateObject("Shell.Application")

    Dim zip_namespace As Object
    Set zip_namespace = shell_app.NameSpace(CVar(ZipFilePath))
    If zip_namespace Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ShellZipExtractor.ExtractZip", Description:="ZIP を開けません。(" & ZipFilePath & ")"
    End If

    Dim destination_namespace As Object
    Set destination_namespace = shell_app.NameSpace(CVar(DestinationDirectoryPath))
    If destination_namespace Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ShellZipExtractor.ExtractZip", Description:="ZIP 展開先を開けません。(" & DestinationDirectoryPath & ")"
    End If

    Call destination_namespace.CopyHere(zip_namespace.Items, C_COPY_HERE_NO_UI)
End Sub
Private Sub IZipExtractor_ExtractZip( _
        ByVal ZipFilePath As String, _
        ByVal DestinationDirectoryPath As String)

    Call ExtractZip(ZipFilePath, DestinationDirectoryPath)
End Sub
