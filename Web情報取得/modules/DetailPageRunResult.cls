VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailPageRunResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 詳細ページ 1 件分の処理結果です。
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pTargetId As String
Private pStatus As String
Private pErrorDescription As String
Private pDetailValues As ArrayObject
Private pDownloadStatus As String
Private pIsOutputExcluded As Boolean

'* 出力行の対象ID。
Public Property Get TargetId() As String
    Call pCheckInit
    TargetId = pTargetId
End Property

'* 取得状態。
Public Property Get Status() As String
    Call pCheckInit
    Status = pStatus
End Property

'* エラー内容。
Public Property Get ErrorDescription() As String
    Call pCheckInit
    ErrorDescription = pErrorDescription
End Property

'* 詳細ページ列定義に対応する抽出値。
Public Property Get DetailValues() As ArrayObject
    Call pCheckInit
    Set DetailValues = pDetailValues
End Property

'* ダウンロード状態。
Public Property Get DownloadStatus() As String
    Call pCheckInit
    DownloadStatus = pDownloadStatus
End Property

'* 出力対象条件により output シートへ書かなかったか否か。
Public Property Get IsOutputExcluded() As Boolean
    Call pCheckInit
    IsOutputExcluded = pIsOutputExcluded
End Property

'* 詳細ページ処理結果を初期化します。
Public Sub Initialize( _
        ByVal TargetId As String, _
        ByVal Status As String, _
        ByVal ErrorDescription As String, _
        ByVal DetailValues As ArrayObject, _
        Optional ByVal DownloadStatus As String = "", _
        Optional ByVal IsOutputExcluded As Boolean = False)

    If DetailValues Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRunResult", Description:="DetailValues は必須です。"
    End If

    pTargetId = TargetId
    pStatus = Status
    pErrorDescription = ErrorDescription
    Set pDetailValues = DetailValues
    pDownloadStatus = DownloadStatus
    pIsOutputExcluded = IsOutputExcluded
    pIsInitialized = True
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailPageRunResult", Description:="初期化されていません。"
    End If
End Sub
