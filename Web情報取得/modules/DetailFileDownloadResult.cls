VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailFileDownloadResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 詳細ページのダウンロード処理結果です。
'!
' #############################################################################

Private pDownloadStatus As String
Private pErrorDescription As String
Private pRequiresDetailError As Boolean
Private pIsInitialized As Boolean

'* ダウンロード状態。
Public Property Get DownloadStatus() As String
    Call pCheckInit
    DownloadStatus = pDownloadStatus
End Property

'* ダウンロード処理のエラー内容。
Public Property Get ErrorDescription() As String
    Call pCheckInit
    ErrorDescription = pErrorDescription
End Property

'* 詳細ページ単位の取得状態を ERROR にする必要があるか。
Public Property Get RequiresDetailError() As Boolean
    Call pCheckInit
    RequiresDetailError = pRequiresDetailError
End Property

'* 詳細ページのダウンロード処理結果を初期化します。
'*
'* @param DownloadStatus ダウンロード状態。
'* @param ErrorDescription ダウンロード処理のエラー内容。
'* @param RequiresDetailError 詳細ページ単位の取得状態を ERROR にする必要がある場合は True。
'*
'* @details
'* DownloadRequired の設定により、ダウンロード失敗を詳細ページ行の ERROR として扱うかを呼び出し側へ伝えます。
Public Sub Initialize( _
        ByVal DownloadStatus As String, _
        ByVal ErrorDescription As String, _
        ByVal RequiresDetailError As Boolean)

    pDownloadStatus = DownloadStatus
    pErrorDescription = ErrorDescription
    pRequiresDetailError = RequiresDetailError
    pIsInitialized = True
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailFileDownloadResult", Description:="初期化されていません。"
    End If
End Sub
