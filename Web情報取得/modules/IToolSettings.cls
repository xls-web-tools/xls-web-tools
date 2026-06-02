VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IToolSettings"
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
'! Web情報取得の settings シート設定を読み取るためのインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IToolSettings", Description:="このクラスはインターフェイス用です。"
End Sub

'* Edge WebDriver 実行ファイルのパス。
Public Property Get WebDriverPath() As String
End Property

'* Edge WebDriver の待ち受け port。
Public Property Get WebDriverPort() As Long
End Property

'* 専用ブラウザプロファイルのパス。
Public Property Get BrowserProfilePath() As String
End Property

'* True の場合は不可視ブラウザで起動します。
Public Property Get Headless() As Boolean
End Property

'* 取得開始 URL。
Public Property Get StartUrl() As String
End Property

'* 取得結果の出力先シート名。
Public Property Get OutputSheetName() As String
End Property

'* 既存行の扱い。
Public Property Get ExistingRowMode() As String
End Property

'* 画面到達や selector 出現を待つ秒数。
Public Property Get TimeoutSeconds() As Long
End Property

'* 画面遷移操作定義の一覧。
Public Property Get TransitionOperations() As ObjectList
End Property

'* 詳細ページ列定義の一覧。
Public Property Get DetailColumnDefinitions() As ObjectList
End Property
