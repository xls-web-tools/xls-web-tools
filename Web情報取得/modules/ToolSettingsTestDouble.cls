VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolSettingsTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IToolSettings

' #############################################################################
'!
'! @brief
'! Web情報取得設定のテストダブルです。
'!
' #############################################################################

Private pWebDriverPath As String
Private pWebDriverPort As Long
Private pBrowserProfilePath As String
Private pHeadless As Boolean
Private pStartUrl As String
Private pOutputSheetName As String
Private pAuthenticatedStartSelector As String
Private pListPageSelector As String
Private pListTransitionOperationName As String
Private pListItemSelector As String
Private pListItemTargetIdSelector As String
Private pDetailTransitionOperationName As String
Private pTargetIdSelector As String
Private pReturnToListOperationName As String
Private pNextPageOperationName As String
Private pNextPageAvailableSelector As String
Private pExistingRowMode As String
Private pTimeoutSeconds As Long
Private pTransitionOperations As ObjectList
Private pDetailColumnDefinitions As ObjectList

'* Edge WebDriver 実行ファイルのパス。
Public Property Get WebDriverPath() As String
    WebDriverPath = pWebDriverPath
End Property
Public Property Let WebDriverPath(ByVal NewValue As String)
    pWebDriverPath = NewValue
End Property
Private Property Get IToolSettings_WebDriverPath() As String
    IToolSettings_WebDriverPath = WebDriverPath
End Property

'* Edge WebDriver の待ち受け port。
Public Property Get WebDriverPort() As Long
    WebDriverPort = pWebDriverPort
End Property
Public Property Let WebDriverPort(ByVal NewValue As Long)
    pWebDriverPort = NewValue
End Property
Private Property Get IToolSettings_WebDriverPort() As Long
    IToolSettings_WebDriverPort = WebDriverPort
End Property

'* 専用ブラウザプロファイルのパス。
Public Property Get BrowserProfilePath() As String
    BrowserProfilePath = pBrowserProfilePath
End Property
Public Property Let BrowserProfilePath(ByVal NewValue As String)
    pBrowserProfilePath = NewValue
End Property
Private Property Get IToolSettings_BrowserProfilePath() As String
    IToolSettings_BrowserProfilePath = BrowserProfilePath
End Property

'* True の場合は不可視ブラウザで起動します。
Public Property Get Headless() As Boolean
    Headless = pHeadless
End Property
Public Property Let Headless(ByVal NewValue As Boolean)
    pHeadless = NewValue
End Property
Private Property Get IToolSettings_Headless() As Boolean
    IToolSettings_Headless = Headless
End Property

'* 取得開始 URL。
Public Property Get StartUrl() As String
    StartUrl = pStartUrl
End Property
Public Property Let StartUrl(ByVal NewValue As String)
    pStartUrl = NewValue
End Property
Private Property Get IToolSettings_StartUrl() As String
    IToolSettings_StartUrl = StartUrl
End Property

'* 取得結果の出力先シート名。
Public Property Get OutputSheetName() As String
    OutputSheetName = pOutputSheetName
End Property
Public Property Let OutputSheetName(ByVal NewValue As String)
    pOutputSheetName = NewValue
End Property
Private Property Get IToolSettings_OutputSheetName() As String
    IToolSettings_OutputSheetName = OutputSheetName
End Property

'* 認証後の起点ページ到達を判定する selector。
Public Property Get AuthenticatedStartSelector() As String
    AuthenticatedStartSelector = pAuthenticatedStartSelector
End Property
Public Property Let AuthenticatedStartSelector(ByVal NewValue As String)
    pAuthenticatedStartSelector = NewValue
End Property
Private Property Get IToolSettings_AuthenticatedStartSelector() As String
    IToolSettings_AuthenticatedStartSelector = AuthenticatedStartSelector
End Property

'* 一覧画面到達を判定する selector。
Public Property Get ListPageSelector() As String
    ListPageSelector = pListPageSelector
End Property
Public Property Let ListPageSelector(ByVal NewValue As String)
    pListPageSelector = NewValue
End Property
Private Property Get IToolSettings_ListPageSelector() As String
    IToolSettings_ListPageSelector = ListPageSelector
End Property

'* 一覧画面への画面遷移操作名。
Public Property Get ListTransitionOperationName() As String
    ListTransitionOperationName = pListTransitionOperationName
End Property
Public Property Let ListTransitionOperationName(ByVal NewValue As String)
    pListTransitionOperationName = NewValue
End Property
Private Property Get IToolSettings_ListTransitionOperationName() As String
    IToolSettings_ListTransitionOperationName = ListTransitionOperationName
End Property
'* 現在ページの一覧項目を数える selector。
Public Property Get ListItemSelector() As String
    ListItemSelector = pListItemSelector
End Property
Public Property Let ListItemSelector(ByVal NewValue As String)
    pListItemSelector = NewValue
End Property
Private Property Get IToolSettings_ListItemSelector() As String
    IToolSettings_ListItemSelector = ListItemSelector
End Property

'* 一覧上の対象IDを読み取る selector。
Public Property Get ListItemTargetIdSelector() As String
    ListItemTargetIdSelector = pListItemTargetIdSelector
End Property
Public Property Let ListItemTargetIdSelector(ByVal NewValue As String)
    pListItemTargetIdSelector = NewValue
End Property
Private Property Get IToolSettings_ListItemTargetIdSelector() As String
    IToolSettings_ListItemTargetIdSelector = ListItemTargetIdSelector
End Property

'* 詳細ページへの画面遷移操作名。
Public Property Get DetailTransitionOperationName() As String
    DetailTransitionOperationName = pDetailTransitionOperationName
End Property
Public Property Let DetailTransitionOperationName(ByVal NewValue As String)
    pDetailTransitionOperationName = NewValue
End Property
Private Property Get IToolSettings_DetailTransitionOperationName() As String
    IToolSettings_DetailTransitionOperationName = DetailTransitionOperationName
End Property

'* 詳細ページ上の対象IDを読み取る selector。
Public Property Get TargetIdSelector() As String
    TargetIdSelector = pTargetIdSelector
End Property
Public Property Let TargetIdSelector(ByVal NewValue As String)
    pTargetIdSelector = NewValue
End Property
Private Property Get IToolSettings_TargetIdSelector() As String
    IToolSettings_TargetIdSelector = TargetIdSelector
End Property

'* 一覧画面へ戻る画面遷移操作名。
Public Property Get ReturnToListOperationName() As String
    ReturnToListOperationName = pReturnToListOperationName
End Property
Public Property Let ReturnToListOperationName(ByVal NewValue As String)
    pReturnToListOperationName = NewValue
End Property
Private Property Get IToolSettings_ReturnToListOperationName() As String
    IToolSettings_ReturnToListOperationName = ReturnToListOperationName
End Property

'* 次ページへ進む画面遷移操作名。
Public Property Get NextPageOperationName() As String
    NextPageOperationName = pNextPageOperationName
End Property
Public Property Let NextPageOperationName(ByVal NewValue As String)
    pNextPageOperationName = NewValue
End Property
Private Property Get IToolSettings_NextPageOperationName() As String
    IToolSettings_NextPageOperationName = NextPageOperationName
End Property

'* 次ページ操作が有効な場合だけ一致する selector。
Public Property Get NextPageAvailableSelector() As String
    NextPageAvailableSelector = pNextPageAvailableSelector
End Property
Public Property Let NextPageAvailableSelector(ByVal NewValue As String)
    pNextPageAvailableSelector = NewValue
End Property
Private Property Get IToolSettings_NextPageAvailableSelector() As String
    IToolSettings_NextPageAvailableSelector = NextPageAvailableSelector
End Property

'* 既存行の扱い。
Public Property Get ExistingRowMode() As String
    ExistingRowMode = pExistingRowMode
End Property
Public Property Let ExistingRowMode(ByVal NewValue As String)
    pExistingRowMode = NewValue
End Property
Private Property Get IToolSettings_ExistingRowMode() As String
    IToolSettings_ExistingRowMode = ExistingRowMode
End Property

'* 画面到達や selector 出現を待つ秒数。
Public Property Get TimeoutSeconds() As Long
    TimeoutSeconds = pTimeoutSeconds
End Property
Public Property Let TimeoutSeconds(ByVal NewValue As Long)
    pTimeoutSeconds = NewValue
End Property
Private Property Get IToolSettings_TimeoutSeconds() As Long
    IToolSettings_TimeoutSeconds = TimeoutSeconds
End Property

'* 画面遷移操作定義の一覧。
Public Property Get TransitionOperations() As ObjectList
    Set TransitionOperations = pTransitionOperations
End Property
Public Property Set TransitionOperations(ByVal NewValue As ObjectList)
    Set pTransitionOperations = NewValue
End Property
Private Property Get IToolSettings_TransitionOperations() As ObjectList
    Set IToolSettings_TransitionOperations = TransitionOperations
End Property

'* 詳細ページ列定義の一覧。
Public Property Get DetailColumnDefinitions() As ObjectList
    Set DetailColumnDefinitions = pDetailColumnDefinitions
End Property
Public Property Set DetailColumnDefinitions(ByVal NewValue As ObjectList)
    Set pDetailColumnDefinitions = NewValue
End Property
Private Property Get IToolSettings_DetailColumnDefinitions() As ObjectList
    Set IToolSettings_DetailColumnDefinitions = DetailColumnDefinitions
End Property

Private Sub Class_Initialize()
    pWebDriverPath = "bin\msedgedriver.exe"
    pWebDriverPort = 9515
    pBrowserProfilePath = "browser-profile"
    pHeadless = True
    pStartUrl = "https://example.test/start"
    pOutputSheetName = "output"
    pAuthenticatedStartSelector = "#authenticated-start"
    pListPageSelector = "#list-page"
    pListTransitionOperationName = "OpenList"
    pListItemSelector = ""
    pListItemTargetIdSelector = ""
    pDetailTransitionOperationName = ""
    pTargetIdSelector = ""
    pReturnToListOperationName = ""
    pNextPageOperationName = "NextPage"
    pNextPageAvailableSelector = ""
    pExistingRowMode = G_WEB_ROW_MODE_SKIP_EXISTING
    pTimeoutSeconds = 30
    Set pTransitionOperations = New_ObjectList("TransitionOperation")
    Set pDetailColumnDefinitions = New_ObjectList("DetailColumnDefinition")
End Sub
