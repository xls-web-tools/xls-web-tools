Attribute VB_Name = "Constructor"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! Web情報取得の引数付きコンストラクター相当の関数です。
'!
' #############################################################################

'* 画面遷移操作定義を生成します。
'*
'* @param OperationName 操作名。
'* @param LocatorType locator 種別。
'* @param LocatorValue locator 値。
'* @param Script click で再現できない場合だけ使う JavaScript。
'* @param WaitConditionName 操作後の待機条件名。
'* @return 初期化済みの画面遷移操作定義。
'*
'* @details
'* New_ 系の処理は生成と Initialize 呼び出しに留め、検証は TransitionOperation.Initialize に委譲します。
Public Function New_TransitionOperation( _
        ByVal OperationName As String, _
        ByVal LocatorType As String, _
        ByVal LocatorValue As String, _
        Optional ByVal Script As String = "", _
        Optional ByVal WaitConditionName As String = "") As TransitionOperation

    Dim result_value As TransitionOperation
    Set result_value = New TransitionOperation
    Call result_value.Initialize(OperationName, LocatorType, LocatorValue, Script:=Script, WaitConditionName:=WaitConditionName)

    Set New_TransitionOperation = result_value
End Function

'* 詳細ページ列定義を生成します。
'*
'* @param OutputColumnName 出力列名。
'* @param Selector DOM selector。
'* @param ExtractType 取得種別。
'* @param AttributeName attribute 取得時の属性名。
'* @param IsRequired True の場合は必須列として扱います。
'* @param ValueExpression 値式。
'* @param BlankMode 空欄時の扱い。
'* @return 初期化済みの詳細ページ列定義。
'*
'* @details
'* New_ 系の処理は生成と Initialize 呼び出しに留め、検証は DetailColumnDefinition.Initialize に委譲します。
Public Function New_DetailColumnDefinition( _
        ByVal OutputColumnName As String, _
        ByVal Selector As String, _
        Optional ByVal ExtractType As String = "InnerText", _
        Optional ByVal AttributeName As String = "", _
        Optional ByVal IsRequired As Boolean = False, _
        Optional ByVal ValueExpression As String = "", _
        Optional ByVal BlankMode As String = "AllowBlank") As DetailColumnDefinition

    Dim result_value As DetailColumnDefinition
    Set result_value = New DetailColumnDefinition
    Call result_value.Initialize( _
            OutputColumnName, _
            Selector, _
            ExtractType:=ExtractType, _
            AttributeName:=AttributeName, _
            IsRequired:=IsRequired, _
            ValueExpression:=ValueExpression, _
            BlankMode:=BlankMode)

    Set New_DetailColumnDefinition = result_value
End Function

'* 詳細ページ処理結果を生成します。
'*
'* @param TargetId 対象ID。
'* @param Status 取得状態。
'* @param ErrorDescription エラー内容。
'* @param DetailValues 詳細ページ列定義に対応する抽出値。
'* @param DownloadStatus ダウンロード状態。
'* @param IsOutputExcluded 出力対象条件により output シートへ書かなかったか否か。
'* @return 初期化済みの詳細ページ処理結果。
'*
'* @details
'* New_ 系の処理は生成と Initialize 呼び出しに留め、実処理は DetailPageRunResult に委譲します。
Public Function New_DetailPageRunResult( _
        ByVal TargetId As String, _
        ByVal Status As String, _
        ByVal ErrorDescription As String, _
        ByVal DetailValues As ArrayObject, _
        Optional ByVal DownloadStatus As String = "", _
        Optional ByVal IsOutputExcluded As Boolean = False) As DetailPageRunResult

    Dim result_value As DetailPageRunResult
    Set result_value = New DetailPageRunResult
    Call result_value.Initialize( _
            TargetId, _
            Status, _
            ErrorDescription, _
            DetailValues, _
            DownloadStatus:=DownloadStatus, _
            IsOutputExcluded:=IsOutputExcluded)

    Set New_DetailPageRunResult = result_value
End Function

'* 詳細ページのダウンロード処理結果を生成します。
'*
'* @param DownloadStatus ダウンロード状態。
'* @param ErrorDescription ダウンロード処理のエラー内容。
'* @param RequiresDetailError 詳細ページ単位の取得状態を ERROR にする必要があるか。
'* @return 初期化済みの詳細ページのダウンロード処理結果。
'*
'* @details
'* New_ 系の処理は生成と Initialize 呼び出しに留め、実処理は DetailFileDownloadResult に委譲します。
Public Function New_DetailFileDownloadResult( _
        ByVal DownloadStatus As String, _
        ByVal ErrorDescription As String, _
        ByVal RequiresDetailError As Boolean) As DetailFileDownloadResult

    Dim result_value As DetailFileDownloadResult
    Set result_value = New DetailFileDownloadResult
    Call result_value.Initialize(DownloadStatus, ErrorDescription, RequiresDetailError)

    Set New_DetailFileDownloadResult = result_value
End Function

'* 詳細ページ runner を生成します。
'*
'* @param SessionClient WebDriver session client。
'* @param Settings Web情報取得の設定。
'* @param FileSystem ファイルシステム操作サービス。DownloadEnabled=True の場合に使用します。
'* @param ZipExtractor ZIP 展開処理。省略時は ShellZipExtractor を使用します。
'* @return 初期化済みの詳細ページ runner。
'*
'* @details
'* New_ 系の処理は生成と Initialize 呼び出しに留め、DownloadEnabled=True の場合だけ詳細ページダウンロード処理を組み立てます。
Public Function New_DetailPageRunner( _
        ByVal SessionClient As WebDriverSessionClient, _
        ByVal Settings As IToolSettings, _
        Optional ByVal FileSystem As IFileSystemService = Nothing, _
        Optional ByVal ZipExtractor As IZipExtractor = Nothing) As DetailPageRunner

    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function New_DetailPageRunner", Description:="Settings は必須です。"
    End If

    Dim file_downloader As DetailFileDownloader
    If Settings.DownloadEnabled Then
        Set file_downloader = New_DetailFileDownloader(SessionClient, Settings, FileSystem, ZipExtractor)
    End If

    Dim result_value As DetailPageRunner
    Set result_value = New DetailPageRunner
    Call result_value.Initialize(SessionClient, Settings, file_downloader)

    Set New_DetailPageRunner = result_value
End Function

'* 出力先シート書き込みを生成します。
'*
'* @param Settings Web情報取得の設定。
'* @return 初期化済みの出力先シート書き込み。
'*
'* @details
'* New_ 系の処理は生成と Initialize 呼び出しに留め、実処理は OutputSheetWriter に委譲します。
Public Function New_OutputSheetWriter(ByVal Settings As IToolSettings) As OutputSheetWriter
    Dim result_value As OutputSheetWriter
    Set result_value = New OutputSheetWriter
    Call result_value.Initialize(Settings)

    Set New_OutputSheetWriter = result_value
End Function

'* 詳細ページのダウンロード処理を生成します。
'*
'* @param SessionClient WebDriver session client。
'* @param Settings Web情報取得の設定。
'* @param FileSystem ファイルシステム操作サービス。省略時は FsSrv を使用します。
'* @param ZipExtractor ZIP 展開処理。省略時は ShellZipExtractor を使用します。
'* @return 初期化済みの詳細ページのダウンロード処理。
'*
'* @details
'* staging 領域の監視と最終保存を分けて扱うため、ファイル保存処理を組み立てて DetailFileDownloader へ渡します。
Public Function New_DetailFileDownloader( _
        ByVal SessionClient As WebDriverSessionClient, _
        ByVal Settings As IToolSettings, _
        Optional ByVal FileSystem As IFileSystemService = Nothing, _
        Optional ByVal ZipExtractor As IZipExtractor = Nothing) As DetailFileDownloader

    Dim actual_file_system As IFileSystemService
    Set actual_file_system = pResolveFileSystem(FileSystem, "Function New_DetailFileDownloader")

    Dim file_saver As DownloadedFileSaver
    Set file_saver = New_DownloadedFileSaver(Settings, actual_file_system, ZipExtractor)

    Dim result_value As DetailFileDownloader
    Set result_value = New DetailFileDownloader
    Call result_value.Initialize(SessionClient, Settings, actual_file_system, file_saver)

    Set New_DetailFileDownloader = result_value
End Function

'* ダウンロード済みファイル保存を生成します。
'*
'* @param Settings Web情報取得の設定。
'* @param FileSystem ファイルシステム操作サービス。
'* @param ZipExtractor ZIP 展開処理。省略時は ShellZipExtractor を使用します。
'* @return 初期化済みのダウンロード済みファイル保存。
'*
'* @details
'* ZIP ファイルの flatten 保存を通常保存と同じ入口で扱うため、ZIP 展開処理を注入して保存処理を生成します。
Public Function New_DownloadedFileSaver( _
        ByVal Settings As IToolSettings, _
        ByVal FileSystem As IFileSystemService, _
        Optional ByVal ZipExtractor As IZipExtractor = Nothing) As DownloadedFileSaver

    Dim actual_zip_extractor As IZipExtractor
    If ZipExtractor Is Nothing Then
        Set actual_zip_extractor = New ShellZipExtractor
    Else
        Set actual_zip_extractor = ZipExtractor
    End If

    Dim result_value As DownloadedFileSaver
    Set result_value = New DownloadedFileSaver
    Call result_value.Initialize(Settings, FileSystem, actual_zip_extractor)

    Set New_DownloadedFileSaver = result_value
End Function

'* WebDriver session client を生成します。
'*
'* @param Client WebDriver HTTP API 呼び出しクライアント。
'* @param Settings Web情報取得の設定。
'* @param FileSystem ファイルシステム操作サービス。
'* @return 初期化済みの WebDriver session client。
'*
'* @details
'* New_ 系の処理は生成と Initialize 呼び出しに留め、実処理は WebDriverSessionClient に委譲します。
Public Function New_WebDriverSessionClient( _
        ByVal Client As IWebDriverClient, _
        ByVal Settings As IToolSettings, _
        Optional ByVal FileSystem As IFileSystemService = Nothing) As WebDriverSessionClient

    Dim actual_file_system As IFileSystemService
    If Not (Settings Is Nothing) Then
        If Settings.DownloadEnabled Then Set actual_file_system = pResolveFileSystem(FileSystem, "Function New_WebDriverSessionClient")
    End If

    Dim result_value As WebDriverSessionClient
    Set result_value = New WebDriverSessionClient
    Call result_value.Initialize(Client, Settings, actual_file_system)

    Set New_WebDriverSessionClient = result_value
End Function

'* WebDriver HTTP client を生成します。
'*
'* @param Settings Web情報取得の設定。
'* @return 初期化済みの WebDriver HTTP client。
'*
'* @details
'* Settings の port と timeout を使って WebDriverClient を初期化します。
Public Function New_WebDriverClient(ByVal Settings As IToolSettings) As WebDriverClient
    Dim result_value As WebDriverClient
    Set result_value = New WebDriverClient
    Call result_value.Initialize(Settings.WebDriverPort, Settings.TimeoutSeconds)

    Set New_WebDriverClient = result_value
End Function

'* WebDriver port probe を生成します。
'*
'* @return 新しい WebDriver port probe。
'*
'* @details
'* 実 port 使用状況確認の既定実装を生成します。
Public Function New_WebDriverPortProbe() As WebDriverPortProbe
    Dim result_value As WebDriverPortProbe
    Set result_value = New WebDriverPortProbe

    Set New_WebDriverPortProbe = result_value
End Function

'* WebDriver process 管理を生成します。
'*
'* @param FileSystem ファイル存在確認に使うサービス。
'* @param PortProbe port 使用状況確認に使う probe。
'* @return 初期化済みの WebDriver process 管理。
'*
'* @details
'* New_ 系の処理は生成と Initialize 呼び出しに留め、実処理は WebDriverProcess に委譲します。
Public Function New_WebDriverProcess( _
        ByVal FileSystem As IFileSystemService, _
        ByVal PortProbe As IWebDriverPortProbe) As WebDriverProcess

    Dim result_value As WebDriverProcess
    Set result_value = New WebDriverProcess
    Call result_value.Initialize(FileSystem, PortProbe)

    Set New_WebDriverProcess = result_value
End Function

'* WebDriver smoke runner を生成します。
'*
'* @param Process WebDriver process 操作。
'* @param SessionClient WebDriver session client。
'* @param Settings Web情報取得の設定。
'* @return 初期化済みの WebDriver smoke runner。
'*
'* @details
'* New_ 系の処理は生成と Initialize 呼び出しに留め、実処理は WebDriverSmokeRunner に委譲します。
Public Function New_WebDriverSmokeRunner( _
        ByVal Process As IWebDriverProcess, _
        ByVal SessionClient As WebDriverSessionClient, _
        ByVal Settings As IToolSettings) As WebDriverSmokeRunner

    Dim result_value As WebDriverSmokeRunner
    Set result_value = New WebDriverSmokeRunner
    Call result_value.Initialize(Process, SessionClient, Settings)

    Set New_WebDriverSmokeRunner = result_value
End Function

'* Web navigation diagnostic runner を生成します。
'*
'* @param Process WebDriver process 操作。
'* @param SessionClient WebDriver session client。
'* @param Settings Web情報取得の設定。
'* @return 初期化済みの Web navigation diagnostic runner。
'*
'* @details
'* New_ 系の処理は生成と Initialize 呼び出しに留め、実処理は WebNavDiagnosticRunner に委譲します。
Public Function New_WebNavDiagnosticRunner( _
        ByVal Process As IWebDriverProcess, _
        ByVal SessionClient As WebDriverSessionClient, _
        ByVal Settings As IToolSettings) As WebNavDiagnosticRunner

    Dim result_value As WebNavDiagnosticRunner
    Set result_value = New WebNavDiagnosticRunner
    Call result_value.Initialize(Process, SessionClient, Settings)

    Set New_WebNavDiagnosticRunner = result_value
End Function

'* Web collection runner を生成します。
'*
'* @param Process WebDriver process 操作。
'* @param SessionClient WebDriver session client。
'* @param Settings Web情報取得の設定。
'* @return 初期化済みの Web collection runner。
'*
'* @details
'* New_ 系の処理は生成と Initialize 呼び出しに留め、実処理は WebCollectionRunner に委譲します。
Public Function New_WebCollectionRunner( _
        ByVal Process As IWebDriverProcess, _
        ByVal SessionClient As WebDriverSessionClient, _
        ByVal Settings As IToolSettings) As WebCollectionRunner

    Dim result_value As WebCollectionRunner
    Set result_value = New WebCollectionRunner
    Call result_value.Initialize(Process, SessionClient, Settings)

    Set New_WebCollectionRunner = result_value
End Function

Private Function pResolveFileSystem( _
        ByVal FileSystem As IFileSystemService, _
        ByVal SourceName As String) As IFileSystemService

    If FileSystem Is Nothing Then
        If FsSrv Is Nothing Then
            Err.Raise Number:=vbObjectError + 1, Source:=SourceName, Description:="FileSystem は必須です。"
        End If
        Set pResolveFileSystem = FsSrv
    Else
        Set pResolveFileSystem = FileSystem
    End If
End Function
