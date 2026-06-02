Attribute VB_Name = "Constructor"
Option Explicit
Option Base 0

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
        Optional ByVal BlankMode As String = "AllowBlank") As DetailColumnDefinition

    Dim result_value As DetailColumnDefinition
    Set result_value = New DetailColumnDefinition
    Call result_value.Initialize( _
            OutputColumnName, _
            Selector, _
            ExtractType:=ExtractType, _
            AttributeName:=AttributeName, _
            IsRequired:=IsRequired, _
            BlankMode:=BlankMode)

    Set New_DetailColumnDefinition = result_value
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
