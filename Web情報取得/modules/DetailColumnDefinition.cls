VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailColumnDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 詳細ページから値を抽出するための列定義です。
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pOutputColumnName As String
Private pSelector As String
Private pExtractType As String
Private pAttributeName As String
Private pIsRequired As Boolean
Private pBlankMode As String

'* 出力列名。
Public Property Get OutputColumnName() As String
    Call pCheckInit
    OutputColumnName = pOutputColumnName
End Property

'* DOM selector。
Public Property Get Selector() As String
    Call pCheckInit
    Selector = pSelector
End Property

'* 取得種別。
Public Property Get ExtractType() As String
    Call pCheckInit
    ExtractType = pExtractType
End Property

'* attribute 取得時の属性名。
Public Property Get AttributeName() As String
    Call pCheckInit
    AttributeName = pAttributeName
End Property

'* True の場合は必須列として扱います。
Public Property Get IsRequired() As Boolean
    Call pCheckInit
    IsRequired = pIsRequired
End Property

'* 空欄時の扱い。
Public Property Get BlankMode() As String
    Call pCheckInit
    BlankMode = pBlankMode
End Property

'* 詳細ページ列定義を初期化します。
'*
'* @param OutputColumnName 出力列名。
'* @param Selector DOM selector。
'* @param ExtractType 取得種別。
'* @param AttributeName attribute 取得時の属性名。
'* @param IsRequired True の場合は必須列として扱います。
'* @param BlankMode 空欄時の扱い。
'*
'* @details
'* OutputColumnName と Selector は必須です。ExtractType、AttributeName、IsRequired、BlankMode は
'* DOM 抽出時の列ごとの振る舞いを指定するために保持します。
Public Sub Initialize( _
        ByVal OutputColumnName As String, _
        ByVal Selector As String, _
        Optional ByVal ExtractType As String = "InnerText", _
        Optional ByVal AttributeName As String = "", _
        Optional ByVal IsRequired As Boolean = False, _
        Optional ByVal BlankMode As String = "AllowBlank")

    If Trim$(OutputColumnName) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailColumnDefinition", Description:="OutputColumnName は必須です。"
    End If
    If Trim$(Selector) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailColumnDefinition", Description:="Selector は必須です。(" & OutputColumnName & ")"
    End If

    pOutputColumnName = OutputColumnName
    pSelector = Selector
    pExtractType = ExtractType
    pAttributeName = AttributeName
    pIsRequired = IsRequired
    pBlankMode = BlankMode
    pIsInitialized = True
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailColumnDefinition", Description:="初期化されていません。"
    End If
End Sub
