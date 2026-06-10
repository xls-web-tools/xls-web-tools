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
Private pValueExpression As String
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

'* 値式。
Public Property Get ValueExpression() As String
    Call pCheckInit
    ValueExpression = pValueExpression
End Property

'* 派生列か否か。
Public Property Get IsDerived() As Boolean
    Call pCheckInit
    IsDerived = (Trim$(pValueExpression) <> "")
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
'* @param ValueExpression 値式。
'* @param BlankMode 空欄時の扱い。
'*
'* @details
'* OutputColumnName は必須です。通常の抽出列は Selector を、派生列は ValueExpression を指定します。
'* ExtractType、AttributeName、IsRequired、BlankMode は DOM 抽出時の列ごとの振る舞いを指定するために保持します。
Public Sub Initialize( _
        ByVal OutputColumnName As String, _
        ByVal Selector As String, _
        Optional ByVal ExtractType As String = "InnerText", _
        Optional ByVal AttributeName As String = "", _
        Optional ByVal IsRequired As Boolean = False, _
        Optional ByVal ValueExpression As String = "", _
        Optional ByVal BlankMode As String = "AllowBlank")

    If Trim$(OutputColumnName) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailColumnDefinition", Description:="OutputColumnName は必須です。"
    End If

    Dim has_selector As Boolean
    has_selector = (Trim$(Selector) <> "")

    Dim has_value_expression As Boolean
    has_value_expression = (Trim$(ValueExpression) <> "")

    If has_selector = has_value_expression Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailColumnDefinition", Description:="Selector と ValueExpression はどちらか一方だけを指定してください。(" & OutputColumnName & ")"
    End If
    If has_value_expression And IsRequired Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailColumnDefinition", Description:="派生列では IsRequired=True を指定できません。(" & OutputColumnName & ")"
    End If

    pOutputColumnName = OutputColumnName
    pSelector = Selector
    pExtractType = ExtractType
    pAttributeName = AttributeName
    pIsRequired = IsRequired
    pValueExpression = ValueExpression
    pBlankMode = BlankMode
    pIsInitialized = True
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailColumnDefinition", Description:="初期化されていません。"
    End If
End Sub
