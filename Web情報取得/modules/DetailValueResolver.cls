VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailValueResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 詳細ページ列定義に対応する抽出値と派生値を解決します。
'!
' #############################################################################

Private pDefinitions As ObjectList
Private pExtractionDefinitions As ObjectList
Private pIsInitialized As Boolean

'* WebDriver から直接抽出する列定義。
Public Property Get ExtractionColumnDefinitions() As ObjectList
    Call pCheckInit
    Set ExtractionColumnDefinitions = pExtractionDefinitions
End Property

'* 値解決を初期化します。
'*
'* @param DetailColumnDefinitions 詳細ページ列定義の一覧。
Public Sub Initialize(ByVal DetailColumnDefinitions As ObjectList)
    If DetailColumnDefinitions Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailValueResolver", Description:="DetailColumnDefinitions は必須です。"
    End If

    Set pDefinitions = DetailColumnDefinitions
    Set pExtractionDefinitions = New_ObjectList("DetailColumnDefinition")

    Dim def_idx As Long
    For def_idx = 0 To pDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(def_idx)
        If Not detail_def.IsDerived Then Call pExtractionDefinitions.Add(detail_def)
    Next def_idx

    pIsInitialized = True
End Sub

'* 通常抽出値から、列定義順の詳細値を解決します。
'*
'* @param ExtractionValues ExtractionColumnDefinitions と同じ順序の抽出値。
'* @return 元の詳細ページ列定義と同じ順序の値。
Public Function ResolveValues(ByVal ExtractionValues As ArrayObject) As ArrayObject
    Call pCheckInit
    If ExtractionValues Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailValueResolver", Description:="ExtractionValues は必須です。"
    End If

    Dim result_values As ArrayObject
    Set result_values = New ArrayObject
    If 0 < pDefinitions.Count Then Call result_values.ReDimArray(0, pDefinitions.Count - 1)

    Dim extraction_idx As Long
    extraction_idx = 0

    Dim def_idx As Long
    For def_idx = 0 To pDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(def_idx)
        If detail_def.IsDerived Then
            Call result_values.Update(def_idx, "")
        Else
            Call result_values.Update(def_idx, CStr(ExtractionValues.Item(extraction_idx)))
            extraction_idx = extraction_idx + 1
        End If
    Next def_idx

    For def_idx = 0 To pDefinitions.Count - 1
        Set detail_def = pDefinitions.Item(def_idx)
        If detail_def.IsDerived Then
            Dim ref_name As String
            ref_name = pParseSimpleColumnReference(detail_def.ValueExpression, detail_def.OutputColumnName)

            Dim ref_idx As Long
            ref_idx = pFindDefinitionIndex(ref_name)
            If ref_idx < 0 Then
                Err.Raise Number:=vbObjectError + 1, Source:="Class DetailValueResolver", Description:="ValueExpression の参照列が見つかりません。(" & detail_def.OutputColumnName & ": " & ref_name & ")"
            End If

            Call result_values.Update(def_idx, CStr(result_values.Item(ref_idx)))
        End If
    Next def_idx

    Set ResolveValues = result_values
End Function

Private Function pParseSimpleColumnReference( _
        ByVal ValueExpression As String, _
        ByVal OutputColumnName As String) As String

    Dim expression_text As String
    expression_text = Trim$(ValueExpression)

    If Len(expression_text) < 3 _
            Or Left$(expression_text, 1) <> "[" _
            Or Right$(expression_text, 1) <> "]" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailValueResolver", Description:="ValueExpression は [OutputColumnName] の形式で指定してください。(" & OutputColumnName & ")"
    End If

    Dim result_value As String
    result_value = Mid$(expression_text, 2, Len(expression_text) - 2)
    If result_value = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailValueResolver", Description:="ValueExpression の参照列名は必須です。(" & OutputColumnName & ")"
    End If

    pParseSimpleColumnReference = result_value
End Function

Private Function pFindDefinitionIndex(ByVal OutputColumnName As String) As Long
    Dim def_idx As Long
    For def_idx = 0 To pDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(def_idx)
        If StrComp(detail_def.OutputColumnName, OutputColumnName, vbBinaryCompare) = 0 Then
            pFindDefinitionIndex = def_idx
            Exit Function
        End If
    Next def_idx

    pFindDefinitionIndex = -1
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class DetailValueResolver", Description:="初期化されていません。"
    End If
End Sub
