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

Private Const C_SOURCE As String = "Class DetailValueResolver"
Private Const C_PARSE_MODE_VALIDATE As Long = 0
Private Const C_PARSE_MODE_EVALUATE As Long = 1
Private Const C_STATE_VISITING As Long = 1
Private Const C_STATE_RESOLVED As Long = 2

Private pDefinitions As ObjectList
Private pExtractionDefinitions As ObjectList
Private pResolvedValues As ArrayObject
Private pResolveState() As Long
Private pExpression As String
Private pPosition As Long
Private pParseMode As Long
Private pCurrentDefinitionIndex As Long
Private pCurrentReferences As Collection
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
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="DetailColumnDefinitions は必須です。"
    End If

    Set pDefinitions = DetailColumnDefinitions
    Set pExtractionDefinitions = New_ObjectList("DetailColumnDefinition")
    Set pResolvedValues = Nothing
    Set pCurrentReferences = Nothing
    pExpression = ""
    pPosition = 1
    pCurrentDefinitionIndex = -1

    Dim def_idx As Long
    For def_idx = 0 To pDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(def_idx)
        If Not detail_def.IsDerived Then Call pExtractionDefinitions.Add(detail_def)
    Next def_idx

    Call pValidateDerivedExpressions

    pIsInitialized = True
End Sub

Private Sub pValidateDerivedExpressions()
    Dim def_idx As Long
    For def_idx = 0 To pDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(def_idx)
        If detail_def.IsDerived Then Call pValidateExpressionReferences(def_idx)
    Next def_idx

    If pDefinitions.Count = 0 Then Exit Sub

    Dim visit_state() As Long
    ReDim visit_state(0 To pDefinitions.Count - 1)

    For def_idx = 0 To pDefinitions.Count - 1
        Set detail_def = pDefinitions.Item(def_idx)
        If detail_def.IsDerived Then Call pVisitDerivedDefinition(def_idx, visit_state)
    Next def_idx
End Sub

Private Sub pValidateExpressionReferences(ByVal DefinitionIndex As Long)
    Dim references As Collection
    Set references = pCollectReferences(DefinitionIndex)

    Dim ref_name As Variant
    For Each ref_name In references
        Dim ref_idx As Long
        ref_idx = pFindDefinitionIndex(CStr(ref_name))
        If ref_idx < 0 Then
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValueExpression が未定義の OutputColumnName を参照しています。(" & CStr(ref_name) & ")"
        End If
        If ref_idx = DefinitionIndex Then
            Dim detail_def As DetailColumnDefinition
            Set detail_def = pDefinitions.Item(DefinitionIndex)
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValueExpression が自己参照しています。(" & detail_def.OutputColumnName & ")"
        End If
    Next ref_name
End Sub

Private Sub pVisitDerivedDefinition( _
        ByVal DefinitionIndex As Long, _
        ByRef VisitState() As Long)

    If VisitState(DefinitionIndex) = C_STATE_RESOLVED Then Exit Sub
    If VisitState(DefinitionIndex) = C_STATE_VISITING Then
        Dim loop_def As DetailColumnDefinition
        Set loop_def = pDefinitions.Item(DefinitionIndex)
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="派生列同士の循環参照があります。(" & loop_def.OutputColumnName & ")"
    End If

    VisitState(DefinitionIndex) = C_STATE_VISITING

    Dim references As Collection
    Set references = pCollectReferences(DefinitionIndex)

    Dim ref_name As Variant
    For Each ref_name In references
        Dim ref_idx As Long
        ref_idx = pFindDefinitionIndex(CStr(ref_name))
        If ref_idx < 0 Then
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValueExpression が未定義の OutputColumnName を参照しています。(" & CStr(ref_name) & ")"
        End If
        If ref_idx = DefinitionIndex Then
            Dim self_def As DetailColumnDefinition
            Set self_def = pDefinitions.Item(DefinitionIndex)
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValueExpression が自己参照しています。(" & self_def.OutputColumnName & ")"
        End If

        Dim ref_def As DetailColumnDefinition
        Set ref_def = pDefinitions.Item(ref_idx)
        If ref_def.IsDerived Then Call pVisitDerivedDefinition(ref_idx, VisitState)
    Next ref_name

    VisitState(DefinitionIndex) = C_STATE_RESOLVED
End Sub

Private Function pCollectReferences(ByVal DefinitionIndex As Long) As Collection
    Dim detail_def As DetailColumnDefinition
    Set detail_def = pDefinitions.Item(DefinitionIndex)

    Set pCurrentReferences = New Collection
    pCurrentDefinitionIndex = DefinitionIndex
    pParseMode = C_PARSE_MODE_VALIDATE
    pExpression = Trim$(detail_def.ValueExpression)
    pPosition = 1

    Call pParseValueExpression
    Call pRequireExpressionEnd

    Set pCollectReferences = pCurrentReferences
    Set pCurrentReferences = Nothing
End Function

'* 通常抽出値から、列定義順の詳細値を解決します。
'*
'* @param ExtractionValues ExtractionColumnDefinitions と同じ順序の抽出値。
'* @return 元の詳細ページ列定義と同じ順序の値。
Public Function ResolveValues(ByVal ExtractionValues As ArrayObject) As ArrayObject
    Call pCheckInit
    If ExtractionValues Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ExtractionValues は必須です。"
    End If

    Set pResolvedValues = New ArrayObject
    If 0 < pDefinitions.Count Then
        Call pResolvedValues.ReDimArray(0, pDefinitions.Count - 1)
        ReDim pResolveState(0 To pDefinitions.Count - 1)
    End If

    Dim extraction_idx As Long
    extraction_idx = 0

    Dim def_idx As Long
    For def_idx = 0 To pDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(def_idx)
        If detail_def.IsDerived Then
            Call pResolvedValues.Update(def_idx, "")
        Else
            Call pResolvedValues.Update(def_idx, CStr(ExtractionValues.Item(extraction_idx)))
            pResolveState(def_idx) = C_STATE_RESOLVED
            extraction_idx = extraction_idx + 1
        End If
    Next def_idx

    For def_idx = 0 To pDefinitions.Count - 1
        Set detail_def = pDefinitions.Item(def_idx)
        If detail_def.IsDerived Then Call pResolveDefinitionValue(def_idx)
    Next def_idx

    Set ResolveValues = pResolvedValues
    Set pResolvedValues = Nothing
End Function

Private Function pResolveDefinitionValue(ByVal DefinitionIndex As Long) As String
    If pResolveState(DefinitionIndex) = C_STATE_RESOLVED Then
        pResolveDefinitionValue = CStr(pResolvedValues.Item(DefinitionIndex))
        Exit Function
    End If
    If pResolveState(DefinitionIndex) = C_STATE_VISITING Then
        Dim loop_def As DetailColumnDefinition
        Set loop_def = pDefinitions.Item(DefinitionIndex)
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="派生列同士の循環参照があります。(" & loop_def.OutputColumnName & ")"
    End If

    Dim detail_def As DetailColumnDefinition
    Set detail_def = pDefinitions.Item(DefinitionIndex)

    pResolveState(DefinitionIndex) = C_STATE_VISITING
    pCurrentDefinitionIndex = DefinitionIndex
    pParseMode = C_PARSE_MODE_EVALUATE
    pExpression = Trim$(detail_def.ValueExpression)
    pPosition = 1

    Dim result_value As String
    result_value = pParseValueExpression()
    Call pRequireExpressionEnd

    Call pResolvedValues.Update(DefinitionIndex, result_value)
    pResolveState(DefinitionIndex) = C_STATE_RESOLVED

    pResolveDefinitionValue = result_value
End Function

Private Function pParseValueExpression() As String
    Call pSkipSpaces

    Dim result_value As String
    If pConsumeKeyword("IF") Then
        Call pRequireChar("(")

        Dim condition_value As Boolean
        condition_value = pParseConditionOrExpression()

        Call pRequireChar(",")

        Dim true_value As String
        true_value = pParseValueExpression()

        Call pRequireChar(",")

        Dim false_value As String
        false_value = pParseValueExpression()

        Call pRequireChar(")")

        If pParseMode = C_PARSE_MODE_EVALUATE Then
            If condition_value Then
                result_value = true_value
            Else
                result_value = false_value
            End If
        Else
            result_value = ""
        End If
    ElseIf pPeekChar() = "[" Then
        result_value = pParseColumnReferenceValue()
    ElseIf pPeekChar() = Chr$(34) Then
        result_value = pParseStringLiteral()
    Else
        Call pRaiseSyntaxError("値式は列参照、ダブルクォート文字列リテラル、IF(条件, 値式, 値式) のいずれかで指定してください。")
    End If

    pParseValueExpression = result_value
End Function

Private Function pParseConditionOrExpression() As Boolean
    Dim result_value As Boolean
    result_value = pParseConditionAndExpression()

    Do While pConsumeKeyword("OR")
        Dim right_value As Boolean
        right_value = pParseConditionAndExpression()
        result_value = result_value Or right_value
    Loop

    pParseConditionOrExpression = result_value
End Function

Private Function pParseConditionAndExpression() As Boolean
    Dim result_value As Boolean
    result_value = pParseConditionNotExpression()

    Do While pConsumeKeyword("AND")
        Dim right_value As Boolean
        right_value = pParseConditionNotExpression()
        result_value = result_value And right_value
    Loop

    pParseConditionAndExpression = result_value
End Function

Private Function pParseConditionNotExpression() As Boolean
    Dim result_value As Boolean

    If pConsumeKeyword("NOT") Then
        result_value = Not pParseConditionNotExpression()
    Else
        result_value = pParseConditionPrimary()
    End If

    pParseConditionNotExpression = result_value
End Function

Private Function pParseConditionPrimary() As Boolean
    Dim result_value As Boolean

    Call pSkipSpaces
    If pConsumeChar("(") Then
        result_value = pParseConditionOrExpression()
        Call pRequireChar(")")
    Else
        result_value = pParseConditionComparison()
    End If

    pParseConditionPrimary = result_value
End Function

Private Function pParseConditionComparison() As Boolean
    Dim left_value As String
    left_value = pParseColumnReferenceValue()

    Dim operator_text As String
    operator_text = pParseComparisonOperator()

    Dim right_value As String
    right_value = pParseStringLiteral()

    pParseConditionComparison = pEvaluateComparison(left_value, operator_text, right_value)
End Function

Private Function pParseColumnReferenceValue() As String
    If Not pConsumeChar("[") Then
        Call pRaiseSyntaxError("列参照は [OutputColumnName] の形式で指定してください。")
    End If

    Dim start_pos As Long
    start_pos = pPosition

    Do While pPosition <= Len(pExpression)
        If Mid$(pExpression, pPosition, 1) = "]" Then Exit Do
        pPosition = pPosition + 1
    Loop

    If Len(pExpression) < pPosition Then
        Call pRaiseSyntaxError("列参照の ] が見つかりません。")
    End If

    Dim column_name As String
    column_name = Mid$(pExpression, start_pos, pPosition - start_pos)
    If column_name = "" Then
        Call pRaiseSyntaxError("列参照の OutputColumnName が空です。")
    End If

    pPosition = pPosition + 1
    pParseColumnReferenceValue = pReadColumnValue(column_name)
End Function

Private Function pReadColumnValue(ByVal ColumnName As String) As String
    Dim def_idx As Long
    def_idx = pFindDefinitionIndex(ColumnName)
    If def_idx < 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValueExpression が未定義の OutputColumnName を参照しています。(" & ColumnName & ")"
    End If

    If pParseMode = C_PARSE_MODE_VALIDATE Then
        Call pAddCurrentReference(ColumnName)
        pReadColumnValue = ""
        Exit Function
    End If

    Dim detail_def As DetailColumnDefinition
    Set detail_def = pDefinitions.Item(def_idx)
    If detail_def.IsDerived Then
        pReadColumnValue = pResolveDefinitionValue(def_idx)
    Else
        pReadColumnValue = CStr(pResolvedValues.Item(def_idx))
    End If
End Function

Private Sub pAddCurrentReference(ByVal ColumnName As String)
    If pCurrentReferences Is Nothing Then Exit Sub
    pCurrentReferences.Add ColumnName
End Sub

Private Function pParseComparisonOperator() As String
    Call pSkipSpaces

    Dim result_value As String
    If pConsumeText("==") Then
        result_value = "=="
    ElseIf pConsumeText("!=") Then
        result_value = "!="
    Else
        Call pRaiseSyntaxError("比較演算子は == または != を指定してください。")
    End If

    pParseComparisonOperator = result_value
End Function

Private Function pParseStringLiteral() As String
    If Not pConsumeChar(Chr$(34)) Then
        Call pRaiseSyntaxError("文字列リテラルはダブルクォートで囲んでください。")
    End If

    Dim raw_value As String
    raw_value = ""

    Do While pPosition <= Len(pExpression)
        Dim current_char As String
        current_char = Mid$(pExpression, pPosition, 1)

        If current_char = Chr$(34) Then
            pPosition = pPosition + 1
            pParseStringLiteral = pJsonUnescape(raw_value)
            Exit Function
        End If

        If current_char = "\" Then
            If Len(pExpression) < pPosition + 1 Then
                Call pRaiseSyntaxError("文字列リテラル内のエスケープが完了していません。")
            End If

            Dim escaped_char As String
            escaped_char = Mid$(pExpression, pPosition + 1, 1)
            Select Case escaped_char
                Case "n", "r", "t", "\", Chr$(34)
                Case Else
                    Call pRaiseSyntaxError("未知のエスケープです。(\" & escaped_char & ")")
            End Select

            raw_value = raw_value & Mid$(pExpression, pPosition, 2)
            pPosition = pPosition + 2
        Else
            raw_value = raw_value & current_char
            pPosition = pPosition + 1
        End If
    Loop

    Call pRaiseSyntaxError("文字列リテラルの閉じダブルクォートが見つかりません。")
End Function

Private Function pEvaluateComparison( _
        ByVal LeftValue As String, _
        ByVal OperatorText As String, _
        ByVal RightValue As String) As Boolean

    Dim left_text As String
    left_text = Trim$(LeftValue)

    Dim right_text As String
    right_text = Trim$(RightValue)

    Dim result_value As Boolean
    Select Case OperatorText
        Case "=="
            result_value = (StrComp(left_text, right_text, vbBinaryCompare) = 0)
        Case "!="
            result_value = (StrComp(left_text, right_text, vbBinaryCompare) <> 0)
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="未対応の比較演算子です。(" & OperatorText & ")"
    End Select

    pEvaluateComparison = result_value
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

Private Function pConsumeKeyword(ByVal Keyword As String) As Boolean
    Call pSkipSpaces

    Dim keyword_len As Long
    keyword_len = Len(Keyword)
    If Len(pExpression) - pPosition + 1 < keyword_len Then
        pConsumeKeyword = False
        Exit Function
    End If
    If StrComp(Mid$(pExpression, pPosition, keyword_len), Keyword, vbTextCompare) <> 0 Then
        pConsumeKeyword = False
        Exit Function
    End If

    Dim before_char As String
    If pPosition <= 1 Then
        before_char = ""
    Else
        before_char = Mid$(pExpression, pPosition - 1, 1)
    End If

    Dim after_char As String
    If Len(pExpression) < pPosition + keyword_len Then
        after_char = ""
    Else
        after_char = Mid$(pExpression, pPosition + keyword_len, 1)
    End If

    If pIsWordChar(before_char) Or pIsWordChar(after_char) Then
        pConsumeKeyword = False
        Exit Function
    End If

    pPosition = pPosition + keyword_len
    pConsumeKeyword = True
End Function

Private Function pConsumeText(ByVal TargetText As String) As Boolean
    Call pSkipSpaces

    If StrComp(Mid$(pExpression, pPosition, Len(TargetText)), TargetText, vbBinaryCompare) = 0 Then
        pPosition = pPosition + Len(TargetText)
        pConsumeText = True
    Else
        pConsumeText = False
    End If
End Function

Private Function pConsumeChar(ByVal TargetChar As String) As Boolean
    Call pSkipSpaces

    If pPosition <= Len(pExpression) And Mid$(pExpression, pPosition, 1) = TargetChar Then
        pPosition = pPosition + 1
        pConsumeChar = True
    Else
        pConsumeChar = False
    End If
End Function

Private Function pPeekChar() As String
    Call pSkipSpaces

    If pPosition <= Len(pExpression) Then
        pPeekChar = Mid$(pExpression, pPosition, 1)
    Else
        pPeekChar = ""
    End If
End Function

Private Sub pRequireChar(ByVal TargetChar As String)
    If Not pConsumeChar(TargetChar) Then
        Call pRaiseSyntaxError(TargetChar & " が見つかりません。")
    End If
End Sub

Private Sub pRequireExpressionEnd()
    Call pSkipSpaces
    If pPosition <= Len(pExpression) Then
        Call pRaiseSyntaxError("式の末尾に不正な文字があります。")
    End If
End Sub

Private Sub pSkipSpaces()
    Do While pPosition <= Len(pExpression)
        Dim current_char As String
        current_char = Mid$(pExpression, pPosition, 1)
        If current_char <> " " And current_char <> vbTab And current_char <> vbCr And current_char <> vbLf Then Exit Do
        pPosition = pPosition + 1
    Loop
End Sub

Private Function pIsWordChar(ByVal TargetChar As String) As Boolean
    If TargetChar = "" Then
        pIsWordChar = False
    Else
        pIsWordChar = _
                (("A" <= TargetChar And TargetChar <= "Z") _
                Or ("a" <= TargetChar And TargetChar <= "z") _
                Or ("0" <= TargetChar And TargetChar <= "9") _
                Or TargetChar = "_")
    End If
End Function

Private Function pJsonUnescape(ByVal TargetText As String) As String
    Dim result_value As String
    result_value = TargetText
    result_value = Replace(result_value, "\n", vbLf)
    result_value = Replace(result_value, "\r", vbCr)
    result_value = Replace(result_value, "\t", vbTab)
    result_value = Replace(result_value, "\" & Chr$(34), Chr$(34))
    result_value = Replace(result_value, "\\", "\")

    pJsonUnescape = result_value
End Function

Private Sub pRaiseSyntaxError(ByVal Message As String)
    Dim output_column_name As String
    output_column_name = ""
    If 0 <= pCurrentDefinitionIndex And pCurrentDefinitionIndex < pDefinitions.Count Then
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(pCurrentDefinitionIndex)
        output_column_name = detail_def.OutputColumnName
    End If

    Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValueExpression の構文が不正です。" & Message & " (" & output_column_name & ", 位置: " & CStr(pPosition) & ")"
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="初期化されていません。"
    End If
End Sub
