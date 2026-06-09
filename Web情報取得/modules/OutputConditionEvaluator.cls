VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutputConditionEvaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! OutputConditionExpression を評価するクラスです。
'!
' #############################################################################

Private Const C_SOURCE As String = "Class OutputConditionEvaluator"
Private Const C_EVAL_MODE_VALIDATE As Long = 0
Private Const C_EVAL_MODE_FULL_VALUES As Long = 1
Private Const C_EVAL_MODE_REFERENCED_VALUES As Long = 2

Private pIsInitialized As Boolean
Private pExpression As String
Private pDetailDefinitions As ObjectList
Private pReferencedDefinitions As ObjectList
Private pEvalValues As ArrayObject
Private pEvalMode As Long
Private pPosition As Long

'* 条件式が設定されているか否か。
Public Property Get IsEnabled() As Boolean
    Call pCheckInit
    IsEnabled = (pExpression <> "")
End Property

'* 条件式が参照する詳細ページ列定義を返します。
Public Property Get ReferencedColumnDefinitions() As ObjectList
    Call pCheckInit

    Dim result_value As ObjectList
    Set result_value = New_ObjectList("DetailColumnDefinition")

    Dim def_idx As Long
    For def_idx = 0 To pReferencedDefinitions.Count - 1
        Dim src_def As DetailColumnDefinition
        Set src_def = pReferencedDefinitions.Item(def_idx)
        ' 条件判定用の先読みでは、参照列が読めない場合も空文字列として評価します。
        Call result_value.Add(New_DetailColumnDefinition( _
                src_def.OutputColumnName, _
                src_def.Selector, _
                ExtractType:=src_def.ExtractType, _
                AttributeName:=src_def.AttributeName, _
                IsRequired:=False, _
                BlankMode:="AllowBlank"))
    Next def_idx

    Set ReferencedColumnDefinitions = result_value
End Property

'* 条件式を初期化します。
'*
'* @param Expression 条件式。空の場合は常に出力対象です。
'* @param DetailColumnDefinitions 詳細ページ列定義。
Public Sub Initialize(ByVal Expression As String, ByVal DetailColumnDefinitions As ObjectList)
    If DetailColumnDefinitions Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="DetailColumnDefinitions は必須です。"
    End If

    pExpression = Trim$(Expression)
    Set pDetailDefinitions = DetailColumnDefinitions
    Set pReferencedDefinitions = New_ObjectList("DetailColumnDefinition")
    pEvalMode = C_EVAL_MODE_VALIDATE
    Set pEvalValues = Nothing
    pPosition = 1

    Call pValidateUniqueOutputColumnNames
    If pExpression <> "" Then
        Call pValidateExpression
    End If

    pIsInitialized = True
End Sub

'* 詳細ページ列値に基づき、出力対象か否かを返します。
Public Function ShouldOutput(ByVal DetailColumnValues As ArrayObject) As Boolean
    Call pCheckInit

    If Not IsEnabled Then
        ShouldOutput = True
        Exit Function
    End If

    Set pEvalValues = DetailColumnValues
    pEvalMode = C_EVAL_MODE_FULL_VALUES
    pPosition = 1

    Dim result_value As Boolean
    result_value = pParseOrExpression()
    Call pRequireExpressionEnd

    Set pEvalValues = Nothing
    ShouldOutput = result_value
End Function

'* 参照列だけを読み取った値に基づき、出力対象か否かを返します。
Public Function ShouldOutputByReferencedValues(ByVal ReferencedValues As ArrayObject) As Boolean
    Call pCheckInit

    If Not IsEnabled Then
        ShouldOutputByReferencedValues = True
        Exit Function
    End If

    Set pEvalValues = ReferencedValues
    pEvalMode = C_EVAL_MODE_REFERENCED_VALUES
    pPosition = 1

    Dim result_value As Boolean
    result_value = pParseOrExpression()
    Call pRequireExpressionEnd

    Set pEvalValues = Nothing
    ShouldOutputByReferencedValues = result_value
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="初期化されていません。"
    End If
End Sub

Private Sub pValidateUniqueOutputColumnNames()
    Dim left_idx As Long
    For left_idx = 0 To pDetailDefinitions.Count - 1
        Dim left_def As DetailColumnDefinition
        Set left_def = pDetailDefinitions.Item(left_idx)

        Dim right_idx As Long
        For right_idx = left_idx + 1 To pDetailDefinitions.Count - 1
            Dim right_def As DetailColumnDefinition
            Set right_def = pDetailDefinitions.Item(right_idx)
            If StrComp(left_def.OutputColumnName, right_def.OutputColumnName, vbBinaryCompare) = 0 Then
                Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="OutputColumnName が重複しています。(" & left_def.OutputColumnName & ")"
            End If
        Next right_idx
    Next left_idx
End Sub

Private Sub pValidateExpression()
    pEvalMode = C_EVAL_MODE_VALIDATE
    Set pEvalValues = Nothing
    pPosition = 1

    Call pParseOrExpression
    Call pRequireExpressionEnd
End Sub

Private Function pParseOrExpression() As Boolean
    Dim result_value As Boolean
    result_value = pParseAndExpression()

    Do While pConsumeKeyword("OR")
        Dim right_value As Boolean
        right_value = pParseAndExpression()
        result_value = result_value Or right_value
    Loop

    pParseOrExpression = result_value
End Function

Private Function pParseAndExpression() As Boolean
    Dim result_value As Boolean
    result_value = pParseNotExpression()

    Do While pConsumeKeyword("AND")
        Dim right_value As Boolean
        right_value = pParseNotExpression()
        result_value = result_value And right_value
    Loop

    pParseAndExpression = result_value
End Function

Private Function pParseNotExpression() As Boolean
    Dim result_value As Boolean

    If pConsumeKeyword("NOT") Then
        result_value = Not pParseNotExpression()
    Else
        result_value = pParsePrimary()
    End If

    pParseNotExpression = result_value
End Function

Private Function pParsePrimary() As Boolean
    Dim result_value As Boolean

    Call pSkipSpaces
    If pConsumeChar("(") Then
        result_value = pParseOrExpression()
        Call pRequireChar(")")
    Else
        result_value = pParseComparison()
    End If

    pParsePrimary = result_value
End Function

Private Function pParseComparison() As Boolean
    Dim left_value As String
    left_value = pParseColumnReferenceValue()

    Dim operator_text As String
    operator_text = pParseComparisonOperator()

    Dim right_value As String
    right_value = pParseStringLiteral()

    pParseComparison = pEvaluateComparison(left_value, operator_text, right_value)
End Function

Private Function pParseColumnReferenceValue() As String
    Call pSkipSpaces
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
    def_idx = pFindDefinitionIndex(ColumnName, pDetailDefinitions)
    If def_idx < 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="条件式が未定義の OutputColumnName を参照しています。(" & ColumnName & ")"
    End If

    Select Case pEvalMode
        Case C_EVAL_MODE_VALIDATE
            Call pAddReferencedDefinition(pDetailDefinitions.Item(def_idx))
            pReadColumnValue = ""
        Case C_EVAL_MODE_FULL_VALUES
            pReadColumnValue = Trim$(CStr(pEvalValues.Item(def_idx)))
        Case C_EVAL_MODE_REFERENCED_VALUES
            Dim ref_idx As Long
            ref_idx = pFindDefinitionIndex(ColumnName, pReferencedDefinitions)
            If ref_idx < 0 Then
                Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="参照列値に OutputColumnName が含まれていません。(" & ColumnName & ")"
            End If
            pReadColumnValue = Trim$(CStr(pEvalValues.Item(ref_idx)))
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="未対応の評価モードです。(" & CStr(pEvalMode) & ")"
    End Select
End Function

Private Sub pAddReferencedDefinition(ByVal DetailDefinition As DetailColumnDefinition)
    If 0 <= pFindDefinitionIndex(DetailDefinition.OutputColumnName, pReferencedDefinitions) Then Exit Sub
    Call pReferencedDefinitions.Add(DetailDefinition)
End Sub

Private Function pFindDefinitionIndex(ByVal ColumnName As String, ByVal Definitions As ObjectList) As Long
    Dim result_value As Long
    result_value = -1

    Dim def_idx As Long
    For def_idx = 0 To Definitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = Definitions.Item(def_idx)
        If StrComp(detail_def.OutputColumnName, ColumnName, vbBinaryCompare) = 0 Then
            result_value = def_idx
            Exit For
        End If
    Next def_idx

    pFindDefinitionIndex = result_value
End Function

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
    Call pSkipSpaces
    If Not pConsumeChar(Chr$(34)) Then
        Call pRaiseSyntaxError("比較値はダブルクオートで囲んでください。")
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

    Call pRaiseSyntaxError("文字列リテラルの閉じダブルクオートが見つかりません。")
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

Private Sub pRequireChar(ByVal TargetChar As String)
    If Not pConsumeChar(TargetChar) Then
        Call pRaiseSyntaxError(TargetChar & " が見つかりません。")
    End If
End Sub

Private Sub pRequireExpressionEnd()
    Call pSkipSpaces
    If pPosition <= Len(pExpression) Then
        Call pRaiseSyntaxError("条件式の末尾に不正な文字があります。")
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
    Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="OutputConditionExpression の構文が不正です。" & Message & " (位置: " & CStr(pPosition) & ")"
End Sub
