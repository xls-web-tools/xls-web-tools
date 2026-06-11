VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const C_SOURCE As String = "Class ColumnExpression"
Private Const C_NODE_KIND_BOOLEAN_LITERAL As String = "BooleanLiteral"
Private Const C_NODE_KIND_STRING_LITERAL As String = "StringLiteral"
Private Const C_NODE_KIND_COLUMN_REFERENCE As String = "ColumnReference"
Private Const C_NODE_KIND_COMPARISON As String = "Comparison"
Private Const C_NODE_KIND_AND As String = "And"
Private Const C_NODE_KIND_OR As String = "Or"
Private Const C_NODE_KIND_NOT As String = "Not"
Private Const C_NODE_KIND_IF As String = "If"

Private pExpression As String
Private pExpressionKind As String
Private pSourceName As String
Private pPosition As Long
Private pRootNode As ColumnExpressionNode
Private pReferencedColumnNames As ObjectList
Private pReferencedColumnNameSet As ObjectSet
Private pIsInitialized As Boolean

Public Property Get Expression() As String
    Call pCheckInit
    Expression = pExpression
End Property

Public Property Get ExpressionKind() As String
    Call pCheckInit
    ExpressionKind = pExpressionKind
End Property

Public Property Get SourceName() As String
    Call pCheckInit
    SourceName = pSourceName
End Property

Public Property Get ReferencedColumnNames() As ObjectList
    Call pCheckInit
    Set ReferencedColumnNames = pReferencedColumnNames.CopyList()
End Property

Public Sub Initialize( _
        ByVal Expression As String, _
        ByVal ExpressionKind As String, _
        Optional ByVal SourceName As String = "")

    pExpression = Expression
    pExpressionKind = ExpressionKind
    pSourceName = SourceName
    pPosition = 1
    Set pReferencedColumnNames = New ObjectList
    Set pReferencedColumnNameSet = New ObjectSet

    Select Case pExpressionKind
        Case G_COLUMN_EXPRESSION_KIND_CONDITION
            If Trim$(pExpression) = "" Then
                Set pRootNode = pNewBooleanNode(True)
            Else
                Set pRootNode = pParseConditionOrExpression()
                Call pRequireExpressionEnd
            End If
        Case G_COLUMN_EXPRESSION_KIND_VALUE
            Set pRootNode = pParseValueExpression()
            Call pRequireExpressionEnd
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="未対応の列参照記法種別です。(" & ExpressionKind & ")"
    End Select

    pIsInitialized = True
End Sub

Private Function pParseConditionOrExpression() As ColumnExpressionNode
    Dim left_node As ColumnExpressionNode
    Set left_node = pParseConditionAndExpression()

    Do While pConsumeKeyword("OR")
        Dim right_node As ColumnExpressionNode
        Set right_node = pParseConditionAndExpression()
        Set left_node = pNewBinaryNode(C_NODE_KIND_OR, "OR", left_node, right_node)
    Loop

    Set pParseConditionOrExpression = left_node
End Function

Private Function pParseValueExpression() As ColumnExpressionNode
    If pConsumeKeyword("IF") Then
        Call pRequireChar("(")

        Dim condition_node As ColumnExpressionNode
        Set condition_node = pParseConditionOrExpression()

        Call pRequireChar(",")

        Dim true_node As ColumnExpressionNode
        Set true_node = pParseValueExpression()

        Call pRequireChar(",")

        Dim false_node As ColumnExpressionNode
        Set false_node = pParseValueExpression()

        Call pRequireChar(")")

        Set pParseValueExpression = pNewTernaryNode(C_NODE_KIND_IF, condition_node, true_node, false_node)
        Exit Function
    End If

    Set pParseValueExpression = pParseValueOperandNode()
End Function

Private Function pNewBooleanNode(ByVal Value As Boolean) As ColumnExpressionNode
    If Value Then
        Set pNewBooleanNode = pNewTextNode(C_NODE_KIND_BOOLEAN_LITERAL, "True")
    Else
        Set pNewBooleanNode = pNewTextNode(C_NODE_KIND_BOOLEAN_LITERAL, "False")
    End If
End Function

Private Sub pRequireExpressionEnd()
    Call pSkipSpaces
    If pPosition <= Len(pExpression) Then
        Call pRaiseSyntaxError("末尾に不正な文字があります。")
    End If
End Sub

Public Function EvaluateCondition(ByVal ColumnValues As ObjectDictionary) As Boolean
    Call pCheckInit
    If pExpressionKind <> G_COLUMN_EXPRESSION_KIND_CONDITION Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:=pExpressionLabel() & " は条件として評価できません。"
    End If

    EvaluateCondition = pEvaluateConditionNode(pRootNode, ColumnValues)
End Function

Private Function pEvaluateConditionNode( _
        ByVal ExpressionNode As ColumnExpressionNode, _
        ByVal ColumnValues As ObjectDictionary) As Boolean

    Dim result_value As Boolean
    Select Case ExpressionNode.Kind
        Case C_NODE_KIND_BOOLEAN_LITERAL
            result_value = (ExpressionNode.TextValue = "True")
        Case C_NODE_KIND_COMPARISON
            result_value = pEvaluateComparison( _
                    pEvaluateValueNode(ExpressionNode.LeftNode, ColumnValues), _
                    ExpressionNode.OperatorText, _
                    pEvaluateValueNode(ExpressionNode.RightNode, ColumnValues))
        Case C_NODE_KIND_AND
            result_value = pEvaluateConditionNode(ExpressionNode.LeftNode, ColumnValues)
            If result_value Then
                result_value = pEvaluateConditionNode(ExpressionNode.RightNode, ColumnValues)
            End If
        Case C_NODE_KIND_OR
            result_value = pEvaluateConditionNode(ExpressionNode.LeftNode, ColumnValues)
            If Not result_value Then
                result_value = pEvaluateConditionNode(ExpressionNode.RightNode, ColumnValues)
            End If
        Case C_NODE_KIND_NOT
            result_value = Not pEvaluateConditionNode(ExpressionNode.LeftNode, ColumnValues)
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="条件として評価できない列参照記法ノードです。(" & ExpressionNode.Kind & ")"
    End Select

    pEvaluateConditionNode = result_value
End Function

Public Function EvaluateValue(ByVal ColumnValues As ObjectDictionary) As String
    Call pCheckInit
    If pExpressionKind <> G_COLUMN_EXPRESSION_KIND_VALUE Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:=pExpressionLabel() & " は値として評価できません。"
    End If

    EvaluateValue = pEvaluateValueNode(pRootNode, ColumnValues)
End Function

Private Function pEvaluateValueNode( _
        ByVal ExpressionNode As ColumnExpressionNode, _
        ByVal ColumnValues As ObjectDictionary) As String

    Dim result_value As String
    Select Case ExpressionNode.Kind
        Case C_NODE_KIND_STRING_LITERAL
            result_value = ExpressionNode.TextValue
        Case C_NODE_KIND_COLUMN_REFERENCE
            result_value = pReadColumnValue(ExpressionNode.TextValue, ColumnValues)
        Case C_NODE_KIND_IF
            If pEvaluateConditionNode(ExpressionNode.LeftNode, ColumnValues) Then
                result_value = pEvaluateValueNode(ExpressionNode.MiddleNode, ColumnValues)
            Else
                result_value = pEvaluateValueNode(ExpressionNode.RightNode, ColumnValues)
            End If
        Case Else
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="値として評価できない列参照記法ノードです。(" & ExpressionNode.Kind & ")"
    End Select

    pEvaluateValueNode = result_value
End Function

Private Function pParseConditionAndExpression() As ColumnExpressionNode
    Dim left_node As ColumnExpressionNode
    Set left_node = pParseConditionNotExpression()

    Do While pConsumeKeyword("AND")
        Dim right_node As ColumnExpressionNode
        Set right_node = pParseConditionNotExpression()
        Set left_node = pNewBinaryNode(C_NODE_KIND_AND, "AND", left_node, right_node)
    Loop

    Set pParseConditionAndExpression = left_node
End Function

Private Function pParseConditionNotExpression() As ColumnExpressionNode
    If pConsumeKeyword("NOT") Then
        Set pParseConditionNotExpression = pNewUnaryNode(C_NODE_KIND_NOT, "NOT", pParseConditionNotExpression())
        Exit Function
    End If

    Set pParseConditionNotExpression = pParseConditionPrimary()
End Function

Private Function pParseConditionPrimary() As ColumnExpressionNode
    If pConsumeChar("(") Then
        Set pParseConditionPrimary = pParseConditionOrExpression()
        Call pRequireChar(")")
        Exit Function
    End If

    Set pParseConditionPrimary = pParseConditionComparison()
End Function

Private Function pParseConditionComparison() As ColumnExpressionNode
    Dim left_node As ColumnExpressionNode
    Set left_node = pParseValueOperandNode()

    Dim operator_text As String
    operator_text = pParseComparisonOperator()

    Dim right_node As ColumnExpressionNode
    Set right_node = pParseValueOperandNode()

    Set pParseConditionComparison = pNewBinaryNode(C_NODE_KIND_COMPARISON, operator_text, left_node, right_node)
End Function

Private Function pParseValueOperandNode() As ColumnExpressionNode
    Call pSkipSpaces

    Select Case pPeekChar()
        Case "["
            Set pParseValueOperandNode = pParseColumnReferenceNode()
        Case Chr$(34)
            Set pParseValueOperandNode = pParseStringLiteralNode()
        Case Else
            Call pRaiseSyntaxError("列参照または文字列リテラルが見つかりません。")
    End Select
End Function

Private Function pParseColumnReferenceNode() As ColumnExpressionNode
    Call pSkipSpaces
    If Not pConsumeChar("[") Then
        Call pRaiseSyntaxError("列参照の開始括弧が見つかりません。")
    End If

    Dim column_name As String
    column_name = ""

    Do While pPosition <= Len(pExpression)
        Dim current_char As String
        current_char = Mid$(pExpression, pPosition, 1)

        If current_char = "]" Then
            pPosition = pPosition + 1
            Call pAddReferencedColumnName(column_name)
            Set pParseColumnReferenceNode = pNewTextNode(C_NODE_KIND_COLUMN_REFERENCE, column_name)
            Exit Function
        End If

        column_name = column_name & current_char
        pPosition = pPosition + 1
    Loop

    Call pRaiseSyntaxError("列参照の閉じ括弧が見つかりません。")
End Function

Private Function pParseStringLiteralNode() As ColumnExpressionNode
    Call pSkipSpaces
    If Not pConsumeChar(Chr$(34)) Then
        Call pRaiseSyntaxError("文字列リテラルはダブルクオートで囲んでください。")
    End If

    Dim raw_value As String
    raw_value = ""

    Do While pPosition <= Len(pExpression)
        Dim current_char As String
        current_char = Mid$(pExpression, pPosition, 1)

        If current_char = Chr$(34) Then
            pPosition = pPosition + 1
            Set pParseStringLiteralNode = pNewTextNode(C_NODE_KIND_STRING_LITERAL, pJsonUnescape(raw_value))
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

Private Function pParseComparisonOperator() As String
    If pConsumeText("==") Then
        pParseComparisonOperator = "=="
    ElseIf pConsumeText("!=") Then
        pParseComparisonOperator = "!="
    Else
        Call pRaiseSyntaxError("比較演算子が見つかりません。")
    End If
End Function

Private Function pReadColumnValue( _
        ByVal ColumnName As String, _
        ByVal ColumnValues As ObjectDictionary) As String

    If Not ColumnValues.Exists(ColumnName) Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:=pExpressionLabel() & " が未定義の列を参照しています。(" & ColumnName & ")"
    End If

    pReadColumnValue = CStr(ColumnValues.Item(ColumnName))
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

Private Sub pAddReferencedColumnName(ByVal ColumnName As String)
    If pReferencedColumnNameSet.Exists(ColumnName) Then Exit Sub

    Call pReferencedColumnNameSet.Add(ColumnName)
    Call pReferencedColumnNames.Add(ColumnName)
End Sub

Private Function pNewTextNode( _
        ByVal Kind As String, _
        ByVal TextValue As String) As ColumnExpressionNode

    Dim result_value As ColumnExpressionNode
    Set result_value = New ColumnExpressionNode
    result_value.Kind = Kind
    result_value.TextValue = TextValue

    Set pNewTextNode = result_value
End Function

Private Function pNewUnaryNode( _
        ByVal Kind As String, _
        ByVal OperatorText As String, _
        ByVal ChildNode As ColumnExpressionNode) As ColumnExpressionNode

    Dim result_value As ColumnExpressionNode
    Set result_value = New ColumnExpressionNode
    result_value.Kind = Kind
    result_value.OperatorText = OperatorText
    Set result_value.LeftNode = ChildNode

    Set pNewUnaryNode = result_value
End Function

Private Function pNewBinaryNode( _
        ByVal Kind As String, _
        ByVal OperatorText As String, _
        ByVal LeftNode As ColumnExpressionNode, _
        ByVal RightNode As ColumnExpressionNode) As ColumnExpressionNode

    Dim result_value As ColumnExpressionNode
    Set result_value = New ColumnExpressionNode
    result_value.Kind = Kind
    result_value.OperatorText = OperatorText
    Set result_value.LeftNode = LeftNode
    Set result_value.RightNode = RightNode

    Set pNewBinaryNode = result_value
End Function

Private Function pNewTernaryNode( _
        ByVal Kind As String, _
        ByVal ConditionNode As ColumnExpressionNode, _
        ByVal TrueNode As ColumnExpressionNode, _
        ByVal FalseNode As ColumnExpressionNode) As ColumnExpressionNode

    Dim result_value As ColumnExpressionNode
    Set result_value = New ColumnExpressionNode
    result_value.Kind = Kind
    Set result_value.LeftNode = ConditionNode
    Set result_value.MiddleNode = TrueNode
    Set result_value.RightNode = FalseNode

    Set pNewTernaryNode = result_value
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

Private Function pExpressionLabel() As String
    If pSourceName <> "" Then
        pExpressionLabel = pSourceName
    ElseIf pExpressionKind = G_COLUMN_EXPRESSION_KIND_CONDITION Then
        pExpressionLabel = "OutputConditionExpression"
    ElseIf pExpressionKind = G_COLUMN_EXPRESSION_KIND_VALUE Then
        pExpressionLabel = "ValueExpression"
    Else
        pExpressionLabel = "ColumnExpression"
    End If
End Function

Private Sub pRaiseSyntaxError(ByVal Message As String)
    Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:=pExpressionLabel() & " の構文が不正です。" & Message & " (位置: " & CStr(pPosition) & ")"
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="初期化されていません。"
    End If
End Sub
