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
Private Const C_STATE_VISITING As Long = 1
Private Const C_STATE_RESOLVED As Long = 2

Private pDefinitions As ObjectList
Private pExtractionDefinitions As ObjectList
Private pValueExpressions() As ColumnExpression
Private pResolvedValues As ObjectDictionary
Private pResolveState() As Long
Private pIsInitialized As Boolean

'* WebDriver から直接抽出する列定義。
Public Property Get ExtractionColumnDefinitions() As ObjectList
    Call pCheckInit
    Set ExtractionColumnDefinitions = pExtractionDefinitions
End Property

'* 値解決を初期化します。
'*
'* @param DetailColumnDefinitions 詳細ページ列定義の一覧。
'*
'* @details
'* 詳細ページ列定義を保持し、WebDriver から直接抽出が必要な列と派生式で解決する列を分けて扱えるようにします。
Public Sub Initialize(ByVal DetailColumnDefinitions As ObjectList)
    If DetailColumnDefinitions Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="DetailColumnDefinitions は必須です。"
    End If

    Set pDefinitions = DetailColumnDefinitions
    Set pExtractionDefinitions = New_ObjectList("DetailColumnDefinition")
    Set pResolvedValues = Nothing

    Call pValidateUniqueOutputColumnNames
    Call pInitializeValueExpressions

    Dim def_idx As Long
    For def_idx = 0 To pDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(def_idx)
        If Not detail_def.IsDerived Then Call pExtractionDefinitions.Add(detail_def)
    Next def_idx

    Call pValidateDerivedExpressions

    pIsInitialized = True
End Sub

Private Sub pValidateUniqueOutputColumnNames()
    Dim known_names As ObjectSet
    Set known_names = New ObjectSet

    Dim def_idx As Long
    For def_idx = 0 To pDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(def_idx)

        If known_names.Exists(detail_def.OutputColumnName) Then
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="OutputColumnName が重複しています。(" & detail_def.OutputColumnName & ")"
        End If
        Call known_names.Add(detail_def.OutputColumnName)
    Next def_idx
End Sub

Private Sub pInitializeValueExpressions()
    If pDefinitions.Count = 0 Then Exit Sub

    ReDim pValueExpressions(0 To pDefinitions.Count - 1)

    Dim def_idx As Long
    For def_idx = 0 To pDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(def_idx)
        If detail_def.IsDerived Then
            Set pValueExpressions(def_idx) = New_ColumnExpression( _
                    Expression:=detail_def.ValueExpression, _
                    ExpressionKind:=G_COLUMN_EXPRESSION_KIND_VALUE, _
                    SourceName:="ValueExpression(" & detail_def.OutputColumnName & ")")
        End If
    Next def_idx
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

'* 指定列の解決に必要な列定義を、元の settings 行順で返します。
'*
'* @param TargetDefinitions 解決対象の詳細ページ列定義。
'* @param OptionalExtraction True の場合、直接抽出列を必須扱いにしません。
'* @param IgnoreDerivedBlankMode True の場合、派生列の空欄検証設定を無視します。
'* @return 解決に必要な詳細ページ列定義。
'*
'* @details
'* 派生列は他の列を参照できるため、対象列だけでなく依存する直接抽出列と派生列を settings 上の順序で集めます。
Public Function BuildDependentDefinitions( _
        ByVal TargetDefinitions As ObjectList, _
        Optional ByVal OptionalExtraction As Boolean = False, _
        Optional ByVal IgnoreDerivedBlankMode As Boolean = False) As ObjectList

    Call pCheckInit
    If TargetDefinitions Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="TargetDefinitions は必須です。"
    End If

    Set BuildDependentDefinitions = BuildDependentDefinitionsByNames( _
            pDefinitionNames(TargetDefinitions), _
            OptionalExtraction:=OptionalExtraction, _
            IgnoreDerivedBlankMode:=IgnoreDerivedBlankMode)
End Function

Private Function pDefinitionNames(ByVal DetailDefinitions As ObjectList) As ObjectList
    Dim result_value As ObjectList
    Set result_value = New ObjectList

    Dim def_idx As Long
    For def_idx = 0 To DetailDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = DetailDefinitions.Item(def_idx)
        Call result_value.Add(detail_def.OutputColumnName)
    Next def_idx

    Set pDefinitionNames = result_value
End Function

'* 指定列名の解決に必要な列定義を、元の settings 行順で返します。
'*
'* @param TargetColumnNames 解決対象の OutputColumnName 一覧。
'* @param OptionalExtraction True の場合、直接抽出列を必須扱いにしません。
'* @param IgnoreDerivedBlankMode True の場合、派生列の空欄検証設定を無視します。
'* @return 解決に必要な詳細ページ列定義。
'*
'* @details
'* 指定列名の解決に必要な列定義を、元の settings 行順で返します。
Public Function BuildDependentDefinitionsByNames( _
        ByVal TargetColumnNames As ObjectList, _
        Optional ByVal OptionalExtraction As Boolean = False, _
        Optional ByVal IgnoreDerivedBlankMode As Boolean = False) As ObjectList

    Call pCheckInit
    If TargetColumnNames Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="TargetColumnNames は必須です。"
    End If

    Dim needed() As Boolean
    If 0 < pDefinitions.Count Then ReDim needed(0 To pDefinitions.Count - 1)

    Dim target_idx As Long
    For target_idx = 0 To TargetColumnNames.Count - 1
        Dim target_name As String
        target_name = CStr(TargetColumnNames.Item(target_idx))

        Dim def_idx As Long
        def_idx = pFindDefinitionIndex(target_name)
        If def_idx < 0 Then
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="依存列が未定義の OutputColumnName を参照しています。(" & target_name & ")"
        End If
        Call pMarkDependentDefinition(def_idx, needed)
    Next target_idx

    Dim result_value As ObjectList
    Set result_value = New_ObjectList("DetailColumnDefinition")

    Dim result_idx As Long
    For result_idx = 0 To pDefinitions.Count - 1
        If needed(result_idx) Then
            Dim src_def As DetailColumnDefinition
            Set src_def = pDefinitions.Item(result_idx)
            Call result_value.Add(pCopyDefinition(src_def, OptionalExtraction, IgnoreDerivedBlankMode))
        End If
    Next result_idx

    Set BuildDependentDefinitionsByNames = result_value
End Function

Private Sub pMarkDependentDefinition( _
        ByVal DefinitionIndex As Long, _
        ByRef Needed() As Boolean)

    If Needed(DefinitionIndex) Then Exit Sub
    Needed(DefinitionIndex) = True

    Dim detail_def As DetailColumnDefinition
    Set detail_def = pDefinitions.Item(DefinitionIndex)
    If Not detail_def.IsDerived Then Exit Sub

    Dim references As ObjectList
    Set references = pCollectReferences(DefinitionIndex)

    Dim ref_idx_in_list As Long
    For ref_idx_in_list = 0 To references.Count - 1
        Dim ref_name As String
        ref_name = CStr(references.Item(ref_idx_in_list))

        Dim ref_idx As Long
        ref_idx = pFindDefinitionIndex(ref_name)
        If ref_idx < 0 Then
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="依存列が未定義の OutputColumnName を参照しています。(" & ref_name & ")"
        End If
        Call pMarkDependentDefinition(ref_idx, Needed)
    Next ref_idx_in_list
End Sub

Private Function pCopyDefinition( _
        ByVal DetailDefinition As DetailColumnDefinition, _
        ByVal OptionalExtraction As Boolean, _
        ByVal IgnoreDerivedBlankMode As Boolean) As DetailColumnDefinition

    Dim blank_mode As String
    blank_mode = DetailDefinition.BlankMode

    If DetailDefinition.IsDerived Then
        If IgnoreDerivedBlankMode Then blank_mode = "AllowBlank"
        Set pCopyDefinition = New_DetailColumnDefinition( _
                DetailDefinition.OutputColumnName, _
                "", _
                ValueExpression:=DetailDefinition.ValueExpression, _
                OutputValueType:=DetailDefinition.OutputValueType, _
                BlankMode:=blank_mode)
        Exit Function
    End If

    Dim is_required As Boolean
    is_required = DetailDefinition.IsRequired
    If OptionalExtraction Then
        is_required = False
        blank_mode = "AllowBlank"
    End If

    Set pCopyDefinition = New_DetailColumnDefinition( _
            DetailDefinition.OutputColumnName, _
            DetailDefinition.Selector, _
            ExtractType:=DetailDefinition.ExtractType, _
            AttributeName:=DetailDefinition.AttributeName, _
            IsRequired:=is_required, _
            OutputValueType:=DetailDefinition.OutputValueType, _
            BlankMode:=blank_mode)
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

'* 通常抽出値から、列名で参照できる詳細値を解決します。
'*
'* @param ExtractionValues ExtractionColumnDefinitions と同じ順序の抽出値。
'* @param ValidateDerivedBlankMode True の場合、派生列の空欄検証設定を適用します。
'* @return OutputColumnName から値を引ける辞書。
'*
'* @details
'* 通常抽出値から、列名で参照できる詳細値を解決します。
Public Function ResolveValues( _
        ByVal ExtractionValues As ArrayObject, _
        Optional ByVal ValidateDerivedBlankMode As Boolean = True) As ObjectDictionary

    Call pCheckInit
    If ExtractionValues Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ExtractionValues は必須です。"
    End If

    Set pResolvedValues = New_ObjectDictionary()
    If 0 < pDefinitions.Count Then ReDim pResolveState(0 To pDefinitions.Count - 1)

    Dim extraction_idx As Long
    extraction_idx = 0

    Dim def_idx As Long
    For def_idx = 0 To pDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(def_idx)
        If Not detail_def.IsDerived Then
            Call pResolvedValues.Add(detail_def.OutputColumnName, CStr(ExtractionValues.Item(extraction_idx)))
            pResolveState(def_idx) = C_STATE_RESOLVED
            extraction_idx = extraction_idx + 1
        End If
    Next def_idx

    For def_idx = 0 To pDefinitions.Count - 1
        Set detail_def = pDefinitions.Item(def_idx)
        If detail_def.IsDerived Then
            Dim derived_value As String
            derived_value = pResolveDefinitionValue(def_idx)
            If ValidateDerivedBlankMode Then Call pValidateDerivedBlankMode(detail_def, derived_value)
        End If
    Next def_idx

    Set ResolveValues = pResolvedValues
    Set pResolvedValues = Nothing
End Function

Private Sub pValidateDerivedBlankMode( _
        ByVal DetailDefinition As DetailColumnDefinition, _
        ByVal DerivedValue As String)

    If StrComp(DetailDefinition.BlankMode, "ErrorIfBlank", vbTextCompare) <> 0 Then Exit Sub
    If DerivedValue <> "" Then Exit Sub

    Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="派生列の値が空欄です。(" & DetailDefinition.OutputColumnName & ")"
End Sub

Private Function pResolveDefinitionValue(ByVal DefinitionIndex As Long) As String
    Dim detail_def As DetailColumnDefinition
    Set detail_def = pDefinitions.Item(DefinitionIndex)

    If pResolveState(DefinitionIndex) = C_STATE_RESOLVED Then
        pResolveDefinitionValue = CStr(pResolvedValues.Item(detail_def.OutputColumnName))
        Exit Function
    End If
    If pResolveState(DefinitionIndex) = C_STATE_VISITING Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="派生列同士の循環参照があります。(" & detail_def.OutputColumnName & ")"
    End If

    pResolveState(DefinitionIndex) = C_STATE_VISITING
    Call pResolveReferencedValues(DefinitionIndex)

    Dim result_value As String
    result_value = pValueExpressions(DefinitionIndex).EvaluateValue(pResolvedValues)

    Call pResolvedValues.Add(detail_def.OutputColumnName, result_value)
    pResolveState(DefinitionIndex) = C_STATE_RESOLVED

    pResolveDefinitionValue = result_value
End Function

'* 列名辞書を元の詳細ページ列定義順の値配列に戻します。
'*
'* @param ValuesByName OutputColumnName から値を引ける辞書。
'* @return 元の詳細ページ列定義と同じ順序の値。
'*
'* @details
'* 列名辞書を元の詳細ページ列定義順の値配列に戻します。
Public Function BuildOutputValues(ByVal ValuesByName As ObjectDictionary) As ArrayObject
    Call pCheckInit
    If ValuesByName Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValuesByName は必須です。"
    End If

    Dim result_values As ArrayObject
    Set result_values = New ArrayObject
    If 0 < pDefinitions.Count Then Call result_values.ReDimArray(0, pDefinitions.Count - 1)

    Dim def_idx As Long
    For def_idx = 0 To pDefinitions.Count - 1
        Dim detail_def As DetailColumnDefinition
        Set detail_def = pDefinitions.Item(def_idx)
        If Not ValuesByName.Exists(detail_def.OutputColumnName) Then
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValuesByName に OutputColumnName がありません。(" & detail_def.OutputColumnName & ")"
        End If
        Call result_values.Update(def_idx, CStr(ValuesByName.Item(detail_def.OutputColumnName)))
    Next def_idx

    Set BuildOutputValues = result_values
End Function

Private Sub pResolveReferencedValues(ByVal DefinitionIndex As Long)
    Dim references As ObjectList
    Set references = pCollectReferences(DefinitionIndex)

    Dim ref_idx_in_list As Long
    For ref_idx_in_list = 0 To references.Count - 1
        Dim ref_name As String
        ref_name = CStr(references.Item(ref_idx_in_list))

        Dim ref_idx As Long
        ref_idx = pFindDefinitionIndex(ref_name)
        If ref_idx < 0 Then
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValueExpression が未定義の OutputColumnName を参照しています。(" & ref_name & ")"
        End If

        Dim ref_def As DetailColumnDefinition
        Set ref_def = pDefinitions.Item(ref_idx)
        If ref_def.IsDerived Then Call pResolveDefinitionValue(ref_idx)
    Next ref_idx_in_list
End Sub

Private Sub pValidateExpressionReferences(ByVal DefinitionIndex As Long)
    Dim references As ObjectList
    Set references = pCollectReferences(DefinitionIndex)

    Dim ref_idx_in_list As Long
    For ref_idx_in_list = 0 To references.Count - 1
        Dim ref_name As String
        ref_name = CStr(references.Item(ref_idx_in_list))

        Dim ref_idx As Long
        ref_idx = pFindDefinitionIndex(ref_name)
        If ref_idx < 0 Then
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValueExpression が未定義の OutputColumnName を参照しています。(" & ref_name & ")"
        End If
        If ref_idx = DefinitionIndex Then
            Dim detail_def As DetailColumnDefinition
            Set detail_def = pDefinitions.Item(DefinitionIndex)
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValueExpression が自己参照しています。(" & detail_def.OutputColumnName & ")"
        End If
    Next ref_idx_in_list
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

    Dim references As ObjectList
    Set references = pCollectReferences(DefinitionIndex)

    Dim ref_idx_in_list As Long
    For ref_idx_in_list = 0 To references.Count - 1
        Dim ref_name As String
        ref_name = CStr(references.Item(ref_idx_in_list))

        Dim ref_idx As Long
        ref_idx = pFindDefinitionIndex(ref_name)
        If ref_idx < 0 Then
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValueExpression が未定義の OutputColumnName を参照しています。(" & ref_name & ")"
        End If
        If ref_idx = DefinitionIndex Then
            Dim self_def As DetailColumnDefinition
            Set self_def = pDefinitions.Item(DefinitionIndex)
            Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="ValueExpression が自己参照しています。(" & self_def.OutputColumnName & ")"
        End If

        Dim ref_def As DetailColumnDefinition
        Set ref_def = pDefinitions.Item(ref_idx)
        If ref_def.IsDerived Then Call pVisitDerivedDefinition(ref_idx, VisitState)
    Next ref_idx_in_list

    VisitState(DefinitionIndex) = C_STATE_RESOLVED
End Sub

Private Function pCollectReferences(ByVal DefinitionIndex As Long) As ObjectList
    Set pCollectReferences = pValueExpressions(DefinitionIndex).ReferencedColumnNames
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:=C_SOURCE, Description:="初期化されていません。"
    End If
End Sub
