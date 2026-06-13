VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransitionOperationResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! settings の画面遷移操作名から実行する画面遷移操作を解決します。
'!
' #############################################################################

Private Const C_CHAIN_DELIMITER As String = ">>"

Private pSettings As IToolSettings
Private pIndexTemplate As ListItemIndexTemplate
Private pIsInitialized As Boolean

'* 画面遷移操作 resolver を初期化します。
'*
'* @param Settings Web情報取得の設定。
'*
'* @details
'* OperationName の照合と一覧項目 index 置換の規則をこの Module に閉じます。
Public Sub Initialize(ByVal Settings As IToolSettings)
    If Settings Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperationResolver.Initialize", Description:="Settings は必須です。"
    End If

    Set pSettings = Settings
    Set pIndexTemplate = New_ListItemIndexTemplate()
    pIsInitialized = True
End Sub

'* 一覧画面へ進む画面遷移操作チェーンを解決します。
'*
'* @details
'* 一覧画面へ進む画面遷移操作チェーンを解決します。
Public Function ResolveListTransitionOperation() As TransitionOperationChain
    Call pCheckInit
    Set ResolveListTransitionOperation = pResolveTransitionOperationChain(pSettings.ListTransitionOperationName, "ListTransitionOperationName")
End Function

'* 詳細ページへ進む画面遷移操作チェーンを解決します。
'*
'* @param ItemIndex 0-based の一覧項目 index。
'*
'* @details
'* ActionSelector、ActionScript、WaitSelector には {{index}} / {{rowNumber}} の置換を適用します。
Public Function ResolveDetailTransitionOperation(ByVal ItemIndex As Long) As TransitionOperationChain
    Call pCheckInit
    Set ResolveDetailTransitionOperation = pResolveIndexedTransitionOperationChain(pSettings.DetailTransitionOperationName, "DetailTransitionOperationName", ItemIndex)
End Function

Private Function pResolveIndexedTransitionOperationChain( _
        ByVal OperationChainText As String, _
        ByVal SettingItemName As String, _
        ByVal ItemIndex As Long) As TransitionOperationChain

    Dim base_chain As TransitionOperationChain
    Set base_chain = pResolveTransitionOperationChain(OperationChainText, SettingItemName)

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")

    Dim operation_idx As Long
    For operation_idx = 0 To base_chain.Count - 1
        Dim base_operation As TransitionOperation
        Set base_operation = base_chain.Item(operation_idx)

        Call operations.Add(New_TransitionOperation( _
                base_operation.OperationName, _
                pIndexTemplate.Apply(base_operation.ActionSelector, ItemIndex), _
                ActionInnerText:=base_operation.ActionInnerText, _
                ActionScript:=pIndexTemplate.Apply(base_operation.ActionScript, ItemIndex), _
                WaitSelector:=pIndexTemplate.Apply(base_operation.WaitSelector, ItemIndex)))
    Next operation_idx

    Set pResolveIndexedTransitionOperationChain = New_TransitionOperationChain(operations)
End Function

'* 一覧画面へ戻る画面遷移操作チェーンを解決します。
'*
'* @details
'* 一覧画面へ戻る画面遷移操作チェーンを解決します。
Public Function ResolveReturnToListOperation() As TransitionOperationChain
    Call pCheckInit
    Set ResolveReturnToListOperation = pResolveTransitionOperationChain(pSettings.ReturnToListOperationName, "ReturnToListOperationName")
End Function

'* 次ページへ進む画面遷移操作チェーンを解決します。
'*
'* @details
'* 次ページへ進む画面遷移操作チェーンを解決します。
Public Function ResolveNextPageOperation() As TransitionOperationChain
    Call pCheckInit
    Set ResolveNextPageOperation = pResolveTransitionOperationChain(pSettings.NextPageOperationName, "NextPageOperationName")
End Function

Private Function pResolveTransitionOperationChain( _
        ByVal OperationChainText As String, _
        ByVal SettingItemName As String) As TransitionOperationChain

    If pContainsChainDelimiter(OperationChainText) And pHasTransitionOperationName(OperationChainText) Then
        Call pRaiseChainResolutionError( _
                SettingItemName, _
                OperationChainText, _
                OperationChainText, _
                "OperationName に >> は使用できません。")
    End If

    Dim operations As ObjectList
    Set operations = New_ObjectList("TransitionOperation")

    Dim operation_names As Variant
    operation_names = Split(OperationChainText, C_CHAIN_DELIMITER)

    Dim operation_idx As Long
    For operation_idx = LBound(operation_names) To UBound(operation_names)
        Dim element_name As String
        element_name = Trim$(CStr(operation_names(operation_idx)))
        If element_name = "" Then
            Call pRaiseChainResolutionError( _
                    SettingItemName, _
                    OperationChainText, _
                    element_name, _
                    "画面遷移操作チェーンに空の要素があります。")
        End If

        Dim target_operation As TransitionOperation
        Set target_operation = pFindTransitionOperation(element_name)
        If target_operation Is Nothing Then
            Call pRaiseChainResolutionError( _
                    SettingItemName, _
                    OperationChainText, _
                    element_name, _
                    "TransitionOperations に OperationName が見つかりません。")
        End If

        Call operations.Add(target_operation)
    Next operation_idx

    Set pResolveTransitionOperationChain = New_TransitionOperationChain(operations)
End Function

Private Function pContainsChainDelimiter(ByVal OperationName As String) As Boolean
    pContainsChainDelimiter = (0 < InStr(1, OperationName, C_CHAIN_DELIMITER, vbBinaryCompare))
End Function

Private Function pHasTransitionOperationName(ByVal OperationName As String) As Boolean
    Dim operations As ObjectList
    Set operations = pSettings.TransitionOperations

    Dim operation_idx As Long
    For operation_idx = 0 To operations.Count - 1
        Dim target_operation As TransitionOperation
        Set target_operation = operations.Item(operation_idx)
        If StrComp(target_operation.OperationName, OperationName, vbTextCompare) = 0 Then
            pHasTransitionOperationName = True
            Exit Function
        End If
    Next operation_idx

    pHasTransitionOperationName = False
End Function

Private Function pFindTransitionOperation(ByVal OperationName As String) As TransitionOperation
    Dim operations As ObjectList
    Set operations = pSettings.TransitionOperations

    Dim operation_idx As Long
    For operation_idx = 0 To operations.Count - 1
        Dim target_operation As TransitionOperation
        Set target_operation = operations.Item(operation_idx)
        If StrComp(target_operation.OperationName, OperationName, vbTextCompare) = 0 Then
            Set pFindTransitionOperation = target_operation
            Exit Function
        End If
    Next operation_idx

    Set pFindTransitionOperation = Nothing
End Function

Private Sub pRaiseChainResolutionError( _
        ByVal SettingItemName As String, _
        ByVal OperationChainText As String, _
        ByVal ElementName As String, _
        ByVal Message As String)

    Err.Raise Number:=vbObjectError + 1, _
            Source:="Class TransitionOperationResolver.pResolveTransitionOperationChain", _
            Description:=Message _
                    & " SettingItemName=" & SettingItemName _
                    & ", OperationChain=" & OperationChainText _
                    & ", ElementName=" & ElementName
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperationResolver.pCheckInit", Description:="初期化されていません。"
    End If
End Sub
