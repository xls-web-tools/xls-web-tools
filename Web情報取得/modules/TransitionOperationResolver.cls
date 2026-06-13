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

'* 一覧画面へ進む画面遷移操作を解決します。
'*
'* @details
'* 一覧画面へ進む画面遷移操作を解決します。
Public Function ResolveListTransitionOperation() As TransitionOperation
    Call pCheckInit
    Set ResolveListTransitionOperation = pFindTransitionOperation(pSettings.ListTransitionOperationName)
End Function

'* 詳細ページへ進む画面遷移操作を解決します。
'*
'* @param ItemIndex 0-based の一覧項目 index。
'*
'* @details
'* ActionSelector、ActionScript、WaitSelector には {{index}} / {{rowNumber}} の置換を適用します。
Public Function ResolveDetailTransitionOperation(ByVal ItemIndex As Long) As TransitionOperation
    Call pCheckInit
    Set ResolveDetailTransitionOperation = pResolveIndexedTransitionOperation(pSettings.DetailTransitionOperationName, ItemIndex)
End Function

Private Function pResolveIndexedTransitionOperation( _
        ByVal OperationName As String, _
        ByVal ItemIndex As Long) As TransitionOperation

    Dim base_operation As TransitionOperation
    Set base_operation = pFindTransitionOperation(OperationName)

    Set pResolveIndexedTransitionOperation = New_TransitionOperation( _
            base_operation.OperationName, _
            pIndexTemplate.Apply(base_operation.ActionSelector, ItemIndex), _
            ActionInnerText:=base_operation.ActionInnerText, _
            ActionScript:=pIndexTemplate.Apply(base_operation.ActionScript, ItemIndex), _
            WaitSelector:=pIndexTemplate.Apply(base_operation.WaitSelector, ItemIndex))
End Function

'* 一覧画面へ戻る画面遷移操作を解決します。
'*
'* @details
'* 一覧画面へ戻る画面遷移操作を解決します。
Public Function ResolveReturnToListOperation() As TransitionOperation
    Call pCheckInit
    Set ResolveReturnToListOperation = pFindTransitionOperation(pSettings.ReturnToListOperationName)
End Function

'* 次ページへ進む画面遷移操作を解決します。
'*
'* @details
'* 次ページへ進む画面遷移操作を解決します。
Public Function ResolveNextPageOperation() As TransitionOperation
    Call pCheckInit
    Set ResolveNextPageOperation = pFindTransitionOperation(pSettings.NextPageOperationName)
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

    Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperationResolver.pFindTransitionOperation", Description:="TransitionOperations に OperationName が見つかりません。(" & OperationName & ")"
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperationResolver.pCheckInit", Description:="初期化されていません。"
    End If
End Sub
