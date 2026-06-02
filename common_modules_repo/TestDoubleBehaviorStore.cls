VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDoubleBehaviorStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! テストダブルのスタブ設定と呼び出し履歴を管理するクラスです。
'!
'! @details
'! RecordCall は、各 TestDouble で正常に成立した呼び出しだけを記録する用途です。strict stub の未登録エラーや RaiseIfError による明示エラーは呼び出し履歴に自動記録されません。
'! 照合引数からキーを生成する操作は対象メソッドの ObjectKeyMode を固定します。HasReturn、HasOutput、RaiseIfError、GetCallCount、GetLatestCall などの読み取り操作でも固定されます。
'!
' #############################################################################

Private pKeyBuilder As TestDoubleVariantKeyBuilder
Private pReturns As Dictionary
Private pOutputs As Dictionary
Private pErrors As Dictionary
Private pCalls As Collection
Private pCallsByKey As Dictionary
Private pLatestByKey As Dictionary
Private pDefaultObjectKeyMode As Long
Private pObjectKeyModesByMethod As Dictionary
Private pLockedMethods As Dictionary

'* 記録済みの呼び出し件数。
Public Property Get Count() As Long
    Count = pCalls.Count
End Property

'* メソッド別の指定がない場合に使う既定のオブジェクト キー化モード。
Public Property Get DefaultObjectKeyMode() As Long
    DefaultObjectKeyMode = pDefaultObjectKeyMode
End Property
Public Property Let DefaultObjectKeyMode(ByVal NewValue As Long)
    Call pValidateObjectKeyMode(NewValue)
    If 0 < pLockedMethods.Count Then
        Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "キー生成に使用済みのため、既定 ObjectKeyMode は変更できません。"
    End If

    pDefaultObjectKeyMode = NewValue
End Property

Private Sub Class_Initialize()
    Set pKeyBuilder = New TestDoubleVariantKeyBuilder
    Set pReturns = New Dictionary
    Set pOutputs = New Dictionary
    Set pErrors = New Dictionary
    Set pCalls = New Collection
    Set pCallsByKey = New Dictionary
    Set pLatestByKey = New Dictionary
    Set pObjectKeyModesByMethod = New Dictionary
    Set pLockedMethods = New Dictionary

    pObjectKeyModesByMethod.CompareMode = TextCompare
    pLockedMethods.CompareMode = TextCompare
    pDefaultObjectKeyMode = G_OBJECT_KEY_MODE_REFERENCE
End Sub

'* メソッド別のオブジェクト キー化モードを設定します。
Public Sub SetObjectKeyMode(ByVal MethodName As String, ByVal ObjectKeyMode As Long)
    Call pValidateMethodName(MethodName)
    Call pValidateObjectKeyMode(ObjectKeyMode)

    Dim normalized_method_name As String
    normalized_method_name = pNormalizeMethodName(MethodName)

    If pLockedMethods.Exists(normalized_method_name) Then
        Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "キー生成に使用済みのメソッドの ObjectKeyMode は変更できません。(" & MethodName & ")"
    End If

    pObjectKeyModesByMethod.Item(normalized_method_name) = ObjectKeyMode
End Sub

Private Sub pValidateObjectKeyMode(ByVal ObjectKeyMode As Long)
    Select Case ObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE, G_OBJECT_KEY_MODE_I_EQUATABLE, G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            Exit Sub
        Case Else
            Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "未対応のオブジェクト キー化モードです。(" & CStr(ObjectKeyMode) & ")"
    End Select
End Sub

'* 戻り値を設定します。
Public Sub SetReturn(ByVal MethodName As String, ByRef ReturnValue As Variant, ParamArray KeyArgs() As Variant)
    Call SetReturnFromArray(MethodName, ReturnValue, KeyArgs)
End Sub

'* ParamArray と同じ形式の配列から戻り値を設定します。
Public Sub SetReturnFromArray(ByVal MethodName As String, ByRef ReturnValue As Variant, ByVal KeyArgs As Variant)
    Call pSetStoredValue(pReturns, pBuildMethodKey(MethodName, KeyArgs), ReturnValue)
End Sub

'* 戻り値が設定されているかを返します。
Public Function HasReturn(ByVal MethodName As String, ParamArray KeyArgs() As Variant) As Boolean
    HasReturn = HasReturnFromArray(MethodName, KeyArgs)
End Function

'* ParamArray と同じ形式の配列から戻り値の設定有無を返します。
Public Function HasReturnFromArray(ByVal MethodName As String, ByVal KeyArgs As Variant) As Boolean
    HasReturnFromArray = pReturns.Exists(pBuildMethodKey(MethodName, KeyArgs))
End Function

'* 戻り値を返します。
Public Function GetReturn(ByVal MethodName As String, ParamArray KeyArgs() As Variant) As Variant
    Dim key_args_arr As Variant
    key_args_arr = KeyArgs

    If IsObject(GetReturnFromArray(MethodName, key_args_arr)) Then
        Set GetReturn = GetReturnFromArray(MethodName, key_args_arr)
    Else
        GetReturn = GetReturnFromArray(MethodName, key_args_arr)
    End If
End Function

'* ParamArray と同じ形式の配列から戻り値を返します。
Public Function GetReturnFromArray(ByVal MethodName As String, ByVal KeyArgs As Variant) As Variant
    Dim store_key As String
    store_key = pBuildMethodKey(MethodName, KeyArgs)

    If IsObject(pGetStoredItem(pReturns, store_key, "戻り値が設定されていません。(" & MethodName & ")")) Then
        Set GetReturnFromArray = pGetStoredItem(pReturns, store_key, "戻り値が設定されていません。(" & MethodName & ")")
    Else
        GetReturnFromArray = pGetStoredItem(pReturns, store_key, "戻り値が設定されていません。(" & MethodName & ")")
    End If
End Function

'* 戻り値が未設定の場合は既定値を返します。既定値を返す場合でもキー生成を行うため、対象メソッドの ObjectKeyMode は固定されます。
Public Function GetReturnOrDefault(ByVal MethodName As String, ByVal DefaultValue As Variant, ParamArray KeyArgs() As Variant) As Variant
    Dim key_args_arr As Variant
    key_args_arr = KeyArgs

    If IsObject(GetReturnOrDefaultFromArray(MethodName, DefaultValue, key_args_arr)) Then
        Set GetReturnOrDefault = GetReturnOrDefaultFromArray(MethodName, DefaultValue, key_args_arr)
    Else
        GetReturnOrDefault = GetReturnOrDefaultFromArray(MethodName, DefaultValue, key_args_arr)
    End If
End Function

'* ParamArray と同じ形式の配列から戻り値または既定値を返します。既定値を返す場合でもキー生成を行うため、対象メソッドの ObjectKeyMode は固定されます。
Public Function GetReturnOrDefaultFromArray(ByVal MethodName As String, ByVal DefaultValue As Variant, ByVal KeyArgs As Variant) As Variant
    If HasReturnFromArray(MethodName, KeyArgs) Then
        If IsObject(GetReturnFromArray(MethodName, KeyArgs)) Then
            Set GetReturnOrDefaultFromArray = GetReturnFromArray(MethodName, KeyArgs)
        Else
            GetReturnOrDefaultFromArray = GetReturnFromArray(MethodName, KeyArgs)
        End If
    Else
        If IsObject(DefaultValue) Then
            Set GetReturnOrDefaultFromArray = DefaultValue
        Else
            GetReturnOrDefaultFromArray = DefaultValue
        End If
    End If
End Function

'* ByRef 出力値を設定します。
Public Sub SetOutput(ByVal MethodName As String, ByVal OutputName As String, ByRef OutputValue As Variant, ParamArray KeyArgs() As Variant)
    Call SetOutputFromArray(MethodName, OutputName, OutputValue, KeyArgs)
End Sub

'* ParamArray と同じ形式の配列から ByRef 出力値を設定します。
Public Sub SetOutputFromArray(ByVal MethodName As String, ByVal OutputName As String, ByRef OutputValue As Variant, ByVal KeyArgs As Variant)
    Call pValidateOutputName(OutputName)
    Call pSetStoredValue(pOutputs, pBuildOutputKey(MethodName, OutputName, KeyArgs), OutputValue)
End Sub

'* ByRef 出力値が設定されているかを返します。
Public Function HasOutput(ByVal MethodName As String, ByVal OutputName As String, ParamArray KeyArgs() As Variant) As Boolean
    HasOutput = HasOutputFromArray(MethodName, OutputName, KeyArgs)
End Function

'* ParamArray と同じ形式の配列から ByRef 出力値の設定有無を返します。
Public Function HasOutputFromArray(ByVal MethodName As String, ByVal OutputName As String, ByVal KeyArgs As Variant) As Boolean
    Call pValidateOutputName(OutputName)
    HasOutputFromArray = pOutputs.Exists(pBuildOutputKey(MethodName, OutputName, KeyArgs))
End Function

'* ByRef 出力値を返します。
Public Function GetOutput(ByVal MethodName As String, ByVal OutputName As String, ParamArray KeyArgs() As Variant) As Variant
    Dim key_args_arr As Variant
    key_args_arr = KeyArgs

    If IsObject(GetOutputFromArray(MethodName, OutputName, key_args_arr)) Then
        Set GetOutput = GetOutputFromArray(MethodName, OutputName, key_args_arr)
    Else
        GetOutput = GetOutputFromArray(MethodName, OutputName, key_args_arr)
    End If
End Function

'* ParamArray と同じ形式の配列から ByRef 出力値を返します。
Public Function GetOutputFromArray(ByVal MethodName As String, ByVal OutputName As String, ByVal KeyArgs As Variant) As Variant
    Call pValidateOutputName(OutputName)

    Dim store_key As String
    store_key = pBuildOutputKey(MethodName, OutputName, KeyArgs)

    If IsObject(pGetStoredItem(pOutputs, store_key, "出力値が設定されていません。(" & MethodName & "." & OutputName & ")")) Then
        Set GetOutputFromArray = pGetStoredItem(pOutputs, store_key, "出力値が設定されていません。(" & MethodName & "." & OutputName & ")")
    Else
        GetOutputFromArray = pGetStoredItem(pOutputs, store_key, "出力値が設定されていません。(" & MethodName & "." & OutputName & ")")
    End If
End Function

'* ByRef 出力値が未設定の場合は既定値を返します。
Public Function GetOutputOrDefault(ByVal MethodName As String, ByVal OutputName As String, ByVal DefaultValue As Variant, ParamArray KeyArgs() As Variant) As Variant
    Dim key_args_arr As Variant
    key_args_arr = KeyArgs

    If IsObject(GetOutputOrDefaultFromArray(MethodName, OutputName, DefaultValue, key_args_arr)) Then
        Set GetOutputOrDefault = GetOutputOrDefaultFromArray(MethodName, OutputName, DefaultValue, key_args_arr)
    Else
        GetOutputOrDefault = GetOutputOrDefaultFromArray(MethodName, OutputName, DefaultValue, key_args_arr)
    End If
End Function

'* ParamArray と同じ形式の配列から ByRef 出力値または既定値を返します。
Public Function GetOutputOrDefaultFromArray(ByVal MethodName As String, ByVal OutputName As String, ByVal DefaultValue As Variant, ByVal KeyArgs As Variant) As Variant
    If HasOutputFromArray(MethodName, OutputName, KeyArgs) Then
        If IsObject(GetOutputFromArray(MethodName, OutputName, KeyArgs)) Then
            Set GetOutputOrDefaultFromArray = GetOutputFromArray(MethodName, OutputName, KeyArgs)
        Else
            GetOutputOrDefaultFromArray = GetOutputFromArray(MethodName, OutputName, KeyArgs)
        End If
    Else
        If IsObject(DefaultValue) Then
            Set GetOutputOrDefaultFromArray = DefaultValue
        Else
            GetOutputOrDefaultFromArray = DefaultValue
        End If
    End If
End Function

'* 指定引数で呼ばれたときに送出するエラーを設定します。
Public Sub SetError(ByVal MethodName As String, ByVal ErrorNumber As Long, ByVal ErrorSource As String, ByVal ErrorDescription As String, ParamArray KeyArgs() As Variant)
    Call SetErrorFromArray(MethodName, ErrorNumber, ErrorSource, ErrorDescription, KeyArgs)
End Sub

'* ParamArray と同じ形式の配列から送出するエラーを設定します。
Public Sub SetErrorFromArray(ByVal MethodName As String, ByVal ErrorNumber As Long, ByVal ErrorSource As String, ByVal ErrorDescription As String, ByVal KeyArgs As Variant)
    Dim err_item As Dictionary
    Set err_item = New Dictionary
    err_item.Item("Number") = ErrorNumber
    err_item.Item("Source") = ErrorSource
    err_item.Item("Description") = ErrorDescription

    Set pErrors.Item(pBuildMethodKey(MethodName, KeyArgs)) = err_item
End Sub

'* エラーが設定されているかを返します。
Public Function HasError(ByVal MethodName As String, ParamArray KeyArgs() As Variant) As Boolean
    HasError = HasErrorFromArray(MethodName, KeyArgs)
End Function

'* ParamArray と同じ形式の配列からエラー設定有無を返します。
Public Function HasErrorFromArray(ByVal MethodName As String, ByVal KeyArgs As Variant) As Boolean
    HasErrorFromArray = pErrors.Exists(pBuildMethodKey(MethodName, KeyArgs))
End Function

'* 設定済みのエラーがあれば送出します。エラーを送出しても呼び出し履歴には記録しません。
Public Sub RaiseIfError(ByVal MethodName As String, ParamArray KeyArgs() As Variant)
    Call RaiseIfErrorFromArray(MethodName, KeyArgs)
End Sub

'* ParamArray と同じ形式の配列から設定済みのエラーがあれば送出します。エラーを送出しても呼び出し履歴には記録しません。
Public Sub RaiseIfErrorFromArray(ByVal MethodName As String, ByVal KeyArgs As Variant)
    Dim store_key As String
    store_key = pBuildMethodKey(MethodName, KeyArgs)

    If Not pErrors.Exists(store_key) Then Exit Sub

    Dim err_item As Dictionary
    Set err_item = pErrors.Item(store_key)
    Err.Raise CLng(err_item.Item("Number")), CStr(err_item.Item("Source")), CStr(err_item.Item("Description"))
End Sub

'* 正常に成立した呼び出しを記録します。エラー呼び出しを記録する仕組みは持ちません。
Public Sub RecordCall(ByVal MethodName As String, ParamArray Args() As Variant)
    Call RecordCallFromArray(MethodName, Args)
End Sub

'* ParamArray と同じ形式の配列から正常に成立した呼び出しを記録します。エラー呼び出しを記録する仕組みは持ちません。
Public Sub RecordCallFromArray(ByVal MethodName As String, ByVal Args As Variant)
    Call RecordCallWithKeyArgsFromArray(MethodName, Args, Args)
End Sub

'* 照合引数と呼び出し引数を分けて、正常に成立した呼び出しを記録します。
Public Sub RecordCallWithKeyArgs(ByVal MethodName As String, ByVal KeyArgs As Variant, ParamArray Args() As Variant)
    Call RecordCallWithKeyArgsFromArray(MethodName, KeyArgs, Args)
End Sub

'* 配列指定で照合引数と呼び出し引数を分けて、正常に成立した呼び出しを記録します。
Public Sub RecordCallWithKeyArgsFromArray(ByVal MethodName As String, ByVal KeyArgs As Variant, ByVal Args As Variant)
    Call pValidateKeyArgsArray(KeyArgs)
    If Not IsArray(Args) Then
        Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "呼び出し引数は配列で指定してください。"
    End If

    Dim object_key_mode As Long
    Dim key_args_part As String
    key_args_part = pBuildKeyArgsPart(MethodName, KeyArgs, object_key_mode)

    Dim record_item As TestDoubleCallRecord
    Set record_item = New TestDoubleCallRecord
    record_item.MethodName = MethodName
    record_item.CallIndex = pCalls.Count
    record_item.ObjectKeyMode = object_key_mode
    Call record_item.StoreArguments(Args)

    Call pCalls.Add(record_item)

    Dim store_key As String
    store_key = pBuildStoreKey(MethodName, object_key_mode, key_args_part)

    Set pLatestByKey.Item(store_key) = record_item

    Dim key_calls As Collection
    If pCallsByKey.Exists(store_key) Then
        Set key_calls = pCallsByKey.Item(store_key)
    Else
        Set key_calls = New Collection
        Call pCallsByKey.Add(store_key, key_calls)
    End If
    Call key_calls.Add(record_item)
End Sub

Private Sub pValidateKeyArgsArray(ByVal KeyArgs As Variant)
    If Not IsArray(KeyArgs) Then
        Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "照合引数は配列で指定してください。"
    End If
End Sub

'* 指定された照合引数で呼び出しが記録されているかを返します。
Public Function HasCall(ByVal MethodName As String, ParamArray KeyArgs() As Variant) As Boolean
    HasCall = HasCallFromArray(MethodName, KeyArgs)
End Function

'* ParamArray と同じ形式の配列から呼び出し有無を返します。
Public Function HasCallFromArray(ByVal MethodName As String, ByVal KeyArgs As Variant) As Boolean
    Dim object_key_mode As Long
    Dim key_args_part As String
    key_args_part = pBuildKeyArgsPart(MethodName, KeyArgs, object_key_mode)

    HasCallFromArray = pCallsByKey.Exists(pBuildStoreKey(MethodName, object_key_mode, key_args_part))
End Function

'* 指定メソッド名の呼び出しが記録されているかを返します。
Public Function HasCallAll(ByVal MethodName As String) As Boolean
    HasCallAll = 0 < GetCallCountAll(MethodName)
End Function

'* 指定された照合引数で呼ばれた回数を返します。
Public Function GetCallCount(ByVal MethodName As String, ParamArray KeyArgs() As Variant) As Long
    GetCallCount = GetCallCountFromArray(MethodName, KeyArgs)
End Function

'* ParamArray と同じ形式の配列から呼び出し回数を返します。
Public Function GetCallCountFromArray(ByVal MethodName As String, ByVal KeyArgs As Variant) As Long
    Dim object_key_mode As Long
    Dim key_args_part As String
    key_args_part = pBuildKeyArgsPart(MethodName, KeyArgs, object_key_mode)

    Dim store_key As String
    store_key = pBuildStoreKey(MethodName, object_key_mode, key_args_part)

    If pCallsByKey.Exists(store_key) Then
        GetCallCountFromArray = pCallsByKey.Item(store_key).Count
    Else
        GetCallCountFromArray = 0
    End If
End Function

'* 指定メソッド名の呼び出し回数を返します。
Public Function GetCallCountAll(ByVal MethodName As String) As Long
    Call pValidateMethodName(MethodName)

    Dim normalized_method_name As String
    normalized_method_name = pNormalizeMethodName(MethodName)

    Dim result_count As Long
    result_count = 0

    Dim call_idx As Long
    For call_idx = 1 To pCalls.Count
        Dim record_item As TestDoubleCallRecord
        Set record_item = pCalls.Item(call_idx)
        If pNormalizeMethodName(record_item.MethodName) = normalized_method_name Then
            result_count = result_count + 1
        End If
    Next call_idx

    GetCallCountAll = result_count
End Function

'* 0 ベースの呼び出し番号から記録を返します。
Public Function GetCall(ByVal CallIndex As Long) As TestDoubleCallRecord
    If CallIndex < 0 Or pCalls.Count <= CallIndex Then
        Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "呼び出し番号が範囲外です。(" & CallIndex & ")"
    End If

    Set GetCall = pCalls.Item(CallIndex + 1)
End Function

'* 指定された照合引数で呼ばれた呼び出し履歴を呼び出し順で返します。
Public Function GetCalls(ByVal MethodName As String, ParamArray KeyArgs() As Variant) As ObjectList
    Set GetCalls = GetCallsFromArray(MethodName, KeyArgs)
End Function

'* ParamArray と同じ形式の配列から、呼び出し履歴を呼び出し順で返します。
Public Function GetCallsFromArray(ByVal MethodName As String, ByVal KeyArgs As Variant) As ObjectList
    Dim object_key_mode As Long
    Dim key_args_part As String
    key_args_part = pBuildKeyArgsPart(MethodName, KeyArgs, object_key_mode)

    Dim store_key As String
    store_key = pBuildStoreKey(MethodName, object_key_mode, key_args_part)

    Dim result_list As ObjectList
    Set result_list = New ObjectList

    If pCallsByKey.Exists(store_key) Then
        Call pAddCallsToObjectList(result_list, pCallsByKey.Item(store_key))
    End If

    Set GetCallsFromArray = result_list
End Function

Private Sub pAddCallsToObjectList(ByVal TargetList As ObjectList, ByVal SourceCalls As Collection)
    Dim call_idx As Long
    For call_idx = 1 To SourceCalls.Count
        Call TargetList.Add(SourceCalls.Item(call_idx))
    Next call_idx
End Sub

'* 指定メソッド名の呼び出し履歴を呼び出し順で返します。
Public Function GetCallsAll(ByVal MethodName As String) As ObjectList
    Call pValidateMethodName(MethodName)

    Dim normalized_method_name As String
    normalized_method_name = pNormalizeMethodName(MethodName)

    Dim result_list As ObjectList
    Set result_list = New ObjectList

    Dim call_idx As Long
    For call_idx = 1 To pCalls.Count
        Dim record_item As TestDoubleCallRecord
        Set record_item = pCalls.Item(call_idx)
        If pNormalizeMethodName(record_item.MethodName) = normalized_method_name Then
            Call result_list.Add(record_item)
        End If
    Next call_idx

    Set GetCallsAll = result_list
End Function

'* 指定された照合引数で呼ばれた最新の呼び出し記録を返します。
Public Function GetLatestCall(ByVal MethodName As String, ParamArray KeyArgs() As Variant) As TestDoubleCallRecord
    Set GetLatestCall = GetLatestCallFromArray(MethodName, KeyArgs)
End Function

'* ParamArray と同じ形式の配列から最新の呼び出し記録を返します。
Public Function GetLatestCallFromArray(ByVal MethodName As String, ByVal KeyArgs As Variant) As TestDoubleCallRecord
    Dim object_key_mode As Long
    Dim key_args_part As String
    key_args_part = pBuildKeyArgsPart(MethodName, KeyArgs, object_key_mode)

    Dim store_key As String
    store_key = pBuildStoreKey(MethodName, object_key_mode, key_args_part)

    If Not pLatestByKey.Exists(store_key) Then
        Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "指定された呼び出しが記録されていません。(" & MethodName & ")"
    End If

    Set GetLatestCallFromArray = pLatestByKey.Item(store_key)
End Function

'* 指定メソッド名の最新の呼び出し記録を返します。
Public Function GetLatestCallAll(ByVal MethodName As String) As TestDoubleCallRecord
    Call pValidateMethodName(MethodName)

    Dim normalized_method_name As String
    normalized_method_name = pNormalizeMethodName(MethodName)

    Dim call_idx As Long
    For call_idx = pCalls.Count To 1 Step -1
        Dim record_item As TestDoubleCallRecord
        Set record_item = pCalls.Item(call_idx)
        If pNormalizeMethodName(record_item.MethodName) = normalized_method_name Then
            Set GetLatestCallAll = record_item
            Exit Function
        End If
    Next call_idx

    Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "指定された呼び出しが記録されていません。(" & MethodName & ")"
End Function

Private Function pBuildMethodKey(ByVal MethodName As String, ByVal KeyArgs As Variant) As String
    Dim object_key_mode As Long
    Dim key_args_part As String
    key_args_part = pBuildKeyArgsPart(MethodName, KeyArgs, object_key_mode)

    pBuildMethodKey = pBuildStoreKey(MethodName, object_key_mode, key_args_part)
End Function

Private Function pBuildOutputKey(ByVal MethodName As String, ByVal OutputName As String, ByVal KeyArgs As Variant) As String
    Dim object_key_mode As Long
    Dim key_args_part As String
    key_args_part = pBuildKeyArgsPart(MethodName, KeyArgs, object_key_mode)

    pBuildOutputKey = pNormalizeMethodName(MethodName) & vbTab _
            & "ObjectKeyMode(" & CStr(object_key_mode) & ")" & vbTab _
            & OutputName & vbTab & key_args_part
End Function

Private Function pBuildKeyArgsPart( _
        ByVal MethodName As String, _
        ByVal KeyArgs As Variant, _
        ByRef ObjectKeyMode As Long) As String

    Call pValidateMethodName(MethodName)
    Call pValidateKeyArgsArray(KeyArgs)

    Dim normalized_method_name As String
    normalized_method_name = pNormalizeMethodName(MethodName)
    Call pLockMethod(normalized_method_name)

    ObjectKeyMode = pGetObjectKeyModeByNormalizedName(normalized_method_name)
    pKeyBuilder.ObjectKeyMode = ObjectKeyMode

    pBuildKeyArgsPart = pKeyBuilder.BuildKeyFromArray(KeyArgs)
End Function

Private Function pBuildStoreKey(ByVal MethodName As String, ByVal ObjectKeyMode As Long, ByVal KeyArgsPart As String) As String
    pBuildStoreKey = pNormalizeMethodName(MethodName) & vbTab _
            & "ObjectKeyMode(" & CStr(ObjectKeyMode) & ")" & vbTab _
            & KeyArgsPart
End Function

Private Function pGetObjectKeyModeByNormalizedName(ByVal NormalizedMethodName As String) As Long
    If pObjectKeyModesByMethod.Exists(NormalizedMethodName) Then
        pGetObjectKeyModeByNormalizedName = CLng(pObjectKeyModesByMethod.Item(NormalizedMethodName))
    Else
        pGetObjectKeyModeByNormalizedName = pDefaultObjectKeyMode
    End If
End Function

Private Sub pLockMethod(ByVal NormalizedMethodName As String)
    pLockedMethods.Item(NormalizedMethodName) = True
End Sub

Private Function pNormalizeMethodName(ByVal MethodName As String) As String
    pNormalizeMethodName = LCase$(MethodName)
End Function

Private Sub pValidateMethodName(ByVal MethodName As String)
    If MethodName = "" Then
        Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "メソッド名は空文字列にできません。"
    End If
    If InStr(1, MethodName, vbTab, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "メソッド名にタブ文字は使用できません。(" & MethodName & ")"
    End If
End Sub

Private Sub pValidateOutputName(ByVal OutputName As String)
    If OutputName = "" Then
        Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "出力名は空文字列にできません。"
    End If
    If InStr(1, OutputName, vbTab, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", "出力名にタブ文字は使用できません。(" & OutputName & ")"
    End If
End Sub

Private Sub pSetStoredValue(ByVal TargetDict As Dictionary, ByVal StoreKey As String, ByRef NewValue As Variant)
    Dim value_holder As Dictionary
    Set value_holder = New Dictionary

    If IsObject(NewValue) Then
        Set value_holder.Item("Value") = NewValue
    Else
        value_holder.Item("Value") = NewValue
    End If

    Set TargetDict.Item(StoreKey) = value_holder
End Sub

Private Function pGetStoredItem(ByVal TargetDict As Dictionary, ByVal StoreKey As String, ByVal ErrorDescription As String) As Variant
    If Not TargetDict.Exists(StoreKey) Then
        Err.Raise vbObjectError + 1, "Class TestDoubleBehaviorStore", ErrorDescription
    End If

    Dim value_holder As Dictionary
    Set value_holder = TargetDict.Item(StoreKey)

    If IsObject(value_holder.Item("Value")) Then
        Set pGetStoredItem = value_holder.Item("Value")
    Else
        pGetStoredItem = value_holder.Item("Value")
    End If
End Function
