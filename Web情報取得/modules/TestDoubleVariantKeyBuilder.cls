VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDoubleVariantKeyBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Variant 値からテストダブル用のキーを生成するクラスです。
'!
' #############################################################################

Private pObjectKeyMode As Long

'* トップレベルのオブジェクト引数に適用するキー化モードを返します。
Public Property Get ObjectKeyMode() As Long
    ObjectKeyMode = pObjectKeyMode
End Property
'* トップレベルのオブジェクト引数に適用するキー化モードを設定します。
Public Property Let ObjectKeyMode(ByVal NewValue As Long)
    Call pValidateObjectKeyMode(NewValue)
    pObjectKeyMode = NewValue
End Property

Private Sub Class_Initialize()
    pObjectKeyMode = G_OBJECT_KEY_MODE_REFERENCE
End Sub

'* 指定された引数群からキーを生成します。
'*
'* @details
'* 0 件の引数は空文字列としてキー化します。
Public Function BuildKey(ParamArray Args() As Variant) As String
    BuildKey = BuildKeyFromArray(Args)
End Function

'* ParamArray と同じ形式の配列からキーを生成します。
'*
'* @details
'* Args には初期化済みの 1 次元配列を指定してください。
'* 非配列、多次元配列、未初期化の動的配列はエラーになります。
'* ObjectKeyMode はトップレベルのオブジェクト引数にだけ適用します。
'* 配列引数は `G_OBJECT_KEY_MODE_REFERENCE` でキー化するため、配列要素内のオブジェクトには ObjectKeyMode を適用しません。
Public Function BuildKeyFromArray(ByVal Args As Variant) As String
    BuildKeyFromArray = GetTypedMultiKeyByModeFromArray(pObjectKeyMode, Args)
End Function

'* 1 つの値からキーを生成します。
'*
'* @details
'* ObjectKeyMode はトップレベルのオブジェクト値にだけ適用します。
'* 配列値は `G_OBJECT_KEY_MODE_REFERENCE` でキー化するため、配列要素内のオブジェクトには ObjectKeyMode を適用しません。
Public Function BuildItemKey(ByVal Value As Variant) As String
    BuildItemKey = pBuildItemKey(Value)
End Function

Private Function pBuildItemKey(ByVal Value As Variant) As String
    Dim result_value As String
    If IsArray(Value) Then
        result_value = GetTypedValueKey(Value, G_OBJECT_KEY_MODE_REFERENCE)
    ElseIf IsObject(Value) Then
        result_value = pBuildObjectItemKey(Value)
    Else
        result_value = GetTypedValueKey(Value, G_OBJECT_KEY_MODE_REFERENCE)
    End If

    pBuildItemKey = result_value
End Function

Private Function pBuildObjectItemKey(ByVal Value As Variant) As String
    Dim effective_object_key_mode As Long
    Select Case pObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE
            effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE

        Case G_OBJECT_KEY_MODE_I_EQUATABLE
            If Value Is Nothing Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            ElseIf TypeOf Value Is IEquatable Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_I_EQUATABLE
            Else
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            End If

        Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            If Value Is Nothing Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            ElseIf TypeOf Value Is IDuplicateCheckable Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            Else
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            End If

        Case Else
            Err.Raise vbObjectError + 1, "Class TestDoubleVariantKeyBuilder", "未対応のオブジェクト キー化モードです。(" & CStr(pObjectKeyMode) & ")"
    End Select

    pBuildObjectItemKey = GetTypedValueKey(Value, effective_object_key_mode)
End Function

Private Sub pValidateObjectKeyMode(ByVal NewValue As Long)
    Select Case NewValue
        Case G_OBJECT_KEY_MODE_REFERENCE, G_OBJECT_KEY_MODE_I_EQUATABLE, G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            Exit Sub
        Case Else
            Err.Raise vbObjectError + 1, "Class TestDoubleVariantKeyBuilder", "未対応のオブジェクト キー化モードです。(" & CStr(NewValue) & ")"
    End Select
End Sub
