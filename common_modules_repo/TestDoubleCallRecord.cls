VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDoubleCallRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! テストダブルの 1 回分の呼び出し記録です。
'!
' #############################################################################

Private pMethodName As String
Private pCallIndex As Long
Private pObjectKeyMode As Long
Private pArguments As Collection

'* 呼び出されたメソッド名。
Public Property Get MethodName() As String
    MethodName = pMethodName
End Property
Public Property Let MethodName(ByVal NewValue As String)
    pMethodName = NewValue
End Property

'* 呼び出し履歴内の 0 ベースの呼び出し番号。
Public Property Get CallIndex() As Long
    CallIndex = pCallIndex
End Property
Public Property Let CallIndex(ByVal NewValue As Long)
    pCallIndex = NewValue
End Property

'* 呼び出し照合に使ったオブジェクト キー化モード。
Public Property Get ObjectKeyMode() As Long
    ObjectKeyMode = pObjectKeyMode
End Property
Public Property Let ObjectKeyMode(ByVal NewValue As Long)
    pObjectKeyMode = NewValue
End Property

'* 呼び出し引数の件数を返します。
Public Property Get ArgumentCount() As Long
    ArgumentCount = pArguments.Count
End Property

Private Sub Class_Initialize()
    Set pArguments = New Collection
End Sub

'* 呼び出し引数を保存します。
Public Sub StoreArguments(ByVal Args As Variant)
    If Not IsArray(Args) Then
        Err.Raise vbObjectError + 1, "Class TestDoubleCallRecord.StoreArguments", "呼び出し引数は配列で指定してください。"
    End If

    Set pArguments = New Collection
    If IsEmptyArray(Args) Then Exit Sub

    Dim arg_item As Variant
    For Each arg_item In Args
        Call pAddArgument(arg_item)
    Next arg_item
End Sub

Private Sub pAddArgument(ByVal NewValue As Variant)
    Dim value_holder As Dictionary
    Set value_holder = New Dictionary

    If IsObject(NewValue) Then
        Set value_holder.Item("Value") = NewValue
    Else
        value_holder.Item("Value") = NewValue
    End If

    Call pArguments.Add(value_holder)
End Sub

'* 0 ベースの引数番号から呼び出し引数を返します。
Public Function GetArgument(ByVal ArgumentIndex As Long) As Variant
    If ArgumentIndex < 0 Or pArguments.Count <= ArgumentIndex Then
        Err.Raise vbObjectError + 1, "Class TestDoubleCallRecord.GetArgument", "呼び出し引数番号が範囲外です。(" & CStr(ArgumentIndex) & ")"
    End If

    Dim value_holder As Dictionary
    Set value_holder = pArguments.Item(ArgumentIndex + 1)

    If IsObject(value_holder.Item("Value")) Then
        Set GetArgument = value_holder.Item("Value")
    Else
        GetArgument = value_holder.Item("Value")
    End If
End Function
