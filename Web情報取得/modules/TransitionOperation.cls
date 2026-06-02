VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransitionOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! onClick 由来の画面遷移を WebDriver で再現するための操作定義です。
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pOperationName As String
Private pLocatorType As String
Private pLocatorValue As String
Private pScript As String
Private pWaitConditionName As String

'* 操作名。
Public Property Get OperationName() As String
    Call pCheckInit
    OperationName = pOperationName
End Property

'* locator 種別。
Public Property Get LocatorType() As String
    Call pCheckInit
    LocatorType = pLocatorType
End Property

'* locator 値。
Public Property Get LocatorValue() As String
    Call pCheckInit
    LocatorValue = pLocatorValue
End Property

'* click で再現できない場合だけ使う JavaScript。
Public Property Get Script() As String
    Call pCheckInit
    Script = pScript
End Property

'* 操作後の待機条件名。
Public Property Get WaitConditionName() As String
    Call pCheckInit
    WaitConditionName = pWaitConditionName
End Property

'* 操作定義を初期化します。
'*
'* @param OperationName 操作名。
'* @param LocatorType locator 種別。
'* @param LocatorValue locator 値。
'* @param Script click で再現できない場合だけ使う JavaScript。
'* @param WaitConditionName 操作後の待機条件名。
'*
'* @details
'* OperationName は必須です。LocatorType を指定する場合は LocatorValue も必須です。
'* LocatorType と Script がどちらも空の場合は、実行可能な画面遷移操作にならないためエラーにします。
Public Sub Initialize( _
        ByVal OperationName As String, _
        ByVal LocatorType As String, _
        ByVal LocatorValue As String, _
        Optional ByVal Script As String = "", _
        Optional ByVal WaitConditionName As String = "")

    If Trim$(OperationName) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperation", Description:="OperationName は必須です。"
    End If
    If Trim$(LocatorType) = "" And Trim$(Script) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperation", Description:="LocatorType または Script は必須です。(" & OperationName & ")"
    End If
    If Trim$(LocatorType) <> "" And Trim$(LocatorValue) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperation", Description:="LocatorValue は必須です。(" & OperationName & ")"
    End If

    pOperationName = OperationName
    pLocatorType = LocatorType
    pLocatorValue = LocatorValue
    pScript = Script
    pWaitConditionName = WaitConditionName
    pIsInitialized = True
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperation", Description:="初期化されていません。"
    End If
End Sub
