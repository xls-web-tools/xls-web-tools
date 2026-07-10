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
Private pActionSelector As String
Private pActionInnerText As String
Private pActionScript As String
Private pWaitSelector As String

'* 操作名。
Public Property Get OperationName() As String
    Call pCheckInit
    OperationName = pOperationName
End Property

'* 操作対象の CSS selector。
Public Property Get ActionSelector() As String
    Call pCheckInit
    ActionSelector = pActionSelector
End Property

'* 操作対象の innerText 絞り込み値。
Public Property Get ActionInnerText() As String
    Call pCheckInit
    ActionInnerText = pActionInnerText
End Property

'* click で再現できない場合だけ使う JavaScript。
Public Property Get ActionScript() As String
    Call pCheckInit
    ActionScript = pActionScript
End Property

'* 操作後に到達を確認する CSS selector。
Public Property Get WaitSelector() As String
    Call pCheckInit
    WaitSelector = pWaitSelector
End Property

'* 操作定義を初期化します。
'*
'* @param OperationName 操作名。
'* @param ActionSelector 操作対象の CSS selector。
'* @param ActionInnerText 操作対象の innerText 絞り込み値。
'* @param ActionScript click で再現できない場合だけ使う JavaScript。
'* @param WaitSelector 操作後に到達を確認する CSS selector。
'*
'* @details
'* OperationName と WaitSelector は必須です。ActionSelector と ActionScript がどちらも空の場合は、実行可能な画面遷移操作にならないためエラーにします。
'* ActionInnerText を指定する場合は ActionSelector も必須です。
Public Sub Initialize( _
        ByVal OperationName As String, _
        ByVal ActionSelector As String, _
        Optional ByVal ActionInnerText As String = "", _
        Optional ByVal ActionScript As String = "", _
        Optional ByVal WaitSelector As String = "")

    If Trim$(OperationName) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperation.Initialize", Description:="OperationName は必須です。"
    End If
    If Trim$(ActionInnerText) <> "" And Trim$(ActionSelector) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperation.Initialize", Description:="ActionInnerText を指定する場合、ActionSelector は必須です。(" & OperationName & ")"
    End If
    If Trim$(ActionSelector) = "" And Trim$(ActionScript) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperation.Initialize", Description:="ActionSelector または ActionScript は必須です。(" & OperationName & ")"
    End If
    If Trim$(WaitSelector) = "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperation.Initialize", Description:="WaitSelector は必須です。(" & OperationName & ")"
    End If

    pOperationName = OperationName
    pActionSelector = ActionSelector
    pActionInnerText = ActionInnerText
    pActionScript = ActionScript
    pWaitSelector = WaitSelector
    pIsInitialized = True
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperation.pCheckInit", Description:="初期化されていません。"
    End If
End Sub
