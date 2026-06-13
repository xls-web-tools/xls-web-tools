VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransitionOperationChain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 画面遷移操作チェーン記法から解決された操作列です。
'!
' #############################################################################

Private pOperations As ObjectList
Private pIsInitialized As Boolean

'* チェーン内の画面遷移操作数。
Public Property Get Count() As Long
    Call pCheckInit
    Count = pOperations.Count
End Property

'* 指定されたインデックスの画面遷移操作を返します。
'*
'* @param ItemIndex 0-based の操作インデックス。
'* @return 指定された画面遷移操作。
'*
'* @details
'* 指定されたインデックスの画面遷移操作を返します。
Public Function Item(ByVal ItemIndex As Long) As TransitionOperation
    Call pCheckInit
    Set Item = pOperations.Item(ItemIndex)
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperationChain.pCheckInit", Description:="初期化されていません。"
    End If
End Sub

'* 画面遷移操作チェーンを初期化します。
'*
'* @param Operations 実行順の画面遷移操作一覧。
'*
'* @details
'* 画面遷移操作チェーンを初期化します。
Public Sub Initialize(ByVal Operations As ObjectList)
    If Operations Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperationChain.Initialize", Description:="Operations は必須です。"
    End If
    If Operations.Count = 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TransitionOperationChain.Initialize", Description:="Operations は 1 件以上必要です。"
    End If

    Set pOperations = Operations
    pIsInitialized = True
End Sub
