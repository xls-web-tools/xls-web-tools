VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 進捗状況を管理するクラスです。タスクの進行状態を視覚的に表示し、更新します。
'!
' #############################################################################

Private Const C_MSG_COMPLETE_BLOCK As String = "■"
Private Const C_MSG_UNCOMPLETE_BLOCK As String = "□"

Private pTaskName As String
Private pBlockCount As Long
Private pStartValue As Double
Private pTotalValue As Double
Private pProcessedValue As Double
Private pIsComplete As Boolean
Private pIntervalPercent As Integer
Private pNextUpdate As Integer

'* タスク名。
Public Property Get TaskName() As String
    TaskName = pTaskName
End Property
Public Property Let TaskName(ByVal NewValue As String)
    pTaskName = NewValue
End Property

'* ブロック数。(デフォルトは 10)
'* 進捗バーのセグメント数です。
Public Property Get BlockCount() As Long
    BlockCount = pBlockCount
End Property
Public Property Let BlockCount(ByVal NewValue As Long)
    If NewValue < 1 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ProgressStatus", Description:="ブロック数は正の整数である必要があります。(" & NewValue & ")"
        Exit Property
    End If
    pBlockCount = NewValue
End Property

'* 開始値。(デフォルトは 0)
'* このプロパティを変更すると、進捗値も変更されます。
Public Property Get StartValue() As Double
    StartValue = pStartValue
End Property
Public Property Let StartValue(ByVal NewValue As Double)
    If TotalValue <= NewValue Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ProgressStatus", Description:="開始値は終了値より小さい必要があります。(新しい値: " & NewValue & " / TotalValue: " & pTotalValue & ")"
        Exit Property
    End If
    Dim value_diff As Double
    value_diff = pStartValue - NewValue
    pStartValue = NewValue
    pProcessedValue = pProcessedValue - value_diff
    pNextUpdate = 0
End Property

'* 終了値。(デフォルトは 100)
Public Property Get TotalValue() As Double
    TotalValue = pTotalValue
End Property
Public Property Let TotalValue(ByVal NewValue As Double)
    If NewValue <= StartValue Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ProgressStatus", Description:="終了値は開始値より大きい必要があります。(新しい値: " & NewValue & " / StartValue: " & pStartValue & ")"
        Exit Property
    End If
    pTotalValue = NewValue
End Property

'* 現在の進捗値。(デフォルトは 0)
Public Property Get ProcessedValue() As Double
    ProcessedValue = pProcessedValue
End Property
Public Property Let ProcessedValue(ByVal NewValue As Double)
    pProcessedValue = NewValue
    Call pUpdateStatus
End Property

'* タスクが完了しているか否か。
Public Property Get IsComplete() As Boolean
    IsComplete = pIsComplete
End Property

Private Sub Class_Initialize()
    Call pInitializeStartAndTotal
    pBlockCount = 10
    pProcessedValue = pStartValue
    pIsComplete = False
    pIntervalPercent = 3
    pNextUpdate = 0
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

'* 指定された範囲でループを処理するための設定を行います。
'*
'* @param FinishIndex ループの終了インデックス。
'* @param Optional StartIndex ループの開始インデックス。デフォルトは 0。
'*
'* @details
'* StartIndex と FinishIndex を設定し、インデックスを ProcessedValue に設定することにより進捗が更新されるようにして、ループ処理に対応する進捗を管理します。
Public Sub SetForLoop(ByVal FinishIndex As Long, Optional ByVal StartIndex As Long = 0)
    Dim start_value As Double
    start_value = CDbl(StartIndex) - 1

    Dim total_value As Double
    total_value = CDbl(FinishIndex)

    If total_value <= start_value Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ProgressStatus", Description:="終了値は開始値より大きい必要があります。(FinishIndex: " & FinishIndex & " / StartIndex: " & StartIndex & ")"
        Exit Sub
    End If

    pStartValue = start_value
    pTotalValue = total_value
    pProcessedValue = pStartValue
    pIsComplete = False
    pNextUpdate = 0
End Sub

'* 進捗状況の監視を開始します。
'*
'* @return 進捗監視が正常に開始されたかどうかを示すブール値。
'*
'* @details
'* 進捗の初期値を設定し、進捗状況の表示を開始します。
Public Function Start() As Boolean
    If pProcessedValue < pStartValue + 1 Then
        pProcessedValue = pStartValue + 1
        If pTotalValue < pProcessedValue Then pProcessedValue = pTotalValue
        Call pUpdateStatus
        Start = True
    Else
        Start = False
    End If
End Function

'* タスクを完了とマークします。
'*
'* @details
'* 進捗値を最大値に設定し、タスクを完了状態にします。
Public Sub MarkAsComplete()
    pProcessedValue = pTotalValue
    Call pUpdateStatus
End Sub

'* 進捗状況をリセットします。
'*
'* @param Optional ResetTaskName タスク名もリセットするかどうか。デフォルトは True。
'* @param Optional ResetStartAndTotalValue 開始値と終了値もリセットするかどうか。デフォルトは True。
'*
'* @details
'* 進捗状況を初期状態に戻し、必要に応じてタスク名や値もリセットします。
Public Sub Reset(Optional ByVal ResetTaskName As Boolean = True, Optional ByVal ResetStartAndTotalValue As Boolean = True)
    If ResetTaskName Then pTaskName = ""
    If ResetStartAndTotalValue Then Call pInitializeStartAndTotal
    pProcessedValue = pStartValue
    pIsComplete = False
    pNextUpdate = 0
    Application.StatusBar = False
End Sub

'* 現在の進捗情報を手動で表示します。
'*
'* @details
'* 進捗状況を強制的に更新し、ステータスバーに表示します。
Public Sub DisplayProgress()
    Call pDisplayProgress(Force:=True)
End Sub

Private Sub pDisplayProgress(Optional ByVal Force As Boolean = False)
    Dim prog_rate As Double
    Dim prog_percent As Integer

    prog_rate = (pProcessedValue - pStartValue) / (pTotalValue - pStartValue)

    If prog_rate < 0 Then
        prog_rate = 0
    ElseIf 1 < prog_rate Then
        prog_rate = 1
    End If

    prog_percent = Int((prog_rate + 0.0001) * 100)

    If Force Or pNextUpdate <= prog_percent Then
        pNextUpdate = pGetNextUpdate(prog_percent)
        Application.StatusBar = pGetStatusString(prog_rate, prog_percent)
        DoEvents
    End If
End Sub

Private Sub pInitializeStartAndTotal()
    pStartValue = 0
    pTotalValue = 100
End Sub

Private Sub pUpdateStatus()
    If pTotalValue <= pProcessedValue Then
        pIsComplete = True
    Else
        pIsComplete = False
    End If


    If pIsComplete Then
        Application.StatusBar = False
        Exit Sub
    End If

    Call pDisplayProgress
End Sub

Private Function pGetNextUpdate(ByVal ProgressPercent As Integer) As Integer
    Dim result_value As Long

    result_value = (ProgressPercent \ pIntervalPercent + 1) * pIntervalPercent

    If 100 < result_value Then
        result_value = 100
    End If

    pGetNextUpdate = result_value
End Function

Private Function pGetStatusString(ByVal ProgressRate As Double, ByVal ProgressPercent As Integer) As String
    Dim result_value As String
    Dim comp_blocks As Long
    Dim item_idx As Long

    result_value = pTaskName & "[" & ProgressPercent & "%]"

    comp_blocks = Int(pBlockCount * (ProgressRate + 0.0001))
    For item_idx = 1 To comp_blocks
        result_value = result_value & C_MSG_COMPLETE_BLOCK
    Next item_idx

    comp_blocks = pBlockCount - comp_blocks
    For item_idx = 1 To comp_blocks
        result_value = result_value & C_MSG_UNCOMPLETE_BLOCK
    Next item_idx

    pGetStatusString = result_value
End Function
