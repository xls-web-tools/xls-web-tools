VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! デバッグ用の情報を保持するクラスです。
'!
' #############################################################################

Private pTaskStack As ObjectList
Private pLocationStack As ObjectList
Private pMessage As String

'* デバッグ用のメッセージ。任意の文字列を保持するためのプロパティです。
Public Property Get Message() As String
    Message = pMessage
End Property
Public Property Let Message(ByVal NewValue As String)
    pMessage = NewValue
End Property

Private Sub Class_Initialize()
    Set pTaskStack = New ObjectList
    Set pLocationStack = New ObjectList
End Sub

'* デバッグ用情報をリセットします。
'*
'* @details
'* 現在保持しているタスクスタック、位置スタック、およびデバッグメッセージを初期化します。
Public Sub Reset()
    Set pTaskStack = New ObjectList
    Set pLocationStack = New ObjectList
    pMessage = ""
End Sub

'* Excel のアドレス情報を指定して、タスクを開始します。
'*
'* @param TaskName タスクの名前。
'* @param RangeBounds ワークシート範囲情報 (省略可能)。
'*
'* @details
'* 指定された Excel の位置情報を基にタスクを開始します。
Public Sub StartTask( _
        ByVal TaskName As String, _
        Optional ByVal RangeBounds As WorksheetRangeBounds = Nothing)

    If RangeBounds Is Nothing Then
        Call StartTaskWithLocationString(TaskName, WbSrv.GetThisWorkbookName())
    Else
        Call StartTaskWithLocationString(TaskName, RangeBounds.ToString())
    End If
End Sub

'* 任意の文字列を場所情報として指定して、タスクを開始します。
'*
'* @param TaskName タスクの名前。
'* @param LocationName 任意の場所情報。
'*
'* @details
'* 場所情報を任意の文字列で指定し、タスクを開始します。
Public Sub StartTaskWithLocationString( _
        ByVal TaskName As String, _
        ByVal LocationName As String)

    Call pTaskStack.PushItem(TaskName)
    Call pLocationStack.PushItem(LocationName)
End Sub

'* タスクを終了します。
'*
'* @param TaskName タスクの名前 (省略時はスタックの最後のタスクを終了)。
'* @param RaiseError エラー発生時に例外をスローするか (省略時は False)。
'*
'* @return Boolean タスクが正常に終了した場合は True、それ以外は False。
'*
'* @details
'* 指定されたタスク名をスタックから削除します。タスク名が省略された場合、最後に開始されたタスクを終了します。
'* 該当タスクが存在しない場合は RaiseError が True であればエラーを発生させます。
Public Function FinishTask(Optional ByVal TaskName As String = "", Optional ByVal RaiseError As Boolean = False) As Boolean
    If TaskName = "" Then
        If 0 < pTaskStack.Count Then
            Call pFinishTaskCore
            FinishTask = True
        Else
            If RaiseError Then
                Err.Raise vbObjectError + 1, "Class DebugInformation.FinishTask", "タスクが開始されていません。"
            Else
                FinishTask = False
            End If
        End If
        Exit Function
    End If

    Dim target_idx As Long

    target_idx = pTaskStack.GetIndexByItem(TaskName, Reverse:=True)

    If target_idx < 0 Then
        If RaiseError Then
            Err.Raise vbObjectError + 1, "Class DebugInformation.FinishTask", "終了しようとしたタスクが存在しません。(" & TaskName & ")"
        Else
            FinishTask = False
        End If
        Exit Function
    End If

    Dim item_idx As Long
    For item_idx = pTaskStack.Count - 1 To target_idx Step -1
        Call pFinishTaskCore
    Next item_idx
    FinishTask = True
End Function

Private Sub pFinishTaskCore()
    Call pTaskStack.PopItem
    Call pLocationStack.PopItem
End Sub

'* デバッグ情報をまとめた文字列を生成します。
'*
'* @param LineSep 行の区切り文字 (省略時は vbCrLf)。
'* @param LocationOnly 場所情報のみを含む場合は True (省略時は False)。
'* @return String デバッグ情報をまとめた文字列。
'*
'* @details
'* 現在保持しているタスクとその位置情報を文字列形式でまとめて返します。
Public Function BuildMessageLines(Optional ByVal LineSep As String = vbCrLf, Optional ByVal LocationOnly As Boolean = False) As String
    Dim result_value As String

    result_value = pMessage

    Dim item_idx As Long
    For item_idx = 0 To pTaskStack.Count - 1
        result_value = Strip(result_value)
        result_value = result_value & LineSep & pBuildMessageByIndex(item_idx, LocationOnly)
    Next item_idx

    BuildMessageLines = Strip(result_value)
End Function

'* 最後に開始されたタスクに関するメッセージを生成します。
'*
'* @param LocationOnly 場所情報のみを含む場合は True (省略時は False)。
'* @return String 最後のタスクのメッセージ。
'*
'* @details
'* 現在のタスクスタックの最後に追加されたタスクの情報を文字列で返します。
Public Function BuildCurrentMessage(Optional ByVal LocationOnly As Boolean = False) As String
    If 0 < pTaskStack.Count Then
        BuildCurrentMessage = pBuildMessageByIndex(pTaskStack.Count - 1, LocationOnly)
    Else
        BuildCurrentMessage = ""
    End If
End Function

Private Function pBuildMessageByIndex(ByVal ItemIndex As Long, ByVal LocationOnly As Boolean) As String
    Dim result_value As String

    If Not LocationOnly Then result_value = result_value & pTaskStack.Item(ItemIndex)

    Dim loc_str As String
    loc_str = pLocationStack.Item(ItemIndex)
    If loc_str <> "" Then result_value = result_value & "@" & loc_str

    If result_value <> "" Then result_value = "*" & result_value

    pBuildMessageByIndex = result_value
End Function

