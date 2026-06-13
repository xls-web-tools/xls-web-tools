VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CounterSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Counter をまとめるクラスです。
'!
' #############################################################################

Private pDict As Dictionary

Private Sub Class_Initialize()
    Set pDict = New Dictionary
End Sub

'* 名前を指定してカウンターを取得します。
'*
'* @param CounterObjectName カウンターオブジェクトの名前。
'* @return Counter 指定された名前に対応するカウンターオブジェクト。
'* @details
'* 指定された名前のカウンターオブジェクトを取得します。
Public Function GetCounter(ByVal CounterObjectName As String) As Counter
    Call pCheckExists(CounterObjectName)
    Set GetCounter = pDict(CounterObjectName)
End Function

'* 名前を指定してカウンターを追加します。
'*
'* @param CounterObjectName カウンターオブジェクトの名前。
'* @param CounterObject 追加するカウンターオブジェクト。
'* @details
'* 名前とカウンターオブジェクトを指定してセットに追加します。
Public Sub AddCounter(ByVal CounterObjectName As String, ByVal CounterObject As Counter)
    If CounterObject Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class CounterSet.AddCounter", Description:="CounterObject が Nothing です。"
    End If

    Call pDict.Add(CounterObjectName, CounterObject)
End Sub

'* 名前と初期値を指定して、新しいカウンターを生成し、セットに追加します。
'*
'* @param CounterObjectName カウンターオブジェクトの名前。
'* @param InitialCount カウンターの初期値 (デフォルトは 0)。
'* @param CountStep カウンターの間隔 (デフォルトは 1)。
'* @return Counter 生成されたカウンターオブジェクト。
'* @details
'* 名前、初期値、間隔を指定して新しいカウンターを生成し、セットに追加します。
Public Function AddNewCounter(ByVal CounterObjectName As String, Optional ByVal InitialCount As Long = 0, Optional ByVal CountStep As Long = 1) As Counter
    Dim count_obj As Counter
    Set count_obj = New Counter
    Call count_obj.Initialize(InitialCount, CountStep)
    Call AddCounter(CounterObjectName, count_obj)
    Set AddNewCounter = count_obj
End Function

'* 名前を指定してカウンターを削除します。
'*
'* @param CounterObjectName カウンターオブジェクトの名前。
'* @return Counter 削除されたカウンターオブジェクト。
'* @details
'* 指定された名前のカウンターを削除し、削除されたカウンターオブジェクトを返します。
Public Function RemoveCounter(ByVal CounterObjectName As String) As Counter
    Call pCheckExists(CounterObjectName)

    Dim result_value As Counter
    Set result_value = pDict(CounterObjectName)
    Call pDict.Remove(CounterObjectName)
    Set RemoveCounter = result_value
End Function

Private Sub pCheckExists(ByVal CounterObjectName As String)
    If Not pDict.Exists(CounterObjectName) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class CounterSet.pCheckExists", Description:="Counter が存在しません。(" & CounterObjectName & ")"
    End If
End Sub
