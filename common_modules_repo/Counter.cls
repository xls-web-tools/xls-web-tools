VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! カウントするためのクラスです。
'!
' #############################################################################

Private pCount As Long
Private pCountStep As Long
Private pStopWhenMax As Boolean
Private pMaxCount As Long

'* 現在のカウント (デフォルトは 0)
Public Property Get Count() As Long
    Count = pCount
End Property

'* カウントの間隔 (デフォルトは 1)
Public Property Get CountStep() As Long
    CountStep = pCountStep
End Property

'* MaxCount に達したらカウントをやめるか否か (HasNext が False を返すようになるか)
'* デフォルトは False であり、常に HasNext が True を返します。
Public Property Get StopWhenMax() As Boolean
    StopWhenMax = pStopWhenMax
End Property
Public Property Let StopWhenMax(ByVal NewValue As Boolean)
    pStopWhenMax = NewValue
End Property

'* カウントを止める最大値
Public Property Get MaxCount() As Long
    MaxCount = pMaxCount
End Property
Public Property Let MaxCount(ByVal NewValue As Long)
    pMaxCount = NewValue
End Property

Private Sub Class_Initialize()
    pCountStep = 1
End Sub

'* カウンタを初期化します。
'*
'* @param InitialCount 初期値 (デフォルトは 0)。
'* @param CountStepNumber カウント間隔 (デフォルトは 1)。
'*
'* @details
'* カウンタを初期化します。
Public Sub Initialize(Optional ByVal InitialCount As Long = 0, Optional ByVal CountStepNumber As Long = 1)
    If CountStepNumber = 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class Counter", Description:="CountStepNumber は 0 以外である必要があります。(" & CountStepNumber & ")"
        Exit Sub
    End If

    pCount = InitialCount
    pCountStep = CountStepNumber
End Sub

'* カウントを進められるかどうかを確認します。
'*
'* @return Boolean 次に進める場合は True、それ以外は False。
'*
'* @details
'* カウントを進められるかどうかを確認します。
Public Function HasNext() As Boolean
    If Not StopWhenMax Then
        HasNext = True
        Exit Function
    End If

    If 0 < pCountStep Then
        HasNext = CDbl(pCountStep) <= CDbl(pMaxCount) - CDbl(pCount)
    ElseIf pCountStep < 0 Then
        HasNext = CDbl(pMaxCount) - CDbl(pCount) <= CDbl(pCountStep)
    Else
        HasNext = False
    End If
End Function

'* カウントを次に進め、現在のカウント値を返します。
'*
'* @return Long 進んだ後のカウント値。
'* @details
'* カウントを次に進め、現在のカウント値を返します。HasNext が False の場合はエラーを発生させます。
Public Function GoNext() As Long
    If Not HasNext() Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class Counter", Description:="次のカウントはありません。(count: " & pCount & ", step: " & pCountStep & ", max: " & pMaxCount & ")"
    End If

    pCount = pCount + pCountStep
    GoNext = pCount
End Function
