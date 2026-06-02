VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFileEntityTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITextFileEntity

' #############################################################################
'!
'! @brief
'! テキスト ファイルのオープン、クローズ、書き込みなどを行うための lightweight fake です。
'!
'! @details
'! ファイル実体は操作せず、初期化、open/read/write/close のライフサイクル状態だけを管理します。
'! ReadLine は strict stub とし、読み取り行を Store に明示登録してください。
'! OpenFile / WriteLine / CloseFile は Store.SetError による任意エラー注入を扱いません。
'! ファイル open の任意失敗をテストする場合は TextFileServiceTestDouble.GetTextFileEntity 側でエラー設定するか、対象処理を分けてテストしてください。
'!
' #############################################################################

Private pFilePath As String
Private pIsInitialized As Boolean
Private pIsOpen As Boolean
Private pAsRead As Boolean
Private pAsWrite As Boolean
Private pAsAppend As Boolean
Private pGetReadLock As Boolean
Private pGetWriteLock As Boolean

'* 読み取り回数。ReadLine によりインクリメントされ、CloseFile によりリセットされます。
Public ReadCount As Long

'* 書き込み回数。WriteLine によりインクリメントされます。リセットはされません。
Public WriteCount As Long

'* スタブ設定と正常に成立した呼び出し履歴。
'* この lightweight fake の状態違反エラーは呼び出し履歴に残しません。
Public Store As TestDoubleBehaviorStore

'* ファイルのパス。
Public Property Get FilePath() As String
    Call pCheckInit
    FilePath = pFilePath
End Property
Private Property Get ITextFileEntity_FilePath() As String
    ITextFileEntity_FilePath = FilePath
End Property

'* 読み取り用か。(未オープンの場合は False)
Public Property Get AsRead() As Boolean
    Call pCheckInit
    If pIsOpen Then
        AsRead = pAsRead
    Else
        AsRead = False
    End If
End Property
Private Property Get ITextFileEntity_AsRead() As Boolean
    ITextFileEntity_AsRead = AsRead
End Property

'* 書き込み用か。(未オープンの場合は False)
Public Property Get AsWrite() As Boolean
    Call pCheckInit
    If pIsOpen Then
        AsWrite = pAsWrite
    Else
        AsWrite = False
    End If
End Property
Private Property Get ITextFileEntity_AsWrite() As Boolean
    ITextFileEntity_AsWrite = AsWrite
End Property

'* 追記か。(未オープンの場合は False)
Public Property Get AsAppend() As Boolean
    Call pCheckInit
    If pIsOpen Then
        AsAppend = pAsAppend
    Else
        AsAppend = False
    End If
End Property
Private Property Get ITextFileEntity_AsAppend() As Boolean
    ITextFileEntity_AsAppend = AsAppend
End Property

'* 読み取りロックを取得しているか。(未オープンの場合は False)
Public Property Get GetReadLock() As Boolean
    Call pCheckInit
    If pIsOpen Then
        GetReadLock = pGetReadLock
    Else
        GetReadLock = False
    End If
End Property
Private Property Get ITextFileEntity_GetReadLock() As Boolean
    ITextFileEntity_GetReadLock = GetReadLock
End Property

'* 書き込みロックを取得しているか。(未オープンの場合は False)
Public Property Get GetWriteLock() As Boolean
    Call pCheckInit
    If pIsOpen Then
        GetWriteLock = pGetWriteLock
    Else
        GetWriteLock = False
    End If
End Property
Private Property Get ITextFileEntity_GetWriteLock() As Boolean
    ITextFileEntity_GetWriteLock = GetWriteLock
End Property

'* 現在ファイルがオープンされているか。
Public Property Get IsOpen() As Boolean
    Call pCheckInit
    IsOpen = pIsOpen
End Property
Private Property Get ITextFileEntity_IsOpen() As Boolean
    ITextFileEntity_IsOpen = IsOpen
End Property

' ファイルの最後か。(未オープンまたは書き込み/追記モードの場合は True)
Public Property Get IsEndOfFile() As Boolean
    Call pCheckInit
    If Not pIsOpen Or pAsWrite Then
        IsEndOfFile = True
    Else
        IsEndOfFile = Not Store.HasReturn("ReadLine", ReadCount)
    End If
End Property
Private Property Get ITextFileEntity_IsEndOfFile() As Boolean
    ITextFileEntity_IsEndOfFile = IsEndOfFile
End Property

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* Initialize の spy only です。初期化パスを記録します。
Public Sub Initialize(ByVal FilePath As String)
    If pIsOpen Then Err.Raise vbObjectError + 1, "Class TextFileEntityTestDouble", "オープン中のファイルは再初期化できません。(" & pFilePath & ")"
    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class TextFileEntityTestDouble", "すでに初期化されています。(" & pFilePath & ")"

    Call Store.RecordCall("Initialize", FilePath)
    pFilePath = FilePath
    pIsInitialized = True
End Sub

'* OpenFile の lightweight fake です。引数に基づいて状態を更新し、呼び出しを記録します。
Public Sub OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)

    Call pCheckInit
    If pIsOpen Then Err.Raise vbObjectError, "Class TextFileEntityTestDouble", "ファイルは既に開かれています。(" & pFilePath & ")"
    If AsAppend And Not AsWrite Then Err.Raise vbObjectError, "Class TextFileEntityTestDouble", "AsWrite が False のときに AsAppend が True となりました。"

    Call Store.RecordCall("OpenFile", AsWrite, AsAppend, GetReadLock, GetWriteLock, Force)

    pAsRead = Not AsWrite
    pAsWrite = AsWrite
    pAsAppend = AsAppend
    pGetReadLock = GetReadLock
    pGetWriteLock = GetWriteLock
    pIsOpen = True
End Sub
Private Sub ITextFileEntity_OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)

    Call OpenFile(AsWrite:=AsWrite, AsAppend:=AsAppend, GetReadLock:=GetReadLock, GetWriteLock:=GetWriteLock, Force:=Force)
End Sub

'* ReadLine の strict stub です。読み取り行は Store に明示登録してください。
Public Function ReadLine() As String
    Call pCheckOpened
    If pAsWrite Then Err.Raise vbObjectError, "Class TextFileEntityTestDouble", "書き込み用で開かれています。(" & pFilePath & ")"

    If Not Store.HasReturn("ReadLine", ReadCount) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TextFileEntityTestDouble", Description:="読み取り行が登録されていません。(read_count: " & ReadCount & ")"
    End If

    Dim result_value As String
    result_value = Store.GetReturn("ReadLine", ReadCount)
    Call Store.RecordCallWithKeyArgs("ReadLine", Array(ReadCount))
    ReadCount = ReadCount + 1

    ReadLine = result_value
End Function
Private Function ITextFileEntity_ReadLine() As String
    ITextFileEntity_ReadLine = ReadLine()
End Function

'* WriteLine の spy only です。書き込み行を記録し、状態に基づいて WriteCount を更新します。
Public Sub WriteLine(ByVal Expression As String)
    Call pCheckOpened
    If pAsRead Then Err.Raise vbObjectError, "Class TextFileEntityTestDouble", "読み取り用で開かれています。(" & pFilePath & ")"

    Call Store.RecordCall("WriteLine", Expression)
    WriteCount = WriteCount + 1
End Sub
Private Sub ITextFileEntity_WriteLine(ByVal Expression As String)
    Call WriteLine(Expression)
End Sub

'* CloseFile の lightweight fake です。0 引数呼び出しとして記録し、open 状態なら状態をリセットします。
Public Sub CloseFile()
    Call pCheckInit
    Call Store.RecordCall("CloseFile")

    If Not pIsOpen Then Exit Sub

    ReadCount = 0
    Call pResetOpenState
End Sub
Private Sub ITextFileEntity_CloseFile()
    Call CloseFile
End Sub

Private Sub pResetOpenState()
    pIsOpen = False
    pAsRead = False
    pAsWrite = False
    pAsAppend = False
    pGetReadLock = False
    pGetWriteLock = False
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class TextFileEntityTestDouble", "初期化前です。"
End Sub

Private Sub pCheckOpened()
    Call pCheckInit
    If Not pIsOpen Then Err.Raise vbObjectError + 1, "Class TextFileEntityTestDouble", "ファイルがオープンされていません。"
End Sub
