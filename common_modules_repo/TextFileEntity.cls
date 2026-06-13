VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFileEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITextFileEntity

' #############################################################################
'!
'! @brief
'! テキスト ファイルのオープン、クローズ、書き込みなどを行うための実装です。
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pIsOpen As Boolean
Private pIsFirst As Boolean
Private pFilePath As String
Private pFileDesc As Integer
Private pAsRead As Boolean
Private pAsWrite As Boolean
Private pAsAppend As Boolean
Private pGetReadLock As Boolean
Private pGetWriteLock As Boolean

' ファイルのパス。
Public Property Get FilePath() As String
    Call pCheckInit
    FilePath = pFilePath
End Property
Private Property Get ITextFileEntity_FilePath() As String
    ITextFileEntity_FilePath = FilePath
End Property

' 読み取り用か。(未オープンの場合は False)
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

' 書き込み用か。(未オープンの場合は False)
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

' 追記か。(未オープンの場合は False)
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

' 読み取りロックを取得しているか。(未オープンの場合は False)
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

' 書き込みロックを取得しているか。(未オープンの場合は False)
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

' 現在ファイルがオープンされているか。
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
    If Not pIsOpen Then
        IsEndOfFile = True
    ElseIf pAsWrite Then
        IsEndOfFile = True
    Else
        IsEndOfFile = EOF(pFileDesc)
    End If
End Property
Private Property Get ITextFileEntity_IsEndOfFile() As Boolean
    ITextFileEntity_IsEndOfFile = IsEndOfFile
End Property

'* 各値を設定します。
'*
'* @param FilePath ファイルのパス。
'*
'* @details
'* 各値を設定します。初期化済みインスタンスは再初期化できません。
Public Sub Initialize(ByVal FilePath As String)
    If pIsOpen Then Err.Raise vbObjectError + 1, "Class TextFileEntity.Initialize", "オープン中のファイルは再初期化できません。(" & pFilePath & ")"
    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class TextFileEntity.Initialize", "すでに初期化されています。(" & pFilePath & ")"

    pFilePath = pGetAbsolutePath(FilePath)
    pIsInitialized = True
End Sub

Private Function pGetAbsolutePath(ByVal PathLikeString As String) As String
    Call pRequireFileSystemService("pGetAbsolutePath")
    pGetAbsolutePath = FsSrv.GetAbsolutePath(PathLikeString)
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    If pIsOpen Then
        Close #pFileDesc
        Call pResetOpenState
    End If
    On Error GoTo 0
End Sub

' ファイルをオープンします。
'
' @param FilePath ファイルのパス。
' @param AsWrite [省略可] 書き込み用として開くかどうか。デフォルトは False であり、読み取り用として開きます。
' @param AsAppend [省略可] 追記として開くかどうか。AsWrite が False のときに AsAppend を True にすると失敗します。
' @param GetReadLock [省略可] ファイルの読み取りロックを取得するか。
' @param GetWriteLock [省略可] ファイルの書き込みロックを取得するか。
'
' @details
' ファイルをオープンします。
Public Sub OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)

    Call pCheckInit
    If pIsOpen Then Err.Raise vbObjectError, "Class TextFileEntity.OpenFile", "ファイルは既に開かれています。(" & pFilePath & ")"
    If AsAppend And Not AsWrite Then Err.Raise vbObjectError, "Class TextFileEntity.OpenFile", "AsWrite が False のときに AsAppend が True となりました。"

    If Dir(pFilePath, vbNormal Or vbHidden Or vbSystem Or vbDirectory) <> "" And AsWrite And Not AsAppend Then
        If Not Force Then Err.Raise vbObjectError + 1, "Class TextFileEntity.OpenFile", "すでにファイルが存在しています。(" & pFilePath & ")"
    End If

    pGetReadLock = GetReadLock
    pGetWriteLock = GetWriteLock
    pAsRead = Not AsWrite
    pAsWrite = AsWrite
    pAsAppend = AsAppend

    pFileDesc = FreeFile()

    If pAsWrite Then
        If pAsAppend Then
            If pGetReadLock Then
                Open pFilePath For Append Lock Read Write As #pFileDesc
            ElseIf pGetWriteLock Then
                Open pFilePath For Append Lock Write As #pFileDesc
            Else
                Open pFilePath For Append Shared As #pFileDesc
            End If
        Else
            If pGetReadLock Then
                Open pFilePath For Output Lock Read Write As #pFileDesc
            ElseIf pGetWriteLock Then
                Open pFilePath For Output Lock Write As #pFileDesc
            Else
                Open pFilePath For Output Shared As #pFileDesc
            End If
        End If
    Else
        If pGetReadLock Then
            Open pFilePath For Input Lock Read Write As #pFileDesc
        ElseIf pGetWriteLock Then
            Open pFilePath For Input Lock Write As #pFileDesc
        Else
            Open pFilePath For Input Shared As #pFileDesc
        End If
    End If

    pIsOpen = True
End Sub
Private Sub ITextFileEntity_OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)

    Call OpenFile(AsWrite:=AsWrite, AsAppend:=AsAppend, GetReadLock:=GetReadLock, GetWriteLock:=GetWriteLock, Force:=Force)
End Sub

' ファイルから 1 行読み取ります。
'
' @return 読み取った文字列。
'
' @details
' ファイルから文字列を 1 行分読み取ります。
Public Function ReadLine() As String
    Call pCheckOpened
    If pAsWrite Then Err.Raise vbObjectError, "Class TextFileEntity.ReadLine", "書き込み用で開いています。(" & pFilePath & ")"

    Dim result_value As String
    Line Input #pFileDesc, result_value
    ReadLine = result_value
End Function
Private Function ITextFileEntity_ReadLine() As String
    ITextFileEntity_ReadLine = ReadLine()
End Function

' 文字列をファイルに書き込みます。
'
' @param Expression 書き込む文字列。
'
' @details
' 文字列をファイルに書き込みます。末尾に改行記号が自動で付与されます。
Public Sub WriteLine(ByVal Expression As String)
    Call pCheckOpened
    If pAsRead Then Err.Raise vbObjectError, "Class TextFileEntity.WriteLine", "読み取り用で開いています。(" & pFilePath & ")"

    Print #pFileDesc, Expression
End Sub
Private Sub ITextFileEntity_WriteLine(ByVal Expression As String)
    Call WriteLine(Expression)
End Sub

' ファイルをクローズします。
'
' @details
' ファイルをクローズします。未オープンの場合は何もしません。
Public Sub CloseFile()
    Call pCheckInit
    If Not IsOpen Then Exit Sub

    Close #pFileDesc
    Call pResetOpenState
End Sub
Private Sub ITextFileEntity_CloseFile()
    Call CloseFile
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class TextFileEntity.pCheckInit", "未初期化状態です。"
End Sub

Private Sub pCheckOpened()
    If Not pIsOpen Then Err.Raise vbObjectError + 1, "Class TextFileEntity.pCheckOpened", "ファイルがオープンされていません。"
End Sub

Private Sub pResetOpenState()
    pFileDesc = 0
    pIsOpen = False
    pAsRead = False
    pAsWrite = False
    pAsAppend = False
    pGetReadLock = False
    pGetWriteLock = False
End Sub

Private Sub pRequireFileSystemService(ByVal MethodName As String)
    If FsSrv Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TextFileEntity" & "." & MethodName, Description:="InitializeCommonService が呼ばれていません。"
    End If
End Sub
