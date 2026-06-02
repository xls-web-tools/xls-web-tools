VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITextFileEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#Interface

' #############################################################################
'!
'! @brief
'! テキスト ファイルのオープン、クローズ、書き込みなどを行うためのインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface ITextFileEntity", Description:="このクラスはインターフェイス用です。"
End Sub

'* ファイルのパス。
Public Property Get FilePath() As String
End Property

'* 読み取り用か。(未オープンの場合は False)
Public Property Get AsRead() As Boolean
End Property

'* 書き込み用か。(未オープンの場合は False)
Public Property Get AsWrite() As Boolean
End Property

'* 追記か。(未オープンの場合は False)
Public Property Get AsAppend() As Boolean
End Property

'* 読み取りロックを取得しているか。(未オープンの場合は False)
Public Property Get GetReadLock() As Boolean
End Property

'* 書き込みロックを取得しているか。(未オープンの場合は False)
Public Property Get GetWriteLock() As Boolean
End Property

'* 現在ファイルがオープンされているか。
Public Property Get IsOpen() As Boolean
End Property

'* ファイルの最後か。(未オープンまたは書き込み/追記モードの場合は True)
Public Property Get IsEndOfFile() As Boolean
End Property

'* ファイルをオープンします。
'*
'* @param FilePath ファイルのパス。
'* @param AsWrite [省略可] 書き込み用として開くかどうか。デフォルトは False であり、読み取り用として開きます。
'* @param AsAppend [省略可] 追記として開くかどうか。AsWrite が False のときに AsAppend を True にすると失敗します。
'* @param GetReadLock [省略可] ファイルの読み取りロックを取得するか。
'* @param GetWriteLock [省略可] ファイルの書き込みロックを取得するか。
'*
'* @details
'* ファイルをオープンします。
Public Sub OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)
End Sub

'* ファイルから 1 行読み取ります。
'*
'* @return 読み取った文字列。
'*
'* @details
'* ファイルから文字列を 1 行分読み取ります。
Public Function ReadLine() As String
End Function

'* 文字列をファイルに書き込みます。
'*
'* @param Expression 書き込む文字列。
'*
'* @details
'* 文字列をファイルに書き込みます。末尾に改行記号が自動で付与されます。
Public Sub WriteLine(ByVal Expression As String)
End Sub

'* ファイルをクローズします。
'*
'* @details
'* ファイルをクローズします。未オープンの場合は何もしません。
Public Sub CloseFile()
End Sub
