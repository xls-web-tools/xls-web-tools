VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUserInputSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#Interface

' #############################################################################
'!
'! @brief
'! 入力シートにアクセスするためのインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IUserInputSheet", Description:="このクラスはインターフェイス用です。"
End Sub

'* 入力シート内で項目名と値を読み取る対象範囲。
Public Property Get InputArea() As WorksheetRangeBounds
End Property

'* 指定された項目名に対応する項目値の範囲を取得します。
'*
'* @param FirstItemName 第一項目名。
'* @param SecondItemName 第二項目名。省略時は第一項目名の値範囲を取得します。
'* @return 対応する項目値の範囲。見つからない場合は Nothing。
'*
'* @details
'* 入力領域上の項目見出しに対応する値範囲を返します。
Public Function GetItemRange( _
        ByVal FirstItemName As String, _
        Optional ByVal SecondItemName As String = "") As WorksheetRangeBounds
End Function
