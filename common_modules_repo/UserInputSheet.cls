VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserInputSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IUserInputSheet

' #############################################################################
'!
'! @brief
'! ユーザー インターフェイス用シートにアクセスするためのクラスです。
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pRangeBounds As WorksheetRangeBounds

'* 入力シート内で項目名と値を読み取る対象範囲。
Public Property Get InputArea() As WorksheetRangeBounds
    Call pCheckInit
    Set InputArea = pRangeBounds.CopyObject()
End Property
Private Property Get IUserInputSheet_InputArea() As WorksheetRangeBounds
    Set IUserInputSheet_InputArea = InputArea
End Property

'* 各プロパティを設定します。
'*
'* @param InputArea 入力シート内で項目名と値を読み取る対象範囲。
'*
'* @details
'* 各プロパティを設定します。
Public Sub Initialize(ByVal InputArea As WorksheetRangeBounds)
    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class UserInputSheet.Initialize", "すでに初期化済みです。"
    If InputArea Is Nothing Then Err.Raise vbObjectError + 1, "Class UserInputSheet.Initialize", "入力領域が指定されていません。"
    If InputArea.IsEmpty Then Err.Raise vbObjectError + 1, "Class UserInputSheet.Initialize", "入力領域が空範囲です。"
    If Not WbSrv.WorksheetExists(InputArea.WorksheetName, Book:=InputArea.WorkbookName) Then Err.Raise vbObjectError + 1, "Class UserInputSheet.Initialize", "対象シートが見つかりません。(" & InputArea.WorksheetName & "@" & InputArea.WorkbookName & ")"

    Set pRangeBounds = InputArea.CopyObject()
    pIsInitialized = True
End Sub

'* 指定された項目名に対応する項目値の範囲を取得します。
'*
'* @param FirstItemName 第一項目名
'* @param SecondItemName 第二項目名 (省略可能)
'* @return 対応する項目値の範囲 (Range オブジェクト)
'*
'* @details
'* 第一項目名の位置を基準に、行または列の範囲を取得します。
'* 第二項目名が省略されている場合、第一項目名の行全体の範囲を返します。
'* 第二項目名が指定されている場合、その列に対応する範囲を返します。
Public Function GetItemRange( _
        ByVal FirstItemName As String, _
        Optional ByVal SecondItemName As String = "") As WorksheetRangeBounds

    Call pCheckInit

    Dim used_range As WorksheetRangeBounds
    Set used_range = WsSrv.GetUsedRangeBounds(pRangeBounds)

    ' 1 列目を第一項目名の列とし、残りと分割する
    Dim first_col As WorksheetRangeBounds
    Set first_col = used_range.TransformAbsolute(FinishColumn:=used_range.Column)

    Dim remain_range As WorksheetRangeBounds
    Set remain_range = used_range.TransformAbsolute(Column:=used_range.Column + 1)

    ' 第一項目名の列を名前で検索する
    Dim found_cells As ObjectList
    Set found_cells = WsSrv.Find(FirstItemName, first_col)

    ' もし空のリストが返ってきたら、見つからなかったということなので Nothingで終了
    If found_cells.Count = 0 Then
        Set GetItemRange = Nothing
        Exit Function
    End If

    ' 見つかった第一項目の右横の行
    Dim second_row As WorksheetRangeBounds
    Dim first_found_cell As WorksheetRangeBounds
    Set first_found_cell = found_cells.Item(0)

    Set second_row = remain_range.TransformAbsolute(Row:=first_found_cell.Row, FinishRow:=first_found_cell.Row)

    ' 切り詰める
    Set second_row = WsSrv.GetUsedRangeBounds(second_row)

    ' 第二項目名が指定されていないときは、second_row は値
    If SecondItemName = "" Then
        If second_row.IsEmpty Then
            ' 空範囲の時は、1 セルにして返す。
            Set GetItemRange = second_row.TransformAbsolute(FinishRow:=second_row.Row, FinishColumn:=second_row.Column)
        Else
            Set GetItemRange = second_row
        End If
        Exit Function
    End If

    ' 第二項目名を探す。
    Set found_cells = WsSrv.Find(SecondItemName, second_row)

    ' もし空のリストが返ってきたら、見つからなかったということなので Nothing で終了
    If found_cells.Count = 0 Then
        Set GetItemRange = Nothing
        Exit Function
    End If

    ' 次の第一項目を見つけ、第二項目の値の行数を算出
    Dim first_remain_col As WorksheetRangeBounds
    Dim second_found_cell As WorksheetRangeBounds
    Set second_found_cell = found_cells.Item(0)

    Set first_remain_col = first_col.TransformAbsolute(Row:=second_found_cell.Row + 1)
    Dim finish_row_idx As Long
    finish_row_idx = second_found_cell.Row
    Dim enum_obj As IEnumerator
    Set enum_obj = first_remain_col.GetEnumerator()
    Do While enum_obj.MoveNext()
        Dim cell_item As WorksheetRangeBounds
        Set cell_item = enum_obj.Current
        If WsSrv.IsEmptyCell(cell_item, IgnoreEmptyString:=True) Then
            finish_row_idx = finish_row_idx + 1
        Else
            Exit Do
        End If
    Loop
    finish_row_idx = MinLng(finish_row_idx, used_range.FinishRow)

    ' 結果の範囲を生成
    Dim result_value As WorksheetRangeBounds
    Set result_value = remain_range.TransformAbsolute(Row:=second_found_cell.Row + 1, Column:=second_found_cell.Column, FinishRow:=finish_row_idx, FinishColumn:=second_found_cell.Column)
    Set result_value = WsSrv.GetUsedRangeBounds(result_value, IgnoreEmpty:=True)
    Set GetItemRange = result_value
End Function
Private Function IUserInputSheet_GetItemRange( _
        ByVal FirstItemName As String, _
        Optional ByVal SecondItemName As String = "") As WorksheetRangeBounds

    Set IUserInputSheet_GetItemRange = GetItemRange(FirstItemName, SecondItemName:=SecondItemName)
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class UserInputSheet.pCheckInit", "未初期化状態です。"
End Sub
