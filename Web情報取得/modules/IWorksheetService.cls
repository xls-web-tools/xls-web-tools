VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorksheetService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'#Interface

' #############################################################################
'!
'! @brief
'! ワークシートへのアクセスを管理するサービスのインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IWorksheetService", Description:="このクラスはインターフェイス用です。"
End Sub

'* 範囲内を検索します
'*
'* @param What 検索する文字列。空文字列は許容し、Excel の空文字検索と同じ意味で扱います。
'* @param RangeBounds 検索対象の範囲。
'* @param LookInValues [省略可] 値を検索対象とするか。デフォルトは True であり、値を検索対象とします。False にしたときは数式を検索対象とします。
'* @param LookAtWhole [省略可] セル内容が完全に同一であるものを検索するか。デフォルトは True であり、完全一致で検索します。False としたときは部分一致で検索します。
'* @param MatchCase [省略可] 大文字と小文字を区別するか。デフォルトは True です。True にした場合は区別して検索します。False としたときは区別せずに検索します。
'* @param MatchByte [省略可] 全角と半角を区別するか。デフォルトは True です。True にした場合は区別して検索します。False としたときは区別せずに検索します。
'* @return 見つかったセルの WorksheetRangeBounds を格納した ObjectList。
'*
'* @details
'* 範囲内を検索します。検索結果がない場合と空範囲の場合は、Count = 0 の ObjectList を返します。
'* 戻り値の各要素は WorksheetRangeBounds です。
'* 呼び出し後は、Excel UI および後続の Range.Find が共有する検索状態を値検索・部分一致の標準状態へ戻します。
Public Function Find( _
        ByVal What As String, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList
End Function

'* 範囲内をソートします。
'*
'* @param RangeBounds ソート対象の範囲。
'* @param SortKeyAndOrder [可変長] ソートのキーとなる相対列番号と、昇順降順指定の組。列番号は RangeBounds の左端を 1 とし、1 から RangeBounds.ColumnCount までを指定します。並び順は xlAscending または xlDescending です。
'*
'* @details
'* 範囲内をソートします。SortKeyAndOrder は列番号と並び順を 1 組として指定します。
Public Sub Sort(ByVal RangeBounds As WorksheetRangeBounds, ParamArray SortKeyAndOrder() As Variant)
End Sub

'* 範囲をアクティブにします。
'*
'* @details
'* 範囲をアクティブにします。
Public Sub ActivateRange(ByVal RangeBounds As WorksheetRangeBounds)
End Sub

'* セルに文字列を書き込みます。
'*
'* @param RangeBounds 書き込む対象のセル範囲。単一セルでない範囲はエラーです。
'* @param Expression 書き込む文字列。
'* @param NumberFormat 表示形式。(既定値は設定変更なし)
'* @param AsFormula 数式として書き込む場合は True (既定値は False)
'* @param ClearWhenEmpty Expression が空文字列の場合にセルをクリアするか (既定値は True)
'* @param IgnoreEmpty Expression が空文字列の場合に何もしない場合は True (既定値は False)
'* @param TypeConvert 数値や日付など、適切な型に変換して書き込む場合は True (既定値は True)
'*
'* @details
'* 指定されたセルに文字列または数式を設定します。
'* TypeConvert が True の場合、数値文字列は整数として表せる場合は整数、それ以外は小数として扱います。
'* Expression が空文字列の場合、`IgnoreEmpty` や `ClearWhenEmpty` に基づいて動作が変わります。
Public Sub WriteCell( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal Expression As String, _
        Optional ByVal NumberFormat As String = "", _
        Optional ByVal AsFormula As Boolean = False, _
        Optional ByVal ClearWhenEmpty As Boolean = True, _
        Optional ByVal IgnoreEmpty As Boolean = False, _
        Optional ByVal TypeConvert As Boolean = True)
End Sub

'* 範囲に値を書き込みます。
'*
'* @param RangeBounds 書き込む対象の範囲情報。空範囲または 1 セルの場合は左上アンカーとして扱い、ValuesArray と同じサイズに拡張します。複数セル範囲の場合は ValuesArray と同じ行数・列数である必要があります。
'* @param ValuesArray 書き込む 2 次元配列。
'*
'* @details
'* 指定された範囲に 2 次元配列を書き込みます。宛先範囲は空範囲、1 セル、または ValuesArray と同形の範囲だけを受け付けます。
Public Sub WriteRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal ValuesArray As Variant)
End Sub

'* 範囲に配列 (CSE) 数式を書き込みます。
'*
'* @param RangeBounds 書き込む対象の範囲情報。
'* @param FormulaArray 配列 (CSE) 数式。
'*
'* @details
'* 範囲に配列 (CSE) 数式を書き込みます。
Public Sub WriteArrayFormula(ByVal RangeBounds As WorksheetRangeBounds, ByVal FormulaArray As String)
End Sub

'* セルから値を読み取ります。
'*
'* @param RangeBounds 読み取る対象のセル範囲。空範囲は空文字列を返し、単一セルでない範囲はエラーです。
'* @param Expression [出力][省略可] 読み取った文字列。戻り値と同じです。空範囲では空文字列です。
'* @param NumberFormat [出力][省略可] 表示形式 (NumberFormatLocal)。空範囲では空文字列です。
'* @param GetText [省略可] Value ではなく Text を読み取るか否か。
'* @return 読み取った文字列。Expression と同じです。
'*
'* @details
'* セルから値を読み取ります。空範囲では Expression と NumberFormat を空文字列にして空文字列を返します。
Public Function ReadCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByRef Expression As String = "", Optional ByRef NumberFormat As String = "", Optional ByVal GetText As Boolean = False) As String
End Function

'* 範囲から値を読み取ります。
'*
'* @param RangeBounds 読み取る対象の範囲情報。
'* @return 読み取った配列。
'*
'* @details
'* 指定された範囲から 2 次元配列を読み取ります。
'* 内部的には Range.Value から = で読み取るため、細かな仕様は Excel に準じます。
'* Excel の仕様と異なる点として、セル範囲を指定しても、2 次元配列にして返します。これにより、処理の一元化が可能となります。
Public Function ReadRange(ByVal RangeBounds As WorksheetRangeBounds) As Variant()
End Function

'* 数式を評価します。
'*
'* @param ContextRangeBounds 数式を評価するシートを示す範囲情報。
'* @param Formula 評価する数式。
'* @return 評価結果。
'*
'* @details
'* Formula を ContextRangeBounds のシート上で評価し、結果を返します。
Public Function EvaluateFormula(ByVal ContextRangeBounds As WorksheetRangeBounds, ByVal Formula As String) As Variant
End Function

'* XLOOKUP の結果を取得します。
'*
'* @param LookupValue 検索値。
'* @param LookupRangeBounds 検索範囲。
'* @param ReturnRangeBounds 戻り値範囲。
'* @param IfNotFound [省略可] 見つからない場合に返す値。
'* @param MatchMode [省略可] 一致モード。
'* @param SearchMode [省略可] 検索モード。
'* @return 検索結果。
'*
'* @details
'* Excel の XLOOKUP 関数を指定範囲に対して評価し、結果を返します。
Public Function XLookup( _
        ByVal LookupValue As Variant, _
        ByVal LookupRangeBounds As WorksheetRangeBounds, _
        ByVal ReturnRangeBounds As WorksheetRangeBounds, _
        Optional ByVal IfNotFound As Variant = "", _
        Optional ByVal MatchMode As Long = 0, _
        Optional ByVal SearchMode As Long = 1) As Variant
End Function

'* セルをコピーします。
'*
'* @param SourceRangeBounds コピー元のセル範囲。単一セルでない範囲はエラーです。
'* @param DestinationRangeBounds コピー先のセル範囲。単一セルでない範囲はエラーです。
'* @param AsValue [省略可] 値を文字列としてコピーするか否か。(既定値は False)
'* @param CopyNumberFormat [省略可] 表示形式をコピーするか否か。(既定値は False)
'*
'* @details
'* 指定された単一セルの値または数式をコピーします。
'* 数式のコピーでは配列数式にも対応します。
Public Sub CopyCell( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
End Sub

'* 範囲をコピーします。
'*
'* @param SourceRangeBounds コピー元の範囲。
'* @param DestinationRangeBounds コピー先の範囲。空範囲または 1 セルの場合は左上アンカーとして扱い、コピー元と同じサイズに拡張します。複数セル範囲の場合はコピー元と同じ行数・列数である必要があります。
'* @param AsValue [省略可] 値を文字列としてコピーするか否か。(既定値は False)
'* @param CopyNumberFormat [省略可] 表示形式をコピーするか否か。(既定値は False)
'*
'* @details
'* 指定された範囲の値または数式をコピーします。
Public Sub CopyRange( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
End Sub

'* セルが空かどうかを調べます。
'*
'* @param RangeBounds 対象となるセルまたは範囲。
'* @param IgnoreEmptyString [省略可]空白文字列を無視するか否か。デフォルト値は False。
'* @return セルが空かどうか。
'*
'* @details
'* セルが空かどうかを調べます。IgnoreEmptyString を True に指定した場合、数式が入っていても計算結果が空文字列であれば、空であると判定されます。
Public Function IsEmptyCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
End Function

'* セルに数式が入力されているかを調べます。
'*
'* @param RangeBounds 対象となるセル。
'* @return 数式が入力されているか。
'*
'* @details
'* 指定されたセルに数式が入力されているかを返します。範囲がセルでない場合は実装側でエラーとします。
Public Function HasFormula(ByVal RangeBounds As WorksheetRangeBounds) As Boolean
End Function

'* 実際に使用されている範囲を取得します。
'*
'* @param RangeBounds 使用範囲を取得する対象の範囲。
'* @param IgnoreEmpty [省略可] 空文字列を無視するか否か。デフォルト値は True。
'* @param IsUsedStartRow [省略可] RangeBounds の開始行を使用済みと見なすか否か。デフォルト値は True。
'* @param IsUsedStartColumn [省略可] RangeBounds の開始列を使用済みと見なすか否か。デフォルト値は True。
'* @param GetRawRange [省略可] 値による切り詰めを行わず、RangeBounds と UsedRange の重なりを基準にした範囲を返すか否か。
'* @return 実際に使用されている範囲
'*
'* @details
'* 指定した範囲内の、実際に使用されている (値が入っている) 範囲を返します。
'* 書式のみが設定された、値の入っていないセルは未使用と見なされます。
'* IgnoreEmpty が True の場合、数式が存在しても、空文字列を結果とするセルは無視されます。
'* GetRawRange が True の場合、値による空行・空列の切り詰めを行わず、Excel の UsedRange と RangeBounds の重なりを基準にした範囲を返します。
'* IsUsedStartRow / IsUsedStartColumn が True の場合、戻り値の開始行 / 開始列は RangeBounds の開始行 / 開始列を保持します。
Public Function GetUsedRangeBounds( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal IgnoreEmpty As Boolean = True, _
        Optional ByVal IsUsedStartRow As Boolean = True, _
        Optional ByVal IsUsedStartColumn As Boolean = True, _
        Optional ByVal GetRawRange As Boolean = False) As WorksheetRangeBounds
End Function

'* 範囲を指定してクリアします。
'*
'* @param RangeBounds クリアする対象の範囲情報。
'* @param ClearAll [省略可] すべてクリアします。このオプションを True にすると、他のオプションの False 指定は無視されます。
'* @param ClearContents [省略可] セルの内容をクリアします。デフォルトは True です。
'* @param ClearNumberFormats [省略可] セルの表示書式をクリアします。
'* @param ClearColors [省略可] セルの文字色および背景色をクリアします。
'* @param ClearComments [省略可] コメントをクリアします。
'* @param ClearHyperlinks [省略可] ハイパーリンクをクリアします。
'*
'* @details
'* 範囲を指定してクリアします。
Public Sub ClearRange( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal ClearAll As Boolean = False, _
        Optional ByVal ClearContents As Boolean = True, _
        Optional ByVal ClearNumberFormats As Boolean = False, _
        Optional ByVal ClearColors As Boolean = False, _
        Optional ByVal ClearComments As Boolean = False, _
        Optional ByVal ClearHyperlinks As Boolean = False)
End Sub

'* 範囲内の重複行を削除します。
'*
'* @param RangeBounds 重複削除対象の範囲。
'* @param DuplicateColumns 重複判定に使う列番号または列番号配列。RangeBounds の左端を 1 とする相対列番号を指定します。
'*
'* @details
'* Excel の RemoveDuplicates と同じ判定で、指定範囲内の重複行を削除します。DuplicateColumns は 1 から RangeBounds.ColumnCount までの値を 1 つ以上指定します。
Public Sub RemoveDuplicates(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant)
End Sub

'* 行を挿入します。
'*
'* @param RangeBounds 挿入する行範囲。行全体の範囲だけを受け付けます。
'*
'* @details
'* RangeBounds.Row から RangeBounds.FinishRow までのシート行を挿入します。空範囲やセル範囲はエラーにします。
Public Sub InsertRows(ByVal RangeBounds As WorksheetRangeBounds)
End Sub

'* 行を削除します。
'*
'* @param RangeBounds 削除する行範囲。行全体の範囲だけを受け付けます。
'*
'* @details
'* RangeBounds.Row から RangeBounds.FinishRow までのシート行を削除します。空範囲やセル範囲はエラーにします。
Public Sub DeleteRows(ByVal RangeBounds As WorksheetRangeBounds)
End Sub

'* すべてのデータを表示します。
'*
'* @param RangeBounds クリアする対象シートの情報。列や行の範囲情報は無視されます。
'*
'* @details
'* Sheet.ShowAllData、Rows.Hidden = False、Columns.Hidden = False します。
Public Sub SetAllDataVisible(ByVal RangeBounds As WorksheetRangeBounds)
End Sub

'* 表示するアウトライン レベルを設定します。
'*
'* @param RangeBounds アウトライン レベルを設定する対象シートの情報。列や行の範囲情報は無視されます。
'* @param RowLevels [省略可] 行のアウトライン レベル。0 は指定なし、1 から 8 は表示レベルです。
'* @param ColumnLevels [省略可] 列のアウトライン レベル。0 は指定なし、1 から 8 は表示レベルです。
'*
'* @details
'* 構成済みのアウトラインに対して、表示するアウトライン レベルを設定します。RowLevels と ColumnLevels が両方 0 の場合は何もしません。
Public Sub SetSheetOutlineLevel(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal RowLevels As Long = 0, Optional ByVal ColumnLevels As Long = 0)
End Sub

'* シートのタブ色を設定します。
'*
'* @param RangeBounds タブ色を設定する対象シートの情報。列や行の範囲情報は無視されます。
'* @param TabColorIndex [省略可] タブの色インデックス。デフォルトは無色です。
'* @param TabColor [省略可] タブの色の値。RGB(,,) 関数で値を設定します。
'*
'* @details
'* シートのタブ色を設定します。インデックスと値を同時に指定した場合、値指定が優先されます。
Public Sub SetSheetTabColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal TabColorIndex As Long = xlColorIndexNone, _
        Optional ByVal TabColor As Long = -1)
End Sub

'* 範囲の色を設定します。
'*
'* @param RangeBounds 色を設定する対象シートの情報。
'* @param FontColorIndex [省略可] 文字の色インデックス。デフォルトは無色です。
'* @param FontColor [省略可] 文字の色の値。RGB(,,) 関数で値を設定します。
'* @param InteriorColorIndex [省略可] 背景の色インデックス。デフォルトは無色です。
'* @param InteriorColor [省略可] 背景の色の値。RGB(,,) 関数で値を設定します。
'*
'* @details
'* 範囲の色を設定します。インデックスと値を同時に指定した場合、値指定が優先されます。
Public Sub SetRangeColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal FontColorIndex As Long = 0, _
        Optional ByVal FontColor As Long = -1, _
        Optional ByVal InteriorColorIndex As Long = xlColorIndexNone, _
        Optional ByVal InteriorColor As Long = -1)
End Sub

'* テキストを折り返して表示するかを設定する。
'*
'* @param RangeBounds 設定する対象の範囲情報。
'* @param WrapText テキストを折り返して表示するか。
'*
'* @details
'* テキストを折り返して表示するかを設定する。
Public Sub SetWrapText(ByVal RangeBounds As WorksheetRangeBounds, ByVal WrapText As Boolean)
End Sub

'* テキストを縮小して全体を表示するかを設定する。
'*
'* @param RangeBounds 設定する対象の範囲情報。
'* @param ShrinkToFit テキストを縮小して全体を表示するか。
'*
'* @details
'* テキストを縮小して全体を表示するかを設定する。
Public Sub SetShrinkToFit(ByVal RangeBounds As WorksheetRangeBounds, ByVal ShrinkToFit As Boolean)
End Sub

'* セル内の配置を設定する。
'*
'* @param RangeBounds 設定する対象の範囲情報。
'* @param HorizontalAlignment 水平方向の配置。G_ALIGNMENT_NO_CHANGE の場合は変更しない。
'* @param VerticalAlignment 垂直方向の配置。G_ALIGNMENT_NO_CHANGE の場合は変更しない。
'* @param Orientation 文字の方向。G_ALIGNMENT_NO_CHANGE の場合は変更しない。
'* @param IndentLevel インデント レベル。G_ALIGNMENT_NO_CHANGE の場合は変更しない。
'*
'* @details
'* セル内の配置を設定する。
'* G_ALIGNMENT_NO_CHANGE を指定した項目は既存設定を変更しない。
Public Sub SetAlignment( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal HorizontalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal VerticalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal Orientation As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal IndentLevel As Long = G_ALIGNMENT_NO_CHANGE)
End Sub
