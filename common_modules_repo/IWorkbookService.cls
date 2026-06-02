VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorkbookService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#Interface

' #############################################################################
'!
'! @brief
'! Excel ファイルのオープン、クローズ、書き込みなどを管理するサービスのインターフェイスです。
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IWorkbookService", Description:="このクラスはインターフェイス用です。"
End Sub

'* ThisWorkbook のブック名を得ます。
'*
'* @return ThisWorkbook のブック名。
'*
'* @details
'* 呼び出し元が ThisWorkbook を直接参照せずに、自ブック名を取得するために使用します。
Public Function GetThisWorkbookName() As String
End Function

'* ThisWorkbook を基準にしたディレクトリ パスを得ます。
'*
'* @return 絶対パスのディレクトリ パス。
'*
'* @details
'* ThisWorkbook が保存済みの場合は、その保存先ディレクトリを返します。
'* ThisWorkbook が未保存の場合は、CurDir の結果を絶対パスにして返します。
Public Function GetThisWorkbookDirectoryPath() As String
End Function

'* ワークブックが存在するかを確認します。
'*
'* @param Book ワークブック名。
'* @return ワークブックが存在するか。
'*
'* @details
'* ワークブック名が存在するかを確認します。
Public Function WorkbookExists(ByVal Book As String) As Boolean
End Function

'* 開いているすべてのワークブックの一覧を得ます。
'*
'* @return ワークブックの一覧。
'*
'* @details
'* ThisWorkbook を含む、開いているすべてのワークブック名の一覧を、文字列の配列で得ます。
Public Function GetAllWorkbooks() As String()
End Function

'* ThisWorkbook 以外のワークブックの一覧を得ます。
'*
'* @return ワークブックの一覧。
'*
'* @details
'* ThisWorkbook を除くワークブック名の一覧を、文字列の配列で得ます。
Public Function GetOtherWorkbooks() As String()
End Function

'* ワークシートが存在するかを確認します。
'*
'* @param Sheet ワークシート名。
'* @param Book ワークブック名。省略した場合は ThisWorkbook 内を確認します。
'* @return ワークシートが存在するか。
'*
'* @details
'* ワークブック内にワークシートが存在するかを確認します。
'* 指定された名前のワークブックがない場合もエラーにはならず False が返ります。
Public Function WorksheetExists(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
End Function

'* ワークシートの一覧を得ます。
'*
'* @param Book ワークブック名。省略した場合は ThisWorkbook 内を確認します。
'* @return ワークシートの一覧。
'*
'* @details
'* 対象ブックの ActiveSheet を含む、ワークブック内のすべてのワークシート名の一覧を、文字列の配列で得ます。
Public Function GetAllWorksheets(Optional ByVal Book As String = "") As String()
End Function

'* 対象ブックの ActiveSheet 以外のワークシートの一覧を得ます。
'*
'* @param Book ワークブック名。省略した場合は ThisWorkbook 内を確認します。
'* @return ワークシートの一覧。
'*
'* @details
'* 対象ブックの ActiveSheet を除くワークシート名の一覧を、文字列の配列で得ます。
Public Function GetOtherWorksheets(Optional ByVal Book As String = "") As String()
End Function

'* ワークブックを開きます。
'*
'* @param FilePath [省略可] ファイルパス。
'* @return ブック名。
'*
'* @details
'* ワークブックを開きます。ファイルパスを省略すると新規作成となります。
'* 元ファイルはロックしません。
Public Function OpenWorkbook(Optional ByVal FilePath As String = "") As String
End Function

'* ワークブックを保存します。
'*
'* @param Book 対象のブック名。
'* @param FilePath 保存先ファイル パス。
'* @param Force [省略可] 強制的に保存するか否か。
'* @return 保存後のブック名。
'*
'* @details
'* ワークブックを保存します。Force が False のときに、すでにファイルが存在している場合、確認メッセージが表示されます。
Public Function SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
End Function

'* ワークブックが最終保存後に変更されていないか確認します。
'*
'* @param Book 対象のブック名。
'* @return 最終保存後に変更されていないか否か。
'*
'* @details
'* Excel の Workbook.Saved と同じ意味で、保存先パスの有無は判定しません。
Public Function IsSaved(ByVal Book As String) As Boolean
End Function

'* ワークブックが保存先パスを持つか確認します。
'*
'* @param Book 対象のブック名。
'* @return 保存先パスを持つか否か。
'*
'* @details
'* 新規未保存ブックのように保存先パスを持たない場合は False を返します。変更有無は判定しません。
Public Function HasPath(ByVal Book As String) As Boolean
End Function

'* ワークブックを閉じます。
'*
'* @param Book 対象のブック名。
'* @param Force [省略可] 強制的に閉じるか否か。True の場合は確認なしで変更を破棄して閉じます。
'*
'* @details
'* ワークブックを閉じます。Force が False のときには、確認メッセージが表示されます。
Public Sub CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
End Sub

'* ワークブック全体を検索します
'*
'* @param What 検索する文字列。
'* @param Book [省略可] 検索対象のワークブック。省略された場合は、ThisWorkbook.Name が指定されたものとします。
'* @param LookInValues [省略可] 値を検索対象とするか。デフォルトは True であり、値を検索対象とします。False にしたときは数式を検索対象とします。
'* @param LookAtWhole [省略可] セル内容が完全に同一であるものを検索するか。デフォルトは True であり、完全一致で検索します。False としたときは部分一致で検索します。
'* @param MatchCase [省略可] 大文字と小文字を区別するか。デフォルトは True です。True にした場合は区別して検索します。False としたときは区別せずに検索します。
'* @param MatchByte [省略可] 全角と半角を区別するか。デフォルトは True です。True にした場合は区別して検索します。False としたときは区別せずに検索します。
'* @return 見つかったセルの WorksheetRangeBounds を格納した ObjectList。検索結果がない場合は Count = 0 です。
'*
'* @details
'* ワークブック全体を検索します
Public Function Find( _
        ByVal What As String, _
        Optional ByVal Book As String = "", _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList
End Function

'* ワークシートを追加します。
'*
'* @param Sheet [省略可] 追加するワークシート名。省略したときは自動で命名されます。
'* @param Book [省略可] ワークブック名。省略したときは ThisWorkbook のワークブック名が使用されます。
'* @param SheetIndex [省略可] ワークシート番号。省略したときはアクティブなシートが指定されたものとします。
'* @param Before [省略可] 前に追加するか。省略したときは指定されたワークシート番号の後ろに追加されます。
'* @return 追加されたワークシートの名前。
'*
'* @details
'* ワークシートを追加します。処理後は呼び出し前のアクティブブックと対象ブックのアクティブシートを復元します。
Public Function AddWorksheet( _
        Optional ByVal Sheet As String = "", _
        Optional ByVal Book As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String
End Function

'* ワークシートを削除します。
'*
'* @param Sheet [省略可] 削除するワークシート名。省略したときはアクティブなシートが指定されたものとします。
'* @param Book [省略可] ワークブック名。省略したときは ThisWorkbook のワークブック名が使用されます。
'*
'* @details
'* ワークシートを削除します。処理後は呼び出し前のアクティブブックと対象ブックのアクティブシートを可能な範囲で復元します。
Public Sub RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
End Sub

'* ワークシートを複製します。
'*
'* @param SourceWorksheetName [省略可] 複製するワークシート名。省略したときはアクティブなシートが指定されたものとします。
'* @param SourceWorkbookName [省略可] 複製するワークシートがあるワークブック名。省略したときは ThisWorkbook のワークブック名が使用されます。
'* @param DestinationWorksheetName [省略可] 複製されたワークシートにつける名前。省略したときは自動で命名されます。
'* @param DestinationWorkbookName [省略可] 複製先のワークブック名。省略したときは ThisWorkbook のワークブック名が使用されます。
'* @param SheetIndex [省略可] ワークシート番号。省略したときは複製先のワークブックのアクティブなシートが指定されたものとします。
'* @param Before [省略可] 複製したシートを対象の前に追加するか。省略したときは指定されたワークシート番号の後ろに追加されます。
'* @return 追加されたワークシートの名前。
'*
'* @details
'* ワークシートを複製します。
Public Function CopyWorksheet( _
        Optional ByVal SourceWorksheetName As String = "", _
        Optional ByVal SourceWorkbookName As String = "", _
        Optional ByVal DestinationWorksheetName As String = "", _
        Optional ByVal DestinationWorkbookName As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String
End Function

'* ワークシートをアクティブにします。
'*
'* @param Sheet [省略可] アクティブにするワークシート名。
'* @param Book [省略可] ワークブック名。省略したときは ThisWorkbook のワークブック名が使用されます。
'*
'* @details
'* 指定したブックとワークシートをアクティブにします。
Public Sub ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
End Sub

'* VBA コンポーネントを削除します。
'*
'* @param ComponentNames 削除対象のコンポーネント名、またはその配列。
'* @param ComponentType [省略可] 削除対象の VBIDE コンポーネント種別。0 の場合は種類を限定しません。
'* @param Book [省略可] ワークブック名。省略したときは ThisWorkbook のワークブック名が使用されます。
'*
'* @details
'* 対象ブックの VBProject から、指定された名前の VBA コンポーネントを削除します。
Public Sub RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
End Sub
