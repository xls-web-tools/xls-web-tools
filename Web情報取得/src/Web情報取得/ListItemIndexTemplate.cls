VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListItemIndexTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 一覧項目 index に依存するテンプレート文字列を解決します。
'!
' #############################################################################

'* 一覧項目 index の置換トークンを実値へ置き換えます。
'*
'* @param TemplateText 置換対象の文字列。
'* @param ItemIndex 0-based の一覧項目 index。
'* @return 置換後の文字列。
'*
'* @details
'* {{index}} は 0-based、{{rowNumber}} は 1-based の値に置換します。
Public Function Apply( _
        ByVal TemplateText As String, _
        ByVal ItemIndex As Long) As String

    Dim result_value As String
    result_value = TemplateText
    result_value = Replace(result_value, "{{index}}", CStr(ItemIndex))
    result_value = Replace(result_value, "{{rowNumber}}", CStr(ItemIndex + 1))

    Apply = result_value
End Function
