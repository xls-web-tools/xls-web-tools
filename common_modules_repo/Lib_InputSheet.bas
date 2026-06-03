Attribute VB_Name = "Lib_InputSheet"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! UserInputSheet のファクトリを提供するモジュールです。
'!
' #############################################################################

'* UserInputSheet インスタンスを新規作成します。
'*
'* @param InputArea 入力シート内で項目名と値を読み取る対象範囲。
'* @return 新しい UserInputSheet インスタンス
'*
'* @details
'* UserInputSheet インスタンスを新規作成します。
Public Function New_InputSheet(ByVal InputArea As WorksheetRangeBounds) As UserInputSheet
    Dim result_value As UserInputSheet
    Set result_value = New UserInputSheet
    Call result_value.Initialize(InputArea)
    Set New_InputSheet = result_value
End Function
