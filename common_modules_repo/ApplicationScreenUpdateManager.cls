VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationScreenUpdateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Excel の実行環境設定を一時退避する one-shot のスコープ管理クラスです。
'! インスタンス生成時点で ScreenUpdating、EnableEvents、DisplayAlerts、Calculation をバックアップし、Restore でその時点へ戻します。
'! DisableUpdates は退避済みの設定を変更する補助であり、DisableUpdates を呼ばない場合でも Restore は有効です。
'! Restore 後の同一インスタンスは再利用できません。再度設定を退避する場合は、新しいインスタンスを生成します。
'! ネストして使う場合は、生成と逆順の LIFO 順で Restore します。順序違いの復元は呼び出し側の誤りとして扱います。
'! StatusBar は ProgressStatus の責務であり、このクラスでは退避・復元しません。
'! インスタンス破棄時には保険として Restore を試行しますが、破棄時の復元エラーは握りつぶします。
'! Esc キー入力による中断や End ステートメントによる中断など、インスタンスが破棄されない場合は、Restore メソッドを明示的に呼び出す必要があります。
'!
' #############################################################################

Private pSavedScreenUpdating As Boolean
Private pSavedEnableEvents As Boolean
Private pSavedDisplayAlerts As Boolean
Private pSavedCalculation As Integer
Private pIsBackedUp As Boolean

'* このオブジェクトに記録された ScreenUpdating の値。
Public Property Get SavedScreenUpdating() As Boolean
    Call pEnsureBackupExists
    SavedScreenUpdating = pSavedScreenUpdating
End Property

'* このオブジェクトに記録された EnableEvents の値。
Public Property Get SavedEnableEvents() As Boolean
    Call pEnsureBackupExists
    SavedEnableEvents = pSavedEnableEvents
End Property

'* このオブジェクトに記録された DisplayAlerts の値。
Public Property Get SavedDisplayAlerts() As Boolean
    Call pEnsureBackupExists
    SavedDisplayAlerts = pSavedDisplayAlerts
End Property

'* このオブジェクトに記録された Calculation の値。
Public Property Get SavedCalculation() As Integer
    Call pEnsureBackupExists
    SavedCalculation = pSavedCalculation
End Property

Private Sub Class_Initialize()
    Call pBackup
End Sub

' このオブジェクトに、各種設定をバックアップします。
'
' @details
' 現在の Excel の設定 (ScreenUpdating、EnableEvents、DisplayAlerts、Calculation) を記録し、
' Restore メソッドで元の設定に戻せるようにします。
Private Sub pBackup()
    pSavedScreenUpdating = Application.ScreenUpdating
    pSavedEnableEvents = Application.EnableEvents
    pSavedDisplayAlerts = Application.DisplayAlerts
    pSavedCalculation = Application.Calculation
    pIsBackedUp = True
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Call Restore
    On Error GoTo 0
End Sub

'* 画面アップデートやイベントのトリガーなどを停止します。
'*
'* @param StopScreenUpdating 画面の更新を停止するかどうか (デフォルトは True)
'* @param StopEvents イベントのトリガーを停止するかどうか (デフォルトは True)
'* @param StopDisplayAlerts 警告ダイアログの表示を停止するかどうか (デフォルトは True)
'* @param StopCalculation 計算を手動モードにするかどうか (デフォルトは True)
'*
'* @details
'* インスタンス生成時点の退避内容は変更せず、指定された Excel 設定だけを作業用に変更します。
'* Restore 後の同一インスタンスでは使用できません。
Public Sub DisableUpdates( _
            Optional ByVal StopScreenUpdating As Boolean = True, _
            Optional ByVal StopEvents As Boolean = True, _
            Optional ByVal StopDisplayAlerts As Boolean = True, _
            Optional ByVal StopCalculation As Boolean = True)
    Call pEnsureBackupExists
    If StopScreenUpdating Then Application.ScreenUpdating = False
    If StopEvents Then Application.EnableEvents = False
    If StopDisplayAlerts Then Application.DisplayAlerts = False
    If StopCalculation Then Application.Calculation = xlCalculationManual
End Sub

'* このオブジェクトに保存された設定を明示的にリストアします。
'*
'* @details
'* インスタンス生成時点で保存された設定を Excel に復元します。
'* DisableUpdates を呼んでいない場合でも、生成後に変わった設定を生成時点へ戻します。
'* 復元後は退避状態を破棄するため、同じインスタンスの再利用や二重復元はエラーになります。
Public Sub Restore()
    Call pEnsureBackupExists
    Application.ScreenUpdating = pSavedScreenUpdating
    Application.EnableEvents = pSavedEnableEvents
    Application.DisplayAlerts = pSavedDisplayAlerts
    Application.Calculation = pSavedCalculation
    pIsBackedUp = False
End Sub

Private Sub pEnsureBackupExists()
    If Not pIsBackedUp Then
        Err.Raise vbObjectError + 1, "Class ApplicationScreenUpdateManager", "バックアップされた設定がありません。"
    End If
End Sub
