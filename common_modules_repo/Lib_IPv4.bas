Attribute VB_Name = "Lib_IPv4"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! IPv4 関連の関数などを集めた標準モジュールです。
'! 他のツールとも共用されるため、このツールで使用しないものも含まれます。
'!
' #############################################################################

Private Const C_OCTET_SEP_RE As String = "\."
Private Const C_OCTET_SEP As String = "."
Private Const C_MASK_SEP As String = "/"
Private Const C_MASK_SEP_RE As String = "[/_]"

'* IPv4 アドレス文字列の 1 オクテットを表す正規表現
'*
'* @note 10 進表現のみです。
Public Const G_IPV4_OCTET_RE As String = "(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])"

'* IPv4 マスク文字列の 1 オクテットを表す正規表現
'*
'* @note 10 進表現のみです。
Public Const G_IPV4_MOCTET_RE As String = "(0|128|192|224|240|248|252|254|255)"

'* IPv4 アドレス文字列のマスク長を表す正規表現
Public Const G_IPV4_MASKLEN_RE As String = "(3[0-2]|[1-2][0-9]|[0-9])"

'* IPv4 アドレス文字列全体を表す正規表現 (ex. 192.168.1.2)
'*
'* @details
'* 「(第 1)\.(第 2)\.(第 3)\.(第4)」となっています。サブマッチを使うときの参考にしてください。
Public Const G_IPV4_ADDR_RE As String = G_IPV4_OCTET_RE & C_OCTET_SEP_RE & G_IPV4_OCTET_RE & C_OCTET_SEP_RE & G_IPV4_OCTET_RE & C_OCTET_SEP_RE & G_IPV4_OCTET_RE

'* IPv4 マスク文字列全体を表す正規表現 (ex. 255.255.255.0)
'*
'* @details
'* 「((第 1 が 0 ～ 254、あとは 0)|(最初は 255、第 2 が 0 ～ 254、あとは 0)|(255 が連なり、第 3 が 0 ～ 254、あとは 0)|(255 が連なり最後が 0 ～ 255))」
'* となっています｡サブマッチを使うときの参考にしてください｡
Public Const G_IPV4_MASK_RE As String = "(" & _
        "(0|128|192|224|240|248|252|254)" & C_OCTET_SEP_RE & "(0)" & C_OCTET_SEP_RE & "(0)" & C_OCTET_SEP_RE & "(0)" & "|" & _
        "(255)" & C_OCTET_SEP_RE & "(0|128|192|224|240|248|252|254)" & C_OCTET_SEP_RE & "(0)" & C_OCTET_SEP_RE & "(0)" & "|" & _
        "(255)" & C_OCTET_SEP_RE & "(255)" & C_OCTET_SEP_RE & "(0|128|192|224|240|248|252|254)" & C_OCTET_SEP_RE & "(0)" & "|" & _
        "(255)" & C_OCTET_SEP_RE & "(255)" & C_OCTET_SEP_RE & "(255)" & C_OCTET_SEP_RE & "(0|128|192|224|240|248|252|254|255)" & _
        ")"

'* IPv4 ネットワーク アドレス文字列を表す正規表現 (ex. 192.168.1.0/255.255.255.0 や 192.168.1.0/24)
'*
'* @details
'* 「G_IPV4_ADDR_RE & "(|" & C_MASK_SEP_RE & G_IPV4_MASK_RE & "|" & C_MASK_SEP_RE & G_IPV4_MASKLEN_RE & ")"」
'* となっています｡サブマッチを使うときの参考にしてください｡
Public Const G_IPV4_NW_RE As String = G_IPV4_ADDR_RE & "(|" & C_MASK_SEP_RE & G_IPV4_MASK_RE & "|" & C_MASK_SEP_RE & G_IPV4_MASKLEN_RE & ")"

'* より広い (マスク長が 1 短い) ネットワークアドレスを取得します。
'*
'* @param NetworkAddressValue ネットワークアドレスの値
'* @param MaskLength 現在のマスク長
'* @return より広いネットワークアドレスの値
'*
'* @details
'* 指定されたネットワークアドレスとマスク長を基に、マスク長が 1 短いネットワークアドレスを計算します。
'* マスク長が 0 の場合や、指定されたアドレスがネットワークアドレスでない場合はエラーとなります。
Public Function ExpandNetwork(ByVal NetworkAddressValue As Long, ByVal MaskLength As Integer) As Long
    Dim result_value As Long
    Dim mask_value As Long

    If MaskLength = 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function ExpandNetwork", Description:="これ以上広げられません。(" & ConvertToIpAddress(NetworkAddressValue) & "/" & MaskLength & ")"
        Exit Function
    End If

    mask_value = ConvertFromMaskLength(MaskLength)
    If Not IsNetwork(NetworkAddressValue, mask_value) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function ExpandNetwork", Description:="ネットワーク アドレスではありません。(" & ConvertToIpAddress(NetworkAddressValue) & "/" & MaskLength & ")"
        Exit Function
    End If

    mask_value = ConvertFromMaskLength(MaskLength - 1)
    result_value = GetNetworkAddress(NetworkAddressValue, mask_value)

    ExpandNetwork = result_value
End Function

'* より狭い (マスク長が 1 長い) ネットワークアドレス (サブネット) を取得します。
'*
'* @param NetworkAddressValue ネットワークアドレスの値
'* @param MaskLength 現在のマスク長
'* @return 2 つのサブネットアドレスの配列
'*
'* @details
'* 指定されたネットワークアドレスとマスク長を基に、マスク長が 1 長い 2 つのサブネットアドレスを計算します。
'* マスク長が 32 に達している場合や、指定されたアドレスがネットワークアドレスでない場合はエラーとなります。
Public Function NarrowNetwork(ByVal NetworkAddressValue As Long, ByVal MaskLength As Integer) As Long()
    Dim result_value() As Long
    ReDim result_value(0 To 1) As Long
    Dim mask_value As Long

    If 29 < MaskLength And MaskLength < 33 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function NarrowNetwork", Description:="これ以上狭くできません。(" & ConvertToIpAddress(NetworkAddressValue) & "/" & MaskLength & ")"
        Exit Function
    End If

    mask_value = ConvertFromMaskLength(MaskLength)
    If Not IsNetwork(NetworkAddressValue, mask_value) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function NarrowNetwork", Description:="ネットワーク アドレスではありません。(" & ConvertToIpAddress(NetworkAddressValue) & "/" & MaskLength & ")"
        Exit Function
    End If

    result_value(0) = NetworkAddressValue
    result_value(1) = NetworkAddressValue Or GetMaskValue(MaskLength)

    NarrowNetwork = result_value
End Function

'* IP アドレスまたはネットワークアドレスを適切な形式に整形します。
'*
'* @param AddressLike 整形する対象の文字列。IP アドレス形式を想定しています。
'* @return 整形後のアドレス文字列。
'*
'* @details
'* この関数は、入力されたアドレス文字列の特定の文字（アンダースコア）をスラッシュに置換し、
'* 標準的な CIDR 表記に整形します。入力されたアドレスがネットワークマスクを含まない場合、
'* "/32" が末尾に追加されます。これにより、個々の IP アドレスを表す形式になります。
'*
'* 具体的には、入力された文字列内の "_" を "/" に置換し、さらに入力形式が単一の IP アドレスの場合には
'* 末尾に "/32" を追加して、CIDR ノーテーションに準拠した形式に整形します。ネットワークマスクが既に
'* 指定されている場合は、そのままの形式を返します。
Public Function WellFormedAddress(ByVal AddressLike As String) As String
    Dim result_value As String

    result_value = Replace(AddressLike, "_", "/")

    If result_value <> "" And result_value Like "?*.?*.?*.?*" And Not result_value Like "?*.?*.?*.?*/?*" Then
        result_value = result_value & "/32"
    End If

    WellFormedAddress = result_value
End Function

'* 「IP アドレス/マスク」という形式の文字列を解析し、値に変換します。
'*
'* @param IpAddressValue [出力] 解析された IP アドレスの値
'* @param MaskValue [出力] 解析されたマスクの値
'* @param MaskLength [出力] 解析されたマスク長
'* @param IpAddressAndMask 解析する「IP アドレス/マスク」形式の文字列
'*
'* @details
'* 「IP アドレス/マスク」という形式の文字列を解析し、IP アドレス値、マスク値、およびマスク長に変換します。
'* マスク部分は、CIDR 表記とドットデシマル形式の両方を受け付けます。
'* 不正な形式の文字列が渡された場合はエラーが発生します。
Public Sub ParseIpAddressAndMask(ByRef IpAddressValue As Long, ByRef MaskValue As Long, ByRef MaskLength As Integer, ByVal IpAddressAndMask As String)
    Dim ip_arr() As String
    Dim ip_addr_value As Long
    Dim mask_value As Long
    Dim mask_length As Integer

    ip_arr = Split(IpAddressAndMask, C_MASK_SEP)

    If 1 < UBound(ip_arr) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Sub ParseIpAddressAndMask", Description:="正しい形式ではありません。(" & IpAddressAndMask & ")"
    End If

    ip_addr_value = ConvertFromIpAddress(ip_arr(0))

    'If IsIpAddress(ip_arr(1)) Then
    If Not IsInteger(ip_arr(1)) Then
        mask_value = ConvertFromIpAddress(ip_arr(1))
        If Not IsValidMaskValue(mask_value) Then
            Err.Raise Number:=vbObjectError + 1, Source:="Sub ParseIpAddressAndMask", Description:="正しい形式ではありません。(" & IpAddressAndMask & ")"
            Exit Sub
        End If
        mask_length = ConvertToMaskLength(mask_value)
    Else
        mask_length = CInt(ip_arr(1))
        mask_value = ConvertFromMaskLength(mask_length)
    End If

    IpAddressValue = ip_addr_value
    MaskValue = mask_value
    MaskLength = mask_length
End Sub

'* 与えられた IP アドレス値が、ネットワークアドレスかどうかを確認します。
'*
'* @param IpAddressValue 確認する IP アドレスの値
'* @param MaskValue 確認に使用するマスクの値
'* @return ネットワークアドレスであれば True、それ以外の場合は False
'*
'* @details
'* 指定された IP アドレス値とマスク値を基に、ネットワークアドレスであるかを確認します。
Public Function IsNetwork(ByVal IpAddressValue As Long, ByVal MaskValue As Long) As Boolean
    Dim result_value As Boolean
    Dim host_addr As Long

    host_addr = GetHostAddress(IpAddressValue, MaskValue)

    If host_addr = 0 Then
        IsNetwork = True
    Else
        IsNetwork = False
    End If
End Function

''
'' IP アドレスの形式を満たしているか (ドットデシマル形式になっているか) を確認します。
''
'Function IsIpAddress(IpAddress As String) As Long
'    Dim regex_obj As New RegExp
'    Dim match_collection As MatchCollection
'
'    regex_obj.Pattern = "^" & IPV4_OCTET_RE & OCTET_SEP_RE & IPV4_OCTET_RE & OCTET_SEP_RE & IPV4_OCTET_RE & OCTET_SEP_RE & IPV4_OCTET_RE & "$"
'    regex_obj.Global = False
'
'    Set match_collection = regex_obj.Execute(IpAddress)
'    If 0 < match_collection.count Then
'        IsIpAddress = True
'    Else
'        IsIpAddress = False
'    End If
'End Function

'* マスクが先頭から連続して 1 かを確認します。
'*
'* @param MaskValue 確認するマスク値
'* @return マスクが先頭から連続して 1 であれば True、それ以外は False
'*
'* @details
'* 指定されたマスク値が、先頭から連続して 1 となっている有効なマスクであるかを確認します。
'* 無効なマスク (途中に 0 を含む場合) では False を返します。
Public Function IsValidMaskValue(ByVal MaskValue As Long) As Boolean
    Dim found_zero As Boolean
    Dim item_idx As Integer

    If MaskValue > 0 Then
        ' 正の数は、先頭が 0 で、それ以降に 1 がある。
        IsValidMaskValue = False
        Exit Function
    End If

    If MaskValue = 0 Then
        ' 桁があふれるので固定値で返す
        ' すべて 0
        IsValidMaskValue = True
        Exit Function
    End If

    ' 32ビットすべてをチェック
    For item_idx = 30 To 0 Step -1
        If (MaskValue And (2& ^ item_idx)) <> 0 Then
            ' ビットが立っていたら
            If found_zero Then
                ' すでに 0 が見つかっていたら
                IsValidMaskValue = False
                Exit Function
            End If
        Else
            ' 初めて 0 が見つかった
            found_zero = True
        End If
    Next item_idx

    ' すべて 1
    IsValidMaskValue = True
End Function

'* IP アドレスのホスト部を取得します。
'*
'* @param IpAddressValue 対象の IP アドレス値
'* @param MaskValue 使用するネットマスク値
'* @return ホスト部を示す値
'*
'* @details
'* 指定された IP アドレス値にネットマスク値を適用して、ホスト部を抽出します。
Public Function GetHostAddress(ByVal IpAddressValue As Long, ByVal MaskValue As Long) As Long
    GetHostAddress = IpAddressValue And Not MaskValue
End Function

'* IP アドレスのネットワーク部を取得します。
'*
'* @param IpAddressValue 対象の IP アドレス値
'* @param MaskValue 使用するネットマスク値
'* @return ネットワーク部を示す値
'*
'* @details
'* 指定された IP アドレス値にネットマスク値を適用して、ネットワーク部を抽出します。
Public Function GetNetworkAddress(ByVal IpAddressValue As Long, ByVal MaskValue As Long) As Long
    GetNetworkAddress = IpAddressValue And MaskValue
End Function

'* IP アドレス文字列を IP アドレス値に変換します。
'*
'* @param IpAddress 変換対象の IP アドレス文字列
'* @return 対応する IP アドレス値
'*
'* @details
'* ドットデシマル形式の IP アドレス文字列を、Long 型の IP アドレス値に変換します。
'* 不正な形式の場合、エラーが発生します。
Public Function ConvertFromIpAddress(ByVal IpAddress As String) As Long
    Dim regex_obj As RegExp
    Dim match_collection As MatchCollection
    Dim match_item As Match
    Dim item_idx As Integer
    Dim ipv4_re_obj As String
    Dim octet_value As Long
    Dim high_bit As Long
    Dim result_value As Long

    Set regex_obj = New RegExp

    regex_obj.Pattern = "^" & G_IPV4_ADDR_RE & "$"
    regex_obj.Global = False

    Set match_collection = regex_obj.Execute(IpAddress)

    If 0 < match_collection.Count Then
        Set match_item = match_collection(0)
        If match_item.SubMatches.Count = 4 Then
            octet_value = CInt(match_item.SubMatches(0))

            ' 桁あふれを防ぐ処置
            If octet_value < 128 Then
                high_bit = 0
            Else
                octet_value = &H7F And octet_value
                high_bit = &H80000000
            End If

            result_value = high_bit Or (octet_value * 16777216 + CLng(match_item.SubMatches(1)) * 65536 + CLng(match_item.SubMatches(2)) * 256 + CLng(match_item.SubMatches(3)))

            ' 正常終了したら抜ける。
            ConvertFromIpAddress = result_value
            Exit Function
        End If
    End If

    ' 正常終了しなかったらエラー
    Err.Raise Number:=vbObjectError + 1, Source:="Function ConvertFromIpAddress", Description:="アドレスの形式が不正です。(" & IpAddress & ")"
End Function

'* IP アドレス値を IP アドレス文字列に変換します。
'*
'* @param IpAddressValue 変換対象の IP アドレス値
'* @return 対応する IP アドレス文字列
'*
'* @details
'* 指定された Long 型の IP アドレス値を、ドットデシマル形式の文字列に変換します。
Public Function ConvertToIpAddress(ByVal IpAddressValue As Long) As String
    Dim result_value As String
    Dim ip_value As Long
    Dim mask_value As Long
    Dim high_bit As Long

    ' 下位 8 ビットだけ抜き出すマスクを作成する
    mask_value = GetMaskValue(24, 32)
    'mask_value = BitRight(&HFFFFFFFF, 24)

    ' 最上位ビットだけ別処理
    If 0 <= IpAddressValue Then
        ip_value = IpAddressValue
        high_bit = 0
    Else
        ip_value = &H7FFFFFFF And IpAddressValue
        high_bit = &H80
    End If

    ' 第 4 オクテット
    result_value = ip_value And mask_value
    ip_value = ip_value \ 256

    ' 第 3 オクテット
    result_value = "" & (ip_value And mask_value) & C_OCTET_SEP & result_value
    ip_value = ip_value \ 256

    ' 第 2 オクテット
    result_value = "" & (ip_value And mask_value) & C_OCTET_SEP & result_value
    ip_value = ip_value \ 256

    ' 第 1 オクテット
    result_value = "" & (high_bit Or ip_value) & C_OCTET_SEP & result_value

    ConvertToIpAddress = result_value
End Function

'* マスク値を反転します。
'*
'* @param MaskValue 反転対象のマスク値
'* @return 反転後のマスク値
'*
'* @details
'* 指定されたマスク値をビット単位で反転します。
Public Function InvertMaskValue(ByVal MaskValue As Long) As Long
    InvertMaskValue = Not MaskValue
End Function

'* マスク長からマスク値に変換します。
'*
'* @param MaskLength 変換対象のマスク長 (0 ～ 32 の範囲内)
'* @return 対応するマスク値
'*
'* @details
'* 指定されたマスク長を 32 ビットのマスク値に変換します。
'* 無効なマスク長 (範囲外) の場合はエラーを発生させます。
Public Function ConvertFromMaskLength(ByVal MaskLength As Integer) As Long
    If MaskLength < 0 Or 32 < MaskLength Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function ConvertFromMaskLength", Description:="マスク長が範囲外です。(" & MaskLength & ")"
        Exit Function
    End If

    ConvertFromMaskLength = GetMaskValue(0, MaskLength)
    'ConvertFromMaskLength = BitLeft(&HFFFFFFFF, (32 - MaskLength))
End Function

'* マスク値からマスク長に変換します。
'*
'* @param MaskValue 変換対象のマスク値
'* @return 対応するマスク長
'*
'* @details
'* 指定されたマスク値をマスク長に変換します。
'* 無効なマスク値 (連続した 1 ではない場合) の場合はエラーを発生させます。
Public Function ConvertToMaskLength(ByVal MaskValue As Long) As Integer
    Dim found_zero As Boolean
    Dim mask_length As Integer
    Dim item_idx As Integer

    If MaskValue > 0 Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertToMaskLength", Description:="不正なマスク値です。(" & MaskValue & ")"
        Exit Function
    End If

    If MaskValue = 0 Then
        ' 桁があふれるので固定値で返す
        ConvertToMaskLength = 0
        Exit Function
    End If

    ' 32ビットすべてをチェック
    mask_length = 1
    For item_idx = 30 To 0 Step -1
        If (MaskValue And (2& ^ item_idx)) <> 0 Then
            If found_zero Then
                Err.Raise vbObjectError + 1, Source:="ConvertToMaskLength", Description:="不正なマスク値です。(" & MaskValue & ")"
            End If
            mask_length = mask_length + 1
        Else
            found_zero = True
        End If
    Next item_idx

    ConvertToMaskLength = mask_length
End Function

'* 指定したビット範囲を 1 にしたマスクを作成します。
'*
'* @param Start 開始ビット (0 オリジン、含む)
'* @param Finish 終了ビット (0 オリジン、含まない)。省略時は Start + 1。
'* @return 作成したマスク値
'*
'* @details
'* 開始ビットから終了ビットまでを 1 にし、それ以外を 0 にした 32 ビットのマスク値を作成します。
'* 範囲外のビット指定の場合はエラーを発生させます。
Public Function GetMaskValue(ByVal Start As Integer, Optional ByVal Finish As Integer = -1) As Long
    Dim result_value As Long
    Dim start_idx As Long
    Dim high_bit As Long
    Dim bit_length As Integer

    If Finish = -1 Then
        Finish = Start + 1
    End If

    If Finish < Start Or Start < 0 Or 32 < Finish Then
        Err.Raise vbObjectError + 1, Source:="Function GenerateMask", Description:="マスクの開始位置と終了位置が不正です。(" & Start & ", " & Finish & ")"
    End If


    ' 最上位ビットの処理
    If Start = 0 And Start <> Finish Then
        high_bit = &H80000000
        start_idx = 1
    Else
        high_bit = 0
        start_idx = Start
    End If

    bit_length = Finish - start_idx

    result_value = (2& ^ bit_length - 1) * (2& ^ (32 - Finish)) Or high_bit

    GetMaskValue = result_value
End Function
