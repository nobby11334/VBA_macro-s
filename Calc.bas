Attribute VB_Name = "Calc"
Private Const STR_REGEXP As String = "VBScript.RegExp"
'Private Declare Function SafeArrayAllocDescriptor Lib "oleaut32" _
'   (ByVal cDims As Long, _
'    ByRef ppsaOut() As Any) As Long


Sub tttt()
    FromBitCountTo255 27
End Sub


'@brief     プレフィックス（/24など）を「255.255.255.0」の表示に変える
'@author    Hiroki Nobumoto
'@date      2010.11.08
'@param     /24などの"24"の部分の10進数の値
'@return    文字列（255.255.255.0等）
Public Function FromBitCountTo255(ByVal tgtval As Long) As String
    Dim i As Long
    Dim counter As Long
    counter = tgtval
    Dim bitstr As String
    bitstr = ""
    For i = 1 To 32 Step 1
        DoEvents
        If counter > 0 Then
            bitstr = bitstr & 1
        Else
            bitstr = bitstr & 0
        End If
        
        If (i Mod 8) = 0 And i < 32 Then
            bitstr = bitstr & "."
        End If
        counter = counter - 1
    Next i
    
    Dim splitres() As String
    splitres = Split(bitstr, ".")
    
    Dim res As String
    res = ""
    Dim imax As Long
    imax = UBound(splitres)
    For i = 0 To imax Step 1
        res = res & CStr(FromBinToDecimal(splitres(i)))
        If i < 3 Then
            res = res & "."
        End If
    Next i
    FromBitCountTo255 = res
    
End Function


'@brief     10進数を2進数（文字列）にする。
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param     tgtVal:ビット値にしたい10進数
'@return    ビット変換した結果（文字列8桁固定）
Public Function FromDecimalToBin(ByVal tgtval As Long) As String
    Dim sp As Long
    Dim md As Long
    Dim strBit As String
    If tgtval < 0 Then
        FromDecimalToBin = "00000000"
        Exit Function
    Else
        sp = tgtval
    End If
    strBit = ""
    
    Do While sp > 0
        DoEvents
        md = sp Mod 2
        sp = sp \ 2
        If md = 0 Then
            strBit = "0" & strBit
        ElseIf md = 1 Then
            strBit = "1" & strBit
        End If
    Loop
    
    
    '桁数8桁に補正
    Dim strlength As Long
    strlength = Len(strBit)
    
    Dim i As Long
    For i = 1 To (8 - strlength) Step 1
        DoEvents
        strBit = "0" & strBit
    Next i
    FromDecimalToBin = strBit
End Function





'@brief     2進数値を10進数にする。
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param     tgtVal：10進数にしたい2進数
'@return    ビット変換した結果（文字列8桁固定）
Public Function FromBinToDecimal(ByVal tgtval As String) As Long

    Dim beki As Long
    Dim dec As Long             '10進数を格納する変数。
    Dim bitval(7) As String     '引数の二進数を一桁に対して1要素へ格納する為の配列
    dec = 0
    
    '配列に11111111のそれぞれの桁の値を分割して右の文字から順番に格納する。
    Dim i As Integer
    Dim j As Integer
    j = 8
    For i = 0 To 7 Step 1
        DoEvents
        bitval(i) = Mid(tgtval, j, 1)
        j = j - 1
    Next i
    
    For i = 0 To 7 Step 1
        DoEvents
        If bitval(i) = "1" Then
            dec = dec + 2 ^ i
        Else
            '--何もしない--
        End If
    Next i
    
    FromBinToDecimal = dec

End Function



'@brief     2進数値のビット数を返す。
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param     tgtVal:ビット数にしたい2進数
'@return    ビット変換した結果（文字列8桁固定）
Public Function FromBinToBit(ByVal tgtval As String) As Integer
    Dim er As Object
    Set er = CreateObject(STR_REGEXP)
    er.Pattern = "(1+)(0+)"
    er.Global = True
    If er.test(tgtval) = True Then
        FromBinToBit = 0
        Exit Function
    End If

    
    Dim beki As Long
    Dim totalbit As Integer     '2進数のビット値（11111111→8ビット）
    Dim bitval(7) As String     '引数の二進数を一桁に対して1要素へ格納する為の配列
    totalbit = 0
    
    '配列に11111111のそれぞれの桁の値を分割して右の文字から順番に格納する。
    Dim i As Integer
    Dim j As Integer
    j = 8
    For i = 0 To 7 Step 1
        DoEvents
        bitval(i) = Mid(tgtval, j, 1)
        j = j - 1
    Next i
    
    For i = 0 To 7 Step 1
        DoEvents
        If bitval(i) = "1" Then
            totalbit = totalbit + 1
        Else
            '--何もしない--
        End If
    Next i
    
    FromBinToBit = totalbit

End Function


'@brief     10進数値を2進数にした時のビット数を返す。
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param     tgtVal:ビット数にしたい10進数
'@return    ビット数（string型）
Public Function FromDecToBit(ByVal tgtval As String) As Integer
    FromDecToBit = FromBinToBit(FromDecimalToBin(tgtval))
End Function


'@brief     引数のIPアドレスとサブネットマスクからをネットワークアドレスを算出して返す。
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param1    IPアドレス          書式:(\d+).(\d+).(\d+).(\d+)
'@param2    サブネットマスク    書式:(\d+).(\d+).(\d+).(\d+)
'@return    ビット数（string型）
Public Function FromIpAddressToNetworkAddress(ByVal strIPaddr As String, ByVal strSubnetmask As String) As String

    If NobbyString.FindRegExp(strIPaddr, REGEXP_IPADDR) = False Then
        FromIpAddressToNetworkAddress = "－"
        Exit Function
    End If
    
    If NobbyString.FindRegExp(strSubnetmask, REGEXP_IPADDR) = False Then
        FromIpAddressToNetworkAddress = "－"
        Exit Function
    End If
    
    Dim splitipaddr() As String
    splitipaddr() = Split(strIPaddr, ".")
    
    Dim splitsubnetmask() As String
    splitsubnetmask() = Split(strSubnetmask, ".")
    
    Dim res As String
    res = CStr(splitipaddr(0) And splitsubnetmask(0)) & "."
    res = res & CStr(splitipaddr(1) And splitsubnetmask(1)) & "."
    res = res & CStr(splitipaddr(2) And splitsubnetmask(2)) & "."
    res = res & CStr(splitipaddr(3) And splitsubnetmask(3))
    FromIpAddressToNetworkAddress = res
End Function


'@brief     引数のサブネットマスクからワイルドカードマスクを算出して返す。
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param2    サブネットマスク    書式:(\d+).(\d+).(\d+).(\d+)
'@return    ビット数（string型）
Public Function FromSubnetmaskToWildcardmask(ByVal strSubnetmask As String) As String

    If NobbyString.FindRegExp(strSubnetmask, REGEXP_IPADDR) = False Then
        FromSubnetmaskToWildcardmask = "－"
        Exit Function
    End If
    
    
    Dim splitsubnetmask() As String
    splitsubnetmask() = Split(strSubnetmask, ".")
    
    Dim res As String
    Dim allBit As String
    allBit = "255"
    res = CStr(allBit Xor splitsubnetmask(0)) & "."
    res = res & CStr(allBit Xor splitsubnetmask(1)) & "."
    res = res & CStr(allBit Xor splitsubnetmask(2)) & "."
    res = res & CStr(allBit Xor splitsubnetmask(3))
    FromSubnetmaskToWildcardmask = res
End Function



'@brief     IPアドレスとサブネットマスク2組が同じネットワークの中にあるかどうかを判定する。
'@author    Hiroki Nobumoto
'@date      2009.10.06
'@param1    比較対象IPアドレスA（アドレスBよりレンジが広い）
'@param2    比較対照IPアドレスAのサブネットマスク（アドレスBよりレンジが広い）
'@param3    比較対象IPアドレスB（アドレスAよりレンジが狭い）
'@param4    比較対照IPアドレスBのサブネットマスク（アドレスAよりレンジが狭い）
'@return    true AがBの範囲にいる。　false AがBの範囲にいない。
Public Function ChkIpAddressRange(ByVal ipaddrA As String, ByVal subnetmaskA As String, ByVal ipaddrB As String, ByVal subnetmaskB As String) As Boolean

    'レンジの深い方がABのどっちか判定
    Dim subnetoqtetsA() As String
    Dim subnetoqtetsB() As String
    subnetoqtetsA = Split(subnetmaskA, ".")
    subnetoqtetsB = Split(subnetmaskB, ".")
    Dim chkwhitchmax As Long
    Dim counter As Long
    chkwhitchmax = UBound(subnetoqtetsA)
    
    Dim netaddrA As String
    Dim netaddrB As String
    netaddrA = FromIpAddressToNetworkAddress(ipaddrA, subnetmaskA)
    netaddrB = FromIpAddressToNetworkAddress(ipaddrB, subnetmaskB)
    
    Dim targetipaddr As String
    Dim targetsubnet As String
    Dim notdeepnetaddr As String
    For counter = 0 To chkwhitchmax Step 1
        DoEvents
        If CLng(subnetoqtetsA(counter)) > CLng(subnetoqtetsB(counter)) Then
            targetipaddr = netaddrA
            targetsubnet = subnetmaskB
            notdeepnetaddr = netaddrB
        ElseIf CLng(subnetoqtetsB(counter)) > CLng(subnetoqtetsA(counter)) Then
            targetipaddr = netaddrB
            targetsubnet = subnetmaskA
            notdeepnetaddr = netaddrA
        Else
            targetipaddr = netaddrB
            targetsubnet = subnetmaskA
            notdeepnetaddr = netaddrA
        End If
    Next counter
    
    Dim targetnetaddr As String
    Dim targetbcastaddr As String
    IpAddressRange targetipaddr, targetsubnet, targetnetaddr, targetbcastaddr
    
    Dim bossnetworkaddr As String
    bossnetworkaddr = FromIpAddressToNetworkAddress(targetipaddr, targetsubnet)
    
    If notdeepnetaddr = bossnetworkaddr Then
        ChkIpAddressRange = True
    Else
        ChkIpAddressRange = False
    End If
    
End Function



'@brief     IPアドレスとサブネットマスクからネットワークアドレスとブロードキャストアドレスを算出する。（ホストアドレスレンジ）
'@author    Hiroki Nobumoto
'@date      2009.10.06
'@param1    範囲を知りたいIPアドレス
'@param2    範囲を知りたいIPアドレスのサブネットマスク
'@param3    出力結果（ネットワークアドレス）
'@param4    出力結果（ブロードキャストアドレス）
'@return
Public Sub IpAddressRange(ByVal ipaddr As String, ByVal subnetMask As String, ByRef netaddr As String, ByRef bcastaddr As String)

    Dim ipoqtets() As String
    oqtets = Split(ipaddr, ".")
    
    Dim maskoqtets() As String
    maskoqtets = Split(subnetMask, ".")
    
    
    'サブネットマスクからホストアドレスサブネットのＩＰアドレスのオクテット番号とホストアドレス範囲を算出する。
    Dim maskoqtetnum As Long
    Dim maskoqtetnummax As Long
    Dim maskrange As Long
    maskoqtetnummax = UBound(maskoqtets)
    For maskoqtetnum = 0 To maskoqtetnummax Step 1
        DoEvents
        If CLng(maskoqtets(maskoqtetnum)) < 255 Then
            maskrange = 256 - maskoqtets(maskoqtetnum)
            Exit For
        End If
    Next maskoqtetnum
    
    '↑ここでホストアドレスの範囲（ホストアドレスの幅）と第何オクテットがサブネット化されてるかが分かる。
    'オクテット番号：maskoqtetnum　の  ホストアドレス範囲：maskrange
    
    'maskoqtetnumとmaskrangeからホストＩＰアドレスの範囲を算出する。
    netaddr = FromIpAddressToNetworkAddress(ipaddr, subnetMask)
    
    Dim netaddroqtets() As String
    netaddroqtets = Split(netaddr, ".")
    Dim onlyoqtetbcastaddr As String
    onlyoqtetbcastaddr = CStr(CLng((netaddroqtets(maskoqtetnum)) + CLng(maskrange)) - 1)
    
    Dim ipoqtetnum As Long
    Dim ipoqtetnummax As Long
    ipoqtetnummax = maskoqtetnummax
    For ipoqtetnum = 0 To ipoqtetnummax Step 1
        DoEvents
        If ipoqtetnum = maskoqtetnum Then
            bcastaddr = bcastaddr & "." & onlyoqtetbcastaddr
        ElseIf ipoqtetnum = ipoqtetnummax Then
            bcastaddr = bcastaddr & "." & CStr(255 - maskoqtets(maskoqtetnummax))
        Else
            bcastaddr = bcastaddr & "." & netaddroqtets(ipoqtetnum)
        End If
    Next ipoqtetnum
    
End Sub

Public Sub calcTest()
    MsgBox FromDecToBit("255")
'    Dim teststr() As String
'    SafeArrayAllocDescriptor 1, teststr()
'    NobbyString.AppendArray teststr(), "5"
'    NobbyString.AppendArray teststr(), "1"
'    NobbyString.AppendArray teststr(), "3"
'    NobbyString.AppendArray teststr(), "2"
'    NobbyString.AppendArray teststr(), "6"
'    NobbyString.AppendArray teststr(), "7"
'    NobbyString.AppendArray teststr(), "4"
'
'    Dim sortres() As String
'    SafeArrayAllocDescriptor 1, sortres()
'    BubleSortAtString teststr(), sortres()
End Sub


'@brief     第一引数の文字列配列（数字の配列）をソートする。
'@author    Hiroki Nobumoto
'@date      2009.11.16
'@param1    ソート対象文字列配列　文字列の数字
'@param2    ソート結果
Public Function BubleSortAtString(ByRef target() As String, ByRef res() As String) As Boolean
    If UBound(target) < 0 Then
        MsgBox "ソート対象データがありません。(BubleSortAtString)"
        BubleSortAtString = False
        Exit Function
    End If
    
    Dim localTarget() As String
    ReDim localTarget(UBound(target))
    localTarget = target
    
    Dim targetMax As Long
    Dim targetBuf As Long
    targetMax = UBound(localTarget)
    Dim j As Long
    For j = targetMax To 0 Step -1
        For i = 0 To j - 1 Step 1
            If CLng(localTarget(i)) > CLng(localTarget(i + 1)) Then
                targetBuf = localTarget(i)
                localTarget(i) = localTarget(i + 1)
                localTarget(i + 1) = targetBuf
            End If
        Next i
    Next j
    ReDim res(UBound(target))
    res = localTarget
    BubleSortAtString = True

End Function


'@brief     文字列に指定した桁数に固定する為に空いてる桁に文字をつける
'@author    Hiroki Nobumoto
'@date      2010.11.11
'@param     対象の文字列
'@param     変更したい桁数
'@param     あいてる桁に埋める文字
Public Function AppChar(ByVal target As String, ByVal ketasu As Long, ByVal appendChar As String) As String
    Dim i As Long
    Dim umeru As String
    umeru = ""
    
    Dim imax As Long
    imax = ketasu - Len(target)
    
    For i = 0 To imax Step 1
        DoEvents
        umeru = umeru & appendChar
    Next i
    
    AppChar = umeru & target
End Function












