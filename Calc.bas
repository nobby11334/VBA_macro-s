Attribute VB_Name = "Calc"
Private Const STR_REGEXP As String = "VBScript.RegExp"
'Private Declare Function SafeArrayAllocDescriptor Lib "oleaut32" _
'   (ByVal cDims As Long, _
'    ByRef ppsaOut() As Any) As Long


Sub tttt()
    FromBitCountTo255 27
End Sub


'@brief     �v���t�B�b�N�X�i/24�Ȃǁj���u255.255.255.0�v�̕\���ɕς���
'@author    Hiroki Nobumoto
'@date      2010.11.08
'@param     /24�Ȃǂ�"24"�̕�����10�i���̒l
'@return    ������i255.255.255.0���j
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


'@brief     10�i����2�i���i������j�ɂ���B
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param     tgtVal:�r�b�g�l�ɂ�����10�i��
'@return    �r�b�g�ϊ��������ʁi������8���Œ�j
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
    
    
    '����8���ɕ␳
    Dim strlength As Long
    strlength = Len(strBit)
    
    Dim i As Long
    For i = 1 To (8 - strlength) Step 1
        DoEvents
        strBit = "0" & strBit
    Next i
    FromDecimalToBin = strBit
End Function





'@brief     2�i���l��10�i���ɂ���B
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param     tgtVal�F10�i���ɂ�����2�i��
'@return    �r�b�g�ϊ��������ʁi������8���Œ�j
Public Function FromBinToDecimal(ByVal tgtval As String) As Long

    Dim beki As Long
    Dim dec As Long             '10�i�����i�[����ϐ��B
    Dim bitval(7) As String     '�����̓�i�����ꌅ�ɑ΂���1�v�f�֊i�[����ׂ̔z��
    dec = 0
    
    '�z���11111111�̂��ꂼ��̌��̒l�𕪊����ĉE�̕������珇�ԂɊi�[����B
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
            '--�������Ȃ�--
        End If
    Next i
    
    FromBinToDecimal = dec

End Function



'@brief     2�i���l�̃r�b�g����Ԃ��B
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param     tgtVal:�r�b�g���ɂ�����2�i��
'@return    �r�b�g�ϊ��������ʁi������8���Œ�j
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
    Dim totalbit As Integer     '2�i���̃r�b�g�l�i11111111��8�r�b�g�j
    Dim bitval(7) As String     '�����̓�i�����ꌅ�ɑ΂���1�v�f�֊i�[����ׂ̔z��
    totalbit = 0
    
    '�z���11111111�̂��ꂼ��̌��̒l�𕪊����ĉE�̕������珇�ԂɊi�[����B
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
            '--�������Ȃ�--
        End If
    Next i
    
    FromBinToBit = totalbit

End Function


'@brief     10�i���l��2�i���ɂ������̃r�b�g����Ԃ��B
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param     tgtVal:�r�b�g���ɂ�����10�i��
'@return    �r�b�g���istring�^�j
Public Function FromDecToBit(ByVal tgtval As String) As Integer
    FromDecToBit = FromBinToBit(FromDecimalToBin(tgtval))
End Function


'@brief     ������IP�A�h���X�ƃT�u�l�b�g�}�X�N������l�b�g���[�N�A�h���X���Z�o���ĕԂ��B
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param1    IP�A�h���X          ����:(\d+).(\d+).(\d+).(\d+)
'@param2    �T�u�l�b�g�}�X�N    ����:(\d+).(\d+).(\d+).(\d+)
'@return    �r�b�g���istring�^�j
Public Function FromIpAddressToNetworkAddress(ByVal strIPaddr As String, ByVal strSubnetmask As String) As String

    If NobbyString.FindRegExp(strIPaddr, REGEXP_IPADDR) = False Then
        FromIpAddressToNetworkAddress = "�|"
        Exit Function
    End If
    
    If NobbyString.FindRegExp(strSubnetmask, REGEXP_IPADDR) = False Then
        FromIpAddressToNetworkAddress = "�|"
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


'@brief     �����̃T�u�l�b�g�}�X�N���烏�C���h�J�[�h�}�X�N���Z�o���ĕԂ��B
'@author    Hiroki Nobumoto
'@date      2009.06.09
'@param2    �T�u�l�b�g�}�X�N    ����:(\d+).(\d+).(\d+).(\d+)
'@return    �r�b�g���istring�^�j
Public Function FromSubnetmaskToWildcardmask(ByVal strSubnetmask As String) As String

    If NobbyString.FindRegExp(strSubnetmask, REGEXP_IPADDR) = False Then
        FromSubnetmaskToWildcardmask = "�|"
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



'@brief     IP�A�h���X�ƃT�u�l�b�g�}�X�N2�g�������l�b�g���[�N�̒��ɂ��邩�ǂ����𔻒肷��B
'@author    Hiroki Nobumoto
'@date      2009.10.06
'@param1    ��r�Ώ�IP�A�h���XA�i�A�h���XB��背���W���L���j
'@param2    ��r�Ώ�IP�A�h���XA�̃T�u�l�b�g�}�X�N�i�A�h���XB��背���W���L���j
'@param3    ��r�Ώ�IP�A�h���XB�i�A�h���XA��背���W�������j
'@param4    ��r�Ώ�IP�A�h���XB�̃T�u�l�b�g�}�X�N�i�A�h���XA��背���W�������j
'@return    true A��B�͈̔͂ɂ���B�@false A��B�͈̔͂ɂ��Ȃ��B
Public Function ChkIpAddressRange(ByVal ipaddrA As String, ByVal subnetmaskA As String, ByVal ipaddrB As String, ByVal subnetmaskB As String) As Boolean

    '�����W�̐[������AB�̂ǂ���������
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



'@brief     IP�A�h���X�ƃT�u�l�b�g�}�X�N����l�b�g���[�N�A�h���X�ƃu���[�h�L���X�g�A�h���X���Z�o����B�i�z�X�g�A�h���X�����W�j
'@author    Hiroki Nobumoto
'@date      2009.10.06
'@param1    �͈͂�m�肽��IP�A�h���X
'@param2    �͈͂�m�肽��IP�A�h���X�̃T�u�l�b�g�}�X�N
'@param3    �o�͌��ʁi�l�b�g���[�N�A�h���X�j
'@param4    �o�͌��ʁi�u���[�h�L���X�g�A�h���X�j
'@return
Public Sub IpAddressRange(ByVal ipaddr As String, ByVal subnetMask As String, ByRef netaddr As String, ByRef bcastaddr As String)

    Dim ipoqtets() As String
    oqtets = Split(ipaddr, ".")
    
    Dim maskoqtets() As String
    maskoqtets = Split(subnetMask, ".")
    
    
    '�T�u�l�b�g�}�X�N����z�X�g�A�h���X�T�u�l�b�g�̂h�o�A�h���X�̃I�N�e�b�g�ԍ��ƃz�X�g�A�h���X�͈͂��Z�o����B
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
    
    '�������Ńz�X�g�A�h���X�͈̔́i�z�X�g�A�h���X�̕��j�Ƒ扽�I�N�e�b�g���T�u�l�b�g������Ă邩��������B
    '�I�N�e�b�g�ԍ��Fmaskoqtetnum�@��  �z�X�g�A�h���X�͈́Fmaskrange
    
    'maskoqtetnum��maskrange����z�X�g�h�o�A�h���X�͈̔͂��Z�o����B
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


'@brief     �������̕�����z��i�����̔z��j���\�[�g����B
'@author    Hiroki Nobumoto
'@date      2009.11.16
'@param1    �\�[�g�Ώە�����z��@������̐���
'@param2    �\�[�g����
Public Function BubleSortAtString(ByRef target() As String, ByRef res() As String) As Boolean
    If UBound(target) < 0 Then
        MsgBox "�\�[�g�Ώۃf�[�^������܂���B(BubleSortAtString)"
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


'@brief     ������Ɏw�肵�������ɌŒ肷��ׂɋ󂢂Ă錅�ɕ���������
'@author    Hiroki Nobumoto
'@date      2010.11.11
'@param     �Ώۂ̕�����
'@param     �ύX����������
'@param     �����Ă錅�ɖ��߂镶��
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












