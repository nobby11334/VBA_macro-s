Attribute VB_Name = "NobbyString"
Private Const STR_REGEXP_CALL As String = "VBScript.RegExp"
    
Public arrayappendflg As Boolean



'@brief     �������̕������������ŕ������đ�O�����Ŏw�肵���ԍ��̕������Ԃ��B
'@date      2009.08.27
'@author    Hiroki Nobumoto
Public Function GetSplit(ByVal strSource As String, ByVal strSplitter As String, ByVal intGetNum As Integer) As String
    Dim splitStrs() As String
    splitStrs() = Split(strSource, strSplitter)
    GetSplit = splitStrs(intGetNum)
End Function

'@brief     �z��Ƀf�[�^���i�[���ėv�f����ǉ�����B
'@date      2009.08.16
'@author    Hiroki Nobumoto
'@param     strTarget() : �f�[�^��ǉ������Ώ۔z��ւ̎Q��
'@param     inputdata : �ǉ�����f�[�^
Public Sub AppendArray(ByRef strTarget() As String, ByVal inputdata As String)
    Dim res() As String
'    If UBound(strTarget) < 0 Then
    If UBound(strTarget) = 0 And strTarget(0) = "" Then
        ReDim res(0)
        res(0) = ""
        ReDim strTarget(0)
    Else
        ReDim res(UBound(strTarget))
        res = strTarget
        ReDim Preserve strTarget(UBound(strTarget) + 1)
    End If
    
    If UBound(res) = 0 Then
        If res(0) = "" Then
            '---�Ȃɂ����Ȃ�---
        Else
            ReDim Preserve res(UBound(res) + 1)
        End If
    Else
        ReDim Preserve res(UBound(res) + 1)
    End If
    res(UBound(res)) = inputdata
    strTarget(UBound(strTarget)) = res(UBound(res))
End Sub

'@brief     ��̔z����}�[�W����B
'@auth      Hiroki Nobumoto
'@date      2009.10.02
'@param1    �}�[�W����擪���̔z��
'@param2    �}�[�W����㑤�̔z��
'@param3    �}�[�W��������
Public Sub MargeStr(ByRef strA() As String, ByRef strB() As String, ByRef strAB() As String)

    If UBound(strA) < 0 Or UBound(strB) < 0 Then
        MsgBox "�}�[�W���錳�̃f�[�^������܂���B"
        Exit Sub
    End If
    
    ReDim strAB(0)
    strAB(0) = ""
    Dim strAlength As Long
    Dim strBlength As Long
    strAlength = UBound(strA())
    strBlength = UBound(strB())
    
    Dim localStrA() As String
    Dim localStrB() As String
    ReDim localStrA(strAlength)
    ReDim localStrB(strBlength)
    localStrA = strA
    localStrB = strB
    Dim i As Long
    For i = 0 To strBlength Step 1
        DoEvents
        ReDim Preserve localStrA(strAlength + i + 1)
        localStrA(UBound(localStrA)) = localStrB(i)
    Next i
    ReDim strAB(UBound(localStrA))
    strAB = localStrA
End Sub



'@brief     �z�񂪋���ۂ̏ꍇ-1��Ԃ�ubound
'@author    Hiroki Nobumoto
'@date       2009.10.03
'@param     �Ώ۔z��i����ۂł��j
Function local_UBound(ByRef pArray)
On Error Resume Next
    local_UBound = UBound(pArray)
    Select Case Err.Number
      Case 0
      Case 9
           local_UBound = -1
      Case Else
           MsgBox Err.Description & "(" & Err.Number & ")", vbOKOnly + vbCritical, Err.Source
           End
    End Select
End Function

'@brief     ��������string�z�񂩂�������Ŏw�肵���R�}���h�Ŏ擾�����X�e�[�^�X�����𒊏o���đ�O�����֕Ԃ��B
'@date      2009.08.27
'@author    Hiroki Nobumoto
'@param     ByRef strall() As String        : ���o�������e�L�X�g���܂ރt�@�C������擾�����e�L�X�g�S��
'@param     ByVal strKey As String          : �Ώۃt�@�C�����璊�o�������e�L�X�g�̒��o�J�n�L�[���[�h�i��ӂȃL�[���[�h���w�肷��B�j
'@param     ByVal strKeyEnd As String       : �Ώۃt�@�C�����璊�o�������e�L�X�g�̒��o�I���L�[���[�h
'@param     ByRef strTargetData() As String : ���o�����e�L�X�g���i�[����string�z��ւ̎Q��
'@return    true: ��������  false:  �����Ώۂ������B
Public Function GetTxtPart(ByRef strAll() As String, ByVal strKey As String, ByVal strKeyEnd As String, ByRef strTargetData() As String) As Boolean

    '-----�G���[�`�F�b�N-----��
    If IsArray(strAll) = False Then
        GetTxtPart = False
        Exit Function
    End If
    
    If UBound(strAll) < 1 Then
        GetTxtPart = False
        Exit Function
    End If
    
    If strKey = NO_TEXT Then
        GetTxtPart = False
        Exit Function
    End If
    
    If strKeyEnd = NO_TEXT Then
        GetTxtPart = False
        Exit Function
    End If
    '-----------------------��
    
    '-----�S�e�L�X�g��ΏۂɃ`�F�b�N�J�n-----��
    Dim execflg As Boolean          'true: �ǂݎ�����������strTargetData()�Ƀo�b�t�@���ėǂ��B
                                    'false: �ǂݎ�����������strTargetData()�Ƀo�b�t�@���Ȃ��B
    execflg = False
    Dim txtrownummax As Long
    txtrownummax = UBound(strAll)
    
    Dim txtrownum As Long
    
    ReDim strTargetData(1)
    For txtrownum = 0 To txtrownummax Step 1
        DoEvents
        If execflg = True Then
        
            NobbyString.AppendArray strTargetData(), strAll(txtrownum)
            
            '�o�b�t�@�I���ۊm�F
            If InStr(strAll(txtrownum + 1), strKeyEnd) > 0 Or txtrownum = txtRowMax Then
            
                execflg = False
            
            Else
            
                '�Ȃɂ����Ȃ�
                
            End If
            
        Else
            '�o�b�t�@�J�n�ۊm�F
            If InStr(strAll(txtrownum), strKey) > 0 Then
            
                '�o�b�t�@�ۃt���OON
                execflg = True
            
            Else
            
                '�Ȃɂ����Ȃ�
                
            End If
            
        End If
        
    Next txtrownum
    '---------------------------------------��
    
    
End Function


Public Sub regTest()
    Dim resvals() As String
    ReDim resvals(0)
    resvals(0) = ""
    Dim resval As String
    MsgBox GetRegExp("C       10.81.32.128/25 is directly connected, Vlan3", "\d+.\d+.\d+.\d+/\d+", resvals())
'    MsgBox GetRegExp("B*   0.0.0.0/0 [20/0] via 10.154.48.46, 01:12:08", "[a-zA-Z_0-9/\.]+", resvals())

'    MsgBox FindRegExp("ITEN-r2A1-1009T#sh storm-control", "(\w+)(#*)(sh)(\w*)(\s+)(storm)")
'    MsgBox FindRegExp("ITEN-r2A1-1009T#show ip route sum", "(\w+)(#*)(\s*)(sh)(\w*)(\s+)(ip)(\s+)(route)(\s+)(sum)(\w*)")
'    If GetRegExp("1      Po1(SU)          -        Fa0/21(P)   Fa0/22(P)   Fa0/23(P)", "\w+/\w+\(\w+\)", resvals()) = True Then
'
'    End If
End Sub

'@brief     �������̒��ɑ������Ŏw�肵�����������������B�i���K�\����OK�j
'@aurhot    Hiroki Nobumoto
'@date      2009.08.27
'@param     strTarget : ��������镶����
'@param     strPattern : ����������������p�^�[��
'@return    true : �����Ƀq�b�g�@false :�@�����Ƀq�b�g���Ȃ��B
Function FindRegExp(ByVal strTarget As String, ByVal strPattern As String) As Boolean
    Set re = CreateObject(STR_REGEXP_CALL)
    Dim res As Boolean
    re.Pattern = strPattern
    res = re.test(strTarget)
    Set re = Nothing
    FindRegExp = res
End Function




'@brief     �������̒��ɑ������Ŏw�肵����������������āA���̃p�^�[���Ɉʒu���镶�����Ԃ��B�i���K�\����OK�j
'@aurhot    Hiroki Nobumoto
'@date      2009.08.27
'@param     strTarget : ��������镶����
'@param     strPattern : ����������������p�^�[��
'@param     ���K�\���Ƀ}�b�`����������B�����̏ꍇ���i�[�B
'@return    true : �����Ƀq�b�g�@false :�@�����Ƀq�b�g���Ȃ��B
Function GetRegExp(ByVal strTarget As String, ByVal strPattern As String, ByRef res() As String) As Boolean
    ReDim res(0)
    res(0) = ""
    Set re = CreateObject(STR_REGEXP_CALL)
    Dim matches As Object
    Dim match As Object
    
    re.Pattern = strPattern
    re.Global = True
    Set matches = re.Execute(strTarget)
    
    For Each match In matches
        NobbyString.AppendArray res(), CStr(match.Value)
    Next match
    
    Set re = Nothing
'    If UBound(res) < 0 Then
    If UBound(res) = 0 And res(0) = "" Then
        GetRegExp = False
    Else
        GetRegExp = True
    End If
    Set matches = Nothing


End Function




