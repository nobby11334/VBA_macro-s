Attribute VB_Name = "CMDGEN"
'Private Declare Function SafeArrayAllocDescriptor Lib "oleaut32" _
'   (ByVal cDims As Long, _
'    ByRef ppsaOut() As Any) As Long



Private Const STR_NMSHOSTNAME As String = "tama_ws2"
Private Const STR_NMSUID As String = "nmsop1"
Private Const STR_NMSPASSWD As String = "fbknms1"

'Private Const STR_NMSPROMPT As String = "'=>'"
Private Const STR_NMSPROMPT As String = "']:'"
'Private Const STR_CMD_TELNET As String = "telnet "
Private Const STR_CMD_TELNET As String = "/usr/bin/telnet "
Private Const STR_SWPROMPT_USER As String = "'>'"
Private Const STR_TTL_COMMENT As String = ";"
'Private Const STR_CONNECT_TTL_WS1 As String = "connect '10.239.0.11:23 /nossh /T=1'"
Private Const STR_CONNECT_TTL_WS1 As String = "connect '10.239.0.168:23 /nossh /T=1'"    '�{�Ԃ̃A�h���X�ɏC���\��NMS#1
Private Const STR_CONNECT_TTL_WS2 As String = "connect '10.239.0.43:23 /nossh /T=1'"
Private Const STR_CONNECT_TTL_WS3 As String = "connect '10.239.0.27:23 /nossh /T=1'"
Private Const STR_CONNECT_TTL_WS4 As String = "connect '10.239.0.29:23 /nossh /T=1'"

Private Const STR_FLG_IOS As String = "IOS"
Private Const STR_FLG_PF As String = "PF"
Private Const STR_FLG_NX As String = "NX"


Private Const STR_SWPROMPT_UIDREQUEST As String = "'Username:'"
Private Const STR_SWPROMPT_PASSREQUEST As String = "'Password:'"
Private Const STR_SWCMD_ENABLE As String = "'enable'"
Private Const STR_SWCMD_TER_LEN As String = "'ter len 0'"
Private Const STR_CWCMD_EXIT As String = "'exit'"
Private Const STR_SWPROMPT_ENABLE As String = "'#'"
'Private Const STR_SWLOGIN_ID As String = "IBMSEG20001F"
'Private Const STR_SWLOGIN_PASSWORD As String = "fbk6611c"
'Private Const STR_SWLOGIN_PASSWORD2 As String = "mhbk6611c"

Private Const STR_PFPROMPT_UID As String = "'login:'"
Private Const STR_PFPROMPT_PASSWD As String = "'Password:'"
Private Const STR_PFPROMPT_USERMODE As String = "'>'"
Private Const STR_PFPROMPT_ADMINMODE As String = "'(A)>'"
Private Const STR_PFCOMMAND_ADMINMODE As String = "'admin'"

Private Const STR_SCROLL_BUFFERSIZE As String = "setenv 'ScrollBuffSize' '10'"

'SYSTEMID PW
Private Const STR_SYSID As String = "SYSID"
Private Const STR_SYSPW As String = "SYSPW"


Private Const STR_TTL_SENDLN As String = "sendln "
Private Const STR_TTL_WAIT As String = "wait "
Private Const STR_TTL_CHANGEDIR As String = "changedir "
Private Const STR_TTL_CONNECT As String = "connect "
Private Const STR_TTL_CONNECT_PARAM As String = " /nossh /user=" & STR_NMSUID & " /passwd=" & STR_NMSPASSWD
Private Const STR_TTL_LOGOPEN As String = "logopen "
Private Const ST_TTL_LOGCLOSE As String = "logclose"

Private Const STR_LOG_EXT As String = ".log"



'================�s�ԍ���ԍ��}�b�v
Private Const LNG_COL_NUM_HOSTNAME As Long = 1
Private Const LNG_ROW_NUM_HOSTNAME As Long = 1
Private Const LNG_COL_NUM_ADDR As Long = 2
Private Const LNG_COL_NUM_UID As Long = 3
Private Const LNG_COL_NUM_PW As Long = 4
Private Const LNG_COL_NUM_PROMPT As Long = 6
Private Const LNG_COL_NUM_CMD As Long = 7
Private Const LNG_ROW_NUM_OFFSET_HOSTNAME As Long = 3
Private Const LNG_ROW_NUM_OFFSET_PROMPT As Long = 4
Private Const LNG_ROW_NUM_OFFSET_CMD As Long = 4
Private Const LNG_COL_NUM_TTLFILENAME As Long = 2
Private Const LNG_ROW_NUM_TTLFILENAME As Long = 58

Private Const LNG_COL_NUM_PFPROMPT As Long = 9
Private Const LNG_COL_NUM_PFCOMMAND As Long = 10
Private Const LNG_COL_NUM_OSTYPE As Long = 5

Private Const LNG_ROW_NUM_HOP_MACHINE As Long = 46
Private Const LNG_COL_NUM_HOP_MACHINE As Long = 1
Private Const LNG_COL_NUM_HOP_IPADDR As Long = 2
Private Const LNG_COL_NUM_HOP_UID As Long = 3
Private Const LNG_COL_NUM_HOP_PW As Long = 4

Private Const STR_MSG_END1 As String = "�t�@�C�����쐬���܂����B�t�@�C�����u"
Private Const STR_MSG_END2 As String = "�v"
Private Const STR_MSG_FILEPATH_ERR As String = "�t�@�C����ۑ��ł��܂���ł����B"


'@brief     ���s�{�^��
'@date      2009.08.16
'@author    Hiroki Nobumoto
'@return    nothing
Sub CommandGeneration()
    Dim filename As String
    Dim oyafolder As String
    oyafolder = CStr(Worksheets("main").Cells(1, 13).Value2)
    filename = TtlStruct(MkWorkDir(oyafolder), ActiveSheet.Name)
    If filename = "false" Then
        Exit Sub
    Else
        MsgBox STR_MSG_END1 & folderpath & filename & STR_MSG_END2
    End If

End Sub

'@brief     TeraTerm�}�N���t�@�C���̓��e���\�z����B
'@date      2009.08.16
'@author    Hiroki Nobumoto
'@param     strttlTxt() : TeraTerm�}�N���t�@�C���Ɋi�[����e�L�X�g���i�[����String�z��ւ̎Q��
'@return    true������ɏ����I���@   false�������G���[
Public Function TtlStruct(ByVal folderpath As String, ByVal shtname As String) As Boolean        'ByRef strttlTxt() As String) As Boolean
    Dim strttlTxt() As String

'    SafeArrayAllocDescriptor 1, strttlTxt()
    ReDim strttlTxt(0)
    strttlTxt(0) = ""
    
    With Worksheets(shtname)
        '�z�X�g�����擾
        Dim hostname As String
        Dim ipaddress As String
        Dim uid As String
        Dim pw As String
        Dim ostype As String
        
        Dim hopipaddr As String
        Dim hopuid As String
        Dim hoppw As String
        Dim hopostype As String
        Dim hophostname As String
        Dim hophostrownum As Long
        
        Dim hopcount As Long
        hopcount = 1
        
        Dim logfileCmdName As String
        Dim logfilecmdnameall As String
        Dim doprompt As String
        Dim docommand As String
        Dim cmdrownum As Long
        Dim logopenLine As String
        
        
        hostrownum = LNG_ROW_NUM_OFFSET_HOSTNAME + LNG_ROW_NUM_HOSTNAME
        ipaddress = LNG_COL_NUM_ADDR
        cmdrownum = LNG_ROW_NUM_OFFSET_PROMPT
        
        hostname = .Cells(hostrownum, LNG_COL_NUM_HOSTNAME).Value2
        ipaddress = .Cells(hostrownum, LNG_COL_NUM_ADDR).Value2
        
        'UID�̎擾
        If .Cells(hostrownum, LNG_COL_NUM_UID).Value2 <> NO_TEXT Then
            uid = STR_QUOAT & .Cells(hostrownum, LNG_COL_NUM_UID).Value2 & STR_QUOAT
        Else
            uid = NO_TEXT
        End If
        
        '�p�X���[�h�̎擾
        pw = STR_QUOAT & .Cells(hostrownum, LNG_COL_NUM_PW).Value2 & STR_QUOAT
        
        'IOS��PF�����擾
        ostype = .Cells(hostrownum, LNG_COL_NUM_OSTYPE).Value2
        
        'NMS�ɐڑ�����
        Dim nmshostname As String
        Dim nmsIPloginCommand As String
        nmshostname = .Cells(6, 13).Value2
'        If nmshostname = "tama_ws1" Then
'            nmsIPloginCommand = STR_CONNECT_TTL_WS1
'        ElseIf nmshostname = "tama_ws2" Then
'            nmsIPloginCommand = STR_CONNECT_TTL_WS2
'        ElseIf nmshostname = "tama_ws3" Then
'            nmsIPloginCommand = STR_CONNECT_TTL_WS3
'        ElseIf nmshostname = "tama_ws4" Then
'            nmsIPloginCommand = STR_CONNECT_TTL_WS4
'        Else
'            MsgBox "NMS�[���̑I��������Ă��܂���B�����𒆒f���܂��B"
'            Exit Function
'        End If
        If nmshostname <> "" Then
            Dim getnmsip() As String
            ReDim getnmsip(0)
            If NobbyString.GetRegExp(nmshostname, "\d+.\d+.\d+.\d+", getnmsip()) = True Then
                nmsIPloginCommand = "connect '" & getnmsip(0) & ":23 /nossh /T=1'"
            Else
                MsgBox "NMS�[���̃A�h���X�̏���������������܂���B�����𒆒f���܂��B"
                TtlStruct = False
                Exit Function
            End If
            
        Else
            MsgBox "NMS�[���̑I��������Ă��܂���B�����𒆒f���܂��B"
            TtlStruct = False
            Exit Function
        End If
'        NobbyString.AppendArray strttlTxt(), STR_CONNECT_TTL_WS1

        '//151109 ����NMS�@�ݏoID�Ή�
        Dim rentalID As String
        Dim rentalPW As String
        rentalID = .Cells(2, 13).Value2
        rentalPW = .Cells(3, 13).Value2
        
        If rentalID = "" Then
            MsgBox "�ݏoID�̓��͂�����܂���B�����𒆒f���܂��B"
            TtlStruct = False
            Exit Function
        End If
        
        If rentalPW = "" Then
            MsgBox "�ݏopassword�̓��͂�����܂���B�����𒆒f���܂��B"
            TtlStruct = False
            Exit Function
        End If

        NobbyString.AppendArray strttlTxt(), STR_SCROLL_BUFFERSIZE
        NobbyString.AppendArray strttlTxt(), nmsIPloginCommand
        NobbyString.AppendArray strttlTxt(), "wait 'login:'"
'        NobbyString.AppendArray strttlTxt(), "sendln 'nmsop1'"             '//�ݏoID�Ή�
        NobbyString.AppendArray strttlTxt(), "sendln '" & rentalID & "'"    '//�ݏoID�Ή�
        NobbyString.AppendArray strttlTxt(), "wait 'Password:'"
'        NobbyString.AppendArray strttlTxt(), "sendln 'fbknms1'"            '//�ݏoID�Ή�
        NobbyString.AppendArray strttlTxt(), "sendln '" & rentalPW & "'"    '//�ݏoID�Ή�
'        NobbyString.AppendArray strttlTxt(), "wait '=>'"                   '//�ݏoID�Ή�
        NobbyString.AppendArray strttlTxt(), "wait '#'"                     '//�ݏoID�Ή�
'        Dim teststr As String
'        teststr = "logopen '" & folderpath & CStr(.Cells(58, 2).Value2) & ".log" & "'" & " 0 0"
'        NobbyString.AppendArray strttlTxt(), "logopen '" & folderpath & CStr(.Cells(58, 2).Value2) & ".log" & "'" & " 0 0"
        
        '110622 SYSID�̎擾
        MakeGetSYSIDTTL strttlTxt()
        uid = STR_SYSID
        pw = STR_SYSPW
'        NobbyString.AppendArray strttlTxt(), "wait '=>'"
        NobbyString.AppendArray strttlTxt(), "wait '#'"

        
        
        Dim hostcount As Integer
        hostcount = 0
        Do While hostname <> NO_TEXT
            DoEvents
            '�z�X�g�������x������B
            NobbyString.AppendArray strttlTxt(), STR_TTL_COMMENT
            NobbyString.AppendArray strttlTxt(), STR_TTL_COMMENT & hostname


            '091128�Ή�=======================================================================================================
            'NMS����̃z�b�v��ւ�TELNET
            If ostype = "PF" Then
                hophostrownum = LNG_ROW_NUM_HOP_MACHINE
                hophostname = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_HOP_MACHINE).Value2 & STR_QUOAT
                hopipaddr = .Cells(hophostrownum, LNG_COL_NUM_HOP_IPADDR).Value2

                hopostype = .Cells(hophostrownum, LNG_COL_NUM_OSTYPE).Value2
                If .Cells(hophostrownum, LNG_COL_NUM_UID).Value2 <> NO_TEXT Then
                    hopuid = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_UID).Value2 & STR_QUOAT
                Else
                    hopuid = NO_TEXT
                End If
                hoppw = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_HOP_PW).Value2 & STR_QUOAT

                Do While .Cells(hophostrownum, LNG_COL_NUM_HOP_MACHINE).Value2 <> NO_TEXT
                    NobbyString.AppendArray strttlTxt(), STR_TTL_SENDLN & STR_QUOAT & STR_CMD_TELNET & hopipaddr & STR_QUOAT
                    If MakeLoginCommand(hopostype, hopuid, hoppw, strttlTxt()) = False Then
                        MsgBox "�Z���̓��͂ɕs��������܂��B���O�C��ID����уp�X���[�h�AOS�^�C�v���m�F���ĉ������B"
                        Exit Function
                    End If
                    hophostrownum = hophostrownum + 1
                    hophostname = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_HOP_MACHINE).Value2 & STR_QUOAT
                    hopipaddr = .Cells(hophostrownum, LNG_COL_NUM_HOP_IPADDR).Value2
                    If .Cells(hophostrownum, LNG_COL_NUM_UID).Value2 <> NO_TEXT Then
                        hopuid = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_UID).Value2 & STR_QUOAT
                    Else
                        hopuid = NO_TEXT
                    End If
                    hoppw = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_HOP_PW).Value2 & STR_QUOAT
                    hopcount = hopcount + 1
                Loop
            End If
            
            '�z�b�v�悩��Ώۋ@��ւ�TELNET
            NobbyString.AppendArray strttlTxt(), STR_TTL_SENDLN & STR_QUOAT & STR_CMD_TELNET & ipaddress & STR_QUOAT
            
            '�Ώۋ@��ւ̃��O�C������я����R�}���h����
            'IOS��PF�̂ǂ��炩�𔻕ʂ��ă��O�C���R�}���h�Z�b�g���쐬����B
            uid = STR_SYSID
            pw = STR_SYSPW

            If MakeLoginCommand(ostype, uid, pw, strttlTxt()) = False Then
                MsgBox "�Z���̓��͂ɕs��������܂��B���O�C��ID����уp�X���[�h�AOS�^�C�v���m�F���ĉ������B�����𒆒f���܂��B"
                TtlStruct = False
                Exit Function
            End If
            '================================================================================================================
            'prompt���擾����ttl���쐬
            MakeGetPromptTTL strttlTxt(), hostcount
            
            'log�f�B���N�g�����ړ�
            NobbyString.AppendArray strttlTxt(), "changedir '" & folderpath & "'"
            
            '�R�}���h���s����TTL���쐬
            NobbyString.AppendArray strttlTxt(), ";�R�}���h���s"
            Do
                DoEvents
                
                'OS�^�C�v�ɂ���ēǂݎ��R�}���h���X�g��؂�ւ���B
                '// IOS or NXOS �̏ꍇ��7��ڂ̃R�}���h���X�g�����s����B
'                If ostype = STR_FLG_IOS Then
                If ostype = STR_FLG_IOS Or ostype = STR_FLG_NX Then
'                    doprompt = STR_QUOAT & .Cells(cmdrownum, LNG_COL_NUM_PROMPT).Value2 & STR_QUOAT
                    doprompt = "hostname"
                    docommand = STR_QUOAT & .Cells(cmdrownum, LNG_COL_NUM_CMD).Value2 & STR_QUOAT
                    
                '// PF �̏ꍇ��10��ڂ̃R�}���h���X�g�����s����B
                ElseIf ostype = STR_FLG_PF Then
'                    doprompt = STR_QUOAT & .Cells(cmdrownum, LNG_COL_NUM_PFPROMPT).Value2 & STR_QUOAT
                    doprompt = "hostname"
                    docommand = STR_QUOAT & .Cells(cmdrownum, LNG_COL_NUM_PFCOMMAND).Value2 & STR_QUOAT
                End If
                If docommand = NO_TEXT Or docommand = "''" Then
                    Exit Do
                End If
                '�R�}���h��쐬
                
                '���ꕶ����S���u_�v�ɕϊ�
                logfileCmdName = Replace(docommand, "'", "_")
                logfileCmdName = Replace(logfileCmdName, " ", "_")
                logfileCmdName = Replace(logfileCmdName, "/", "_")
                logfileCmdName = Replace(logfileCmdName, "%", "_")
                logfileCmdName = Replace(logfileCmdName, "|", "_")
                logfileCmdName = Replace(logfileCmdName, "#", "_")
                logfileCmdName = Replace(logfileCmdName, "*", "_")
                logfileCmdName = Replace(logfileCmdName, ">", "_")
                logfileCmdName = Replace(logfileCmdName, "<", "_")
                logfileCmdName = Replace(logfileCmdName, "[", "_")
                logfileCmdName = Replace(logfileCmdName, "]", "_")
                logfileCmdName = Replace(logfileCmdName, "$", "_")
                logfileCmdName = Replace(logfileCmdName, "&", "_")
                logfileCmdName = Replace(logfileCmdName, "+", "_")
                logfileCmdName = Replace(logfileCmdName, ",", "_")
                logfileCmdName = Replace(logfileCmdName, "\", "_")
                logfileCmdName = Replace(logfileCmdName, "(", "_")
                logfileCmdName = Replace(logfileCmdName, ")", "_")

                logopenLine = "logopen '" & CStr(.Cells(1, 13).Value2) & "_" _
                                            & hostname & "_" _
                                            & logfileCmdName _
                                            & ".log" & "'" & " 0 0 1 0 1"


                NobbyString.AppendArray strttlTxt(), logopenLine
                MakeExecuteCommand doprompt, docommand, strttlTxt()
                MakeExecuteCommand doprompt, "''", strttlTxt()
                MakeExecuteCommand doprompt, "''", strttlTxt()
                NobbyString.AppendArray strttlTxt(), "logclose"
                cmdrownum = cmdrownum + 1
'                hostcount = hostcount + 1
            Loop While docommand <> NO_TEXT
            
            'EXIT�R�}���h���� IOS��PF�𔻕ʂ��č쐬����B
            If MakeExitCommand(ostype, hopcount, strttlTxt(), hostcount) = False Then
                MsgBox "OS�^�C�v�̎w�肪����܂���B�����𒆒f���܂��B"
                TtlStruct = False
                Exit Function
            End If
            
            NobbyString.AppendArray strttlTxt(), STR_TTL_WAIT & STR_NMSPROMPT
            '�R�}���h�s�̐擪�s�ԍ��֍s�ԍ�������������B
            cmdrownum = LNG_ROW_NUM_OFFSET_PROMPT

            '���̃z�X�g���̍s�֎Q�ƍs�ԍ����C���N�������g
            hostrownum = hostrownum + 1
            
            '���̃z�X�g����IP�A�h���X���擾
            hostname = .Cells(hostrownum, LNG_COL_NUM_HOSTNAME).Value2
            ipaddress = .Cells(hostrownum, LNG_COL_NUM_ADDR).Value2
'            If .Cells(hostrownum, LNG_COL_NUM_UID).Value2 <> NO_TEXT Then
'                uid = STR_QUOAT & .Cells(hostrownum, LNG_COL_NUM_UID).Value2 & STR_QUOAT
'            Else
'                uid = NO_TEXT
'            End If
'            pw = STR_QUOAT & .Cells(hostrownum, LNG_COL_NUM_PW).Value2 & STR_QUOAT
'            ostype = .Cells(hostrownum, LNG_COL_NUM_OSTYPE).Value2
            hostcount = hostcount + 1
        Loop
        NobbyString.AppendArray strttlTxt(), ":END"
'        NobbyString.AppendArray strttlTxt(), "logclose"
        NobbyString.AppendArray strttlTxt(), "sendln 'exit'"
        NobbyString.AppendArray strttlTxt(), "messagebox '�X�e�[�^�X���O�擾�������I�����܂����B' '��������'"
        'strttlTxt()�̓��e��ttl�t�@�C���֕ۑ�����B
        Dim filename As String
        filename = .Cells(LNG_ROW_NUM_TTLFILENAME, LNG_COL_NUM_TTLFILENAME).Value2
        
        If FileActuators.WriteOnceFile(folderpath & filename, strttlTxt()) < 0 Then
            MsgBox STR_MSG_FILEPATH_ERR
            TtlStruct = False
            Exit Function
        End If
        Dim battxt() As String
'        SafeArrayAllocDescriptor 1, battxt()
        ReDim battxt(0)
        battxt(0) = ""
        
        NobbyString.AppendArray battxt(), """" & CStr(.Cells(59, 2).Value2) & """ /V " & """" & folderpath & CStr(.Cells(58, 2).Value2) & """"
        FileActuators.WriteOnceFile folderpath & filename & ".bat", battxt()
'        MsgBox STR_MSG_END1 & folderpath & filename & STR_MSG_END2
'        TtlStruct = folderpath & filename
        TtlStruct = True
    End With
    
End Function

'@brief     TeraTerm�}�N���t�@�C���̓��e���\�z����B
'@date      2009.08.16
'@author    Hiroki Nobumoto
'@param     strttlTxt() : TeraTerm�}�N���t�@�C���Ɋi�[����e�L�X�g���i�[����String�z��ւ̎Q��
'@return    true ����ɏ����I���@false �����G���[
Public Function CopyRunTftpStruct(ByVal folderpath As String, ByVal shtname As String) As String        'ByRef strttlTxt() As String) As Boolean
    Dim strttlTxt() As String

'    SafeArrayAllocDescriptor 1, strttlTxt()
    ReDim strttlTxt(0)
    strttlTxt(0) = ""
    
    With Worksheets(shtname)
        '�z�X�g�����擾
        Dim hostname As String
        Dim ipaddress As String
        Dim uid As String
        Dim pw As String
        Dim ostype As String
        
        Dim hopipaddr As String
        Dim hopuid As String
        Dim hoppw As String
        Dim hopostype As String
        Dim hophostname As String
        Dim hophostrownum As Long
        
        Dim hopcount As Long
        hopcount = 1
        
        Dim logfileCmdName As String
        Dim logfilecmdnameall As String
        Dim doprompt As String
        Dim docommand As String
        Dim cmdrownum As Long
        Dim logopenLine As String
        
        
        hostrownum = LNG_ROW_NUM_OFFSET_HOSTNAME + LNG_ROW_NUM_HOSTNAME
        ipaddress = LNG_COL_NUM_ADDR
        cmdrownum = LNG_ROW_NUM_OFFSET_PROMPT
        
        hostname = .Cells(hostrownum, LNG_COL_NUM_HOSTNAME).Value2
        ipaddress = .Cells(hostrownum, LNG_COL_NUM_ADDR).Value2
        If .Cells(hostrownum, LNG_COL_NUM_UID).Value2 <> NO_TEXT Then
            uid = STR_QUOAT & .Cells(hostrownum, LNG_COL_NUM_UID).Value2 & STR_QUOAT
        Else
            uid = NO_TEXT
        End If
        pw = STR_QUOAT & .Cells(hostrownum, LNG_COL_NUM_PW).Value2 & STR_QUOAT
        ostype = .Cells(hostrownum, LNG_COL_NUM_OSTYPE).Value2
        
        'NMS�ɐڑ�����
        Dim nmshostname As String
        Dim nmsIPloginCommand As String
        Dim ipaddrNMS As String
        nmshostname = .Cells(6, 13).Value2
        If nmshostname = "tama_ws1" Then
'            ipaddrNMS = "10.239.0.11"
            ipaddrNMS = "10.239.0.168"
            nmsIPloginCommand = STR_CONNECT_TTL_WS1
        ElseIf nmshostname = "tama_ws2" Then
            ipaddrNMS = "10.239.0.43"
            nmsIPloginCommand = STR_CONNECT_TTL_WS2
        ElseIf nmshostname = "tama_ws3" Then
            ipaddrNMS = "10.239.0.27"
            nmsIPloginCommand = STR_CONNECT_TTL_WS3
        ElseIf nmshostname = "tama_ws4" Then
            ipaddrNMS = "10.239.0.29"
            nmsIPloginCommand = STR_CONNECT_TTL_WS4
        Else
            MsgBox "NMS�[���̑I��������Ă��܂���B�����𒆒f���܂��B"
            Exit Function
        End If
'        NobbyString.AppendArray strttlTxt(), STR_CONNECT_TTL_WS1
        NobbyString.AppendArray strttlTxt(), STR_SCROLL_BUFFERSIZE
        NobbyString.AppendArray strttlTxt(), nmsIPloginCommand
        NobbyString.AppendArray strttlTxt(), "wait 'login:'"
'        NobbyString.AppendArray strttlTxt(), "sendln 'nmsop1'"             '// �ݏoID�Ή�
        NobbyString.AppendArray strttlTxt(), "sendln '" & rentalID & "'"    '// �ݏoID�Ή�
        
        NobbyString.AppendArray strttlTxt(), "wait 'Password:'"
'        NobbyString.AppendArray strttlTxt(), "sendln 'fbknms1'"            '// �ݏoID�Ή�
        NobbyString.AppendArray strttlTxt(), "sendln '" & rentalPW & "'"    '// �ݏoID�Ή�
'        NobbyString.AppendArray strttlTxt(), "wait '=>'"                   '// �ݏoID�Ή�
        NobbyString.AppendArray strttlTxt(), "wait '#'"                     '// �ݏoID�Ή�

'        Dim teststr As String
'        teststr = "logopen '" & folderpath & CStr(.Cells(58, 2).Value2) & ".log" & "'" & " 0 0"
'        NobbyString.AppendArray strttlTxt(), "logopen '" & folderpath & CStr(.Cells(58, 2).Value2) & ".log" & "'" & " 0 0"
        'SYSID�̎擾
        MakeGetSYSIDTTL strttlTxt()
        uid = STR_SYSID
        pw = STR_SYSPW
        NobbyString.AppendArray strttlTxt(), "wait '=>'"
        Dim hostcount As Integer
        hostcount = 0
        Do While hostname <> NO_TEXT
            DoEvents
            '�z�X�g�������x������B
            NobbyString.AppendArray strttlTxt(), STR_TTL_COMMENT
            NobbyString.AppendArray strttlTxt(), STR_TTL_COMMENT & hostname


            '091128�Ή�=======================================================================================================
            'NMS����̃z�b�v��ւ�TELNET
            If ostype = "PF" Then
                hophostrownum = LNG_ROW_NUM_HOP_MACHINE
                hophostname = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_HOP_MACHINE).Value2 & STR_QUOAT
                hopipaddr = .Cells(hophostrownum, LNG_COL_NUM_HOP_IPADDR).Value2

                hopostype = .Cells(hophostrownum, LNG_COL_NUM_OSTYPE).Value2
                If .Cells(hophostrownum, LNG_COL_NUM_UID).Value2 <> NO_TEXT Then
                    hopuid = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_UID).Value2 & STR_QUOAT
                Else
                    hopuid = NO_TEXT
                End If
                hoppw = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_HOP_PW).Value2 & STR_QUOAT

                Do While .Cells(hophostrownum, LNG_COL_NUM_HOP_MACHINE).Value2 <> NO_TEXT
                    NobbyString.AppendArray strttlTxt(), STR_TTL_SENDLN & STR_QUOAT & STR_CMD_TELNET & hopipaddr & STR_QUOAT
                    If MakeLoginCommand(hopostype, hopuid, hoppw, strttlTxt()) = False Then
                        MsgBox "�Z���̓��͂ɕs��������܂��B���O�C��ID����уp�X���[�h�AOS�^�C�v���m�F���ĉ������B"
                        Exit Function
                    End If
                    hophostrownum = hophostrownum + 1
                    hophostname = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_HOP_MACHINE).Value2 & STR_QUOAT
                    hopipaddr = .Cells(hophostrownum, LNG_COL_NUM_HOP_IPADDR).Value2
                    If .Cells(hophostrownum, LNG_COL_NUM_UID).Value2 <> NO_TEXT Then
                        hopuid = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_UID).Value2 & STR_QUOAT
                    Else
                        hopuid = NO_TEXT
                    End If
                    hoppw = STR_QUOAT & .Cells(hophostrownum, LNG_COL_NUM_HOP_PW).Value2 & STR_QUOAT
                    hopcount = hopcount + 1
                Loop
            End If
            
            '�z�b�v�悩��Ώۋ@��ւ�TELNET
            NobbyString.AppendArray strttlTxt(), STR_TTL_SENDLN & STR_QUOAT & STR_CMD_TELNET & ipaddress & STR_QUOAT
            
            '�Ώۋ@��ւ̃��O�C������я����R�}���h����
            'IOS��PF�̂ǂ��炩�𔻕ʂ��ă��O�C���R�}���h�Z�b�g���쐬����B
            uid = STR_SYSID
            pw = STR_SYSPW
            If MakeLoginCommand(ostype, uid, pw, strttlTxt()) = False Then
                MsgBox "�Z���̓��͂ɕs��������܂��B���O�C��ID����уp�X���[�h�AOS�^�C�v���m�F���ĉ������B"
                Exit Function
            End If
            '================================================================================================================
            'prompt���擾����ttl���쐬
            MakeGetPromptTTL strttlTxt(), hostcount
            
            
            '�R�}���h���s����TTL���쐬
            NobbyString.AppendArray strttlTxt(), ";running-config�ۑ����s"
'            Do
'                DoEvents
'
                'OS�^�C�v�ɂ���ēǂݎ��R�}���h���X�g��؂�ւ���B
'                If ostype = STR_FLG_IOS Then                           '// NX�Ή�
                If ostype = STR_FLG_IOS Or ostype = STR_FLG_NX Then     '// NX�Ή�
'                    doprompt = STR_QUOAT & .Cells(cmdrownum, LNG_COL_NUM_PROMPT).Value2 & STR_QUOAT
                    doprompt = "hostname"
                    docommand = STR_QUOAT & "copy run tftp" & STR_QUOAT
                ElseIf ostype = STR_FLG_PF Then
'                    doprompt = STR_QUOAT & .Cells(cmdrownum, LNG_COL_NUM_PFPROMPT).Value2 & STR_QUOAT
'                    doprompt = "hostname"
'                    docommand = STR_QUOAT & .Cells(cmdrownum, LNG_COL_NUM_PFCOMMAND).Value2 & STR_QUOAT
                End If
                If docommand = NO_TEXT Or docommand = "''" Then
                    Exit Do
                End If
                '�R�}���h��쐬

                MakeExecuteCommand doprompt, docommand, strttlTxt()
                MakeExecuteCommand "'Address or name of remote host'", "'" & ipaddrNMS & "'", strttlTxt()
                MakeExecuteCommand "'Destination filename'", "'" & "/pd/network/tgkk/" & hostname & ".cfg" & "'", strttlTxt()
'                NobbyString.AppendArray strttlTxt(), "logclose"
'                cmdrownum = cmdrownum + 1
''                hostcount = hostcount + 1
'            Loop While docommand <> NO_TEXT
            
            'EXIT�R�}���h���� IOS��PF�𔻕ʂ��č쐬����B
            If MakeExitCommand(ostype, hopcount, strttlTxt(), hostcount) = False Then
                MsgBox "OS�^�C�v�̎w�肪����܂���B"
            End If
            
            NobbyString.AppendArray strttlTxt(), STR_TTL_WAIT & STR_NMSPROMPT
            '�R�}���h�s�̐擪�s�ԍ��֍s�ԍ�������������B
            cmdrownum = LNG_ROW_NUM_OFFSET_PROMPT

            '���̃z�X�g���̍s�֎Q�ƍs�ԍ����C���N�������g
            hostrownum = hostrownum + 1
            
            '���̃z�X�g����IP�A�h���X���擾
            hostname = .Cells(hostrownum, LNG_COL_NUM_HOSTNAME).Value2
            ipaddress = .Cells(hostrownum, LNG_COL_NUM_ADDR).Value2
'            If .Cells(hostrownum, LNG_COL_NUM_UID).Value2 <> NO_TEXT Then
'                uid = STR_QUOAT & .Cells(hostrownum, LNG_COL_NUM_UID).Value2 & STR_QUOAT
'            Else
'                uid = NO_TEXT
'            End If
'            pw = STR_QUOAT & .Cells(hostrownum, LNG_COL_NUM_PW).Value2 & STR_QUOAT
            ostype = .Cells(hostrownum, LNG_COL_NUM_OSTYPE).Value2
            hostcount = hostcount + 1
        Loop
        NobbyString.AppendArray strttlTxt(), ":END"
'        NobbyString.AppendArray strttlTxt(), "logclose"
        NobbyString.AppendArray strttlTxt(), "sendln 'exit'"
        NobbyString.AppendArray strttlTxt(), "messagebox 'config�ۑ��������I�����܂����B' '��������'"
        'strttlTxt()�̓��e��ttl�t�@�C���֕ۑ�����B
        Dim filename As String
        filename = .Cells(LNG_ROW_NUM_TTLFILENAME, LNG_COL_NUM_TTLFILENAME).Value2
        
        If FileActuators.WriteOnceFile(folderpath & filename, strttlTxt()) < 0 Then
            MsgBox STR_MSG_FILEPATH_ERR
            Exit Function
        End If
        Dim battxt() As String
'        SafeArrayAllocDescriptor 1, battxt()
        ReDim battxt(0)
        battxt(0) = ""
        NobbyString.AppendArray battxt(), """" & CStr(.Cells(59, 2).Value2) & """ /V " & """" & folderpath & CStr(.Cells(58, 2).Value2) & """"
        FileActuators.WriteOnceFile folderpath & filename & ".bat", battxt()
'        MsgBox STR_MSG_END1 & folderpath & filename & STR_MSG_END2
        CopyRunTftpStruct = folderpath & filename
    End With
    
End Function

'@brief     ���[�J��PC��ō�ƃf�B���N�g�����쐬����B
'@date      2009.08.16
'@author    Hiroki Nobumoto
'@return    �쐬�����f�B���N�g���t���p�X
Function MkWorkDir(ByVal oyafolder As String) As String

    Dim folderpath As String
    folderpath = FileActuators.GetFolderPath
    If folderpath = "" Then
        MkWorkDir = "�t�H���_���쐬�ł��܂���ł����B"
        Exit Function
    End If
    
    Dim getday As Variant
    getday = Date
    
    Dim strGetdate As String
    strGetdate = Format(getday, "yymmdd")
    folderpath = folderpath & STR_YEN & oyafolder & STR_YEN
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(folderpath) = True Then
        
    Else
        MkDir folderpath
    End If
    MkWorkDir = folderpath

End Function

'@brief     IOS��PF�̃��O�C�������terlen�܂ł̑���R�}���h�Z�b�g�����
'@date      2009.11.26
'@author    Hiroki Nobumoto
'@param1    OS�̃^�C�v�iIOS or PF�j
'@param2    USER ID
'@param3    PASSWORD
'@param5    �R�}���h�Z�b�g��ǉ�����string�z��
'@return
Private Function MakeLoginCommand(ByVal ostype As String, ByVal uid As String, ByVal pw As String, ByRef res() As String) As Boolean

'    If ostype = STR_FLG_IOS Or ostype = STR_FLG_PF Then                            '// NXOS�Ή�
    If ostype = STR_FLG_IOS Or ostype = STR_FLG_PF Or ostype = STR_FLG_NX Then      '// NXOS�Ή�
    
    Else
        MakeLoginCommand = False
        Exit Function
    End If
    With ActiveSheet
        If ostype = STR_FLG_IOS Then
           'IOS uid���o�^����Ă����ꍇ�̓��O�C����UID����͂���B
            If uid <> NO_TEXT Then
                NobbyString.AppendArray res(), STR_TTL_WAIT & STR_SWPROMPT_UIDREQUEST
                NobbyString.AppendArray res(), STR_TTL_SENDLN & uid
            End If
            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_SWPROMPT_PASSREQUEST
            NobbyString.AppendArray res(), STR_TTL_SENDLN & pw
        
            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_SWPROMPT_USER
            NobbyString.AppendArray res(), STR_TTL_SENDLN & STR_SWCMD_ENABLE
        
            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_SWPROMPT_PASSREQUEST
            NobbyString.AppendArray res(), STR_TTL_SENDLN & pw
            
            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_SWPROMPT_ENABLE
            NobbyString.AppendArray res(), STR_TTL_SENDLN & STR_SWCMD_TER_LEN
            MakeLoginCommand = True
        
'        ElseIf ostype = STR_FLG_PF Then                            '// NXOS�Ή� ���O�C������у��O�C�����PF�Ɠ����B
        ElseIf ostype = STR_FLG_PF Then      '// NXOS�Ή� ���O�C������у��O�C�����PF�Ɠ����B
            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_PFPROMPT_UID
            NobbyString.AppendArray res(), STR_TTL_SENDLN & uid
            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_PFPROMPT_PASSWD
            NobbyString.AppendArray res(), STR_TTL_SENDLN & pw
            'ACS�F�؂���Ƃ����Ȃ�ADMIN���[�h�Ń��O�C������ׁA�ȉ�4�s�̓��[�J���F�؎��ȊO�ł͕s�v
            '2009/11/29 nobumoto
'            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_PFPROMPT_USERMODE
'            NobbyString.AppendArray res(), STR_TTL_SENDLN & STR_PFCOMMAND_ADMINMODE
'            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_PFPROMPT_PASSWD
'            NobbyString.AppendArray res(), STR_TTL_SENDLN & pw
            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_PFPROMPT_ADMINMODE
            NobbyString.AppendArray res(), STR_TTL_SENDLN & "'set pager disable current'"
            MakeLoginCommand = True
        
        ElseIf ostype = STR_FLG_NX Then
            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_PFPROMPT_UID
            NobbyString.AppendArray res(), STR_TTL_SENDLN & uid
            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_PFPROMPT_PASSWD
            NobbyString.AppendArray res(), STR_TTL_SENDLN & pw
            'ACS�F�؂���Ƃ����Ȃ�ADMIN���[�h�Ń��O�C������ׁA�ȉ�4�s�̓��[�J���F�؎��ȊO�ł͕s�v
            '2009/11/29 nobumoto
'            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_PFPROMPT_USERMODE
'            NobbyString.AppendArray res(), STR_TTL_SENDLN & STR_PFCOMMAND_ADMINMODE
'            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_PFPROMPT_PASSWD
'            NobbyString.AppendArray res(), STR_TTL_SENDLN & pw
            NobbyString.AppendArray res(), STR_TTL_WAIT & "'#'"
            NobbyString.AppendArray res(), STR_TTL_SENDLN & "'ter len 0 '"
            MakeLoginCommand = True
        End If
        
        
        
        
       
    End With
End Function



'@brief     IOS��PF��EXIT����R�}���h�Z�b�g�����
'@date      2009.11.26
'@author    Hiroki Nobumoto
'@param1    OS�̃^�C�v�iIOS or PF�j
'@param2    USER ID
'@param3    PASSWORD
'@param4    �R�}���h�Z�b�g��ǉ�����string�z��
'@return
Private Function MakeExitCommand(ByVal ostype As String, ByVal hopcount As Long, ByRef res() As String, ByVal hostcount As Integer) As Boolean

    Dim i As Long
    If ostype <> STR_FLG_IOS And ostype <> STR_FLG_PF And ostype <> STR_FLG_NX Then
        MakeExitCommand = False
        Exit Function
    End If
    For i = 1 To hopcount Step 1
        DoEvents
        If ostype = STR_FLG_IOS Or ostype = STR_FLG_NX Or i <> 1 Then
            NobbyString.AppendArray res(), STR_TTL_WAIT & "hostname"
        ElseIf ostype = STR_FLG_PF Then
            NobbyString.AppendArray res(), STR_TTL_WAIT & STR_PFPROMPT_ADMINMODE
        End If
        NobbyString.AppendArray res(), ":NEXT" & CStr(hostcount)
        NobbyString.AppendArray res(), STR_TTL_SENDLN & STR_CWCMD_EXIT
    Next i
    MakeExitCommand = True
End Function




'@brief     ����R�}���h���P�Z�b�g���
'@date      2009.11.26
'@author    Hiroki Nobumoto
'@param1    prompt
'@param2    command
'@param3    �R�}���h�Z�b�g��ǉ�����string�z��
'@return
Private Function MakeExecuteCommand(ByVal doprompt As String, ByVal docommand As String, ByRef res() As String) As Boolean

    NobbyString.AppendArray res(), STR_TTL_WAIT & doprompt
    NobbyString.AppendArray res(), STR_TTL_SENDLN & docommand

End Function

'@brief     �v�����v�g���擾����ttl���쐬����B
'@author    Hiroki Nobumoto
'@date      2010.12.09
'@param     �쐬�����������ǉ�����string�z��
Private Sub MakeGetPromptTTL(ByRef strttlTxt() As String, ByVal hostcount As Integer)
    NobbyString.AppendArray strttlTxt(), ";recieve�ɂ܂��ϐ��̏�����"
    NobbyString.AppendArray strttlTxt(), "sendln ''"
    NobbyString.AppendArray strttlTxt(), "recvln"
    NobbyString.AppendArray strttlTxt(), "sendln ''"
    NobbyString.AppendArray strttlTxt(), "recvln"
    NobbyString.AppendArray strttlTxt(), "sendln ''"
    NobbyString.AppendArray strttlTxt(), "recvln"
    NobbyString.AppendArray strttlTxt(), "sendln ''"
    NobbyString.AppendArray strttlTxt(), "recvln"
    NobbyString.AppendArray strttlTxt(), ""
    NobbyString.AppendArray strttlTxt(), ";�W���o�́i�R���\�[���o�́j�̎擾"
    NobbyString.AppendArray strttlTxt(), "result=1"
    NobbyString.AppendArray strttlTxt(), "count = 10"
    NobbyString.AppendArray strttlTxt(), "hostname=''"
    NobbyString.AppendArray strttlTxt(), "while count>0"
    NobbyString.AppendArray strttlTxt(), "  count = count - 1"
    NobbyString.AppendArray strttlTxt(), "  sendln''"
    NobbyString.AppendArray strttlTxt(), "  recvln"
    NobbyString.AppendArray strttlTxt(), "  strcompare inputstr ''"
    NobbyString.AppendArray strttlTxt(), "  if result!=0 then"
    NobbyString.AppendArray strttlTxt(), "      hostname=inputstr"
    NobbyString.AppendArray strttlTxt(), "      break"
    NobbyString.AppendArray strttlTxt(), "  endif"
    NobbyString.AppendArray strttlTxt(), "endwhile"
    NobbyString.AppendArray strttlTxt(), "if count>0 then"
    NobbyString.AppendArray strttlTxt(), "  goto GETUSERPROMPT" & CStr(hostcount)
    NobbyString.AppendArray strttlTxt(), "endif"
    NobbyString.AppendArray strttlTxt(), ":ERRENDMARK" & CStr(hostcount)
    NobbyString.AppendArray strttlTxt(), "messagebox '�z�X�g���̎擾���ł��܂���ł����B' 'error'"
    NobbyString.AppendArray strttlTxt(), "goto NEXT" & CStr(hostcount)
    NobbyString.AppendArray strttlTxt(), ""
    NobbyString.AppendArray strttlTxt(), ":GETUSERPROMPT" & CStr(hostcount)


End Sub







'@brief     �v�����v�g���擾����ttl���쐬����B
'@author    Hiroki Nobumoto
'@date      2010.12.09
'@param     �쐬�����������ǉ�����string�z��
Private Sub MakeGetSYSIDTTL(ByRef strttlTxt() As String)
    NobbyString.AppendArray strttlTxt(), ";SYSID�ɂ܂��ϐ��̏�����"
    NobbyString.AppendArray strttlTxt(), "result=1"
    NobbyString.AppendArray strttlTxt(), "count=0"
    NobbyString.AppendArray strttlTxt(), STR_SYSID & "=''"
    NobbyString.AppendArray strttlTxt(), "while count<10"
    NobbyString.AppendArray strttlTxt(), "  count=count+1"
    NobbyString.AppendArray strttlTxt(), "  sendln'/fbknms/appl/proc/ebxr810.sh'"
    NobbyString.AppendArray strttlTxt(), "  recvln"
    NobbyString.AppendArray strttlTxt(), "  recvln"
    NobbyString.AppendArray strttlTxt(), "  strcompare inputstr ''"
    NobbyString.AppendArray strttlTxt(), "  if result!=0 then"
    NobbyString.AppendArray strttlTxt(), "      " & STR_SYSID & "=inputstr"
'    NobbyString.AppendArray strttlTxt(), "      messagebox SYSID 'SYSID'"
    NobbyString.AppendArray strttlTxt(), "      goto GETPASSWORD"
    NobbyString.AppendArray strttlTxt(), "      endif"
    NobbyString.AppendArray strttlTxt(), "  if count>5 goto END"
    NobbyString.AppendArray strttlTxt(), "endwhile"
    NobbyString.AppendArray strttlTxt(), ":GETPASSWORD"
    NobbyString.AppendArray strttlTxt(), "result=1"
    NobbyString.AppendArray strttlTxt(), "count=0"
    NobbyString.AppendArray strttlTxt(), STR_SYSPW & "=''"
    NobbyString.AppendArray strttlTxt(), "while count<10"
    NobbyString.AppendArray strttlTxt(), "  count=count+1"
    NobbyString.AppendArray strttlTxt(), "  sendln'/fbknms/appl/proc/ebxr812.sh'"
    NobbyString.AppendArray strttlTxt(), "  recvln"
    NobbyString.AppendArray strttlTxt(), "  recvln"
    NobbyString.AppendArray strttlTxt(), "  strcompare inputstr ''"
    NobbyString.AppendArray strttlTxt(), "  if result!=0 then"
    NobbyString.AppendArray strttlTxt(), "      " & STR_SYSPW & "=inputstr"
'    NobbyString.AppendArray strttlTxt(), "      messagebox SYSPW 'SYSPW'"
    NobbyString.AppendArray strttlTxt(), "      goto LOGINALL"
    NobbyString.AppendArray strttlTxt(), "      endif"
    NobbyString.AppendArray strttlTxt(), "  if count>5 goto END"
    NobbyString.AppendArray strttlTxt(), "endwhile"
    NobbyString.AppendArray strttlTxt(), ":LOGINALL"


End Sub















