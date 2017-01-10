Attribute VB_Name = "NobbyString"
Private Const STR_REGEXP_CALL As String = "VBScript.RegExp"
    
Public arrayappendflg As Boolean



'@brief     第一引数の文字列を第二引数で分割して第三引数で指定した番号の文字列を返す。
'@date      2009.08.27
'@author    Hiroki Nobumoto
Public Function GetSplit(ByVal strSource As String, ByVal strSplitter As String, ByVal intGetNum As Integer) As String
    Dim splitStrs() As String
    splitStrs() = Split(strSource, strSplitter)
    GetSplit = splitStrs(intGetNum)
End Function

'@brief     配列にデータを格納して要素数を追加する。
'@date      2009.08.16
'@author    Hiroki Nobumoto
'@param     strTarget() : データを追加される対象配列への参照
'@param     inputdata : 追加するデータ
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
            '---なにもしない---
        Else
            ReDim Preserve res(UBound(res) + 1)
        End If
    Else
        ReDim Preserve res(UBound(res) + 1)
    End If
    res(UBound(res)) = inputdata
    strTarget(UBound(strTarget)) = res(UBound(res))
End Sub

'@brief     二つの配列をマージする。
'@auth      Hiroki Nobumoto
'@date      2009.10.02
'@param1    マージする先頭側の配列
'@param2    マージする後側の配列
'@param3    マージした結果
Public Sub MargeStr(ByRef strA() As String, ByRef strB() As String, ByRef strAB() As String)

    If UBound(strA) < 0 Or UBound(strB) < 0 Then
        MsgBox "マージする元のデータがありません。"
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



'@brief     配列が空っぽの場合-1を返すubound
'@author    Hiroki Nobumoto
'@date       2009.10.03
'@param     対象配列（空っぽでも可）
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

'@brief     第一引数のstring配列から第二引数で指定したコマンドで取得したステータスだけを抽出して第三引数へ返す。
'@date      2009.08.27
'@author    Hiroki Nobumoto
'@param     ByRef strall() As String        : 抽出したいテキストを含むファイルから取得したテキスト全部
'@param     ByVal strKey As String          : 対象ファイルから抽出したいテキストの抽出開始キーワード（一意なキーワードを指定する。）
'@param     ByVal strKeyEnd As String       : 対象ファイルから抽出したいテキストの抽出終了キーワード
'@param     ByRef strTargetData() As String : 抽出したテキストを格納するstring配列への参照
'@return    true: 処理成功  false:  処理対象が無い。
Public Function GetTxtPart(ByRef strAll() As String, ByVal strKey As String, ByVal strKeyEnd As String, ByRef strTargetData() As String) As Boolean

    '-----エラーチェック-----↓
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
    '-----------------------↑
    
    '-----全テキストを対象にチェック開始-----↓
    Dim execflg As Boolean          'true: 読み取った文字列をstrTargetData()にバッファして良い。
                                    'false: 読み取った文字列はstrTargetData()にバッファしない。
    execflg = False
    Dim txtrownummax As Long
    txtrownummax = UBound(strAll)
    
    Dim txtrownum As Long
    
    ReDim strTargetData(1)
    For txtrownum = 0 To txtrownummax Step 1
        DoEvents
        If execflg = True Then
        
            NobbyString.AppendArray strTargetData(), strAll(txtrownum)
            
            'バッファ終了可否確認
            If InStr(strAll(txtrownum + 1), strKeyEnd) > 0 Or txtrownum = txtRowMax Then
            
                execflg = False
            
            Else
            
                'なにもしない
                
            End If
            
        Else
            'バッファ開始可否確認
            If InStr(strAll(txtrownum), strKey) > 0 Then
            
                'バッファ可否フラグON
                execflg = True
            
            Else
            
                'なにもしない
                
            End If
            
        End If
        
    Next txtrownum
    '---------------------------------------↑
    
    
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

'@brief     第一引数の中に第二引数で指定した文字列を検索する。（正規表現もOK）
'@aurhot    Hiroki Nobumoto
'@date      2009.08.27
'@param     strTarget : 検索される文字列
'@param     strPattern : 検索したい文字列パターン
'@return    true : 検索にヒット　false :　検索にヒットしない。
Function FindRegExp(ByVal strTarget As String, ByVal strPattern As String) As Boolean
    Set re = CreateObject(STR_REGEXP_CALL)
    Dim res As Boolean
    re.Pattern = strPattern
    res = re.test(strTarget)
    Set re = Nothing
    FindRegExp = res
End Function




'@brief     第一引数の中に第二引数で指定した文字列を検索して、そのパターンに位置する文字列を返す。（正規表現もOK）
'@aurhot    Hiroki Nobumoto
'@date      2009.08.27
'@param     strTarget : 検索される文字列
'@param     strPattern : 検索したい文字列パターン
'@param     正規表現にマッチした文字列。複数の場合も格納可。
'@return    true : 検索にヒット　false :　検索にヒットしない。
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




