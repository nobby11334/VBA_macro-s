Attribute VB_Name = "FileActuators"
Private Const STR_ERR_MSG_TARGET_FILE As String = "参照ファイル："

'@brief     ファイル選択ダイアログを開き、選択したファイルのパスを返す
'@author    Hiroki Nobumoto
'@date      2008/02/20
Function GetFilePath() As String

    Dim OpenFileName As Variant
    Dim filenameFullpath As String
    
    OpenFileName = Application.GetOpenFilename(TITLE_FILESELECTDIALOG)
    If OpenFileName <> False Then
        filenameFullpath = CStr(OpenFileName)
        Debug.Print filenameFullpath
        
        GetFilePath = filenameFullpath
    Else
    
        GetFilePath = "error file read"
        
    End If
    
End Function

'@brief     フォルダ選択ダイアログを表示して、選択したフォルダパスを返す。
'@author    Hiroki Nobumoto
'@date      2008/12/15
'@return    選択したフォルダのフルパス。ダイアログでキャンセルしたら何もない。
Function GetFolderPath() As String
    Dim ShellApp As Object
    Dim oFolder As Object
    Set ShellApp = CreateObject("Shell.Application")
    Set oFolder = ShellApp.BrowseForFolder(0, "フォルダ選択", 1)
    
    If oFolder Is Nothing Then
        GetFolderPath = ""
        Exit Function
    End If
    
    GetFolderPath = oFolder.items.Item.path


End Function


'@brief     引数の2ファイルをマージする。第一引数が先頭行側
'@author    Hiroki Nobumoto
'@date      2009/11/09
'@param1    マージするファイルの先頭行側のファイルパス
'@param2    マージするファイルの後側のファイルパス
'@param3    マージしたファイルのファイルパス
'@return    マージしたファイルの名称
Public Function MargeToFile(ByVal fileA As String, ByVal fileB As String, ByVal fileC As String) As String
    If fileA = "" Then
        MargeToFile = "error"
        Exit Function
    End If
    
    If fileB = "" Then
        MargeToFile = "error"
        Exit Function
    End If
    
    If fileC = "" Then
        MargeToFile = "error"
        Exit Function
    End If
    
    Dim fileAtxt() As String
    Dim fileBtxt() As String
    Dim fileCtxt() As String
    ReDim fileAtxt(0)
    ReDim fileBtxt(0)
    
    ReadOnceFile fileA, fileAtxt()
    ReadOnceFile fileB, fileBtxt()
    
    Dim fileArowMax As Long
    Dim fileBrowMax As Long
    
    fileArowMax = UBound(fileAtxt)
    fileBrowMax = UBound(fileBtxt)
    
    Dim i As Long
    Dim j As Long
    Dim max As Long
    max = fileArowMax + fileBrowMax + 1
    j = 0
    
    For i = 1 To max Step 1
        DoEvents
        ReDim Preserve fileAtxt(fileArowMax + i)
        fileAtxt(fileArowMax + i) = fileBtxt(j)
        j = j + 1
        If j > fileBrowMax Then
            Exit For
        End If
    Next i
    
    If FileActuators.WriteOnceFile(fileC, fileAtxt()) < 0 Then
        MargeToFile = "error"
        Exit Function
    End If
    
    MargeToFile = fileC

End Function


'@brief     ファイル作成
'@author    Hiroki Nobumoto
'@date      2008/02/26
'@param     path        : 作成するファイルのパス
'@param     document()  : 書き込む内容への参照
'@return    0: 正常終了　-1: 書き込む内容が無いエラー
Function WriteOnceFile(path As String, ByRef document() As String) As Integer
'    Dim fstreamWrite
    Dim factualWrite
    
    Dim dirpath As String
    Dim splitpath() As String
    
    Dim docRowsCount As Long    '書き込むドキュメント行数格納用
    Dim i As Integer
    
'    On Error GoTo errhandler
    
'    Set fstreamWrite = CreateObject("Scripting.FileSystemObject")
    Dim fstreamWrite As New Scripting.FileSystemObject
    

    If InStr(path, STR_YEN) = 0 Then
        WriteOnceFile = -1
        Exit Function
    End If
    splitpath() = Split(path, STR_YEN)
    dirpath = Replace(path, splitpath(UBound(splitpath)), "")
    
    If fstreamWrite.FolderExists(dirpath) Then
    
    Else
        MkDir (dirpath)
    End If
    fstreamWrite.CreateTextFile path
    Set factualWrite = fstreamWrite.Getfile(path)
    Dim filestreamResultWrite
    Set filestreamResultWrite = factualWrite.OpenAsTextStream(2, -2)
    
    If IsArray(document) = False Then
        WriteOnceFile = -1
        Exit Function
    End If

    docRowsCount = UBound(document)

    For i = 0 To docRowsCount Step 1
        filestreamResultWrite.WriteLine document(i)
    Next i
    
    WriteOnceFile = 0

errhandler:
    
    If Err.Number > 0 Then
        ErrorProcess.DisplayErrorMessage Err.Number, Err.Description, "WriteOnceFile"
    End If
    
End Function

'@brief     ファイル読み込み
'@author    Hiroki Nobumoto
'@date      2008/02/26
'@param     path        : 読み込むファイルのパス
'@param     document()  : 読み込む内容を格納する配列への参照
'@return    0: 正常終了　-1: 指定したパスが見つかりません。
Function ReadOnceFile(path As String, ByRef document() As String)

'    Dim fstreamread
    Dim factualread
    Dim i As Integer        '行数カウンター
        
'    On Error GoTo errhandler
    
    ReDim document(0)
    
'    Set fstreamread = CreateObject("Scripting.FileSystemObject")
    Dim fstreamread As New Scripting.FileSystemObject
    

    Set factualread = fstreamread.Getfile(path)
    Dim fstreamResultRead
    Set fstreamResultRead = factualread.OpenAsTextStream(1, -2)
    
    i = 0
    Do While fstreamResultRead.AtEndOfStream = False
        ReDim Preserve document(i)
        document(i) = fstreamResultRead.ReadLine
        i = i + 1
    Loop
    ReadOnceFile = 0

errhandler:
    'エラー処理　エラーメッセージ作成＆メッセージボックス表示
    If Err.Number > 0 Then
        ErrorProcess.DisplayErrorMessage Err.Number, Err.Description, STR_ERR_MSG_TARGET_FILE + path
    End If
End Function
