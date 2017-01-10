Attribute VB_Name = "FileActuators"
Private Const STR_ERR_MSG_TARGET_FILE As String = "�Q�ƃt�@�C���F"

'@brief     �t�@�C���I���_�C�A���O���J���A�I�������t�@�C���̃p�X��Ԃ�
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

'@brief     �t�H���_�I���_�C�A���O��\�����āA�I�������t�H���_�p�X��Ԃ��B
'@author    Hiroki Nobumoto
'@date      2008/12/15
'@return    �I�������t�H���_�̃t���p�X�B�_�C�A���O�ŃL�����Z�������牽���Ȃ��B
Function GetFolderPath() As String
    Dim ShellApp As Object
    Dim oFolder As Object
    Set ShellApp = CreateObject("Shell.Application")
    Set oFolder = ShellApp.BrowseForFolder(0, "�t�H���_�I��", 1)
    
    If oFolder Is Nothing Then
        GetFolderPath = ""
        Exit Function
    End If
    
    GetFolderPath = oFolder.items.Item.path


End Function


'@brief     ������2�t�@�C�����}�[�W����B���������擪�s��
'@author    Hiroki Nobumoto
'@date      2009/11/09
'@param1    �}�[�W����t�@�C���̐擪�s���̃t�@�C���p�X
'@param2    �}�[�W����t�@�C���̌㑤�̃t�@�C���p�X
'@param3    �}�[�W�����t�@�C���̃t�@�C���p�X
'@return    �}�[�W�����t�@�C���̖���
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


'@brief     �t�@�C���쐬
'@author    Hiroki Nobumoto
'@date      2008/02/26
'@param     path        : �쐬����t�@�C���̃p�X
'@param     document()  : �������ޓ��e�ւ̎Q��
'@return    0: ����I���@-1: �������ޓ��e�������G���[
Function WriteOnceFile(path As String, ByRef document() As String) As Integer
'    Dim fstreamWrite
    Dim factualWrite
    
    Dim dirpath As String
    Dim splitpath() As String
    
    Dim docRowsCount As Long    '�������ރh�L�������g�s���i�[�p
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

'@brief     �t�@�C���ǂݍ���
'@author    Hiroki Nobumoto
'@date      2008/02/26
'@param     path        : �ǂݍ��ރt�@�C���̃p�X
'@param     document()  : �ǂݍ��ޓ��e���i�[����z��ւ̎Q��
'@return    0: ����I���@-1: �w�肵���p�X��������܂���B
Function ReadOnceFile(path As String, ByRef document() As String)

'    Dim fstreamread
    Dim factualread
    Dim i As Integer        '�s���J�E���^�[
        
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
    '�G���[�����@�G���[���b�Z�[�W�쐬�����b�Z�[�W�{�b�N�X�\��
    If Err.Number > 0 Then
        ErrorProcess.DisplayErrorMessage Err.Number, Err.Description, STR_ERR_MSG_TARGET_FILE + path
    End If
End Function
