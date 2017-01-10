Attribute VB_Name = "WorkBookActuator"


'@brief     �Ώۂ̃��[�N�u�b�N�̃��[�N�V�[�g����l���������ăI�u�W�F�N�g�Ƃ��ĕԂ��B
'@author    Hiroki Nobumoto
'@date      2008/03/11
'@param     wantSearchTarget : ���[�N�V�[�g��Ō����������f�[�^
'@param     strWorkBookName : �������������[�N�u�b�N����
'@param     workSheetIndex : �������������[�N�V�[�g���̂܂��̓C���f�b�N�X�ԍ�
'@param     lngRowNo : �����Ƀq�b�g�����l�̊i�[����Ă���Z���̍s�ԍ�
'@param     intColumnNo : �����Ƀq�b�g�����l�̊i�[����Ă���Z���̗�ԍ�
Function FindHostNameCell(wantSearchTarget As Variant, strWorkBookName As String, workSheetIndex As Variant, ByRef lngRowNo As Long, ByRef intColumnNo As Integer) As Boolean

    Dim resultfind As Range

    '�����c�[������擾�����z�X�g�����܂ލs�̔ԍ�����������
    Set resultfind = Workbooks(strWorkBookName).Worksheets(workSheetIndex).Range("k11:k500").Find(wantSearchTarget)
    
    If Not resultfind Is Nothing Then
        
        intColumnNo = resultfind.Column
        lngRowNo = resultfind.Row
        FindHostNameCell = True
        
    Else
    
        intColumnNo = 0
        lngRowNo = 0
        FindHostNameCell = False
        
    End If
End Function


Public Function GetSheetName() As String
    GetSheetName = ActiveSheet.Name
End Function



