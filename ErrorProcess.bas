Attribute VB_Name = "ErrorProcess"
Private Const STR_ERR_NO As String = "�G���[�ԍ��F"


'@brief     �G���[�g���b�v���̃G���[���b�Z�[�W�\������
'@author    Hiroki Nobumoto
'@date      2008/03/12
'@param     errNo : �g���b�v���ꂽ�G���[�ԍ�
'@param     errMsg : �g���b�v���ꂽ�G���[���b�Z�[�W
'@param     errState : �G���[�������̃p�����[�^
Sub DisplayErrorMessage(errNo As Long, errMsg As String, errState As Variant)
    Dim strReturn As String
    strReturn = Chr(13) + Chr(10)
    
    strErrorMessage = STR_ERR_NO
    strErrorMessage = strErrorMessage + CStr(Err.Number)
    strErrorMessage = strErrorMessage + strReturn
    strErrorMessage = strErrorMessage + CStr(Err.Description)
    strErrorMessage = strErrorMessage + strReturn
    strErrorMessage = strErrorMessage + CStr(errState)
    
    MsgBox strErrorMessage
End Sub
