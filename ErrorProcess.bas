Attribute VB_Name = "ErrorProcess"
Private Const STR_ERR_NO As String = "エラー番号："


'@brief     エラートラップ時のエラーメッセージ表示処理
'@author    Hiroki Nobumoto
'@date      2008/03/12
'@param     errNo : トラップされたエラー番号
'@param     errMsg : トラップされたエラーメッセージ
'@param     errState : エラー発生時のパラメータ
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
