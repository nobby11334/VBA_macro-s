Attribute VB_Name = "WindowActuator"
Public Sub AddSheet()
Attribute AddSheet.VB_Description = "�}�N���L�^�� : 2010/12/7  ���[�U�[�� : p010758"
Attribute AddSheet.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
' �}�N���L�^�� : 2010/12/7  ���[�U�[�� : p010758
'

    Dim sheetscount As Long
    
    sheetscount = Worksheets.Count
    Sheets(sheetscount).Copy after:=Sheets(sheetscount)
    Worksheets(sheetscount + 1).Name = CStr(sheetscount + 1)
    Worksheets(sheetscount + 1).Cells(1, 13).Value2 = Worksheets(1).Cells(1, 13).Value2 & "-" & CStr(sheetscount + 1)
    

End Sub

