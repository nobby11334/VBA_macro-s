Attribute VB_Name = "StrConverter"
Private Const NAME_SHEET As String = "Sheet1"
Private Const COLNO_SUP2 As Integer = 1
Private Const COLNO_SUP720 As Integer = 2
Private Const COLNO_MACHPATTERN As Integer = 3
Private Const LBL_CURRENT_ROW_NO As String = "�������̍s�ԍ��F "


Sub FromSup2ToSup720(ByRef fileBuf() As String)


    '�ϊ��e�[�u�����Q�Ƃ��ĕϊ����鏈��
    Dim i As Integer
    Dim j As Integer
    i = 0
    j = 2
    For i = 0 To UBound(fileBuf) Step 1
        DoEvents
        FormMain.Label_CurRowNo.Caption = LBL_CURRENT_ROW_NO & CStr(i)
        Do While Worksheets(NAME_SHEET).Cells(j, COLNO_SUP2).Value2 <> ""
            DoEvents
            If InStr(fileBuf(i), STR_SHARP) <> 0 Then Exit Do
            
            '�����񂪂m�t�k�k���R�����g�s�̏ꍇ�͏������Ȃ�
            If fileBuf(i) <> NO_TEXT And InStr(fileBuf(i), STR_SHARP) = 0 And InStr(fileBuf(i), STR_EXCLAMATIONMARK) = 0 Then
                If InStr(fileBuf(i), Worksheets(NAME_SHEET).Cells(j, COLNO_SUP2).Value2) Then
                    '���S�}�b�`���Ȃ���΂Ȃ�Ȃ��ꍇ
                    If Worksheets(NAME_SHEET).Cells(j, COLNO_MACHPATTERN).Value2 = 1 Then
                        If Len(fileBuf(i)) = Len(Worksheets(NAME_SHEET).Cells(j, COLNO_SUP2).Value2) Then
                            fileBuf(i) = Replace(fileBuf(i), Worksheets(NAME_SHEET).Cells(j, COLNO_SUP2).Value2, Worksheets(NAME_SHEET).Cells(j, COLNO_SUP720).Value2)
                            j = 2
                            Exit Do
                        End If
                    '���S�}�b�`���Ȃ��Ă��ǂ��ꍇ
                    Else
                        fileBuf(i) = Replace(fileBuf(i), Worksheets(NAME_SHEET).Cells(j, COLNO_SUP2).Value2, Worksheets(NAME_SHEET).Cells(j, COLNO_SUP720).Value2)
                        j = 2
                        Exit Do
                    End If
                End If
            End If
            j = j + 1
        Loop
        j = 2
    Next i
End Sub

