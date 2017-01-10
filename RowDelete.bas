Attribute VB_Name = "RowDelete"
Sub SecondRowDeleteSub(ByVal startRowNo As Long, ByVal sheetName As String)
    Dim i As Long
    Dim flgEmp As Boolean
    Worksheets(sheetName).Select
    flgEmp = IsEmpty(Cells(startRowNo, 1).Value2)
    Do While flgEmp = False
        DoEvents
        startRowNo = startRowNo + 1
        Rows(startRowNo).Delete
        flgEmp = IsEmpty(Cells(startRowNo, 1).Value2)
    Loop
End Sub

Sub ColorRowDeleteSub(ByVal startRowNo As Long, ByVal sheetName As String, ByVal delColor As Variant)
    Dim i As Long
    Dim colorNo As Variant
    Dim counter As Long
    Dim flgEmp As Boolean
    Worksheets(sheetName).Select
    flgEmp = IsEmpty(Cells(startRowNo, 1).Value2)
    counter = startRowNo
    Do While flgEmp = False
        DoEvents
        colorNo = Cells(counter, 1).Interior.ColorIndex
        If colorNo = 15 Then
            Rows(counter).Delete
        Else
            counter = counter + 1
        End If
        flgEmp = IsEmpty(Cells(counter, 1).Value2)
    Loop
End Sub

Sub SecondRowdeleteMain()
    SecondRowDeleteSub 10, "EG"
    MsgBox "10�s�ڂ���1�s�Ƃі��ɍs�폜���܂����B"
End Sub

Sub ColorRowDeleteMain()
    ColorRowDeleteSub 10, "EG", 15
    MsgBox "�O���[�A�E�g�̍s�̂ݍ폜���܂����B"
End Sub
