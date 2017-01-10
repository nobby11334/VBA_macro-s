Attribute VB_Name = "WorkBookActuator"


'@brief     対象のワークブックのワークシートから値を検索してオブジェクトとして返す。
'@author    Hiroki Nobumoto
'@date      2008/03/11
'@param     wantSearchTarget : ワークシート上で検索したいデータ
'@param     strWorkBookName : 検索したいワークブック名称
'@param     workSheetIndex : 検索したいワークシート名称またはインデックス番号
'@param     lngRowNo : 検索にヒットした値の格納されているセルの行番号
'@param     intColumnNo : 検索にヒットした値の格納されているセルの列番号
Function FindHostNameCell(wantSearchTarget As Variant, strWorkBookName As String, workSheetIndex As Variant, ByRef lngRowNo As Long, ByRef intColumnNo As Integer) As Boolean

    Dim resultfind As Range

    '現調ツールから取得したホスト名を含む行の番号を検索する
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



