Attribute VB_Name = "MJoin"
'namespace=vba-files\

' 그냥 풀아우터조인 달림
Public Sub DoJoin()

    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strConn As String
    Dim i As Integer
     
    'Dim 시트명_비교
    
    'strSQL = ThisWorkbook.Sheets("SQL").Range("D1")

    strSQL = _
    "SELECT *" & vbCrLf & _
    "FROM [시트1$] [A]" & vbCrLf & _
    "LEFT OUTER JOIN [시트2$] [B]" & vbCrLf & _
    "ON [A].[속성] = [B].[속성]" & vbCrLf & _
    "UNION ALL" & vbCrLf & _
    "SELECT *" & vbCrLf & _
    "FROM [시트1$] [A]" & vbCrLf & _
    "RIGHT OUTER JOIN [시트2$] [B]" & vbCrLf & _
    "ON [A].[속성] = [B].[속성]" & vbCrLf & _
    "WHERE [A].[속성] IS NULL"

    strSQL = Replace(strSQL, "시트1", "전") '하드코딩
    strSQL = Replace(strSQL, "시트2", "후") '하드코딩
    strSQL = Replace(strSQL, "속성", "a") '하드코딩
    
    'Excel을 Database로 사용
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\" & ActiveWorkbook.Name & ";" & "Extended Properties=Excel 12.0;"
      
    rs.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
      
    If Err.Number <> 0 Then
        MsgBox "SQL 호출에서 에러가 발생했습니다. 헤더가 적절하게 설정되어 있지 않을 가능성이 높습니다."
    End If
       
    '시트가 있으면 삭제하는 코드
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("RESULT").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Call Worksheets.Add(After:=Worksheets(Worksheets.Count))
    Worksheets(Worksheets.Count).Name = "RESULT"
    
    If rs.EOF Then
        Call MsgBox("조회된 레코드가 없습니다.") 'FULL OUTER JOIN을 분리함에 따라, 일부 구문이 없을수도 있으므로 주석표시함 (230823)
    Else
        For i = 0 To rs.Fields.Count - 1
            Cells(1, i + 1).Value = rs.Fields(i).Name
        Next
        
        rs.MoveFirst
        
        ActiveSheet.Range("A2").CopyFromRecordset rs
    
    End If
     
    rs.Close
    Set rs = Nothing

End Sub
