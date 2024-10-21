Attribute VB_Name = "MJoin"
'namespace=vba-files\

' �׳� Ǯ�ƿ������� �޸�
Public Sub DoJoin()

    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strConn As String
    Dim i As Integer
     
    'Dim ��Ʈ��_��
    
    'strSQL = ThisWorkbook.Sheets("SQL").Range("D1")

    strSQL = _
    "SELECT *" & vbCrLf & _
    "FROM [��Ʈ1$] [A]" & vbCrLf & _
    "LEFT OUTER JOIN [��Ʈ2$] [B]" & vbCrLf & _
    "ON [A].[�Ӽ�] = [B].[�Ӽ�]" & vbCrLf & _
    "UNION ALL" & vbCrLf & _
    "SELECT *" & vbCrLf & _
    "FROM [��Ʈ1$] [A]" & vbCrLf & _
    "RIGHT OUTER JOIN [��Ʈ2$] [B]" & vbCrLf & _
    "ON [A].[�Ӽ�] = [B].[�Ӽ�]" & vbCrLf & _
    "WHERE [A].[�Ӽ�] IS NULL"

    strSQL = Replace(strSQL, "��Ʈ1", "��") '�ϵ��ڵ�
    strSQL = Replace(strSQL, "��Ʈ2", "��") '�ϵ��ڵ�
    strSQL = Replace(strSQL, "�Ӽ�", "a") '�ϵ��ڵ�
    
    'Excel�� Database�� ���
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\" & ActiveWorkbook.Name & ";" & "Extended Properties=Excel 12.0;"
      
    rs.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
      
    If Err.Number <> 0 Then
        MsgBox "SQL ȣ�⿡�� ������ �߻��߽��ϴ�. ����� �����ϰ� �����Ǿ� ���� ���� ���ɼ��� �����ϴ�."
    End If
       
    '��Ʈ�� ������ �����ϴ� �ڵ�
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("RESULT").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Call Worksheets.Add(After:=Worksheets(Worksheets.Count))
    Worksheets(Worksheets.Count).Name = "RESULT"
    
    If rs.EOF Then
        Call MsgBox("��ȸ�� ���ڵ尡 �����ϴ�.") 'FULL OUTER JOIN�� �и��Կ� ����, �Ϻ� ������ �������� �����Ƿ� �ּ�ǥ���� (230823)
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
