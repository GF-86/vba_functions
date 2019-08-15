Attribute VB_Name = "Module1"

Sub query_db(ByVal query_str As String, ByVal worksheet_name As String)
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strConnString As String
    Dim ws As Worksheet
    Set ws = Worksheets(worksheet_name)
    
        strConnString = "Provider=SQLOLEDB;Data Source=DESKTOP-CLTNVDD\SQLEXPRESS;" _
                        & "Initial Catalog=gtfs;Integrated Security=SSPI;"
        
        Set conn = New ADODB.Connection
        conn.Open strConnString

        Set rs = conn.Execute(query_str)
        
        
        For iCols = 0 To rs.Fields.Count - 1
        ws.Cells(1, iCols + 1).Value = rs.Fields(iCols).Name
        Next
        ws.Range(ws.Cells(1, 1), ws.Cells(1, rs.Fields.Count)).Font.Bold = True
        ws.Range("A2").CopyFromRecordset Data:=rs, MaxRows:=1000000
        rs.Close

End Sub
