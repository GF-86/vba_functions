Attribute VB_Name = "Module3"
Sub ImportCSVFilesAsTabs()
    Dim dialog As FileDialog
    Dim selectedFiles As FileDialogSelectedItems
    Dim wb As Workbook
    Dim fileName As Variant
    Dim ws As Worksheet
    
    ' Create a FileDialog object to select CSV files
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    dialog.AllowMultiSelect = True
    dialog.Title = "Select CSV Files"
    dialog.Filters.Clear
    dialog.Filters.Add "CSV Files", "*.csv"
    
    ' Show the dialog and get selected files
    If dialog.Show = -1 Then
        Set selectedFiles = dialog.SelectedItems
    Else
        Exit Sub
    End If
    
    ' Create a new workbook
    Set wb = Workbooks.Add
    
    ' Loop through selected files
    For Each fileName In selectedFiles
        ' Import CSV data into a new worksheet
        Set ws = wb.Worksheets.Add
        With ws.QueryTables.Add(Connection:= _
            "TEXT;" & fileName, Destination:=Range("$A$1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .Refresh
        End With
        
        ' Rename the worksheet with the file name (excluding extension)
        ws.Name = Left(Left$(Dir(fileName), InStrRev(Dir(fileName), ".") - 1), 30)
    Next fileName
        
    ' Clean up
    Set ws = Nothing
    Set dialog = Nothing
    
    MsgBox "CSV files imported as tabs in a workbook.", vbInformation
End Sub


