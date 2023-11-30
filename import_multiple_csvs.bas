Attribute VB_Name = "Module1"

Sub import_multiple_csvs()
Application.ScreenUpdating = False
Dim fso As FileSystemObject
Dim ts As TextStream
gfn = get_file_names(file_path = "C:\Users\giles.field\Desktop\temp")
Set fso = New FileSystemObject

For i = LBound(gfn) To UBound(gfn) - 1
    If fso.FileExists(gfn(i)) = False Then GoTo line1 Else
    Set ts = fso.OpenTextFile(gfn(i), ForReading, False, TristateUseDefault)
    If i <> LBound(gfn) Then ts.SkipLine Else
    Do While ts.AtEndOfStream = False
        row_cnt = row_cnt + 1
        Cells(row_cnt, 1) = ts.ReadLine '& ", " & gfn(i)
    Loop
line1:
    ts.Close
Next i

Set ts = Nothing
Set fso = Nothing

End Sub

Function get_file_names(ByVal file_path As String) As Variant
'OPENS FILES
Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
    .InitialFileName = file_path
    .ButtonName = "Select"
    .AllowMultiSelect = True
    .Title = "Choose file"
    .InitialView = msoFileDialogViewDetails
    .Filters.Add "CSV", "*.csv", 1
    .Show
    Dim a As Variant
    ReDim a(0 To .SelectedItems.Count)
    i = 0
    For Each objfl In .SelectedItems
        a(i) = objfl

        i = i + 1
    Next objfl
    On Error GoTo 0
End With
get_file_names = a
End Function