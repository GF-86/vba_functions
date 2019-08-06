Attribute VB_Name = "Module1"
Function get_file_name() As String
'OPENS FILES
Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
    .InitialFileName = "C:\Users"
    .ButtonName = "Select"
    .AllowMultiSelect = False
    .Title = "Choose file"
    .InitialView = msoFileDialogViewDetails
    .Filters.Add "KML", "*.kml", 1
    .Show
    For Each objfl In .SelectedItems
        get_file_name = objfl
    Next objfl
    On Error GoTo 0
End With

End Function