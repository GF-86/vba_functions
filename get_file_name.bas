Attribute VB_Name = "Module1"
Function get_file_name(ByVal filter_name As String, ByVal filter_extension As String) As String
'OPENS FILES
Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
    .InitialFileName = "C:\Users"
    .ButtonName = "Select"
    .AllowMultiSelect = False
    .Title = "Choose file"
    .InitialView = msoFileDialogViewDetails
    .Filters.Add filter_name, filter_extension, 1
    .Show
    For Each objfl In .SelectedItems
        get_file_name = objfl
    Next objfl
    On Error GoTo 0
End With
End Function
