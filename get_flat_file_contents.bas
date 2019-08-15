Attribute VB_Name = "Module1"

Function get_flat_file_contents(ByVal file_name As String) As String
Dim fs As FileSystemObject
Dim f As TextStream
Set fs = New FileSystemObject
Set f = fs.openTextFile(file_name, 1, TristateFalse)

get_file_contents = f.ReadAll
f.Close
Set fs = Nothing

End Function
