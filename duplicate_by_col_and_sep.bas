Attribute VB_Name = "Module2"
Sub duplicate_by_col_and_sep()

Dim r As Range
Dim n As Integer
Dim s As String
Dim arr, newarr As Variant
Dim NewSheet As Worksheet

r_cntr = 1

Set r = Application.InputBox(Prompt:="Select the range", Type:=8)
n = Application.InputBox(Prompt:="Input the column number", Type:=1)
s = Application.InputBox(Prompt:="Input the separator", Type:=2)

arr = r
ReDim newarr(LBound(arr, 1) To UBound(arr, 1) * 5, LBound(arr, 2) To UBound(arr, 2))

For i = LBound(arr, 1) To UBound(arr, 1)
    SplitValues = Split(arr(i, n), s)
    If UBound(SplitValues) > 0 And i <> 1 Then
        For Each v In SplitValues
            For j = LBound(arr, 2) To UBound(arr, 2)
                If j = n Then
                    newarr(r_cntr, j) = v
                Else
                    newarr(r_cntr, j) = arr(i, j)
                End If
            Next j
            r_cntr = r_cntr + 1
        Next
    Else
        For j = LBound(arr, 2) To UBound(arr, 2)
            newarr(r_cntr, j) = arr(i, j)
        Next j
        r_cntr = r_cntr + 1
    End If
Next i

Set NewSheet = ThisWorkbook.Sheets.Add
NewSheet.Activate
    
For i = 1 To UBound(newarr, 1)
    For j = 1 To UBound(newarr, 2)
        Cells(i, j).Value = newarr(i, j)
    Next j
Next i

End Sub
