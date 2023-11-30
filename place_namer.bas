Attribute VB_Name = "Module1"
Function place_namer(s As String)

Dim arr(0 To 26, 0 To 1) As String

arr(0, 0) = "High School": arr(0, 1) = "HS"
arr(1, 0) = "Public School": arr(1, 1) = "PS"
arr(2, 0) = "Primary School": arr(2, 1) = "PS"
arr(3, 0) = "Rd": arr(3, 1) = ""
arr(4, 0) = "St": arr(4, 1) = ""
arr(5, 0) = "at": arr(5, 1) = ""
arr(6, 0) = "opp": arr(6, 1) = ""
arr(7, 0) = "Ave": arr(7, 1) = ""
arr(8, 0) = "before": arr(8, 1) = ""
arr(9, 0) = "after": arr(9, 1) = ""
arr(10, 0) = "School": arr(10, 1) = "sch"
arr(11, 0) = "Dr": arr(11, 1) = ""
arr(12, 0) = "Park": arr(12, 1) = "pk"
arr(13, 0) = "The": arr(13, 1) = ""
arr(14, 0) = "Pde": arr(14, 1) = ""
arr(15, 0) = "Pl": arr(15, 1) = ""
arr(16, 0) = "Public": arr(16, 1) = "P"
arr(17, 0) = "Station": arr(17, 1) = "stn"
arr(18, 0) = "Hwy": arr(18, 1) = ""
arr(19, 0) = "Cres": arr(19, 1) = ""
arr(20, 0) = "College": arr(20, 1) = "col"
arr(21, 0) = "Stand": arr(21, 1) = "Std"
arr(22, 0) = "Centre": arr(22, 1) = "Ctr"
arr(23, 0) = "Shops": arr(23, 1) = "shp"
arr(24, 0) = "Primary": arr(24, 1) = ""
arr(25, 0) = "T-Way": arr(25, 1) = ""
arr(26, 0) = ",": arr(26, 1) = " "

s = Replace(s, ",", " ")
arr_s = Split(s, " ")

For i = 0 To 26
    For y = LBound(arr_s) To UBound(arr_s)
        If arr_s(y) = arr(i, 0) Then arr_s(y) = Trim(arr(i, 1)) Else
        'arr_s(y) = Replace(arr_s(y), arr(i, 0), arr(i, 1))
    Next y
Next i

place_namer = Join(arr_s, " ")

End Function
Function plc(ByVal s As String)

s = place_namer(s)
s = Replace(s, Space(2), Space(1))
s = Replace(s, Space(2), Space(1))
s_arr = Split(s, " ")
u = UBound(s_arr): l = LBound(s_arr)

If u Mod 2 = 0 Then nu = 6 / u Else nu = 6 / (u + 1)

nu = Round(nu, 0): u = UBound(s_arr): l = LBound(s_arr)

For i = l To u
    new_s = new_s & Left(s_arr(i), nu)
Next i

If Len(new_s) > 6 Then new_s = Left(new_s, 6) Else

plc = LCase(new_s)

End Function
