Attribute VB_Name = "Module1"
Public Function Haversine(Lat1 As Variant, Lon1 As Variant, Lat2 As Variant, Lon2 As Variant)
Dim R As Integer, dlon As Variant, dlat As Variant, Rad1 As Variant
Dim a As Variant, c As Variant, d As Variant, Rad2 As Variant

R = 6371
dlon = Excel.WorksheetFunction.Radians(Lon2 - Lon1)
dlat = Excel.WorksheetFunction.Radians(Lat2 - Lat1)
Rad1 = Excel.WorksheetFunction.Radians(Lat1)
Rad2 = Excel.WorksheetFunction.Radians(Lat2)
a = Sin(dlat / 2) * Sin(dlat / 2) + Cos(Rad1) * Cos(Rad2) * Sin(dlon / 2) * Sin(dlon / 2)
c = 2 * Excel.WorksheetFunction.Atan2(Sqr(1 - a), Sqr(a))
d = R * c
Haversine = d
End Function
