Attribute VB_Name = "Module1"
Function UTM_LL(x_utm, y_utm, Zone) As String

If IsNumeric(x_utm) And IsNumeric(y_utm) Then Else UTM_LL = "Non Numeric Coordinates": Exit Function

Dim northing, easting, PolarAxis, EquRad, Ecc, Ecc2, Ecc1, k0, c1, c2, c3, c4, arclength, mu, footprintlat, sin1, CC1, TT1, NN1, RR1, DD
Dim Fact1, Fact2, Fact3, Fact4, LoFact1, LoFact2, LoFact3

northing = 10000000 - y_utm
easting = 500000 - x_utm
PolarAxis = 6356752.3142
EquRad = 6378137
Ecc = 8.18191909289069E-02
Ecc2 = 0.006739496756587
k0 = 0.9996
Ecc1 = 1.67922038993736E-03
c1 = 2.51882658972117E-03
c2 = 3.70094905128485E-06
c3 = 7.44781381478857E-09
c4 = 1.7035993382807E-11

arclength = northing / k0
mu = arclength / (EquRad * (1 - Ecc ^ 2 / 4 - 3 * Ecc ^ 4 / 64 - 5 * Ecc ^ 6 / 256))
footprintlat = mu + c1 * Sin(2 * mu) + c2 * Sin(4 * mu) + c3 * Sin(6 * mu) + c4 * Sin(8 * mu)

sin1 = Sin(1)
CC1 = Ecc2 * Cos(footprintlat) ^ 2
TT1 = Tan(footprintlat) ^ 2
NN1 = EquRad / (1 - (Ecc * Sin(footprintlat)) ^ 2) ^ (1 / 2)
RR1 = EquRad * (1 - Ecc * Ecc) / (1 - (Ecc * Sin(footprintlat)) ^ 2) ^ (3 / 2)
DD = easting / (NN1 * k0)

Fact1 = NN1 * Tan(footprintlat) / RR1
Fact2 = DD * DD / 2
Fact3 = (5 + 3 * TT1 + 10 * CC1 - 4 * CC1 * CC1 - 9 * Ecc2) * DD ^ 4 / 24
Fact4 = (61 + 90 * TT1 + 298 * CC1 + 45 * TT1 * TT1 - 252 * Ecc2 - 3 * CC1 * CC1) * DD ^ 6 / 720
LoFact1 = DD
LoFact2 = (1 + 2 * TT1 + CC1) * DD ^ 3 / 6
LoFact3 = (5 - 2 * CC1 + 28 * TT1 - 3 * CC1 ^ 2 + 8 * Ecc2 + 24 * TT1 ^ 2) * DD ^ 5 / 120

deltalong = (LoFact1 - LoFact2 + LoFact3) / Cos(footprintlat)
deltalongdec = deltalong * 180 / 3.14159265358979

decimal_lat = -(180 * (footprintlat - Fact1 * (Fact2 + Fact3 + Fact4)) / 3.14159265358979)
decimal_lon = Zone - deltalongdec


UTM_LL = Round(decimal_lat, 7) & "," & Round(decimal_lon, 7)

End Function
