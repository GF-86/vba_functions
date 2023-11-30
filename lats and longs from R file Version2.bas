Attribute VB_Name = "Module1"
Type acoor
x As String
y As String
End Type

Sub testxml()
Dim xdox As MSXML2.DOMDocument
Dim rvps As MSXML2.IXMLDOMNodeList
Dim rvpsnode As MSXML2.IXMLDOMNode
Dim coor As acoor

Set xdoc = New MSXML2.DOMDocument
xdoc.async = False
xdoc.Load (get_file_name())

Set rvps = xdoc.DocumentElement.SelectNodes("//Transport_Operations_Data/RouteVariantPoints/RouteVariantPoint")

i = 1
For Each rvpsnode In rvps

tsn_attri = False


For x = 0 To (rvpsnode.Attributes.Length - 1)
    If rvpsnode.Attributes(x).BaseName = "TransitStopTSN" Then tsn_attri = True: tsn_no = rvpsnode.Attributes(x).Text Else
Next x

hold = rvpsnode.Attributes(4).Text
Cells(i, 1) = rvpsnode.Attributes(0).Text & ","

If tsn_attri Then coor = return_stp_tsn(xdoc, tsn_no): _
Cells(i, 1) = rvpsnode.Attributes(4).Text & "," & Cells(i, 1) & UTM_LL(coor.x, coor.y, 153) Else _
Cells(i, 1) = rvpsnode.Attributes(3).Text & "," & Cells(i, 1) & UTM_LL(Left(hold, 7) / 10, Right(hold, 8) / 10, 153)
i = i + 1
Next rvpsnode

Set xdoc = Nothing

End Sub

Function return_stp_tsn(ByVal xd As MSXML2.DOMDocument, ByVal tsn As String) As acoor

Dim rvps1 As MSXML2.IXMLDOMNodeList
Dim rvpsnode1 As MSXML2.IXMLDOMNode
Set rvps1 = xd.DocumentElement.SelectNodes("//Transport_Operations_Data/TransitStops/TransitStop")

For Each rvpsnode1 In rvps1

For i = 0 To (rvpsnode1.Attributes.Length - 1)
    If rvpsnode1.Attributes(i).BaseName = "XCoord" Then Exit For Else
Next i


If rvpsnode1.Attributes(0).Text = tsn Then _
return_stp_tsn.x = rvpsnode1.Attributes(i).Text: return_stp_tsn.y = rvpsnode1.Attributes(i + 1).Text: Exit Function Else

Next rvpsnode1

return_stp_tsn.x = "no stop location"
return_stp_tsn.y = "no stop location"

End Function
Function get_file_name() As String
'OPENS FILES
Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
    .ButtonName = "Select"
    .AllowMultiSelect = False
    .Title = "Choose file"
    .InitialView = msoFileDialogViewDetails
    .Filters.Add "XML", "*.xml", 1
    .Show
    For Each objfl In .SelectedItems
        get_file_name = objfl
    Next objfl
    On Error GoTo 0
End With

End Function

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
