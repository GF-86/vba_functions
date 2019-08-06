Attribute VB_Name = "Module1"
Public Function ll_UTM(ByVal lat As Variant, ByVal lon As Variant) As String

Pi = WorksheetFunction.Pi()
false_easting = 500000
false_northing = 10000000
lat_rad = (lat / 180) * Pi
lon_rad = (lon / 180) * Pi

sma = 6378137 'Semi major axis (a) (m)

zone_width = 6 'Zone Width(Degrees)
lcmz1 = -177 'Longitude of the central meridian of zone 1(degrees)
lwez0 = lcmz1 - (1.5 * zone_width) 'Longitude of western edge of zone zero
cmzz = lwez0 + (zone_width / 2) 'Central meridian of zone zero
k0 = 0.9996 'Central Scale factor (K0)

zone_no_real = (lon - lwez0) / zone_width
zone_no = Int(zone_no_real)
central_meridian = zone_no * zone_width + cmzz
diff = lon - central_meridian
diff_long_rad = (diff / 180) * Pi

ecc_2 = 0.006694380023   'eccentricity (e2)
ecc_4 = ecc_2 * ecc_2
ecc_6 = ecc_2 * ecc_4

sin1 = Sin(lat_rad)
sin2 = Sin(2 * lat_rad)
sin4 = Sin(4 * lat_rad)
sin6 = Sin(6 * lat_rad)

a_0 = 1 - (ecc_2 / 4) - ((3 * ecc_4) / 64) - ((5 * ecc_6) / 256)
a_2 = (3 / 8) * (ecc_2 + (ecc_4 / 4) + ((15 * ecc_6) / 128))
a_4 = (15 / 256) * (ecc_4 + ((3 * ecc_6) / 4))
a_6 = (35 * ecc_6) / 3072

'meridian distance
term_1 = sma * a_0 * lat_rad
term_2 = -1 * sma * sin2 * a_2
term_3 = sma * sin4 * a_4
term_4 = -1 * sma * sin6 * a_6
meridian_dist = term_1 + term_2 + term_3 + term_4

'Radii of Curvature
rho = sma * (1 - ecc_2) / (1 - (ecc_2 * sin1 * sin1)) ^ 1.5
nu = sma / (1 - (ecc_2 * sin1 * sin1)) ^ 0.5

'cos latitude powers
cos_1 = Cos(lat_rad)
cos_2 = cos_1 ^ 2
cos_3 = cos_1 ^ 3
cos_4 = cos_1 ^ 4
cos_5 = cos_1 ^ 5
cos_6 = cos_1 ^ 6
cos_7 = cos_1 ^ 7

'diff long powers
diff_lon_1 = diff_long_rad
diff_lon_2 = diff_lon_1 ^ 2
diff_lon_3 = diff_lon_1 ^ 3
diff_lon_4 = diff_lon_1 ^ 4
diff_lon_5 = diff_lon_1 ^ 5
diff_lon_6 = diff_lon_1 ^ 6
diff_lon_7 = diff_lon_1 ^ 7
diff_lon_8 = diff_lon_1 ^ 8

'tan latitude powers
tan_1 = Tan(lat_rad)
tan_2 = tan_1 ^ 2
tan_4 = tan_1 ^ 4
tan_6 = tan_1 ^ 6

'psi powers
psi_1 = nu / rho
psi_2 = psi_1 ^ 2
psi_3 = psi_1 ^ 3
psi_4 = psi_1 ^ 4

'Easting
easting_1 = cos_1 * diff_lon_1 * nu
easting_2 = nu * diff_lon_3 * cos_3 * (psi_1 - tan_2) / 6
easting_3 = nu * diff_lon_5 * cos_5 * (4 * psi_3 * (1 - 6 * tan_2) + psi_2 * (1 + 8 * tan_2) - psi_1 * (2 * tan_2) + tan_4) / 120
easting_4 = nu * diff_lon_7 * cos_7 * (61 - 479 * tan_2 + 179 * tan_4 - tan_6) / 5040
easting_sum = easting_1 + easting_2 + easting_3 + easting_4
e_sum_k0 = k0 * easting_sum
easting = false_easting + e_sum_k0

'Northing
northing_1 = nu * sin1 * diff_lon_2 * cos_1 / 2
northing_2 = nu * sin1 * diff_lon_4 * cos_3 * (4 * psi_2 + psi_1 - tan_2) / 24
northing_3 = nu * sin1 * diff_lon_7 * cos_5 * (8 * psi_4 * (11 - 24 * tan_2) - 28 * psi_3 * (1 - 6 * tan_2) + psi_2 * (1 - 32 * tan_2) - psi_1 * (2 * tan_2) + tan_4) / 720
northing_4 = nu * sin1 * diff_lon_8 * cos_7 * (1385 - 3111 * tan_2 + 543 * tan_4 - tan_6) / 40320
northing_sum = meridian_dist + northing_1 + northing_2 + northing_3 + northing_4
n_sum_k0 = k0 * northing_sum
northing = n_sum_k0 + false_northing

ll_UTM = Round(easting, 1) & ", " & Round(northing, 1) & ", " & zone_no

End Function


