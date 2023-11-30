Attribute VB_Name = "Module1"
Public Function UNIX_CONVERTER(ByVal input1 As Range, Optional ByVal gmtdiff As Double = 10, Optional ByVal days_in_advance As Integer = 0) As Double

UNIX_CONVERTER = (((Date + input1) - DateValue("01/01/1970")) * 86400 - (gmtdiff * 3600)) + (days_in_advance * 86400)

End Function
