Attribute VB_Name = "Module1"
Function FindNumbers(inputString As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim result As String
    
    ' Create a regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True ' Find all matches
        .Pattern = "\b\d{6,}\b" ' Pattern to match 6 or more digits
    End With
    
    ' Find all matches in the input string
    Set matches = regex.Execute(inputString)
    
    ' Construct the result string with comma-separated numbers
    For Each match In matches
        If result <> "" Then result = result & ", "
        result = result & match.Value
    Next match
    
    ' Return the result
    FindNumbers = result
End Function

Function FuzzyLookup(lookupValue As String, lookupRange As Range) As Variant
    Dim closestMatch As String
    Dim maxSimilarity As Double
    Dim currentSimilarity As Double
    Dim cellValue As String
    Dim cell As Range
    
    ' Initialize variables
    closestMatch = ""
    maxSimilarity = 0
    
    ' Loop through each cell in the lookup range
    For Each cell In lookupRange
        cellValue = CStr(cell.Value)
        currentSimilarity = JaroWinklerProximity(lookupValue, cellValue)
        
        ' Update closest match if the current similarity is higher
        If currentSimilarity > maxSimilarity Then
            maxSimilarity = currentSimilarity
            closestMatch = cellValue
        End If
    Next cell
    
    ' Return the closest match
    FuzzyLookup = closestMatch
End Function

Function JaroWinklerProximity(String1 As String, String2 As String) As Double

Dim mWeightThreshold As Double
mWeightThreshold = 0.7
Dim mNumChars As Integer
mNumChars = 4
Dim aString1 As String
aString1 = LCase(String1)
Dim aString2 As String
aString2 = LCase(String2)
Dim lLen1 As Integer
lLen1 = Len(aString1)
Dim lLen2 As Integer
lLen2 = Len(aString2)

If lLen1 = 0 Then
    If lLen2 = 0 Then
        JaroWinklerProximity = 1
        Exit Function
    Else
        JaroWinklerProximity = 0
        Exit Function
    End If
End If
    
Dim lSearchRange As Integer
lSearchRange = WorksheetFunction.Max(1, WorksheetFunction.Max(lLen1, lLen2) / 2)

ReDim lMatched1(lLen1) As Boolean
ReDim lMatched2(lLen2) As Boolean
Dim lNumCommon As Integer
lNumCommon = 0

Dim i As Integer
For i = 1 To lLen1 Step 1
    Dim lStart As Integer
    lStart = WorksheetFunction.Max(1, i - lSearchRange)
    Dim lEnd As Integer
    lEnd = WorksheetFunction.Min(i + lSearchRange, lLen2)

    Dim j As Integer
    For j = lStart To lEnd - 1 Step 1
        If lMatched2(j) Then
            GoTo NextIteration1
        End If
        Dim charAtIndex1 As String
        charAtIndex1 = Mid(aString1, i, 1)
        Dim charAtIndex2 As String
        charAtIndex2 = Mid(aString2, j, 1)
        If charAtIndex1 <> charAtIndex2 Then
            GoTo NextIteration1
        End If
        lMatched1(i) = True
        lMatched2(j) = True
        lNumCommon = lNumCommon + 1
        Exit For
NextIteration1:
        Next j
Next i

If lNumCommon = 0 Then
    JaroWinklerProximity = 0
    Exit Function
End If

Dim lNumHalfTransposed As Integer
lNumHalfTransposed = 0
Dim k As Integer
k = 1
For i = 1 To lLen1 Step 1
    If Not lMatched1(i) Then
        GoTo NextIteration2
    End If
    
    Do While Not lMatched2(k)
        k = k + 1
    Loop

    If Mid(aString1, i, 1) <> Mid(aString2, j, 1) Then
        lNumHalfTransposed = lNumHalfTransposed + 1
    End If
    
    k = k + 1
NextIteration2:
Next

Dim lNumTransposed As Integer
lNumTransposed = lNumHalfTransposed / 2
Dim lNumCommonD As Double
lNumCommonD = lNumCommon
Dim lWeight As Double
lWeight = (lNumCommonD / lLen1 + lNumCommonD / lLen2 + (lNumCommon - lNumTransposed) / lNumCommonD) / 3
If lWeight <= mWeightThreshold Then
    JaroWinklerProximity = lWeight
    Exit Function
End If
Dim lMax As Integer
lMax = WorksheetFunction.Min(mNumChars, WorksheetFunction.Min(Len(aString1), Len(aString2)))
Dim lPos As Integer
lPos = 1

Do While lPos < lMax And Mid(aString1, lPos, 1) = Mid(aString2, lPos, 1)
    lPos = lPos + 1
Loop

If lPos = 1 Then
    JaroWinklerProximity = lWeight
    Exit Function
End If
JaroWinklerProximity = lWeight + 0.1 * lPos * (1# - lWeight)

End Function
