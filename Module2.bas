Attribute VB_Name = "Module2"
Function NearMatch(vLookupValue, rng As Range, iNumChars)
    Dim x As Integer
    Dim sSub As String

    Set rng = rng.Columns(1)
    sSub = Left(vLookupValue, iNumChars)
    For x = 1 To rng.Cells.Count
        If Left(rng.Cells(x), iNumChars) = sSub Then
            NearMatch = rng.Cells(x).Address
            Exit Function
        End If
    Next
    NearMatch = CVErr(xlErrNA)
End Function
