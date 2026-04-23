Attribute VB_Name = "inv_DataProc"
Public Function myDateSql(pDate As String) As String
If IsDate(pDate) Then
    myDateSql = addstring(Format(pDate, "YYYY-MM-DD HH:NN"))
Else
    myDateSql = "null"
End If
End Function
Public Function mRound(ByVal nValue As Variant, Optional nRound As Integer = 2) As Double
Dim cString As String
cString = "##"
If nRound > 0 Then cString = cString & "." & String(nRound, "#")
mRound = Val(Format(nValue, cString))
End Function
Public Function arString(aString As Variant, Optional pSep As String = ",", Optional nCount As Integer = -1) As String
Dim I As Integer
Dim cString As New ChilkatStringBuilder
For I = 0 To IIf(nCount = -1, UBound(aString), nCount)
    cString.Append aString(I) & pSep
Next
cString.Shorten Len(pSep)
arString = cString.GetAsString
End Function
Public Function mTotal(nValue As Double, nRate As Double, nTotal As Double, Optional nRound As Integer = 4) As Double
If nRate = 0 Or nValue = 0 Then
    mTotal = 0
ElseIf Round(nTotal / nValue, nRound) <> nRate Then
    mTotal = nValue * nRate
Else
    mTotal = nTotal
End If
End Function
Public Function mRate(nValue As Double, nTotal As Double, nRate As Double, Optional nRound As Integer = 2) As Double
If nTotal = 0 Or nValue = 0 Then
    mRate = 0
ElseIf Round(nRate * nValue, nRound) <> nTotal Then
    mRate = mRound(nTotal / nValue)
Else
    mRate = nRate
End If
End Function



