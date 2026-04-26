Attribute VB_Name = "increment"
Public Function IncrementString(ByVal currentStr As String) As String
    Dim length As Integer
    Dim numericPart As String
    Dim alphaPart As String
    Dim maxNumeric As Long
    Dim currentNum As Double
    Dim i As Integer
    length = Len(currentStr)
    If IsNumeric(currentStr) Then
        length = Len(currentStr)
        currentNum = Val(currentStr) + 1
        
        ' ??? ????? ????? ???? ?????? (??? 999)? ???? ?? A01
        If Len(CStr(currentNum)) > length Then
            If length = 1 Then
                IncrementString = "A"
            Else
                IncrementString = "A" & Format(1, String(length - 1, "0"))
            End If
        Else
            ' ????? ????? ?? ?????? ??? ??????? (Replicate Zeros)
            IncrementString = Format(currentNum, String(length, "0"))
        End If
        Exit Function
    End If
    
    
    
    ' ??? ????? ????? ?? ?????? ?? ??????
    i = length
    Do While i > 0 And IsNumeric(Mid(currentStr, i, 1))
        i = i - 1
    Loop
    
    alphaPart = Left(currentStr, i)
    numericPart = Right(currentStr, length - i)
    maxNumeric = (10 ^ Len(numericPart)) - 1
    
    ' ?????? ??????: ??? ??? ????? ?????? ??? ?? ???? ??????
    If Val(numericPart) < maxNumeric Then
        currentNum = Val(numericPart) + 1
        IncrementString = alphaPart & Right(String(Len(numericPart), "0") & currentNum, Len(numericPart))
    
    ' ?????? ???????: ?????? ???? ?????? (????? 999 ?? A99)
    Else
        ' ??? ??? ?????? ????? ??? (??? 999) ????? ??? A01
        If alphaPart = "" Then
            IncrementString = "A" & Right(String(length - 1, "0") & "1", length - 1)
        Else
            ' ????? ????? ????? (A -> B ?? Z -> AA)
            alphaPart = IncrementAlpha(alphaPart)
            If Len(alphaPart) > length Then
                IncrementString = Left(currentStr, i)
            Else
                IncrementString = Left(alphaPart & Right(String(length, "0") & "1", length - Len(alphaPart)), length)
            End If
        End If
    End If
End Function
' ???? ?????? ?????? ?????? (A ????? B ? ? Z ????? AA)
Private Function IncrementAlpha(ByVal alpha As String) As String
    Dim lastChar As Integer
    If alpha = "" Then
        IncrementAlpha = "A"
    Else
        lastChar = Asc(Right(alpha, 1))
        If lastChar < 90 Then ' ??? ?? Z
            IncrementAlpha = Left(alpha, Len(alpha) - 1) & Chr(lastChar + 1)
        Else
            IncrementAlpha = IncrementAlpha(Left(alpha, Len(alpha) - 1)) & "A"
        End If
    End If
End Function


