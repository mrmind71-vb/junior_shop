Attribute VB_Name = "Module1"
Public cComp_Name As String, cComp_address As String, cComp_Head1 As String, cComp_Head2 As String, cComp_Phone As String, nCountBoon As Double, nLang_Boon As Double
Public Lookupdata, lManger As Boolean, cPathTemp As String, cPMosm As String, cPMosmD As String, cManBox As String, lSupperVisor As Boolean, cBranchBox As String, cBranchStore As String
Public MdbPath, PublicPath, tempPath As String, tempFile As String, aPublic(), lServerOnLine As Boolean, lServerOnLineShop As Boolean
Public PublicVar As String, publicFlag As Variant, lCust As Integer, sStore As String
Public Secondtitle As String, Firstitle As String, lExirSave As Boolean
Public Firsttitle As String, cMyModel As String, sBranchCode As String, cBranch As String
Public nCountPrint As Byte, cPathShop2 As String, cCodeShop2 As String, sBranch As String, lIsBranchStore As Boolean
Public lRepCust As Byte, pDevice As String, cCust As String, lDisplayCom As Boolean
Public cRepItem As String, cDataShop As String, cBranchBoxVisa As String, cComputerName As String
Public DRepDate1 As Date, sUsername As String, lokaz As Boolean, lokazItem As Boolean
Public DRepDate2 As Date, cAddressBranch As String, cPhoneBranch As String, cBranchSales As String
Public lHeadStore As Boolean
Public cPStore As String, lHeadBox As Boolean, lOneVisa As Boolean, cPathConf As String
Function aTurnValue(pSource, aOld, pNew)
Dim i As Long
For i = 0 To UBound(aOld)
    pOld = aOld(i)
    If pSource = pOld Or (IsNull(pSource) And IsNull(pOld)) Then
         aTurnValue = pNew
         Exit Function
    End If
Next
aTurnValue = pSource
End Function
Function faIndex(Row, col, grd As MSFlexGrid)
faIndex = (Row * grd.Cols) + col
End Function
Function loadGrd(Row, col, grd As MSFlexGrid)
nCount = (Row * grd.Cols) + col
loadGrd = grd.TextArray(nCount)
End Function
Function aDel(aTarget, nItem)
Dim aTemp
ReDim aTemp(UBound(aTarget) - 1, UBound(aTarget, 2))
For i = 0 To UBound(aTarget, 1) - 1
    If i <> nItem Then
        For i2 = 0 To UBound(aTarget, 2) - 1
            aTemp(IIf(i < nItem, i, i - 1), i2) = aTarget(i, i2)
        Next
    End If
Next
aDel = aTemp
End Function
Function MakeQry(csource, cAdd)
MakeQry = IIf(csource = "", cAdd, csource & " and " & cAdd)
End Function
Function MyParn(pValue)
MyParn = "'" & Trim(SQLFixup(pValue & "")) & "'"
End Function
Function MyPar(pValue)
MyPar = "N'" & Trim(SQLFixup(pValue & "")) & "'"
End Function
Function myDateString(pDate)
myDateString = "Date(" & Year(pDate) & "," & Month(pDate) & "," & Day(pDate) & ")"
End Function
Sub MyDelGrd(grd As MSFlexGrid, Row)
If grd.Rows <= 2 Then
    grd.Rows = 1
Else
    grd.RemoveItem (Row)
End If
End Sub
Function ReturnIndexed(aTarget, aIndex, nDim)
Dim aReturn()
If nDim = 1 Then
    ReDim aReturn(UBound(aTarget))
Else
    ReDim aReturn(UBound(aTarget, 1), UBound(aTarget, 2))
End If
If nDim = 1 Then
    For i = 0 To UBound(aTarget) - 1
        aReturn(aIndex(i)) = aTarget(i)
    Next
ElseIf nDim = 2 Then
    For i = 0 To UBound(aTarget, 1) - 1
        For i2 = 0 To UBound(aTarget, 2) - 1
            aReturn(aIndex(i), i2) = aTarget(i, i2)
        Next
    Next
End If
ReturnIndexed = aReturn
End Function
Function IndexArray(myarray)
Dim aReturn()
ReDim aReturn(UBound(myarray))
For i = 0 To UBound(myarray) - 1
    aReturn(i) = itemIndex(myarray, i, nBegin)
Next
IndexArray = aReturn
End Function
Function OneDimArray(ParArray, nDim)
Dim myarray()
ReDim myarray(UBound(ParArray))
For i = LBound(myarray) To UBound(myarray) - 1
    myarray(i) = ParArray(i, nDim)
Next
OneDimArray = myarray
End Function
Function itemIndex(myarray, nItem, nBegin)
Item = myarray(nItem)
NRETURN = nBegin
For i = nBegin To UBound(myarray) + nBegin - 1
    If i < nItem Then
        NRETURN = IIf(Item >= myarray(i), NRETURN + 1, NRETURN)
    Else
        NRETURN = IIf(Item > myarray(i), NRETURN + 1, NRETURN)
    End If
Next
itemIndex = NRETURN
End Function
Function aSearch(myarray, xSearch, Optional mydim, Optional nBegin)
nBegin = IIf(IsMissing(nBegin), 0, nBegin)
If IsMissing(mydim) Then
    For i = nBegin To UBound(myarray)
        If myarray(i) = xSearch Then
            aSearch = i
            Exit Function
        End If
    Next
Else
    For i = nBegin To UBound(myarray)
        If myarray(i, mydim) = xSearch Then
            aSearch = i
            Exit Function
        End If
    Next
End If
aSearch = Null
End Function
Function aScan(myarray, xSearch, Optional nBegin)
For i = nBegin To UBound(myarray)
    If myarray(i) = xSearch Then
    aScan = i
    Exit Function
    End If
Next
aScan = Null
End Function
Function aSearch2(myarray, xSearch, Optional mydim, Optional nBegin)
nBegin = IIf(IsMissing(nBegin), 0, nBegin)
If IsMissing(mydim) Then
    For i = nBegin To UBound(myarray)
        If myarray(i) = xSearch Then
'            aSearch = i
            Exit Function
        End If
    Next
Else
    For i = nBegin To UBound(myarray)
        If myarray(i, mydim) = xSearch Then
            Exit Function
        End If
    Next
End If
End Function
Function aAdd(aTarget, xItem)
Dim aTemp
ReDim aTemp(UBound(aTarget) + 1)
For i = 1 To UBound(aTarget)
      aTemp(i) = aTarget(i)
Next
aTemp(UBound(aTemp)) = xItem
aAdd = aTemp
End Function
Function GrdText(xGrid As MSFlexGrid, nCol As Integer)
nCounter = (xGrid.Row * xGrid.Cols) + nCol
GrdText = xGrid.TextArray(nCounter)
End Function
Function LastInStr(cStr1, cStr2)
    If InStr(1, cStr1, cStr2) = 0 Then
        LastInStr = 0: Exit Function
    Else
        nFound = nFound = InStr(1, cStr1, cStr2)
        Do While True
        nloop = InStr(nloop + 1, cStr1, cStr2)
        If nloop = 0 Then
            Exit Do
        Else
            nFound = nloop
        End If
        Loop
    End If
    LastInStr = nFound
End Function
Function TurnValue(pSource, Optional pOld = "", Optional pNew = Null)
  TurnValue = IIf(pSource = pOld Or (IsNull(pSource) And IsNull(pOld)), pNew, pSource)
End Function
Function RetField(xTable, cField, cSearchStr As String, xReturn As String)
cSearchStr = cField & " = " & "'" & cSearchStr & "'"
xTable.FindFirst cSearchStr
If xTable.NoMatch Then Exit Function
RetField = xTable(xReturn)
End Function
Function StrDel(csource, myarray)
mycounter = UBound(myarray)
For i = 0 To mycounter
csource = StrTran(csource, myarray(i), "")
Next
StrDel = csource
End Function

Function StrTran(csource, cStr1, cStr2)
    nLen = Len(cStr1)
    Do Until InStr(1, csource, cStr1) = 0
        nPosition = InStr(1, csource, cStr1)
        csource = Mid(csource, 1, nPosition - 1) & cStr2 & _
                  Mid(csource, nPosition + nLen)
    Loop
    StrTran = csource
End Function
Function IncRec(cString)
Dim G As Double
Dim cChr As String
nLen = Len(cString)
For G = 0 To nLen - 1
    cChr = Mid(cString, nLen - G, 1)
    If IsNumeric(cChr) And cChr <> "9" Then
        cChr = cChr + 1
        cString = Mid(cString, 1, nLen - (G + 1)) & cChr & Mid(cString, nLen - (G - 1))
        Exit For
    Else
       cChr = IIf(cChr = "9", "0", cChr)
       cString = Mid(cString, 1, nLen - (G + 1)) & cChr & Mid(cString, nLen - (G - 1))
    End If
Next
IncRec = cString
End Function
Function RetNumber(pNumber, myDec As Boolean)
If (pNumber >= 48 And pNumber <= 57) _
    Or pNumber = 8 Or pNumber = 45 _
    Or (myDec = True And pNumber = 46) Then
    RetNumber = pNumber
Else
    RetNumber = 0
End If
End Function
Function ValidQuant(nquant, Pack)
On Error Resume Next
ValidQuant = Abs(nquant) / Pack < 1
If Err.Number > 0 Then
ValidQuant = 0
End If
End Function
Function myiif(cCondition, cField, Optional cFunction As String = "SUM", Optional cElse As Variant = 0, Optional nRound As Long = -1)
If cCondition = "" Then
    myiif = cFunction & "(" & cField & ")"
Else
    myiif = cFunction & " ( case when (" & cCondition & ") THEN " & _
         cField & " else " & "0 end" & ")"
End If
If nRound <> -1 Then
    myiif = "Round(" & myiif & "," & nRound & ")"
End If
End Function
Function myiif_r(cCondition, cField, Optional nRound As Long = 2, Optional cFunction As String = "SUM")
If cCondition = "" Then
    myiif_r = cFunction & "(" & cField & ")"
Else
    myiif_r = cFunction & " ( case when (" & cCondition & ") THEN " & _
         cField & " else " & "0 end" & ")"
End If
myiif_r = "Round(" & myiif_r & "," & nRound & ")"
End Function
Function myiif2(cCondition, cField, Optional cFuction As String = "SUM")
If cCondition = "" Then
    myiif2 = cFuction & "(" & cField & ")"
Else
    myiif2 = cFuction & " ( case when (" & cCondition & ") THEN " & _
         cField & " else " & " NULL  end" & ")"
End If
End Function

Function RetFind(rTable, cFieldFind, cFieldRet, cSearch)
'rTable.FindFirst cFieldFind & "=" & MyParn(cSearch)
'RetFind = IIf(rTable.NoMatch, "", rTable(cFieldRet))
End Function
Function Units(nPart1, nPart2, nPack)
If VarType(nPart1) = vbString Then nPart1 = Val(nPart1)
If VarType(nPart2) = vbString Then nPart2 = Val(nPart2)
If VarType(nPack) = vbString Then nPack = Val(nPack)
Units = (TurnValue(nPart1, Null, 0) * nPack) + TurnValue(nPart2, Null, 0)
End Function
Function retRev(cString) As String
For i = 1 To Len(cString)
    retRev = retRev & Mid(cString, Len(cString) - i + 1, 1)
Next
End Function
Function NameOfDay(xPass)
    cDay = Weekday(xPass)
    Select Case cDay
        Case 1
            NameOfDay = "«·√Õœ"
        Case 2
            NameOfDay = "«·√À‰Ì‰"
        Case 3
            NameOfDay = "«·À·«À«¡"
        Case 4
            NameOfDay = "«·√—»⁄«¡"
        Case 5
            NameOfDay = "«·Œ„Ì”"
        Case 6
            NameOfDay = "«·Ã„⁄…"
        Case 7
            NameOfDay = "«·”» "
   End Select
End Function
Function myQuery(sWhere)
myQuery = sWhere & IIf(sWhere = "", " Where ", " and ")
End Function
Function RetNumber2(cString, pNumber)
cString = cString & Format(Chr(pNumber))
RetNumber2 = IIf(IsNumeric(cString), pNumber, 0)
End Function
Function Unit1Q(Quant, Pack, cType)
If Not IsNull(Quant) And TurnValue(Pack, Null, 0) <> 0 Then
    If cType = "1" Then
        Unit1Q = Fix(TurnValue(Quant, Null, 0) / Pack)
    Else
        Unit1Q = TurnValue(Quant, Null, 0) Mod Pack
    End If
End If
If TurnValue(Pack, Null, 0) = 0 And cType = "2" Then Unit1Q = Quant
End Function
Function QtyToString(Quant, Pack)
If Not IsNull(Quant) And TurnValue(Pack, Null, 0) <> 0 Then
    QtyToString = TurnValue(Quant, Null, 0)
    If Len(QtyToString) = 3 Then QtyToString = " " & QtyToString
    If Len(QtyToString) = 2 Then QtyToString = "  " & QtyToString
    If Len(QtyToString) = 1 Then QtyToString = "   " & QtyToString
    QtyToString = QtyToString & "/" & TurnValue(Quant, Null, 0) Mod Pack
End If
If TurnValue(Pack, Null, 0) = 0 Then QtyToString = Quant
End Function
Function Chr254(cPass)
    Do While True
        nPos = MyInStr(cPass)
        If nPos > 0 Then
             cPass = Mid(cPass, 1, nPos - 1) & " " & Chr(254) & Mid(cPass, nPos + 1)
        Else
            Exit Do
        End If
    Loop
    Chr254 = Chr(254) & cPass & Chr(254)
End Function
Function RetZero(ByVal cString, Optional ByVal nLen As Integer = 6)
cString = Trim(cString)
If Len(cString) >= nLen Then
    RetZero = cString
    Exit Function
End If
nLen = nLen - Len(cString)
RetZero = String(nLen, "0") & cString
End Function
Function DelZero(ByVal pString) As String
Dim cString As String
cString = Trim(pString & "")
If cString = "" Then Exit Function
For i = 1 To Len(cString)
    If Mid(cString, i, 1) <> "0" Then Exit For
Next i
If i > 0 Then cString = Mid(cString, i)
DelZero = Trim(cString)
End Function
Function addstring(pValue)
addstring = IIf(Trim(pValue & "") = "", "null", "'" & Trim(SQLFixup(pValue & "")) & "'")
End Function
Function addvalue(pValue) As String
addvalue = IIf(IsNumeric(pValue & ""), pValue, "null")
End Function
Function RetSeek(pTable, pIndex, pFind, pField) As String
If Trim(pFind) = "" Then Exit Function
Dim loctable As New ADODB.Recordset
loctable.CursorLocation = adUseServer
loctable.index = pIndex
loctable.Open pTable, con, adOpenDynamic, adLockOptimistic, adCmdTableDirect
loctable.Seek pFind, adSeekFirstEQ
If Not loctable.EOF Then RetSeek = loctable(pField) & ""
loctable.Close
Set loctable = Nothing
End Function


Function MyParnAll(parStr, Optional bAsOne As Boolean = False)
Dim aString
If bAsOne Then
    MyParnAll = "'%" & parStr & "%'"
Else
    aString = Split(parStr, " ")
    If Not IsEmpty(aString) Then
        For i = 0 To UBound(aString)
            MyParnAll = IIf(MyParnAll = "", "%", "") & MyParnAll & aString(i) & "%"
        Next
    End If
    MyParnAll = "'" & MyParnAll & "'"
End If
End Function


Function MyInStr(cPass)
    Do While True
        MyInStr = InStr(cPass, " 0")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 1")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 2")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 3")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 4")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 5")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 6")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 7")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 8")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 9")
        If MyInStr <> 0 Then Exit Do
        Exit Do
    Loop
End Function
Function TwoLine(cPass)
nPos = InStr(cPass, "-")
If nPos = 0 Then
    TwoLine = cPass
Else
    TwoLine = Mid(cPass, 1, nPos - 1) & String(30 - nPos, " ") & "-" & Mid(cPass, nPos)
End If
End Function
Public Function D_RetItemBalance(cItem, cStore, dDate) As Double
Dim con As Connection
If cItem = "" Then Exit Function
cString = "Select sum(val([IN] & '') - VAL([OUT] & '')) as Balance From file1_11 where item = " & MyParn(cItem) & _
          " and Store = " & MyParn(cStore) & " and Date <= " & DateConv(dDate)
D_RetItemBalance = Val(GetDesca(cString, con) & "")
End Function
Public Sub Inform_OK(sCaption1 As String, Optional sCaption2 As String = "")
InformfrmOK.sCaption1 = sCaption1
InformfrmOK.sCaption2 = sCaption2
InformfrmOK.Show 1
End Sub

Function DateSql(dDate)
DateSql = "DateValue(" & MyParn(Format(dDate, "dd-mm-yy")) & ")"
End Function

Function RetRand(nCount) As Integer
For i = 1 To nCount
    Randomize
    RetRand = RetRand & Int((9 * Rnd) + 1)     ' Generate random value between 1 and 6.
Next
End Function
Sub AcntLookup_end(oForm As Form, oSearch As Form, Optional cActiveName As String = "", Optional cTitle As String = "«” ⁄·«„")
Dim Generalarray(6)
Dim listarray(1, 7)
Dim GrdArray(2, 1)
Set Generalarray(0) = oForm
Generalarray(1) = " SELECT ACC1_10.CODE, ACC1_10.DESCA, ACC1_10.SUBDESCA FROM    ACC1_10 "
Generalarray(2) = ""
Generalarray(3) = 6000
Generalarray(5) = False
Generalarray(6) = cActiveName

listarray(0, 0) = "≈”„ - —ﬁ„ «·Õ”«» "
listarray(0, 1) = "(%%ACC1_10.desca%% or ACC1_10.CODE Like '%cFilter%' )"

listarray(1, 0) = "‰Ê⁄ «·Õ”«» "
listarray(1, 1) = "%%ACC1_10.SUBdesca%%"

GrdArray(0, 0) = "—ﬁ„"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Õ”«»"
GrdArray(1, 1) = 6000

GrdArray(2, 0) = "—∆Ì”Ï"
GrdArray(2, 1) = 2000


searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = cTitle
oSearch.Show 1
End Sub

Sub AcntLookup_Sub(oForm As Form, oSearch As Form, Optional cActiveName As String = "", Optional cTitle As String = "«” ⁄·«„")
Dim Generalarray(6)
Dim listarray(0, 7)
Dim GrdArray(1, 1)
Set Generalarray(0) = oForm
Generalarray(1) = " SELECT ACC1_10.CODE, ACC1_10.DESCA FROM  ACC1_10 WHERE FLAG = 6 "
Generalarray(2) = ""
Generalarray(3) = 6000
Generalarray(5) = False
Generalarray(6) = cActiveName

listarray(0, 0) = "≈”„ - —ﬁ„ «·Õ”«» "
listarray(0, 1) = "(ACC1_10.desca%% or ACC1_10.CODE Like '%cFilter%' )"

GrdArray(0, 0) = "—ﬁ„"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Õ”«»"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = cTitle
oSearch.Show 1
End Sub
Public Sub SendKeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.SendKeys CStr(text), wait
  Set WshShell = Nothing
End Sub
Public Sub AddLod_Data(pUser, pType, pDesca, pCon As ADODB.Connection, Optional pDoc_no = Null, Optional pDate_Doc = Null, Optional pFileName = Null, Optional pDesca2 = Null)
    On Error Resume Next
    If lServerOnLine Then
        pCon.Execute " INSERT INTO LOG_DATA (USERNAME, TYPE, DESCA, doc_no, DATE_DOC, FILENAME,DESCA2,DATE2, [ComputerName]) VALUES     (" & addstring(pUser) & " ," & addvalue(pType) & " ," & addstring(pDesca) & " , " & addstring(pDoc_no) & " , " & addDate(pDate_Doc) & " , " & addstring(pFileName) & "," & addstring(pDesca2) & ", " & addDate(Date) & "," & addstring(cComputerName) & ")"
    End If
    Err.Clear
End Sub
