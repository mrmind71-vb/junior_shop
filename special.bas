Attribute VB_Name = "Special"
Public nusercode As String
Public cUserName As String
Public cSalesMan As String
Public dSalesDate As String, aPassword As Variant, aAddModel As Variant, dMaxDate As Date
Public aPrinterBar As Variant
Function itemCost(cItem, Optional pDate As String = "") As Single
Dim con As New ADODB.Connection
If con.State = adStateOpen Then con.Close
openCon con

Dim cString As String
cString = "Select FILE7_20.PRICE FROM  " & _
          "FILE7_20 INNER JOIN FILE7_20H ON FILE7_20.DOC_NO = FILE7_20H.DOC_NO " & _
          " WHERE ITEM = " & MyParn(cItem)
          
If IsDate(pDate) Then
    cString = cString & turnFound(cString) & " FILE7_20H.Date <= " & DateSq(pDate)
End If
cString = cString & " " & " ORDER BY FILE7_20H.DATE DESC"
itemCost = Round(Val(GetDesca(cString, con) & ""), 2)
If itemCost = 0 Then itemCost = Val(GetDesca("select cost from file1_10 where item = " & MyParn(cItem), con))
End Function
Sub ModelLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "Select TOP 1000 FILE1_10h.MODEL ,FILE1_10h.DescA,File1_10h.MODELFACT0 ,FILE1_10h.MOSM ,FILE1_10h.SUPP , fact.DESCA   FROM FILE1_10h INNER JOIN fact ON FILE1_10h.FACT = fact.CODE  LEFT JOIN MOSM ON MOSM.MOSM = FILE1_10H.MOSM "
Generalarray(2) = "order by MOSM.DATE DESC  , FILE1_10H.MODEL "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ „ÊœÌ·"
listarray(0, 1) = "(%%File1_10h.MODELFACT0+FACT.DESCA%%)"

listarray(1, 0) = "«·„ÊœÌ·"
listarray(1, 1) = "(%%File1_10h.DESCA%%)"

'listarray(1, 0) = "„Ê”„ ° „Ê—œ "
'listarray(1, 1) = "(%%File1_10h.MOSM & FILE1_10h.SUPP%%)"

GrdArray(0, 0) = "„ÊœÌ·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·’‰› "
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·„Ê’„ "
GrdArray(3, 1) = 800

GrdArray(4, 0) = " „ﬂ » Ã„·…"
GrdArray(4, 1) = 800

GrdArray(5, 0) = " «·„’‰⁄"
GrdArray(5, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Sub ModelNoLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
'                                        0               1                    2                   3           4            5               6
Generalarray(1) = "Select TOP 100 FILE1_10h.MODELNO,FILE1_10h.DescA,File1_10h.MODELFACT0,FILE1_10h.MOSM,fact.DESCA  FROM FILE1_10h INNER JOIN fact ON FILE1_10h.FACT = fact.CODE"
Generalarray(2) = " GROUP BY FILE1_10h.MODELNO,FILE1_10h.DescA,File1_10h.MODELFACT0,FILE1_10h.MOSM,fact.DESCA order by file1_10h.MOSM "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ „ÊœÌ·"
listarray(0, 1) = "(%%File1_10h.MODELFACT0 + FACT.DESCA%%)"

listarray(1, 0) = "«·„ÊœÌ·"
listarray(1, 1) = "(%%File1_10h.DESCA%%)"

'listarray(1, 0) = "„Ê”„ ° „Ê—œ "
'listarray(1, 1) = "(%%File1_10h.MOSM & FILE1_10h.SUPP%%)"

GrdArray(0, 0) = "„ÊœÌ·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·’‰› "
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(2, 1) = 1000

GrdArray(3, 0) = "«·„Ê’„ "
GrdArray(3, 1) = 800

GrdArray(4, 0) = " «·„’‰⁄"
GrdArray(4, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Sub ItemsLookupAll(oForm As Form, oSearch As Form, Optional pOrderNo As String = "")
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(9, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "SELECT  Top 1000 item , file1_10.mosm , fact.desca , file1_10.modelfact0 , file1_10.supp,file1_10.desca , file1_10.scal , file1_10.color , PRICE , PRICE_2 from file1_10 inner join fact on fact.code = file1_10.fact  WHERE ISSTOP = 0 "
If pOrderNo <> "" Then
    Generalarray(1) = Generalarray(1) & " AND ITEM IN (SELECT ITEM FROM FILE6_50 WHERE DOC_NO = " & MyParn(pOrderNo) & ")"
End If
Generalarray(2) = "ORDER BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_SCAL "
Generalarray(3) = 12000
Generalarray(5) = False

listarray(0, 0) = "≈”„ «·’‰› ° «·„’‰⁄ ° —ﬁ„ „ÊœÌ· & „ﬁ«” "
listarray(0, 1) = "(%%FILE1_10.DESCA+file1_10.MODELFACT+FACT.DESCA%%)"

listarray(1, 0) = "„Ê”„ "
listarray(1, 1) = "(%%File1_10.MOSM%%)"

listarray(2, 0) = "„ﬁ«”"
listarray(2, 1) = "%%File1_10.SCAL%%"

listarray(3, 0) = "«··Ê‰"
listarray(3, 1) = "(%%File1_10.COLOR%%)"
 
GrdArray(0, 0) = "»«—ﬂÊœ"
GrdArray(0, 1) = 900

GrdArray(1, 0) = "„Ê”„"
GrdArray(1, 1) = 800

GrdArray(2, 0) = "«·„’‰⁄"
GrdArray(2, 1) = 1800

GrdArray(3, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = "„ﬂ »"
GrdArray(4, 1) = 800

GrdArray(5, 0) = " «·’‰›"
GrdArray(5, 1) = 4500

GrdArray(6, 0) = "«·„ﬁ«”"
GrdArray(6, 1) = 1300

GrdArray(7, 0) = "«··Ê‰"
GrdArray(7, 1) = 800

GrdArray(8, 0) = " «·”⁄—"
GrdArray(8, 1) = 1000

GrdArray(9, 0) = " ”⁄— «Êﬂ«“ÊÌ‰"
GrdArray(9, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Sub ItemsLookup(oForm As Form, oSearch As Form, Optional pOrderNo As String = "")
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(9, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "SELECT  Top 1000 item , file1_10.mosm , fact.desca , file1_10.modelfact0 , file1_10.supp,file1_10.desca , file1_10.scal , file1_10.color , PRICE , PRICE_2 from file1_10 inner join fact on fact.code = file1_10.fact  WHERE ISSTOP = 0 "
If pOrderNo <> "" Then
    Generalarray(1) = Generalarray(1) & " AND ITEM IN (SELECT ITEM FROM FILE6_50 WHERE DOC_NO = " & MyParn(pOrderNo) & ")"
End If
Generalarray(2) = "ORDER BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_SCAL "
Generalarray(3) = 12000
Generalarray(5) = False

listarray(0, 0) = "≈”„ «·’‰› ° «·„’‰⁄ ° —ﬁ„ „ÊœÌ· & „ﬁ«” "
listarray(0, 1) = "(%%FILE1_10.DESCA+file1_10.MODELFACT+FACT.DESCA%%)"

listarray(1, 0) = "„Ê”„ "
listarray(1, 1) = "(%%File1_10.MOSM%%)"

listarray(2, 0) = "„ﬁ«”"
listarray(2, 1) = "%%File1_10.SCAL%%"

listarray(3, 0) = "«··Ê‰"
listarray(3, 1) = "(%%File1_10.COLOR%%)"
 
GrdArray(0, 0) = "»«—ﬂÊœ"
GrdArray(0, 1) = 900

GrdArray(1, 0) = "„Ê”„"
GrdArray(1, 1) = 800

GrdArray(2, 0) = "«·„’‰⁄"
GrdArray(2, 1) = 1800

GrdArray(3, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = "„ﬂ »"
GrdArray(4, 1) = 800

GrdArray(5, 0) = " «·’‰›"
GrdArray(5, 1) = 4500

GrdArray(6, 0) = "«·„ﬁ«”"
GrdArray(6, 1) = 1300

GrdArray(7, 0) = "«··Ê‰"
GrdArray(7, 1) = 800

GrdArray(8, 0) = " «·”⁄—"
GrdArray(8, 1) = 1000

GrdArray(9, 0) = " ”⁄— «Êﬂ«“ÊÌ‰"
GrdArray(9, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Sub ItemsLook(oForm As Form, oSearch As Search_abd, Optional pOrderNo As String = "", Optional pOffer As String = "", Optional bUnload As Boolean = False)
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(9, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "SELECT  Top 1000 item," & _
                  "file1_10.mosm," & _
                  "fact.desca," & _
                  "file1_10.modelfact0," & _
                  "file1_10.supp," & _
                  "file1_10.desca," & _
                  "file1_10.scal," & _
                  "file1_10.color," & _
                  "PRICE," & _
                  "PRICE_2 " & _
                  "from file1_10 " & _
                  "inner join fact on fact.code = file1_10.fact" & _
                  " WHERE ISSTOP = 0 "
If pOrderNo <> "" Then
    Generalarray(1) = Generalarray(1) & " AND ITEM IN (SELECT ITEM FROM FILE6_50 WHERE DOC_NO = " & MyParn(pOrderNo) & ")"
End If

If pOffer <> "" Then
    Generalarray(1) = Generalarray(1) & " AND FILE1_10.MODELNO IN (SELECT MODELNO FROM FILE0_90 WHERE FILE0_90.DOC_NO = " & MyParn(pOffer) & ")"
End If


Generalarray(2) = "ORDER BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_SCAL "
Generalarray(3) = 12000
Generalarray(5) = False

listarray(0, 0) = "≈”„ «·’‰› ° «·„’‰⁄ ° —ﬁ„ „ÊœÌ· & „ﬁ«” "
listarray(0, 1) = "(%%FILE1_10.DESCA%% OR %%file1_10.MODELFACT%% OR %%FACT.DESCA%%)"

listarray(1, 0) = "„Ê”„ "
listarray(1, 1) = "(%%File1_10.MOSM%%)"

listarray(2, 0) = "„ﬁ«”"
listarray(2, 1) = "%%File1_10.SCAL%%"

listarray(3, 0) = "«··Ê‰"
listarray(3, 1) = "(%%File1_10.COLOR%%)"
 
GrdArray(0, 0) = "»«—ﬂÊœ"
GrdArray(0, 1) = 900

GrdArray(1, 0) = "„Ê”„"
GrdArray(1, 1) = 800

GrdArray(2, 0) = "«·„’‰⁄"
GrdArray(2, 1) = 1800

GrdArray(3, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = "„ﬂ »"
GrdArray(4, 1) = 800

GrdArray(5, 0) = " «·’‰›"
GrdArray(5, 1) = 4500

GrdArray(6, 0) = "«·„ﬁ«”"
GrdArray(6, 1) = 1300

GrdArray(7, 0) = "«··Ê‰"
GrdArray(7, 1) = 800

GrdArray(8, 0) = " «·”⁄—"
GrdArray(8, 1) = 1000

GrdArray(9, 0) = " ”⁄— «Êﬂ«“ÊÌ‰"
GrdArray(9, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.bUnload = bUnload
oSearch.bNoRef = True
oSearch.sCaption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Sub CustLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = oForm

Generalarray(1) = "Select FILE3_10.CODE , FILE3_10.DESCA , FILE3_50.DESCA  FROM FILE3_10  INNER JOIN FILE3_50  ON FILE3_10.[GROUP] = FILE3_50.CODE  "
Generalarray(2) = "order by file3_10.desca "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì· "
listarray(0, 1) = "(%%File3_10.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·⁄„Ì·"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ ⁄„·«¡ √Ã·"
oSearch.Show 1
End Sub
Sub FactLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm

Generalarray(1) = "Select FACT.CODE , FACT.DESCA FROM FACT"
Generalarray(2) = "order by FACT.CODE"
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ«·„Œ ’—-«·«”„"
listarray(0, 1) = "(%%FACT.CODE%% OR %%FACT.DESCA%%)"


GrdArray(0, 0) = "«·ﬂÊœ «·„Œ ’—"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·≈”„"
GrdArray(1, 1) = 7000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "«” ⁄·«„ «·„’«‰⁄"
oSearch.Show 1
End Sub
Sub SuppLookupAll(oForm As Form, oSearch As Form, Optional vSecondary As Variant = Null)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = oForm

Generalarray(1) = "Select FILE4_10.CODE , FILE4_10.DESCA , FILE4_50.DESCA  FROM FILE4_10  left JOIN FILE4_50  ON FILE4_10.[GROUP] = FILE4_50.CODE "
If Not IsNull(vSecondary) Then
    Generalarray(1) = Generalarray(1) & " WHERE FILE4_10.IS_SECONDARY = " & vSecondary
End If
Generalarray(2) = "order by file4_10.desca "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "«·„Ê—œ "
listarray(0, 1) = "(%%File4_10.DESCA%%    )"


GrdArray(0, 0) = "ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·„Ê—œ"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „Ê—œÌ‰  "
oSearch.Show 1
End Sub
Sub CustSUPPLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select FILE3_20.CODE , FILE3_20.DESCA , file3_20.Phone1 , IsOk FROM FILE3_20  "
Generalarray(2) = "order by file3_20.desca "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì· "
listarray(0, 1) = "(%%File3_20.DESCA%%   or %%File3_20.Phone1%%    )"


GrdArray(0, 0) = "ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·⁄„Ì·"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = " ·Ì›Ê‰"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = " "
GrdArray(3, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ ⁄„·«¡ ﬂ«—  Âœ«Ì« "
oSearch.Show 1
End Sub

Function ValidDate() As Boolean
ValidDate = Format("01-12-2000", "dd-mm-yyyy") = "01-12-2000"
If Not ValidDate Then
    cString = "‰Ÿ«„ «· «—ÌŒ €Ì— ’«·Õ" & vbCrLf & _
            "«Ã⁄· «·œÊ·… ›Ï ·ÊÕ… «· Õﬂ„" & vbCrLf & _
            "control Panel" & vbCrLf & _
            "Regional and language options" & vbCrLf & _
            "Egypt"
            
    MsgBox ArbString(cString), vbCritical
'    End

End If
End Function
Function LastBalance(cItem, cStore, pCon As ADODB.Connection) As Double
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
If cItem = "" Or cStore = "" Then Exit Function
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("Store", adVarWChar, adParamInput, 3, cStore)
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adDouble, adParamInput, 20, cItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)

cmdTable.CommandText = "LastBalance"
Set obj = cmdTable.Execute
LastBalance = Val(cmdTable.Parameters(2).Value & "")
Set obj = Nothing
End Function


Function lastColor(cItem As String, cDoc_No As String, pCon As ADODB.Connection) As Variant
Dim obj As New ADODB.Recordset, aRet(1) As Double
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adVarWChar, adParamInput, 15, cItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("doc_no", adVarWChar, adParamInput, 6, cDoc_No)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret1", adDouble, adParamOutput)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret2", adDouble, adParamOutput)

cmdTable.CommandText = "lastcolor"
Set obj = cmdTable.Execute
aRet(0) = Val(cmdTable.Parameters(2).Value & "")
aRet(1) = Val(cmdTable.Parameters(3).Value & "")
lastColor = aRet
Set obj = Nothing
End Function
Function LastCost(cItem, pCon As ADODB.Connection) As Double
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adVarWChar, adParamInput, 15, cItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)

cmdTable.CommandText = "LastCost"
Set obj = cmdTable.Execute
LastCost = Val(cmdTable.Parameters(1).Value & "")
Set obj = Nothing
End Function
Function RetValChr(ByVal sValue) As String
sValue = Trim(sValue)
Dim str1 As String, str2 As String
For i = 1 To Len(sValue)
    If IsNumeric(Right(sValue, i)) Then
        str1 = Val(Right(sValue, i))
        Exit For
    End If
Next
If str1 <> "" Then
    RetValChr = Left(sValue, Len(str1)) & RetZero(str1, 10)
End If
End Function
Function retFile(pModel As String) As String
Dim cString As String
If Len(Mid(pModel, 4, 3) & turn(Mid(pModel, 4, 3), "\")) < 3 Then Exit Function
retFile = App.Path & "\PICT\" & Mid(pModel, 4, 3) & turn(Mid(pModel, 4, 3), "\") & pModel & ".jpg"
End Function
Function RetItemBalance(cItem, cStore, dDate, pCon As ADODB.Connection, Optional pCode As String = "") As Double
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("Store", adVarWChar, adParamInput, 3, cStore)
cmdTable.Parameters.Append cmdTable.CreateParameter("date", adVarWChar, adParamInput, 12, Format(dDate, "YYYY-MM-DD"))
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adDouble, adParamInput, 20, cItem)
If pCode <> "" Then cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adVarWChar, adParamInput, 6, pCode)

cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)
If pCode = "" Then
    cmdTable.CommandText = "retitemBalance"
    Set obj = cmdTable.Execute
    RetItemBalance = Val(cmdTable.Parameters(3).Value & "")
Else
    cmdTable.CommandText = "retitemBalance_CODE"
    Set obj = cmdTable.Execute
    RetItemBalance = Val(cmdTable.Parameters(4).Value & "")
End If
Set obj = Nothing
End Function
Function ItemBalanceNoStore(cItem, dDate, pCon As ADODB.Connection) As Double
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc

cmdTable.Parameters.Append cmdTable.CreateParameter("date", adVarWChar, adParamInput, 12, Format(dDate, "YYYY-MM-DD"))
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adDouble, adParamInput, 20, cItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)

cmdTable.CommandText = "itemBalanceNoStore"
Set obj = cmdTable.Execute
ItemBalanceNoStore = Val(cmdTable.Parameters(2).Value & "")
Set obj = Nothing
End Function
Function ModelFind(pModel As String, pCon As ADODB.Connection) As ADODB.Recordset
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("MODELNO", adVarWChar, adParamInput, 30, pModel)
cmdTable.CommandText = "ModelFind"
Set obj = cmdTable.Execute
Set ModelFind = obj
Set obj = Nothing
End Function
Function card_last(pCode As String, pCon As ADODB.Connection) As String
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adVarWChar, adParamInput, 6, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return1", adVarWChar, adParamOutput, 10)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return2", adVarWChar, adParamOutput, 10)
cmdTable.CommandText = "card_last"
Set rdTable = cmdTable.Execute
'pDoc_no = cmdTable.Parameters(1).Value & ""
card_last = cmdTable.Parameters(2).Value & ""
Set rdTable = Nothing
End Function
Function valid_card(pCode As String, pDate As String, pCon As ADODB.Connection) As String
Dim sDate
sDate = Format(card_last(pCode, pCon), "yyyy-mm-dd")
If sDate = "" Then
    valid_card = "null"
    Exit Function
End If
If sDate < Format(DateAdd("m", -6, pDate), "yyyy-mm-dd") Then
    valid_card = "valid"
End If
valid_card = "ok"
End Function
Function card_Sales(pCode As String, pCon As ADODB.Connection) As Double
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adVarWChar, adParamInput, 6, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adDecimal, adParamOutput)
cmdTable.Parameters(1).Precision = 18
cmdTable.Parameters(1).NumericScale = 2
cmdTable.CommandText = "card_sales"
Set rdTable = cmdTable.Execute
card_Sales = cmdTable.Parameters(1).Value
Set rdTable = Nothing
End Function
Function card_discount(pCode As String, pCon As ADODB.Connection) As Double
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adVarWChar, adParamInput, 6, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adDecimal, adParamOutput)
cmdTable.Parameters(1).Precision = 18
cmdTable.Parameters(1).NumericScale = 2
cmdTable.CommandText = "card_discount"
Set rdTable = cmdTable.Execute
card_discount = cmdTable.Parameters(1).Value
Set rdTable = Nothing
End Function


Function isOffice(pCode As String, pCon As ADODB.Connection) As Boolean
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adVarWChar, adParamInput, 3, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adBoolean, adParamOutput)
cmdTable.CommandText = "isOffice"
Set rdTable = cmdTable.Execute
isOffice = cmdTable.Parameters(1).Value
Set rdTable = Nothing
End Function
Function retCode(pCode As String, pCon As ADODB.Connection) As String
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adVarWChar, adParamInput, 3, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adVarWChar, adParamOutput, 3)
cmdTable.CommandText = "retCode"
Set rdTable = cmdTable.Execute
retCode = cmdTable.Parameters(1).Value
Set rdTable = Nothing
End Function
Function fact_Desc(pCode As String, pCon As ADODB.Connection) As String
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adVarWChar, adParamInput, 3, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adVarWChar, adParamOutput)
cmdTable.CommandText = "fact_desc"
Set rdTable = cmdTable.Execute
fact_Desca = cmdTable.Parameters(1).Value & ""
Set rdTable = Nothing
End Function
Function Item_cost(pBARCODE As String, pCon As ADODB.Connection) As Double
    Item_cost = Val(GetDesca(" SELECT COSTITEM FROM FILE1_10 WHERE ITEM = " & pBARCODE, pCon) & "")
End Function
Function retSysdate() As String
retSysdate = Format(IIf(Val(Format(Time, "hh")) > 4, Date, DateAdd("d", -1, Date)), "dd-mm-yyyy")
End Function
Function Model_Found(pModel, pCon As ADODB.Connection) As Integer
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("Model", adVarWChar, adParamInput, 30, pModel)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adDouble, adParamOutput)

cmdTable.CommandText = "Model_Found"
Set obj = cmdTable.Execute
Model_Found = Val(cmdTable.Parameters(1).Value & "")
Set obj = Nothing
End Function
Function BAL_SUPPL(cCode, cMosm, pCon As ADODB.Connection) As Double
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
If cCode = "" Or cMosm = "" Then Exit Function
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adVarWChar, adParamInput, 6, cCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("MOSM", adVarWChar, adParamInput, 3, cMosm)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)

cmdTable.CommandText = "BAL_SUPPL"
Set obj = cmdTable.Execute
BAL_SUPPL = Val(cmdTable.Parameters(2).Value & "")
Set obj = Nothing
End Function
Function validNumber(ByVal pNumber As String) As Boolean
Dim cString As String
pNumber = Trim(pNumber)
If Not IsNumeric(pNumber) Then Exit Function
If Val(pNumber) <= 0 Then Exit Function
For i = 1 To Len(validNumber)
    cString = Mid(pNumber, i, 1)
    If i = 1 And cString = "0" Then Exit Function
    If Not IsNumeric(cString) Then Exit Function
Next
validNumber = True
'validNumber = Int(pNumber) & "" = pNumber & "" And Val(pNumber) > 0
End Function
Function ValidInt(pString As Variant)
ValidInt = Int(Val(pString & "")) & "" = Trim(pString)
End Function
Function ItemFind(pItem As String, pCon As ADODB.Connection) As ADODB.Recordset
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
If Len(pItem) > 7 Then pItem = 0

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adInteger, adParamInput, 20, pItem)
cmdTable.CommandText = "ItemFind"
Set rdTable = cmdTable.Execute
Set ItemFind = rdTable
Set rdTable = Nothing
End Function
Function ItemFind_BARCODE(pBARCODE As String, pCon As ADODB.Connection) As ADODB.Recordset
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("BARCODE", adChar, adParamInput, 20, pBARCODE)
cmdTable.CommandText = "ItemFind_BARCODE"
Set rdTable = cmdTable.Execute
Set ItemFind_BARCODE = rdTable
Set rdTable = Nothing
End Function

Function validItem(pItem As String, pCon As ADODB.Connection) As Boolean
On Error GoTo myerror
If Not ValidInt(pItem) Then Exit Function
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adInteger, adParamInput, 10, pItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adBoolean, adParamOutput)
cmdTable.CommandText = "validItem"
Set rdTable = cmdTable.Execute
validItem = cmdTable.Parameters(1).Value
Set rdTable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Function ItemFields(ByVal pItem As String, pCon As ADODB.Connection) As Variant
If Not ValidInt(pItem) Then Exit Function
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adInteger, adParamInput, 10, pItem)
cmdTable.CommandText = "ItemFind"
Set rdTable = cmdTable.Execute
If Not (rdTable.EOF And rdTable.BOF) Then
    For i = 0 To rdTable.Fields.Count - 1
        ItemFields = AddFlag(ItemFields, rdTable.Fields(i).Name, rdTable.Fields(i).Value)
    Next
End If
Set rdTable = Nothing
End Function
Function ItemField(pItem As String, pField As String, pCon As ADODB.Connection) As Variant
If Not ValidInt(pItem) Then Exit Function
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc

cmdTable.Parameters.Append cmdTable.CreateParameter("item", adInteger, adParamInput, 10, pItem)
cmdTable.CommandText = "ItemFind"
Set rdTable = cmdTable.Execute
If Not (rdTable.EOF And rdTable.BOF) Then ItemField = rdTable(pField)
Set rdTable = Nothing
End Function
Function QItemIn(cItem, cStore, pCon As ADODB.Connection) As Double
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
If cItem = "" Or cStore = "" Then Exit Function
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("Store", adVarWChar, adParamInput, 3, cStore)
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adDouble, adParamInput, 20, cItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)
cmdTable.CommandText = "RetItemIn"
Set obj = cmdTable.Execute
QItemIn = Val(cmdTable.Parameters(2).Value & "")
Set obj = Nothing
End Function
Sub GrModelLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "Select GRMODEL ,GRDESCA FROM FILE1_10h WHERE GRMODEL IS NOT NULL "
Generalarray(2) = "GROUP BY GRMODEL ,GRDESCA order by GRMODEL "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "„Ã„Ê⁄…- »Ì«‰"
listarray(0, 1) = "(%%GRMODEL+GRDESCA%%)"


GrdArray(0, 0) = "„Ã„Ê⁄… "
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "»Ì«‰ «·„Ã„Ê⁄…"
GrdArray(1, 1) = 4000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „Ã„Ê⁄«  „ÊœÌ·« "
oSearch.Show 1
End Sub

Sub ModelLookupAll_bal(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "Select TOP 1000 FILE1_10h.MODEL ,FILE1_10h.DescA,File1_10h.MODELFACT0 ,FILE1_10h.MOSM ,FILE1_10h.SUPP , fact.DESCA   FROM FILE1_10h INNER JOIN fact ON FILE1_10h.FACT = fact.CODE  LEFT JOIN MOSM ON MOSM.MOSM = FILE1_10H.MOSM "
Generalarray(2) = "order by MOSM.DATE DESC  , FILE1_10H.MODEL "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ „ÊœÌ·"
listarray(0, 1) = "(%%File1_10h.MODELFACT0+FACT.DESCA%%)"

listarray(1, 0) = "«·„ÊœÌ·"
listarray(1, 1) = "(%%File1_10h.DESCA%%)"

'listarray(1, 0) = "„Ê”„ ° „Ê—œ "
'listarray(1, 1) = "(%%File1_10h.MOSM & FILE1_10h.SUPP%%)"

GrdArray(0, 0) = "„ÊœÌ·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·’‰› "
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·„Ê’„ "
GrdArray(3, 1) = 800

GrdArray(4, 0) = " „ﬂ » Ã„·…"
GrdArray(4, 1) = 800

GrdArray(5, 0) = " «·„’‰⁄"
GrdArray(5, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Sub ModelLookupAll_Ser1(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "Select FILE1_10.MODEL , FILE1_10.ITEM , FILE1_10.DESCA , FILE8_51.desca  , FILE8_52.DESCA FROM FILE1_10  LEFT JOIN FILE8_51 ON FILE8_51.CODE = FILE1_10.CHARGE LEFT JOIN FILE8_52 ON FILE8_51.maingroup = FILE8_52.CODE WHERE FILE1_10.ISNOITEM = 1 "
Generalarray(2) = "order by FILE1_10.ITEM "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%File1_10.DESCA%%)"

GrdArray(0, 0) = "„ÊœÌ·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "»«—ﬂÊœ"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "«·»Ì«‰"
GrdArray(2, 1) = 3000

GrdArray(3, 0) = "„’—Ê› "
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "„’—Ê› —∆Ì”Ì… "
GrdArray(4, 1) = 3000


searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «’‰«› Œœ„… "
oSearch.Show 1
End Sub
Sub ModelLookupAll_Ser2(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "Select FILE1_10.MODEL , FILE1_10.ITEM , FILE1_10.DESCA , FILE8_51.desca  , FILE8_52.DESCA FROM FILE1_10  LEFT JOIN FILE8_51 ON FILE8_51.CODE = FILE1_10.CHARGE LEFT JOIN FILE8_52 ON FILE8_51.maingroup = FILE8_52.CODE WHERE FILE1_10.ISNOITEM2 = 1 "
Generalarray(2) = "order by FILE1_10.ITEM "
Generalarray(3) = 3000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%File1_10.DESCA%%)"

GrdArray(0, 0) = "„ÊœÌ·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "»«—ﬂÊœ"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "«·»Ì«‰"
GrdArray(2, 1) = 4000

GrdArray(3, 0) = "„’—Ê›"
GrdArray(3, 1) = 4000

GrdArray(4, 0) = "„’—Ê› —∆”ÌÏ"
GrdArray(4, 1) = 4000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «’‰«› «’Ê· "
oSearch.Show 1
End Sub
Function ItemFind_Fact(pModelFact0 As String, pScal As String, pCon As ADODB.Connection) As ADODB.Recordset
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("MODELFACT0", adChar, adParamInput, 20, pModelFact0)
cmdTable.Parameters.Append cmdTable.CreateParameter("SCAL", adChar, adParamInput, 10, pScal)
cmdTable.CommandText = "ItemFind_Fact"
Set rdTable = cmdTable.Execute
Set ItemFind_Fact = rdTable
Set rdTable = Nothing
End Function
Function ItemFind_Model_S_C(pModel, pCode_Color, pCode_Scal, pCon As ADODB.Connection) As Double
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("MODEL", adVarWChar, adParamInput, 20, pModel)
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE_COLOR", adInteger, adParamInput, 3, pCode_Color)
cmdTable.Parameters.Append cmdTable.CreateParameter("Code_Scal", adInteger, adParamInput, 3, pCode_Scal)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adInteger, adParamOutput)

cmdTable.CommandText = "ItemFind_Model_S_C"
Set obj = cmdTable.Execute
ItemFind_Model_S_C = Val(cmdTable.Parameters(3).Value & "")
Set obj = Nothing
End Function
Function Item_BARCODE(pBARCODE, pCon As ADODB.Connection) As Double
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("BARCODE", adVarWChar, adParamInput, 30, pBARCODE)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adDouble, adParamOutput)
cmdTable.CommandText = "Item_BARCODE"
Set obj = cmdTable.Execute
Item_BARCODE = Val(cmdTable.Parameters(1).Value & "")
Set obj = Nothing
End Function
Function retDOC_NO(pDoc_no As String, pCon As ADODB.Connection) As String
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("DOC_NO", adVarWChar, adParamInput, 15, pDoc_no)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adVarWChar, adParamOutput, 15)
cmdTable.CommandText = "retDOC_NO"
Set rdTable = cmdTable.Execute
retDOC_NO = cmdTable.Parameters(1).Value & ""
Set rdTable = Nothing
End Function
Function SALES_DOC_FIND(ByVal pDoc_no As String, pCon As ADODB.Connection) As Variant
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("DOC_NO", adChar, adParamInput, 20, pDoc_no)
cmdTable.CommandText = "SALES_DOC_FIND"
Set rdTable = cmdTable.Execute
Set SALES_DOC_FIND = rdTable
Set rdTable = Nothing
End Function
Function ItemFind_BARCODE2(pItem As String, pCon As ADODB.Connection) As ADODB.Recordset
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
If Len(pItem) > 8 Then pItem = ""
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("BARCODE2", adChar, adParamInput, 20, pItem)
cmdTable.CommandText = "ItemFind_BARCODE2"
Set rdTable = cmdTable.Execute
Set ItemFind_BARCODE2 = rdTable
Set rdTable = Nothing
End Function
Function RetDEmItem(pItem As String, pstore As String, pCon As ADODB.Connection) As Double
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("ITEM", adVarWChar, adParamInput, 15, pItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("STORE", adVarWChar, adParamInput, 5, pstore)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adDouble, adParamOutput)
cmdTable.CommandText = "RetDEmItem"
Set obj = cmdTable.Execute
RetDEmItem = Val(cmdTable.Parameters(2).Value & "")
Set obj = Nothing
End Function
Function ISPRINTED_sales(pDoc_no As String, pCon As ADODB.Connection) As Boolean
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("DOC_NO", adVarWChar, adParamInput, 15, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adBoolean, adParamOutput)
cmdTable.CommandText = "ISPRINTED_sales"
'Set rdTable = cmdTable.Execute
ISPRINTED_Pr = cmdTable.Parameters(1).Value
Set rdTable = Nothing
End Function
Function Ret_Phone(pPhone As String, pCon As ADODB.Connection) As String
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("PHONE", adVarWChar, adParamInput, 15, pPhone)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adVarWChar, adParamOutput, 200)
cmdTable.CommandText = "Ret_Phone"
Set rdTable = cmdTable.Execute
Ret_Phone = cmdTable.Parameters(1).Value & ""
Set rdTable = Nothing
End Function
Function TransStock(pDoc_no As String, pClosed As Integer, con As ADODB.Connection, Optional ByRef pErrorNumber As Variant, Optional ByRef pErrorMsg As Variant, Optional ByRef pUpdate As Long = 0, Optional ByRef pDelete As Long = 0, Optional ByRef pInsert As Long = 0) As Boolean
Dim cmd As New ADODB.command
aPrm = AddFlag(Empty, "DOC_NO", pDoc_no)
aPrm = AddFlag(aPrm, "CLOSED", pClosed)
Set cmd = mycmdEx("dbo.sp_stock_trans", con, aPrm)

If Not IsNull(cmd.Parameters("@ERROR_NUMBER").Value) Then
     pErrorNumber = cmd.Parameters("@ERROR_NUMBER").Value
     pErrorMsg = cmd.Parameters("@ERROR_MSG").Value
     Exit Function
End If

pUpdate = cmd.Parameters("@COUNT_UPDATE")
If Not IsNull(cmd.Parameters("@COUNT_DELETE")) Then
    pDelete = cmd.Parameters("@COUNT_DELETE")
End If

If Not IsNull(cmd.Parameters("@COUNT_INSERT")) Then
    pInsert = cmd.Parameters("@COUNT_INSERT")
End If
TransStock = True
End Function
Function TransStockAll(pDoc_no As String, con As ADODB.Connection, pClosed As String, Optional ByRef pCountDelete As Long = 0, Optional ByRef pCountInsert As Long = 0, Optional ByRef pCountDeleteMove As Long = 0, Optional ByRef pCountInsertMove As Long = 0) As Boolean
Dim command As New ADODB.command

Dim aPrm As Variant
aPrm = AddFlag(Empty, "DOC_NO", pDoc_no)
aPrm = AddFlag(aPrm, "CLOSED", pClosed)

Set command = cmd("[dbo].[sp_stock_trans_items]", con, adStoredProc, aPrm)
command.Execute

pCountInsert = Val(command.Parameters("@COUNT_INSERT") & "")
pCountDelete = Val(command.Parameters("@COUNT_DELETE") & "")

Set command = cmd("[dbo].[sp_stock_trans_header]", con, adStoredProc, aPrm)
command.Execute

pCountInsert = Val(command.Parameters("@COUNT_INSERT") & "")
pCountDelete = Val(command.Parameters("@COUNT_DELETE") & "")
TransStockAll = True

Set command = cmd("[dbo].[sp_stock_trans_header]", con, adStoredProc, aPrm)
command.Execute

If pClosed = "1" Then
    aPrm = AddFlag(Empty, "DOC_NO", pDoc_no)
    Set command = cmd("[dbo].[sp_stock_cost]", con, adStoredProc, aPrm)
    command.Execute

    ss = command.Parameters("@COUNT_UPDATE") & ""
End If


End Function
Function TransStockItems(pDoc_no As String, con As ADODB.Connection, pClosed As String, Optional ByRef pCountDelete As Long = 0, Optional ByRef pCountInsert As Long = 0) As Boolean
'Dim command As New adodb.command
'Dim aPrm As Variant
'aPrm = AddFlag(aPrm, "DOC_NO", pDoc_no)
'aPrm = AddFlag(aPrm, "CLOSED", pClosed)
'
'Set command = mycmd("[dbo].[sp_stock_trans_items]", con, adStoredProc, aPrm)
'command.Execute
'
'pCountDelete = Val(command.Parameters("@COUNT_INSERT") & "")
'pCountInsert = Val(.Parameters("@COUNT_DELETE") & "")
'TransStockItems = True
End Function
Function TransStockHeader(pDoc_no As String, con As ADODB.Connection, pClosed As String, Optional ByRef pCountDelete As Long = 0, Optional ByRef pCountInsert As Long = 0) As Boolean
Dim command As New ADODB.command
aPrm = AddFlag(Empty, "DOC_NO", pDoc_no)
aPrm = AddFlag(aPrm, "CLOSED", pClosed)
Set command = mycmd("dbo.sp_stock_trans", con, adStoredProc, aPrm)
command.Execute

pCountDelete = Val(command.Parameters("@COUNT_DELETE") & "")
pCountInsert = Val(command.Parameters("@COUNT_INSERT") & "")
TransStockHeader = True
End Function
Public Sub YearsLookup(oForm As Form, oSearch As Form, Optional sControl As String = "", Optional bAddRow As Boolean = False, Optional pWhere As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select vw_years_codes.CODE,vw_years_codes.DESCA " & _
                  " FROM vw_years_codes"
If pWhere <> "" Then Generalarray(1) = Generalarray(1) & " WHERE " & pWhere

Generalarray(2) = " ORDER BY vw_years_codes.CODE DESC"
Generalarray(3) = 2000
Generalarray(5) = True

listarray(0, 0) = "«·”‰…"
listarray(0, 1) = "(**[YEAR]**)"

GrdArray(0, 0) = "«·”‰…"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·”‰…"
GrdArray(1, 1) = 4000

Dim aRow As Variant
If bAddRow Then
    aRow = AddFlag(Empty, "text", "ﬂ· «·”‰Ì‰")
    aRow = AddFlag(aRow, "col", 1)
End If

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.aAddRow = aRow
oSearch.bNoSort = True
oSearch.sCaption = "≈” ⁄·«„ «·”‰Ì‰"
oSearch.sControl = sControl
oSearch.Show 1
End Sub
Public Sub FixRpImage(myForm As Form)
On Error Resume Next
With myForm
.cmdApply.Picture = LoadPicture(App.Path & "\sys_img\preview.jpg")
.cmdExit.Picture = LoadPicture(App.Path & "\sys_img\exit.jpg")
.cmdClear.Picture = LoadPicture(App.Path & "\sys_img\clear.jpg")
Err.Clear
If .cmdApply.Picture = 0 Then .cmdApply.Caption = "⁄—÷ «· ﬁ—Ì—"
If .cmdExit.Picture = 0 Then .cmdExit.Caption = "Œ—ÊÃ"
If .cmdClear.Picture = 0 Then .cmdClear.Caption = "„”Õ"
End With
End Sub
Sub DefineText(myForm As Form)
Dim i As Long
For i = 0 To myForm.Count - 1
    If TypeOf myForm(i) Is TextBox Then
        myForm(i).text = ""
    ElseIf TypeOf myForm(i) Is DataCombo Then
        myForm(i).BoundText = ""
    ElseIf TypeOf myForm(i) Is CheckBox Then
        myForm(i).Value = 0
    End If
Next
End Sub
Sub clientLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT FILE3_10.CODE, FILE3_10.desca, FILE3_50.DESCA FROM  FILE3_10 LEFT JOIN FILE3_50 ON FILE3_10.[GROUP] = FILE3_50.CODE"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by FILE3_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·„Ã„Ê⁄…"
listarray(0, 1) = "(%%FILE3_10.DESCA%% OR %%FILE3_50.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE3_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.sCaption = "≈” ⁄·«„ «·⁄„·«¡"
oSearch.Show 1
End Sub
Sub myClientLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT FILE3_10.CODE, FILE3_10.desca, FILE3_50.DESCA FROM  FILE3_10 LEFT JOIN FILE3_50 ON FILE3_10.[GROUP] = FILE3_50.CODE"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by FILE3_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·„Ã„Ê⁄…"
listarray(0, 1) = "(%%FILE3_10.DESCA%% OR %%FILE3_50.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE3_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.sCaption = "≈” ⁄·«„ «·⁄„·«¡"
oSearch.Show 1
End Sub
Sub suplookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT FILE4_10.CODE," & _
          " FILE4_10.desca," & _
          " FILE4_50.DESCA " & _
          " FROM  FILE4_10 LEFT JOIN FILE4_50 ON FILE4_10.[GROUP] = FILE4_50.CODE"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by FILE4_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·„Ã„Ê⁄…"
listarray(0, 1) = "(%%FILE4_10.DESCA%% OR %%FILE4_50.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·„Ê—œ"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE4_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·„Ê—œÌ‰"
oSearch.Show 1
End Sub
Sub partLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT FILE8_71.CODE," & _
          " FILE8_71.desca" & _
          " FROM  FILE8_71"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE8_71.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%FILE8_71.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·‘—Ìﬂ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·‘—Ìﬂ"
GrdArray(1, 1) = 4000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE4_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·‘—ﬂ«¡"
oSearch.Show 1
End Sub
Sub BankLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm
                 
cString = "SELECT FILE5_10.CODE," & _
          "FILE5_10.desca," & _
          " FILE5_50.DESCA " & _
          " FROM  FILE5_10 " & _
          " LEFT JOIN FILE5_50 ON FILE5_10.[GROUP] = FILE5_50.CODE"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE5_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·„Ã„Ê⁄…"
listarray(0, 1) = "(%%FILE5_10.DESCA%% OR %%FILE5_50.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·»‰ﬂ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·»‰ﬂ"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE5_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·»‰Êﬂ"
oSearch.Show 1
End Sub
Sub BoxLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT FILE0_50.CODE," & _
          "FILE0_50.DESCA," & _
          " branch.desca AS BRANCH_DESCA" & _
          " FROM FILE0_50 INNER JOIN branch ON FILE0_50.BRANCH = branch.code"
    
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE0_50.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%FILE0_50.DESCA%% OR %%BRANCH.DESCA%%)"

listarray(1, 0) = "«·›—⁄"
listarray(1, 1) = "(FILE0_50.BRANCH = 'cFilter')"
listarray(1, 2) = "SELECT BRANCH.CODE,BRANCH.DESCA FROM BRANCH"
listarray(1, 3) = "CODE"
listarray(1, 4) = "DESCA"

GrdArray(0, 0) = "ﬂÊœ «·Œ“‰…"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·›—⁄"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE5_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·»‰Êﬂ"
oSearch.Show 1
End Sub
Sub accountLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT ACCOUNT.CODE," & _
          " ACCOUNT.desca," & _
          " ACCOUNT.WITH_BRANCH " & _
          " FROM  ACCOUNT"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by ACCOUNT.code"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%ACCOUNT.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·Õ”«»"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "≈”„ «·Õ”«»"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "»›—⁄"
GrdArray(2, 1) = 0

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "ACCOUNT.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·Õ”«»« "
oSearch.Show 1
End Sub
Sub BranchLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT Branch.CODE," & _
          " Branch.desca" & _
          " FROM  Branch"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by Branch.code"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·›—⁄"
listarray(0, 1) = "(%%Branch.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·›—⁄"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "≈”„ «·›—⁄"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "BRANCH.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·›—Ê⁄"
oSearch.Show 1
End Sub
Sub StoreLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT file0_40.CODE," & _
          " file0_40.desca" & _
          " FROM  file0_40"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by file0_40.code"
Generalarray(3) = 4500
Generalarray(5) = True

listarray(0, 0) = "«·„Œ“‰"
listarray(0, 1) = "(%%file0_40.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·„Œ“‰"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "≈”„ «·„Œ“‰"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "file0_40.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·„Œ«“‰"
oSearch.Show 1
End Sub
Sub ChargeLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT FILE8_51.CODE," & _
          " FILE8_51.DESCA," & _
          " FILE8_51.desca AS GROUP_DESCA" & _
          " FROM FILE8_51 LEFT JOIN FILE8_52 ON FILE8_51.[MAINGROUP] = FILE8_52.CODE"
    
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE8_51.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·„’—Ê›"
listarray(0, 1) = "(%%FILE8_51.DESCA%% OR %%FILE8_52.DESCA%%)"

listarray(1, 0) = "„Ã„Ê⁄… —∆Ì”Ì…"
listarray(1, 1) = "(FILE8_51.[MAINGROUP] = 'cFilter')"
listarray(1, 2) = "SELECT FILE8_52.CODE,FILE8_52.DESCA FROM FILE8_52"
listarray(1, 3) = "CODE"
listarray(1, 4) = "DESCA"

GrdArray(0, 0) = "ﬂÊœ «·„’—Ê›"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«·„’—Ê›"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE5_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.sCaption = "≈” ⁄·«„ «·„’«—Ì›"
oSearch.Show 1
End Sub
Sub IncomeLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT FILE8_61.CODE," & _
          "FILE8_61.DESCA," & _
          " FILE8_61.desca AS GROUP_DESCA" & _
          " FROM FILE8_61 LEFT JOIN FILE8_62 ON FILE8_61.[MAINGROUP] = FILE8_62.CODE"
    
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE8_61.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«Ì—«œ"
listarray(0, 1) = "(%%FILE8_61.DESCA%% OR %%FILE8_62.DESCA%%)"

listarray(1, 0) = "„Ã„Ê⁄… —∆Ì”Ì…"
listarray(1, 1) = "(FILE8_61.MAINGROUP = 'cFilter')"
listarray(1, 2) = "SELECT FILE8_62.CODE,FILE8_62.DESCA FROM FILE8_62"
listarray(1, 3) = "CODE"
listarray(1, 4) = "DESCA"

GrdArray(0, 0) = "ﬂÊœ «·«Ì—«œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«·«Ì—«œ"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE5_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·«Ì—«œ« "
oSearch.Show 1
End Sub
Sub SalesManLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm
'                       0                   1
cString = "SELECT FILE6_25.CODE, FILE6_25.DESCA, branch.desca" & _
          " FROM   FILE6_25 LEFT OUTER JOIN  branch ON FILE6_25.BRANCH = branch.code"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by FILE6_25.CODE"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·›—⁄-«·ﬂÊœ"
listarray(0, 1) = "(@@FILE6_25.CODE@@4 OR %%FILE6_25.DESCA%% OR %%BRANCH.DESCA%%)"

GrdArray(0, 0) = "ﬂÊœ «·»«∆⁄"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·»«∆⁄"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·›—⁄"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE3_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·»«∆⁄Ì‰"
oSearch.Show 1
End Sub
Public Sub BankVisaLookup(oForm As Form, oSearch As Form, Optional pName As String = "", Optional pFilter As String)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select FILE5_10.CODE,FILE5_10.DESCA" & _
                  " FROM FILE5_10 INNER JOIN VISA_CODES ON FILE5_10.CODE = VISA_CODES.BANK" & _
                  " WHERE VISA_CODES.STOPED = 0"
Generalarray(2) = " GROUP BY FILE5_10.CODE,FILE5_10.DESCA ORDER BY FILE5_10.CODE"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ -«·≈”„"
listarray(0, 1) = "(@@FILE5_10.CODE@@6 or %%FILE5_10.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ »‰Êﬂ «·›Ì“"
oSearch.Show 1
End Sub
Public Sub visaLookup(oForm As Form, oSearch As Search_abd, sBank As String, Optional cFilter As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(6, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT " & _
                  " VISA_CODES.CODE," & _
                  " VISA_CODES.DESCA," & _
                  " FILE5_10.DESCA," & _
                  " VISA_CODES.BANK_NAME," & _
                  " VISA_CODES.RATE_INSTALL," & _
                  " VISA_CODES.RATE," & _
                  " VISA_CODES.WITH_POINTS" & _
                  " FROM  VISA_CODES INNER JOIN FILE5_10 ON VISA_CODES.BANK = FILE5_10.CODE" & _
                  " WHERE VISA_CODES.STOPED = 0" & _
                  " AND VISA_CODES.BANK =  " & MyParn(sBank)

Generalarray(2) = "ORDER BY VISA_CODES.CODE"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ -«·≈”„"
listarray(0, 1) = "(%%FILE5_10.DESCA%%)"

listarray(1, 0) = "«·»‰ﬂ"
listarray(1, 1) = "(VISA_CODES.[BANK_NAME] = 'cFilter')"
listarray(1, 2) = "SELECT DISTINCT BANK_NAME,BANK_NAME AS DESCA FROM VISA_CODES WHERE STOPED = 0 AND BANK =  " & MyParn(sBank)
listarray(1, 3) = "BANK_NAME"
listarray(1, 4) = "BANK_NAME"


GrdArray(0, 0) = "Ê”Ì·… «·œ›⁄"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "Ê”Ì·… «·œ›⁄"
GrdArray(1, 1) = 2500

GrdArray(2, 0) = "«·»‰ﬂ"
GrdArray(2, 1) = 2500

GrdArray(3, 0) = "»‰ﬂ Ê”Ì·… «·œ›⁄"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "‰”»… «·⁄„Ì·"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "‰”»… «·»‰ﬂ"
GrdArray(5, 1) = 1000

GrdArray(6, 0) = "ﬂ«—  ‰ﬁ«ÿ"
GrdArray(6, 1) = 0

Dim aMerge As Variant
aMerge = Array(2, 3)

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.sCaption = "√ﬂÊ«œ Ê”«∆· œ›⁄ «·›Ì“«"
oSearch.aMerge = aMerge
oSearch.bUnload = True
oSearch.Show 1
End Sub
Public Function retItemCode(pItem As String, con As ADODB.Connection) As String
If Trim(pItem) = "" Then Exit Function
retItemCode = myField("Select  [dbo].[fn_ret_item](" & MyParn(pItem) & ")", con) & ""
End Function
Public Function UpdateInvTotal(pDoc_no As String, con As ADODB.Connection) As Boolean
con.Execute "UPDATE FILE6_20H " & _
            " SET FILE6_20H.LATE = FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT" & _
            " FROM FILE6_20H" & _
            " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
End Function
Public Sub OrdersLookup(oForm As Form, oSearch As Search_abd, Optional cFilter As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = oForm
cString = "SELECT FILE6_51H.Doc_No," & _
           "CONVERT(VARCHAR(10),FILE6_51H.DATE,111)," & _
           " FILE3_10.DESCA " & _
           " FROM FILE6_51H LEFT JOIN FILE3_10 ON FILE6_51H.CODE = FILE3_10.CODE"

If pFilter <> "" Then
    cString = cString & " WHERE " & pFilter
End If

Generalarray(1) = cString
Generalarray(2) = "ORDER BY FILE6_51H.[DATE]," & _
                  " FILE6_51H.DOC_NO "
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "%%FILE6_51H.DOC_NO%%"

listarray(1, 0) = " «—ÌŒ «·„” ‰œ"
listarray(1, 1) = "(##[Date]## OR %%FILE3_10.DESCA%%)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1600

GrdArray(2, 0) = "«·⁄„Ì·"
GrdArray(2, 1) = 3000

Dim aFilter As Variant
aFilter = AddFlag(aFilter, "FILTER", True)
aFilter = AddFlag(aFilter, "FIELD", "DOC_NO")

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.sCaption = "«” ⁄·«„ ÿ·»Ì«  «·⁄„·«¡"
oSearch.aFilter = aFilter
oSearch.Show 1
End Sub
Public Sub ChqSupLookup(oForm As Form, oSearch As Search_abd, Optional cFilter As String = "")
Dim Generalarray(5)
Dim listarray(2, 5)
Dim GrdArray(9, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE5_21.SER_NO," & _
                  "FILE5_21.DESCA1," & _
                  "FILE5_21.BNF," & _
                  " FORMAT(FILE5_21.DATE_1 ,'yyyy/M/d')," & _
                  " FORMAT(FILE5_21.DATE_R, 'yyyy/M/d')," & _
                  "[VALUE]," & _
                  "CHK_ID," & _
                  "FILE8_51.DESCA," & _
                  "FILE5_21.DESCA," & _
                  "FILE5_21.ACC" & _
                  " FROM FILE5_21 " & _
                  " LEFT JOIN FILE8_51 ON FILE8_51.CODE = FILE5_21.CHARGE "
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & " WHERE " & cFilter
End If

Generalarray(2) = " ORDER BY FILE5_21.SER_NO"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "„”·”· «Ê „” ›Ìœ -  «—ÌŒ  Õ—Ì—-—ﬁ„ «·‘Ìﬂ- „’—Ê›-»Ì«‰"
listarray(0, 1) = "(%%FILE8_51.desca%% or %%FILE5_21.SER_NO%% or %%FILE5_21.desca1%% or %%FILE5_21.desca%% or chk_id Like '%cFilter%' Or ##Date_R##)"

listarray(1, 0) = " «—ÌŒ «” Õﬁ«ﬁ"
listarray(1, 1) = " ##Date_1##"

listarray(2, 0) = "«·ﬁÌ„…"
listarray(2, 1) = "**[value]**"

GrdArray(0, 0) = "„”·”·"
GrdArray(0, 1) = 800

GrdArray(1, 0) = "‘Ìﬂ „‰"
GrdArray(1, 1) = 2000

GrdArray(2, 0) = "«·„” ›Ìœ"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«” Õﬁ«ﬁ"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = " Õ—Ì—"
GrdArray(4, 1) = 1300

GrdArray(5, 0) = "ﬁÌ„…"
GrdArray(5, 1) = 1200

GrdArray(6, 0) = "—ﬁ„ «·‘Ìﬂ"
GrdArray(6, 1) = 1400

GrdArray(7, 0) = "«·„’—Ê›"
GrdArray(7, 1) = 2000

GrdArray(8, 0) = "«·»Ì«‰"
GrdArray(8, 1) = 2000

GrdArray(9, 0) = "Õ”«»"
GrdArray(9, 1) = 1000

Dim aFilter As Variant
aFilter = AddFlag(aFilter, "FILTER", True)
aFilter = AddFlag(aFilter, "FIELD", "SER_NO")

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.sCaption = "«Ê—«ﬁ œ›⁄"
oSearch.aFilter = aFilter
oSearch.Show 1
End Sub
Public Function retSalesDate(con As ADODB.Connection, Optional pBranch As String = "") As String
Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "SELECT DSALES FROM DSALES"
If pBranch <> "" Then
    cString = cString & " WHERE BRANCH = " & MyParn(pBranch)
End If
Set loctable = cmd(cString, con).Execute
If Not loctable.EOF Then
    retSalesDate = myFormat(loctable!dSales)
End If
End Function
