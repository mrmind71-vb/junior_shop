Attribute VB_Name = "special2"
Public bClient As Boolean
Public bVersion As Boolean
Public rsBranch As ADODB.Recordset
Public rsMall As ADODB.Recordset
Public rsBranches As ADODB.Recordset
Public rsUser As ADODB.Recordset
Public rsAddress As ADODB.Recordset
Public Const userid_vpn = "data_pro"
Public Const password_vpn = "2010"
Public vpn As Boolean
Public Const sBt As String = vbCrLf & vbTab
Public Sub SalesLadingLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(4, 5)
Dim GrdArray(8, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1
cString = "SELECT FILE6_20H.DOC_NO," & _
          "FILE6_20H.DOC_NO2," & _
          "CONVERT(VARCHAR(10),FILE6_20H.DATE,111)," & _
          "FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT," & _
          "FILE6_20H.[NAME]," & _
          "FILE6_20H.PHONE, " & _
          "FILE6_20H.SHIP_NO, " & _
          "FILE6_20H.ONLINE_DOC, " & _
          "SHIP.DESCA" & _
          " FROM  FILE6_20H " & _
          " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE AND FILE0_40.ONLINE = 1" & _
          " LEFT JOIN SHIP ON FILE6_20H.SHIP = SHIP.DESCA" & _
          " LEFT JOIN LADING_BILL ON FILE6_20H.DOC_NO = LADING_BILL.DOC_NO" & _
          " WHERE FILE6_20H.ONLINE = 2" & _
          " AND FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT > 0" & _
          " AND FILE6_20H.ISCLOSED = 1" & _
          " AND (FILE6_20H.SHIP_NO IS NOT NULL)" & _
          " AND (FILE6_20H.SHIP IS NOT NULL)"

'If cBranch <> "" Then
'    cString = cString & " AND FILE6_20H.BRANCH = " & MyParn(cBranch)
'End If

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " AND " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by FILE6_20H.DATE DESC,FILE6_20H.DOC_NO"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "(FILE6_20H.DOC_NO = 'cFilter' OR FILE6_20H.DOC_NO2 = 'cFilter')"

listarray(1, 0) = "«·«”„-«·⁄‰Ê«‰"
listarray(1, 1) = "(%%FILE6_20H.NAME%% OR %%FILE6_20H.ADDRESS%%)"

listarray(2, 0) = "—ﬁ„ «·»Ê·Ì’…-—ﬁ„ «·ÿ·»Ì…"
listarray(2, 1) = "(SHIP_NO LIKE 'cFilter%' OR SHIP LIKE 'cFilter%' OR FILE6_20H.ONLINE_DOC LIKE 'cFilter%')"

listarray(3, 0) = "«· ·Ì›Ê‰"
listarray(3, 1) = "(PHONE LIKE 'cFilter%')"

listarray(4, 0) = "«·‘—ﬂ…"
listarray(4, 1) = "(SHIP.CODE = 'cFilter')"
listarray(4, 2) = "SELECT  CODE,DESCA FROM SHIP"
listarray(4, 3) = "CODE"
listarray(4, 4) = "DESCA"


GrdArray(0, 0) = "ﬂÊœ «·„” ‰œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "—›„ «·›« Ê—…"
GrdArray(1, 1) = 1600

GrdArray(2, 0) = " «—ÌŒ «·›« Ê—…"
GrdArray(2, 1) = 1300

GrdArray(3, 0) = "«·≈Ã„«·Ì"
GrdArray(3, 1) = 1200

GrdArray(4, 0) = "«·«”„"
GrdArray(4, 1) = 2500

GrdArray(5, 0) = "«·⁄‰Ê«‰"
GrdArray(5, 1) = 3500

GrdArray(6, 0) = "«· ·Ì›Ê‰"
GrdArray(6, 1) = 1700

GrdArray(7, 0) = "—ﬁ„ «·»Ê·Ì’…"
GrdArray(7, 1) = 1400

GrdArray(8, 0) = "‘—ﬂ… «·‘Õ‰"
GrdArray(8, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE6_20H.DOC_NO")
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
oSearch.Caption = "≈” ⁄·«„ ›Ê« Ì— «Ê‰ ·«Ì‰"
oSearch.Show 1
End Sub
Public Sub SalesOnlineLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(4, 5)
Dim GrdArray(8, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1
cString = "SELECT FILE6_20H.DOC_NO," & _
          "FILE6_20H.DOC_NO2," & _
          "CONVERT(VARCHAR(10),FILE6_20H.DATE,111)," & _
          "FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT," & _
          "FILE6_20H.[NAME]," & _
          "FILE6_20H.PHONE, " & _
          "FILE6_20H.SHIP_NO, " & _
          "SHIP.DESCA" & _
          " FROM  FILE6_20H " & _
          " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE AND FILE0_40.ONLINE = 1" & _
          " LEFT JOIN SHIP ON FILE6_20H.SHIP = SHIP.DESCA" & _
          " WHERE FILE6_20H.ONLINE = 2"

'If cBranch <> "" Then
'    cString = cString & " AND FILE6_20H.BRANCH = " & MyParn(cBranch)
'End If

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " AND " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by FILE6_20H.DATE DESC,FILE6_20H.DOC_NO"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "(FILE6_20H.DOC_NO = 'cFilter' OR FILE6_20H.DOC_NO2 = 'cFilter')"

listarray(1, 0) = "«·«”„-«·⁄‰Ê«‰"
listarray(1, 1) = "(%%FILE6_20H.NAME%% OR %%FILE6_20H.ADDRESS%%)"

listarray(2, 0) = "—ﬁ„ «·»Ê·Ì’…"
listarray(2, 1) = "(SHIP_NO LIKE 'cFilter%')"

listarray(3, 0) = "«· ·Ì›Ê‰"
listarray(3, 1) = "(PHONE LIKE 'cFilter%')"

listarray(4, 0) = "«·‘—ﬂ…"
listarray(4, 1) = "(SHIP.CODE = 'cFilter')"
listarray(4, 2) = "SELECT  CODE,DESCA FROM SHIP"
listarray(4, 3) = "CODE"
listarray(4, 4) = "DESCA"


GrdArray(0, 0) = "ﬂÊœ «·„” ‰œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "—›„ «·›« Ê—…"
GrdArray(1, 1) = 1600

GrdArray(2, 0) = " «—ÌŒ «·›« Ê—…"
GrdArray(2, 1) = 1300

GrdArray(3, 0) = "«·≈Ã„«·Ì"
GrdArray(3, 1) = 1200

GrdArray(4, 0) = "«·«”„"
GrdArray(4, 1) = 2500

GrdArray(5, 0) = "«·⁄‰Ê«‰"
GrdArray(5, 1) = 3500

GrdArray(6, 0) = "«· ·Ì›Ê‰"
GrdArray(6, 1) = 1700

GrdArray(7, 0) = "—ﬁ„ «·»Ê·Ì’…"
GrdArray(7, 1) = 1400

GrdArray(8, 0) = "‘—ﬂ… «·‘Õ‰"
GrdArray(8, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE6_20H.DOC_NO")
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
oSearch.Caption = "≈” ⁄·«„ ›Ê« Ì— «Ê‰ ·«Ì‰"
oSearch.Show 1
End Sub
Sub ShipLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1
cString = "SELECT SHIP.CODE," & _
          " SHIP.DESCA" & _
          " FROM   SHIP"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by SHIP.CODE"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·ﬂÊœ"
listarray(0, 1) = "(**CODE** OR %%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1500

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ ‘—ﬂ«  «·‘Õ‰"
oSearch.Show 1
End Sub
Sub LadingLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1
cString = "SELECT LADING_CODES.CODE," & _
          " LADING_CODES.DESCA" & _
          " FROM   LADING_CODES"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by LADING_CODES.CODE"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·ﬂÊœ"
listarray(0, 1) = "(**CODE** OR %%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1500

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «‰Ê⁄ »Ê«·’ ‘Õ‰ „”»ﬁ…"
oSearch.Show 1
End Sub
Public Sub ClientOrderLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = oForm
cString = "SELECT FILE6_52H.Doc_No," & _
           " FILE6_52H.PO_NO," & _
           " Convert(Varchar(10),FILE6_52H.DATE,111)," & _
           " FILE3_10.DESCA " & _
           " FROM FILE6_52H LEFT JOIN FILE3_10 ON FILE6_52H.CODE = FILE3_10.CODE"

If pFilter <> "" Then cString = cString & " WHERE " & pFilter

Generalarray(1) = cString
Generalarray(2) = "order by FILE6_52H.[DATE] DESC,FILE6_52H.DOC_NO DESC"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì·"
listarray(0, 1) = "%%FILE3_10.DESCA%%"

listarray(1, 0) = " «—ÌŒ «·„” ‰œ"
listarray(1, 1) = "(##[Date]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1600

GrdArray(2, 0) = "—ﬁ„ «·ÿ·»Ì…"
GrdArray(2, 1) = 1600

GrdArray(3, 0) = "«·⁄„Ì·"
GrdArray(3, 1) = 3000

Dim aFilter As Variant
aFilter = AddFlag(aFilter, "FILTER", True)
aFilter = AddFlag(aFilter, "FIELD", "DOC_NO")

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.sCaption = "«” ⁄·«„  ”·Ì„ ÿ·»Ì«  «·⁄„·«¡"
oSearch.aFilter = aFilter
oSearch.Show 1
End Sub
Public Sub clientLookup2(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = oForm
cString = "SELECT FILE3_10.CODE," & _
           " FILE3_10.DESCA," & _
           " FILE3_50.DESCA" & _
           " FROM FILE3_10 " & _
           " LEFT JOIN FILE3_50 ON FILE3_10.[GROUP] = FILE3_50.CODE"

If pFilter <> "" Then cString = cString & " WHERE " & pFilter

Generalarray(1) = cString
Generalarray(2) = "order by FILE3_10.DESCA"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì·"
listarray(0, 1) = "%%FILE3_10.DESCA%%"

listarray(1, 0) = "«·›—⁄"
listarray(1, 1) = "(FILE3_10.[GROUP] = 'cFilter')"
listarray(1, 2) = "SELECT CODE,DESCA FROM FILE3_50"
listarray(1, 3) = "CODE"
listarray(1, 4) = "DESCA"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 900

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2500

Dim aFilter As Variant
aFilter = AddFlag(aFilter, "FILTER", True)
aFilter = AddFlag(aFilter, "FIELD", "CODE")

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.sCaption = "«” ⁄·«„ «·⁄„·«¡"
oSearch.aFilter = aFilter
oSearch.Show 1
End Sub
Public Sub PurOrderLook(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(2, 4)
Dim GrdArray(7, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE6_50H.DOC_NO," & _
                  "FORMAT(FILE6_50H.[DATE],'yyyy/M/d')," & _
                  "FILE4_10.Desca," & _
                  "FILE6_50H.INV_NO," & _
                  "STORE_BR.DESCA," & _
                  "ISEND_ORDER" & _
                  " FROM FILE6_50H " & _
                  " INNER JOIN file4_10 ON FILE6_50H.CODE = FILE4_10.CODE" & _
                  " INNER JOIN STORE_BR ON FILE6_50H.STORE = STORE_BR.CODE"
If cBranch = "00" And (Not bOpt5) Then
    Generalarray(1) = Generalarray(1) & _
                     " where store in " & _
                     " (select store from USERSHOP AS F where F.code = " & nusercode & " ) "
End If

Generalarray(2) = "ORDER BY DATE DESC," & _
                   "DOC_NO DESC "
Generalarray(3) = 6000
Generalarray(5) = True


listarray(0, 0) = "—ﬁ„ «·›« Ê—…-≈”„ «·„Ê—œ"
listarray(0, 1) = "(FILE6_50H.INV_NO LIKE '%cFilter%'" & _
                  " OR FILE4_10.DESCA like 'cFilter%')"
                  
listarray(1, 0) = "—ﬁ„ «·„” ‰œ- «—ÌŒ «·„” ‰œ"
listarray(1, 1) = "(FILE6_50H.DOC_NO LIKE '%cFilter%'" & _
                  " OR ##DATE##)"
                
listarray(2, 0) = "«·›—⁄"
listarray(2, 1) = "%%STORE_BR.DESCA%%"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1400

GrdArray(2, 0) = "≈”„ «·„Ê—œ"
GrdArray(2, 1) = 3000

GrdArray(3, 0) = "—ﬁ„ ›« Ê—…"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "«·„Œ“‰"
GrdArray(4, 1) = 1500

GrdArray(5, 0) = "„€·ﬁ…"
GrdArray(5, 1) = 1000

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.sCaption = "«” ⁄·«„ ÿ·»Ì«  «·‘—«¡"
oSearch.Show 1
End Sub
Sub AccLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1
cString = "SELECT CODE," & _
          " DESCA" & _
          " FROM   ACC0_10"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by CODE"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·ﬂÊœ"
listarray(0, 1) = "(**CODE** OR %%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1500

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)

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
Public Sub fixSql(con As ADODB.Connection, Optional sMarker As String = "GO")
Dim sb As New ChilkatStringBuilder
Dim fs As New FileSystemObject
Dim success As Long

Dim cString As String

For i = 1 To 10
    success = sb.LoadFile(App.Path & "\sql_fix\fix" & i & ".sql", "utf-8")
    If success = 1 Then
        Do Until sb.length = 0
            cString = sb.GetBefore(sMarker, True)
            If cString <> "" Then
                On Error Resume Next
                con.Execute cString
                Err.Clear
            End If
        Loop
    Else
        Exit Sub
    End If
Next
End Sub
Public Function GetNumbersFromString(ByVal inputString As String) As String
    Dim i As Long
    Dim char As String
    Dim resultString As String

    For i = 1 To Len(inputString)
        char = Mid(inputString, i, 1)
        If IsNumeric(char) Then
            resultString = resultString & char
        End If
    Next i

    GetNumbersFromString = resultString
End Function
Public Sub FillCheckbox(ByRef chkArray As Object, ByVal cSql As String, con As ADODB.Connection)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Set rs = cmd(cSql, con).Execute
        
        
    ' Ensure the recordset is open and contains data
    If Not rs.EOF Then
       Do Until rs.EOF Or i > chkArray.UBound
            chkArray(i).Visible = True
            chkArray(i).Caption = rs.Fields(1).Value
            chkArray(i).Tag = rs.Fields(0).Value
            i = i + 1
            rs.MoveNext
       Loop
    End If
    
    rs.Close
    Set rs = Nothing
End Sub
Public Function getCheckBox(chkArray As Object)
Dim i As Long
For i = 0 To chkArray.UBound
    If chkArray(i).Tag <> "" Then
        If chkArray(i).Value = 1 Then
            getCheckBox = getCheckBox & Tr(getCheckBox, ",") & chkArray(i).Tag
        End If
    End If
Next
End Function
Public Function openConDoc(ByRef con As ADODB.Connection, Optional pCatalog As String = "SHOP_DOCS") As String
Dim cServerName As String
Dim cUserId As String
Dim cPassword As String
Dim cString As String

cServerName = servername_vpn
cUserId = userid_vpn
cPassword = password_vpn

cString = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
                & "catalog=" & pCatalog & ";user id=" & cUserId & ";" & "password=" & cPassword & _
                ";Encrypt=True" & _
                ";TrustServerCertificate=False" & _
                ";Timeout=10"
openConDoc = openCon(con, cString)
End Function
Public Function CountFilesWithExt(DirectoryToSearch As String, Extension As String) As Long
    Dim lCounter As Long
    Dim sTemp As String
    
    lCounter = 0
    ' Start the loop by getting the first file matching the pattern
    sTemp = Dir(DirectoryToSearch & "\*." & Extension)
    
    ' Loop as long as files are found
    Do While sTemp <> ""
        lCounter = lCounter + 1
        ' Get the next file in the directory
        sTemp = Dir
    Loop
    
    CountFilesWithExt = lCounter
End Function
Public Function accountRs(con As ADODB.Connection, Optional id_image As String, Optional id_cash As String) As ADODB.Recordset
Dim aPrm As Variant
If id_image <> "" Then
    aPrm = AddFlag(aPrm, "id_image", id_image)
ElseIf id_cash <> "" Then
    aPrm = AddFlag(aPrm, "id_cash", id_cash)
End If
Set accountRs = cmd("[dbo].[sp_get_account_image]", con, adStoredProc, aPrm).Execute
End Function
Public Sub ZoneLookup(oForm As Form, oSearch As Form, pDoc_no As String, Optional sAddRow As String, Optional bNoRef As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT ZONE FROM FILE0_100 WHERE DOC_NO = " & MyParn(sDoc_no)
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by ZONE"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·›—⁄"
listarray(0, 1) = "%%BRANCH.DESCA%%"

GrdArray(0, 0) = "ﬂÊœ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«·›—⁄"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "BRANCH.CODE")
    oSearch.aFilter = aFilter
End If

oSearch.sCaption = "«” ⁄·«„ «·›—Ê⁄"
Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.bNoRef = bNoRef
oSearch.aAddRow = aRow
oSearch.Show 1
End Sub
Public Function ISODate(ByVal strDate As String) As Date
    Dim dtePart As String
    Dim tmePart As String
    
    ' Extract Date (YYYY-MM-DD) and Time (HH:MM:SS)
    dtePart = Left(strDate, 10)
    tmePart = Mid(strDate, 12, 8)
    
    ' Combine and convert
    ISODate = CDate(dtePart & " " & tmePart)
End Function
Public Sub SearchGrd(pGrid As Object, searchText As String, col As Long, Optional bExct As Boolean = False)
Dim R As Long

' Stop redrawing for better performance
pGrid.Redraw = flexRDNone

For R = pGrid.FixedRows To pGrid.Rows - 1
    ' Check if search text exists in specific column (e.g., Column 1)
    If bExct Then
        If Trim(searchText) = "" Then
            pGrid.RowHeight(R) = -1
        ElseIf LCase$(pGrid.TextMatrix(R, col)) = searchText Then
            ' Reset row height to default (-1) to show it
            pGrid.RowHeight(R) = -1
        Else
            ' Set row height to 0 to hide it
            'pGrid.RowHidden(r) = True
            pGrid.RowHeight(R) = 0
        End If
        
    Else
        If Trim(searchText) = "" Then
            pGrid.RowHeight(R) = -1
        ElseIf InStr(1, LCase$(pGrid.TextMatrix(R, col)), searchText) > 0 Then
            ' Reset row height to default (-1) to show it
            pGrid.RowHeight(R) = -1
        Else
            ' Set row height to 0 to hide it
            'pGrid.RowHidden(r) = True
            pGrid.RowHeight(R) = 0
        End If
    End If
Next R
'salesfrm.grid1.Redraw = flexRDBuffered
' Resume redrawing
pGrid.Redraw = flexRDBuffered
End Sub
Public Function IsDgt(ByVal strData As String) As Boolean
    ' Check if string is empty
    If Len(Trim(strData)) = 0 Then
        IsDgt = False
        Exit Function
    End If
    
    ' Pattern checks if the entire string consists only of 0-9
    ' The String(Len(strData), "#") creates a pattern like "####"
    IsDgt = (Trim(strData) Like String(Len(Trim(strData)), "#"))
End Function
Public Function fixCostStock(con As ADODB.Connection, Optional pDoc_no As String, Optional pItem As String, Optional pDate1 As String = "", Optional pDate2 As String = "", Optional ByRef nUpdated As Long = 0, Optional ByRef sError As String) As Boolean
Dim aPrm As Variant
If pDoc_no <> "" Then
    aPrm = AddFlag(aPrm, "doc_no", pDoc_no)
End If

If IsDgt(pItem) Then
    aPrm = AddFlag(aPrm, "item", pItem)
End If

If IsDate(pDate1) Then
    aPrm = AddFlag(aPrm, "date", myFormat_sp(pDate1))
End If

If IsDate(pDate2) Then
    aPrm = AddFlag(aPrm, "date", myFormat_sp(pDate2))
End If

On Error GoTo myerror
Dim Com As New ADODB.command
Set Com = cmd("[dbo].[sp_stock_cost]", con, adStoredProc, aPrm)
Com.Execute

nUpdated = Nz(Com.Parameters("@COUNT_UPDATE"))

fixCostStock = True
Exit Function
myerror:
sError = Err.Description
Err.Clear
End Function
Public Sub ManLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT FILE6_25.CODE," & _
          "FILE6_25.DESCA," & _
          " branch.desca AS BRANCH_DESCA" & _
          " FROM FILE6_25  " & _
          " INNER JOIN branch ON FILE6_25.BRANCH = branch.code"
    
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE6_25.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%FILE6_25.DESCA%% OR %%BRANCH.DESCA%%)"

listarray(1, 0) = "«·›—⁄"
listarray(1, 1) = "(FILE6_25.BRANCH = 'cFilter')"
listarray(1, 2) = "SELECT BRANCH.CODE,BRANCH.DESCA FROM BRANCH"
listarray(1, 3) = "CODE"
listarray(1, 4) = "DESCA"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«·≈”„"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·›—⁄"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE6_25.CODE")
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
Public Sub ManOnlineLook(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT MAN_ONLINE_CODES.CODE," & _
          "MAN_ONLINE_CODES.DESCA," & _
          " FROM MAN_ONLINE_CODES  "

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by MAN_ONLINE_CODES.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%MAN_ONLINE_CODES.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«·≈”„"
GrdArray(1, 1) = 6000


searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "MAN_ONLINE_CODES.CODE")
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
oSearch.Caption = "≈” ⁄·«„ «·„‰«œÌ»"
oSearch.Show 1
End Sub
Public Sub StagesOnlineLook(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT STAGES_CODES.CODE," & _
          "STAGES_CODES.DESCA" & _
          " FROM STAGES_CODES"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by STAGES_CODES.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·„—Õ·…"
listarray(0, 1) = "(%%STAGES_CODES.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "««·„—Õ·…"
GrdArray(1, 1) = 6000


searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "STAGES_CODES.CODE")
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
oSearch.Caption = "≈” ⁄·«„ „—«Õ· «· ÃÂÌ“"
oSearch.Show 1
End Sub


