Attribute VB_Name = "Special3"
Public pcName As String
Public obj As ChilkatGlobal
Public sOnlineStore As String
Public servername_vpn
Public sStoreOnline As String
Public sBranchOnline As String
Public Function Nz(value As Variant, Optional ValueIfNull As Variant = 0) As Variant
    If IsNull(value) Then
        Nz = ValueIfNull
    Else
        Nz = value
    End If
End Function
Sub SectionLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0
cString = "SELECT FILE1_10SC.CODE," & _
          " FILE1_10SC.desca" & _
          " FROM  FILE1_10SC"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE1_10SC.code"
Generalarray(3) = 4000
Generalarray(5) = True

listarray(0, 0) = "«·ﬁ”„"
listarray(0, 1) = "(%%FILE1_10SC.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·ﬁ”„"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "≈”„ «·ﬁ”„"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE1_10SC.CODE")
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
oSearch.Caption = "≈” ⁄·«„ «·«ﬁ”«„"
oSearch.Show 1
End Sub
Sub OnlineTypeLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "", Optional sControl As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0
cString = "SELECT CODE," & _
          " desca" & _
          " FROM  ONLINE_TYPE_CODES"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by CODE"
Generalarray(3) = 4000
Generalarray(5) = True

listarray(0, 0) = "«·‰Ê⁄"
listarray(0, 1) = "(%%DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·‰Ê⁄"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow
oSearch.sControl = sControl
searchArray = Array(Generalarray, listarray, GrdArray)
'oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·«‰Ê«⁄"
oSearch.Show 1
End Sub
Sub PayTypeLook(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "", Optional sControl As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0
cString = "SELECT Payment_type," & _
          " Payment_type as p" & _
          " FROM  FILE6_90H" & _
          " WHERE PAYMENT_TYPE IS NOT NULL"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " AND " & cWhere

Generalarray(1) = cString

Generalarray(2) = " Group by Payment_type Order by Payment_type"
Generalarray(3) = 4000
Generalarray(5) = True

listarray(0, 0) = "‰Ê⁄ «·”œ«œ"
listarray(0, 1) = "(%%Payment_type%%)"


GrdArray(0, 0) = "‰Ê⁄ «·”œ«œ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "‰Ê⁄ «·”œ«œ"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "Payment_Type")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow
oSearch.sControl = sControl
searchArray = Array(Generalarray, listarray, GrdArray)
'oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «‰Ê«⁄ «·”œ«œ"
oSearch.Show 1
End Sub
Sub FactLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT FACT.CODE," & _
          " FACT.desca" & _
          " FROM  FACT"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FACT.code"
Generalarray(3) = 4000
Generalarray(5) = True

listarray(0, 0) = "«·„’‰⁄"
listarray(0, 1) = "(%%FACT.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·„’‰⁄"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "≈”„ «·„’‰⁄"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FACT.CODE")
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
oSearch.Caption = "≈” ⁄·«„ «·„’«‰⁄"
oSearch.Show 1
End Sub
Public Function fnDateSales(con As ADODB.Connection) As String
If cBranch = "00" Then
    fnDateSales = myFormat(Date)
    Exit Function
End If

Dim locTable As New ADODB.Recordset
Dim cString As String
Set locTable = cmd("SELECT dSales FROM DSALES" & _
            " WHERE BRANCH = " & MyParn(cBranch), con).Execute
If Not locTable.EOF Then
    fnDateSales = myFormat(locTable!dSales)
End If
locTable.Close
End Function
Public Function rsDateBranch(Optional pBranch As String, Optional con As ADODB.Connection) As String
If pBranch = "00" Then
    rsDateSales = myFormat(Date)
Else
    rsDateBranch = rsValue("SELECT dSales FROM DSALES" & _
                " WHERE BRANCH = " & MyParn(pBranch), con) & ""
End If
End Function
Public Function fnBalance(pItem As String, con As ADODB.Connection, Optional pstore As String = "", Optional pDate As String = "", Optional pId As String = "") As Long
Dim cmBalance As New ADODB.command
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "ITEM", pItem)
If pstore <> "" Then aPrm = AddFlag(aPrm, "STORE", pstore)
If IsDate(pDate) Then aPrm = AddFlag(aPrm, "DATE", myFormat_sp(pDate))
If pId <> "" Then aPrm = AddFlag(aPrm, "ID", pId)

Set cmBalance = cmd("dbo.sp_balance", con, adStoredProc, aPrm)
cmBalance.Execute
fnBalance = Val(cmBalance.Parameters("@BALANCE") & "")
Set cmBalance = Nothing
End Function
Public Function rsBalance(pItem As String, Optional pstore As String = "", Optional pDate As String = "", Optional pId As String = "") As Long
Dim locTable As New ADODB.Recordset
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "ITEM", pItem)
If pstore <> "" Then aPrm = AddFlag(aPrm, "STORE", pstore)
If IsDate(pDate) Then aPrm = AddFlag(aPrm, "DATE", myFormat_sp(pDate))
If pId <> "" Then aPrm = AddFlag(aPrm, "ID", pId)

Set locTable = myRs("dbo.sp_balance_rs", , , adStoredProc, aPrm)
If Not locTable.EOF Then
    rsBalance = Val(locTable!balance & "")
End If
locTable.Close
Set locTable = Nothing
End Function
Public Function IsFormOpen(ByVal FormName As String) As Boolean
    Dim frm As Form
    For Each frm In Forms
        If UCase(frm.Name) = UCase(FormName) Then
            IsFormOpen = True
            Exit Function
        End If
    Next
    IsFormOpen = False
End Function

Public Function fnPhoneName(pPhone As String, con As ADODB.Connection) As String
Dim cmdPhone As New ADODB.command
Set cmdPhone = cmd("dbo.sp_cust_phone", con, adStoredProc, AddFlag(Empty, "phone", pPhone))
cmdPhone.Execute

fnPhoneName = cmdPhone.Parameters("@DESCA").value & ""
End Function
Public Function IsValidMobile(ByVal strNumber As String) As Boolean
    Dim i As Integer
    Dim prefix As String
    
    ' 1. ?????? ?? ????? (??? ?? ???? 11 ????? ??????)
    If Len(strNumber) <> 11 Then
        IsValidMobile = False
        Exit Function
    End If
    
    ' 2. ?????? ?? ?? ???? ???????? ????? ???
    For i = 1 To 11
        If Not IsNumeric(Mid(strNumber, i, 1)) Then
            IsValidMobile = False
            Exit Function
        End If
    Next i
    
    ' 3. ?????? ?? ??????? (??? ?? ???? ?? 01 ?? 0 ?? 1 ?? 2 ?? 5)
    prefix = Left(strNumber, 3)
    Select Case prefix
        Case "010", "011", "012", "015"
            IsValidMobile = True
        Case Else
            IsValidMobile = False
    End Select
End Function


Public Function UpdatePhones(pPhone As String, pName As String, pDate As String, con As ADODB.Connection) As Boolean
Dim cmdPhone As New ADODB.command
Set cmdPhone = cmd("dbo.sp_cust_phone", con, adStoredProc, AddFlag(Empty, "phone", pPhone))
cmdPhone.Execute

Dim aInsert As Variant
If Trim(pName) <> "" Then
    If Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Or Trim(pName) <> cmdPhone.Parameters("@DESCA") Then
        aInsert = AddFlag(aInsert, "DESCA", addstring(pName))
    End If
ElseIf (Not IsNull(cmdPhone.Parameters("@DESCA").value)) And Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Then
    aInsert = AddFlag(aInsert, "DESCA", addstring(cmdPhone.Parameters("@DESCA").value))
End If

'If IsNull(cmdPhone.Parameters("@F_DATE").Value) Then
'    aInsert = AddFlag(aInsert, "F_DATE", addDate(pDate))
'ElseIf Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Then
'    aInsert = AddFlag(aInsert, "F_DATE", addDate(cmdPhone.Parameters("@F_DATE").Value))
'End If

If Not IsNull(cmdPhone.Parameters("@F_DATE").value) Then
    If myFormat(cmdPhone.Parameters("@F_DATE").value) < myFormat(pDate) Then
        aInsert = AddFlag(aInsert, "F_DATE", addDate(pDate))
    ElseIf Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Then
        aInsert = AddFlag(aInsert, "F_DATE", addDate(cmdPhone.Parameters("@F_DATE").value))
    End If
ElseIf IsDate(pDate) Then
    aInsert = AddFlag(aInsert, "F_DATE", addDate(pDate))
End If

If Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Then
    aInsert = AddFlag(aInsert, "PHONE", addstring(pPhone))
End If

If Not IsEmpty(aInsert) Then
    con.BeginTrans
    If Val(cmdPhone.Parameters("@COUNT") & "") > 1 Then
        con.Execute "DELETE FROM SUBCUST WHERE PHONE = " & MyParn(pPhone)
    End If
    
    If Val(cmdPhone.Parameters("@COUNT") & "") = 1 Then
        con.Execute addUpdate(aInsert, "SUBCUST", "PHONE = " & MyParn(pPhone))
    Else
        con.Execute addInsert(aInsert, "SUBCUST")
    End If
    con.CommitTrans
End If
UpdatePhones = True
Exit Function
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Public Function UpdateDiscount(pDoc_no As String, con As ADODB.Connection, Optional pDiscount_add As Double, Optional pDiscount_total As Double) As Boolean
Dim cmDiscount As New ADODB.command
Set cmDiscount = cmd("dbo.sp_offer_discount", con, adStoredProc, AddFlag(Empty, "DOC_NO", pDoc_no))
cmDiscount.Execute
If Not IsNull(cmDiscount.Parameters("@OFFER_NO").value) Then
    If pDiscount_add = 0 Then
        con.Execute "UPDATE FILE6_20H " & _
                    "SET FILE6_20H.DISCOUNT_OFFER = " & cmDiscount.Parameters("@DISCOUNT").value & "," & _
                    "FILE6_20H.DISCOUNT = FILE6_20H.DISCOUNT_ADD + " & cmDiscount.Parameters("@DISCOUNT").value & "," & _
                    "FILE6_20H.IS_OFFER = " & IIf(cmDiscount.Parameters("@OFFER_NO").value > 0, "1", "0") & _
                    " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
    Else
        con.Execute "UPDATE FILE6_20H " & _
                    "SET FILE6_20H.DISCOUNT_OFFER = " & cmDiscount.Parameters("@DISCOUNT").value & "," & _
                    "FILE6_20H.DISCOUNT_ADD = " & pDiscount_add & "," & _
                    "FILE6_20H.DISCOUNT = " & pDiscount_add + cmDiscount.Parameters("@DISCOUNT").value & "," & _
                    "FILE6_20H.IS_OFFER = " & IIf(cmDiscount.Parameters("@OFFER_NO").value > 0, "1", "0") & _
                    " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
    End If
Else
    If pDiscount_add = 0 Then
        con.Execute "UPDATE FILE6_20H " & _
                    "SET FILE6_20H.DISCOUNT_OFFER = 0," & _
                    "FILE6_20H.DISCOUNT_Add = 0," & _
                    "FILE6_20H.DISCOUNT = " & pDiscount_total & "," & _
                    "FILE6_20H.IS_OFFER = 0" & _
                    " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
    Else
        con.Execute "UPDATE FILE6_20H " & _
                    "SET FILE6_20H.DISCOUNT_OFFER = 0," & _
                    "FILE6_20H.DISCOUNT_Add = 0," & _
                    "FILE6_20H.DISCOUNT = " & pDiscount_add & " ," & _
                    "FILE6_20H.IS_OFFER = 0" & _
                    " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
    End If
End If

'If Not IsNull(cmDiscount.Parameters("@OFFER_NO").value) Then
'    con.Execute "UPDATE FILE6_20H " & _
'                "SET FILE6_20H.DISCOUNT = FILE6_20H.DISCOUNT_ADD +  " & cmDiscount.Parameters("@DISCOUNT").value & "," & _
'                "FILE6_20H.DISCOUNT_OFFER = " & cmDiscount.Parameters("@DISCOUNT").value & "," & _
'                "FILE6_20H.IS_OFFER = " & IIf(cmDiscount.Parameters("@OFFER_NO").value > 0, "1", "0") & _
'                " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
'ElseIf bDelDiscount Then
'    con.Execute "UPDATE FILE6_20H " & _
'                "SET FILE6_20H.DISCOUNT = DISCOUNT_ADD," & _
'                "FILE6_20H.DISCOUNT_OFFER = 0," & _
'                "FILE6_20H.DISCOUNT_ADD = 0," & _
'                "FILE6_20H.IS_OFFER = 0" & _
'                " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
'End If
End Function
Public Function RoundToNearest5(ByVal value As Double) As Double
    RoundToNearest5 = Int((value / 5) + 0.5) * 5
End Function
Function addInsertUpdate(aInsert As Variant, pTable As String, pCondition As String, pFieldName As String)
Dim cInsert As New ChilkatStringBuilder, cUpdate As New ChilkatStringBuilder

cInsert.Append "INSERT INTO " & pTable & "("
For i = 0 To UBound(aInsert) Step 2
    cInsert.Append aInsert(i) & ","
Next
cInsert.Shorten 1
cInsert.AppendLine ")", 1

cInsert.Append "VALUES("
For i = 0 To UBound(aInsert) Step 2
    cInsert.Append aInsert(i + 1) & ","
Next
cInsert.Shorten 1
cInsert.Append ")"

cUpdate.AppendLine "UPDATE " & pTable & " SET ", 1
For i = 0 To UBound(aInsert) Step 2
    cUpdate.Append aInsert(i) & " = " & aInsert(i + 1) & ","
Next
cUpdate.Shorten 1
cUpdate.AppendLine "", 1
If pCondition <> "" Then
    cUpdate.Append "WHERE " & pCondition
End If

Dim sb As New ChilkatStringBuilder
sb.AppendLine "IF dbo.DOC_" & pTable & "(" & pFieldName & ") = 0", 1
'sb.append "BEGIN"
sb.AppendLine cInsert.GetAsString, 1
'sb.append "END"
sb.AppendLine "ELSE ", 1
'sb.append "BEGIN "
sb.Append cUpdate.GetAsString
'sb.Append "END"

addInsertUpdate = sb.GetAsString()
End Function
Private Function ValidPlus(pGrid As Object) As Boolean
For i = 1 To pGrid.Rows - 1
    If pGrid.ValueMatrix(i, 10) = 0 Then
    ElseIf pGrid.TextMatrix(i, 10) < 0 Then
        Exit Function
    End If
Next
ValidPlus = True
End Function
Private Function ValidMinus(pGrid As Object) As Boolean
For i = 1 To pGrid.Rows - 1
    If pGrid.TextMatrix(i, 10) >= 0 Then
        Exit Function
    End If
Next
ValidMinus = True
End Function
Public Function fn_order_Sates(sOrder_no As String, sType As String) As Variant
Dim locTable As New ADODB.Recordset
Dim aPrm As Variant
'aPrm = AddFlag(Empty, "DOC_NO", sOrder_no)
'aPrm = AddFlag(aPrm, "DOC_NO", sOrder_no)
End Function
' ???? ?? ????? ????? Microsoft Scripting Runtime ?? ????? References
' ?? ???? ??????? ??? Late Binding ??? ?? ???? ?? ????? ?????
Public Function CompareGrids(grid1 As Variant, grid2 As Variant, col1 As Integer, col2 As Integer, col11 As Integer, col12 As Integer) As Boolean
    Dim dict1 As Object
    Dim dict2 As Object
    Dim i As Long
    Dim code As String
    Dim qty As Double
    Dim key As Variant

    ' ????? ?????? ?????? ??????? ??? ???
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    dict1.CompareMode = 1 ' ?????? ?????? ??? ????? ????? ?????? (TextCompare)
    dict2.CompareMode = 1

    ' 1. ????? ????? ????? ?????? ?????? (Grid 1)
    '??? ?????? ?? ???? 1 ????? ?? ???????? Fixed Rows
    For i = grid1.FixedRows To grid1.Rows - 1
        code = Trim(grid1.TextMatrix(i, col1))
        If code <> "" Then
            qty = Val(grid1.TextMatrix(i, col2))
            If dict1.Exists(code) Then
                dict1(code) = dict1(code) + qty
            Else
                dict1.Add code, qty
            End If
        End If
    Next i

    ' 2. ????? ????? ????? ?????? ??????? (Grid 2)
    For i = grid2.FixedRows To grid2.Rows - 1
        code = Trim(grid2.TextMatrix(i, col11))
        If code <> "" Then
            qty = Val(grid2.TextMatrix(i, col12))
            If dict2.Exists(code) Then
                dict2(code) = dict2(code) + qty
            Else
                dict2.Add code, qty
            End If
        End If
    Next i

    ' 3. ???????? ??? ?????????
    ' ?????? ????? ?? ????? ??? ??????? ???????
    If dict1.Count <> dict2.Count Then
        CompareGrids = False
        Exit Function
    End If

    ' ?????? ?? ????? ??? ??????? ??? ???
    For Each key In dict1.Keys
        If Not dict2.Exists(key) Then
            CompareGrids = False ' ????? ??? ????? ?? ?????? ???????
            Exit Function
        ElseIf dict1(key) <> dict2(key) Then
            CompareGrids = False ' ??????? ??? ??????
            Exit Function
        End If
    Next key

    ' ??? ?????? ?? ??????????? ??? ???????? ?????????
    CompareGrids = True
End Function
Public Sub UnloadForms(ExcludeFormName As String)
    Dim i As Integer
    
    ' Loop backwards through the Forms collection
    For i = Forms.Count - 1 To 0 Step -1
        If LCase(Trim(Forms(i).Name)) = LCase(Trim(ExcludeFormName)) Then
            Unload Forms(i)
            Set Forms(i) = Nothing
        End If
    Next i
End Sub
Public Function DetectEncoding(ByVal sFilePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim bytes(3) As Byte
    Dim i As Integer
    
    ' › Õ «·„·› ·ﬁ—«¡… «·»«Ì «  «·√Ê·Ï ›ﬁÿ
    fileNum = FreeFile
    Open sFilePath For Binary Access Read As #fileNum
    
    ' ﬁ—«¡… √Ê· 4 »«Ì 
    For i = 0 To 3
        If Not EOF(fileNum) Then
            Get #fileNum, , bytes(i)
        End If
    Next i
    Close #fileNum
    
    ' «·›Õ’ »‰«¡ ⁄·Ï ⁄·«„…  — Ì» «·»«Ì «  (BOM)
    If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
        DetectEncoding = "UTF-8"
ElsePtr:
    ElseIf bytes(0) = &HFF And bytes(1) = &HFE Then
        DetectEncoding = "UTF-16LE" ' UTF-16 Little Endian
    ElseIf bytes(0) = &HFE And bytes(1) = &HFF Then
        DetectEncoding = "UTF-16BE" ' UTF-16 Big Endian
    Else
        ' ≈–« ·„ ÌÃœ ⁄·«„… „„Ì“…° Ì› —÷ √‰Â ANSI (√Ê ﬂÊœ «· ⁄—Ì» «·„Õ·Ì)
        DetectEncoding = "Windows-1256"
    End If
    
    Exit Function

ErrorHandler:
    DetectEncoding = "Windows-1256" ' «·«› —«÷Ì ⁄‰œ ÕœÊÀ Œÿ√
End Function
Public Function GetTextFileEncoding(ByVal FileName As String) As String
    Dim fileNum As Integer
    Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte
    
    ' Check if file exists
    If Dir(FileName) = "" Then
        GetTextFileEncoding = "File Not Found"
        Exit Function
    End If
    
    ' Open the file in binary mode to read raw bytes
    fileNum = FreeFile()
    Open FileName For Binary Access Read As #fileNum
    
    ' Get the first 4 bytes if the file is large enough
    If LOF(fileNum) >= 1 Then Get #fileNum, 1, b1
    If LOF(fileNum) >= 2 Then Get #fileNum, 2, b2
    If LOF(fileNum) >= 3 Then Get #fileNum, 3, b3
    If LOF(fileNum) >= 4 Then Get #fileNum, 4, b4
    
    Close #fileNum
    
    ' Evaluate the Byte Order Mark (BOM)
    If b1 = &HEF And b2 = &HBB And b3 = &HBF Then
        GetTextFileEncoding = "UTF-8"
    ElseIf b1 = &HFF And b2 = &HFE And b3 = &H0 And b4 = &H0 Then
        GetTextFileEncoding = "UTF-16LE / UTF-32LE"
    ElseIf b1 = &HFF And b2 = &HFE Then
        GetTextFileEncoding = "UTF-16LE (Unicode)"
    ElseIf b1 = &HFE And b2 = &HFF Then
        GetTextFileEncoding = "UTF-16BE (Big-Endian)"
    ElseIf b1 = &H0 And b2 = &H0 And b3 = &HFE And b4 = &HFF Then
        GetTextFileEncoding = "UTF-32BE"
    Else
        ' No BOM present: Defaults to system ANSI code page
        GetTextFileEncoding = "ANSI"
    End If
End Function
