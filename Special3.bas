Attribute VB_Name = "Special3"
Public pcName As String
Public obj As ChilkatGlobal
Public sOnlineStore As String
Public servername_vpn
Public sStoreOnline As String
Public sBranchOnline As String
Public Function Nz(Value As Variant, Optional ValueIfNull As Variant = 0) As Variant
    If IsNull(Value) Then
        Nz = ValueIfNull
    Else
        Nz = Value
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
Sub OnlineTypeLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
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

searchArray = Array(Generalarray, listarray, GrdArray)
'oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·«‰Ê«⁄"
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

Dim loctable As New ADODB.Recordset
Dim cString As String
Set loctable = cmd("SELECT dSales FROM DSALES" & _
            " WHERE BRANCH = " & MyParn(cBranch), con).Execute
If Not loctable.EOF Then
    fnDateSales = myFormat(loctable!dSales)
End If
loctable.Close
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
Dim loctable As New ADODB.Recordset
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "ITEM", pItem)
If pstore <> "" Then aPrm = AddFlag(aPrm, "STORE", pstore)
If IsDate(pDate) Then aPrm = AddFlag(aPrm, "DATE", myFormat_sp(pDate))
If pId <> "" Then aPrm = AddFlag(aPrm, "ID", pId)

Set loctable = myRs("dbo.sp_balance_rs", , , adStoredProc, aPrm)
If Not loctable.EOF Then
    rsBalance = Val(loctable!balance & "")
End If
loctable.Close
Set loctable = Nothing
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

fnPhoneName = cmdPhone.Parameters("@DESCA").Value & ""
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
ElseIf (Not IsNull(cmdPhone.Parameters("@DESCA").Value)) And Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Then
    aInsert = AddFlag(aInsert, "DESCA", addstring(cmdPhone.Parameters("@DESCA").Value))
End If

'If IsNull(cmdPhone.Parameters("@F_DATE").Value) Then
'    aInsert = AddFlag(aInsert, "F_DATE", addDate(pDate))
'ElseIf Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Then
'    aInsert = AddFlag(aInsert, "F_DATE", addDate(cmdPhone.Parameters("@F_DATE").Value))
'End If

If Not IsNull(cmdPhone.Parameters("@F_DATE").Value) Then
    If myFormat(cmdPhone.Parameters("@F_DATE").Value) < myFormat(pDate) Then
        aInsert = AddFlag(aInsert, "F_DATE", addDate(pDate))
    ElseIf Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Then
        aInsert = AddFlag(aInsert, "F_DATE", addDate(cmdPhone.Parameters("@F_DATE").Value))
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
If Not IsNull(cmDiscount.Parameters("@OFFER_NO").Value) Then
    If pDiscount_add = 0 Then
        con.Execute "UPDATE FILE6_20H " & _
                    "SET FILE6_20H.DISCOUNT_OFFER = " & cmDiscount.Parameters("@DISCOUNT").Value & "," & _
                    "FILE6_20H.DISCOUNT = FILE6_20H.DISCOUNT_ADD + " & cmDiscount.Parameters("@DISCOUNT").Value & "," & _
                    "FILE6_20H.IS_OFFER = " & IIf(cmDiscount.Parameters("@OFFER_NO").Value > 0, "1", "0") & _
                    " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
    Else
        con.Execute "UPDATE FILE6_20H " & _
                    "SET FILE6_20H.DISCOUNT_OFFER = " & cmDiscount.Parameters("@DISCOUNT").Value & "," & _
                    "FILE6_20H.DISCOUNT_ADD = " & pDiscount_add & "," & _
                    "FILE6_20H.DISCOUNT = " & pDiscount_add + cmDiscount.Parameters("@DISCOUNT").Value & "," & _
                    "FILE6_20H.IS_OFFER = " & IIf(cmDiscount.Parameters("@OFFER_NO").Value > 0, "1", "0") & _
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
Public Function RoundToNearest5(ByVal Value As Double) As Double
    RoundToNearest5 = Int((Value / 5) + 0.5) * 5
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
Public Function fn_order_Sates(sOrder_No As String, sType As String) As Variant
Dim loctable As New ADODB.Recordset
Dim aPrm As Variant
'aPrm = AddFlag(Empty, "DOC_NO", sOrder_no)
'aPrm = AddFlag(aPrm, "DOC_NO", sOrder_no)
End Function



