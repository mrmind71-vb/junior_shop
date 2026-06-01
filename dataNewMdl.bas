Attribute VB_Name = "dataNewMdl"
Public Enum cmType
adStoredProc = adCmdStoredProc
adText = adCmdText
adTable = adCmdTable
End Enum

Public Enum tbMode
    tbFind = 1
    tbPrevious = 2
    tblast = 3
    tbNext = 4
    tbFirst = 5
End Enum

Public Enum tableMode
Found = 1
NotFound = 2
NoRecords = 3
NotMatch = 3
End Enum
Public Function mycmd(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300) As ADODB.Recordset
Dim loctable As New ADODB.Recordset
Dim cmd As New ADODB.command
cmd.CommandTimeout = nTimeout
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
Set mycmd = cmd.Execute
End Function
Public Function myRs(pString As String, Optional con As ADODB.Connection, Optional ByVal pConString As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300, Optional nConTimeout As Integer = 3) As ADODB.Recordset
Dim bClose As Boolean
If con Is Nothing Then
    Set con = New ADODB.Connection
    If pConString = "" Then
        If openCon(con, , nConTimeout) <> "ok" Then
            Exit Function
        End If
    ElseIf openCon(con, pConString, nConTimeout) <> "ok" Then
        Exit Function
    End If
    bClose = True
End If

Dim cmd As New ADODB.command
cmd.CommandTimeout = nTimeout
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If

Set myRs = New ADODB.Recordset
myRs.CursorLocation = adUseClient
Set myRs = cmd.Execute
Set myRs.ActiveConnection = Nothing
If bClose Then closeCon con
End Function
Public Function rsFunc(pFunction As String, Optional con As ADODB.Connection, Optional pParam1 As String = "", Optional pParam2 As String = "", Optional pParam3 As String = "", Optional pParam4 As String = "", Optional pParam5 As String = "", Optional pParam6 As String = "", Optional pParam7 As String = "", Optional pParam8 As String = "", Optional pParam9 As String = "", Optional pParam10 As String = "") As Variant
Dim cPrm As String, cString As String
If pParam1 <> "" Then cPrm = pParam1
If pParam2 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam2
If pParam3 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam3
If pParam4 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam4
If pParam5 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam5
If pParam6 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam6
If pParam7 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam7
If pParam8 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam8
If pParam9 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam9
If pParam10 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam10
cString = "Select " & pFunction & "(" & cPrm & ") as [value]"
rsFunc = rsValue(cString, con)
End Function
Public Function rsValue(pString As String, Optional con As ADODB.Connection, Optional ByVal pConString As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional pDef As Variant = Null, Optional nTimeout As Integer = 300, Optional nConTimeout = 3) As Variant
On Error GoTo myerror:
Dim loctable As ADODB.Recordset
Set loctable = myRs(pString, con, pConString, pType, aParam, nTimeout)
If Not loctable.EOF Then
    rsValue = loctable.Fields(0).Value
ElseIf Not IsEmpty(pDef) Then
    rsValue = pDef
Else
    rsValue = Null
End If
If loctable.State = adStateOpen Then loctable.Close
Finally:
Set loctable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
rsValue = Null
GoTo Finally
End Function
Public Function rsValues(pString As String, Optional con As ADODB.Connection, Optional ByVal pConString As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300, Optional nConTimeout As Integer = 3) As Variant
Dim loctable As ADODB.Recordset
Set loctable = myRs(pString, con, pConString, pType, aParam, nTimeout, nConTimeout)
If Not (loctable.BOF And loctable.EOF) Then
    For i = 0 To loctable.Fields.Count - 1
        rsValues = AddFlag(rsValues, LCase(loctable.Fields(i).Name), loctable.Fields(i).Value)
    Next
End If
loctable.Close
Set loctable = Nothing
End Function
Public Function myFieldValue(pString As String, pField As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional pDef As Variant = Empty, Optional nTimeout As Integer = 100) As Variant
On Error GoTo myerror:
Dim cmd As New ADODB.command
Dim loctable As ADODB.Recordset
cmd.CommandTimeout = nTimeout
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
Set loctable = cmd.Execute
If Not loctable.EOF Then
    myFieldValue = loctable.Fields(pField).Value
ElseIf Not IsEmpty(pDef) Then
    myFieldValue = pDef
End If
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
myFieldValue = pDef
End Function
Public Function myCmdRecord(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300) As ADODB.Record
Dim loctable As New ADODB.Recordset
Dim cmd As New ADODB.command
cmd.CommandTimeout = nTimeout
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If

Set myCmdRecord = cmd.Execute
End Function

Public Function myPrcString(pProc As String, Optional aParam As Variant = Empty) As String
Dim i As Long

If Not IsEmpty(aParam) Then
    For i = 0 To UBound(aParam) Step 2
        myPrcString = myPrcString & IIf(myPrcString = "", " ", ",") & "@" & aParam(i) & " = " & aParam(i + 1)
    Next
End If
myPrcString = "EXEC " & pProc & myPrcString
End Function
Public Function mycmdEx(pString As String, Optional con As ADODB.Connection, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 600) As ADODB.command
Dim loctable As New ADODB.Recordset
Dim cmd As New ADODB.command
cmd.CommandTimeout = nTimeout
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
       cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
cmd.Execute
Set mycmdEx = cmd
End Function
Public Function myCommand(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300) As ADODB.command
Dim loctable As New ADODB.Recordset
Dim cmd As New ADODB.command
cmd.CommandTimeout = nTimeout
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
cmd.Execute
Set myCommand = cmd
End Function
Public Function cmd(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 1000) As ADODB.command
Set cmd = New ADODB.command
cmd.CommandTimeout = nTimeout
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
End Function
Public Function myField(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional pDef As Variant = Empty, Optional nTimeout As Integer = 100) As Variant
On Error GoTo myerror:
Dim loctable As ADODB.Recordset
Set loctable = cmd(pString, con, pType).Execute
If Not loctable.EOF Then
    myField = loctable.Fields(0).Value
ElseIf Not IsEmpty(pDef) Then
    myField = pDef
End If
loctable.Close
Finally:
Set loctable = Nothing
Exit Function
myerror:
MsgBox Err.Description
'Err.Clear
myField = pDef
GoTo Finally
End Function
Public Function myFields(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 100) As Variant
Dim cmd As New ADODB.command
Dim loctable As ADODB.Recordset
cmd.CommandTimeout = nTimeout
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
Set loctable = cmd.Execute
If Not (loctable.BOF And loctable.EOF) Then
    For i = 0 To loctable.Fields.Count - 1
        myFields = AddFlag(myFields, LCase(loctable.Fields(i).Name), loctable.Fields(i).Value)
    Next
End If
loctable.Close
Set loctable = Nothing
End Function
Public Function openConHr(ByRef pCon As ADODB.Connection, Optional ByVal pString As String = "", Optional ByVal lMsg As Boolean = True) As String
On Error GoTo myerror
Dim cString As String
If pString = "" Then cString = LoadConStringHr Else cString = pString
If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.Open cString
openConHr = "ok"
Exit Function
myerror:
openConHr = Err.Description
Err.Clear
End Function
Public Function LoadConStringHr()
Dim cServerName As String, cUserId As String, cPassword As String, sCatalogHr As String
'If lServerOnLine Or lServerOnLineShop Then

    If Not vpn Then
        cServerName = "junior-sql.database.windows.net"
        cUserId = decrypt("062F8C407C77", "dr")
        cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
        sCatalogHr = "JUNIOR_SHOP"
    Else
        cServerName = servername_vpn
        cUserId = userid_vpn
        cPassword = password_vpn
        sCatalogHr = "JUNIOR_SHOP"
    End If
'Else
'    cServerName = RetSetting("server", App.Path & "\hr\confHr.txt")
'    cUserId = decrypt(RetSetting("userId", App.Path & "\hr\confHr.txt"), "dr")
'    cPassword = decrypt(RetSetting("Password", App.Path & "\hr\confHr.txt"), "dr")
'    sCatalogHr = RetSetting("Catalog", App.Path & "\conf.txt")
'End If
If cServerName = "" Then cServerName = "."
LoadConStringHr = "provider=SQLOLEDB;data source=" & MyParn(cServerName) & ";initial " _
            & "catalog =" & sCatalogHr & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Function myFormat_p(sDate As Variant, Optional bLong As Boolean = False) As String
If Not IsDate(sDate) Then
    myFormat_p = ""
Else
    myFormat_p = Format(sDate & "", "YYYY/M/D") & IIf(bLong, Format(sDate & "", " HH:NN"), "")
End If
End Function
Function TransCount(con As ADODB.Connection) As Integer
Dim loctable As New ADODB.Recordset
Set loctable = mycmd("select @@TRANCOUNT as myCount", con)
On Error GoTo myerror
If Not loctable.EOF Then
    TransCount = loctable!myCount
End If
Finally:
loctable.Close
Set loctable = Nothing
Exit Function
On Error GoTo myerror:
myerror:
TransCount = -1
Err.Clear
GoTo Finally
End Function
Public Function clearTemp(con As ADODB.Connection)
con.Execute " If(OBJECT_ID('tempdb..#temp') Is Not Null)" & _
           " Begin" & _
           " Drop Table #Temp" & _
           " End"
End Function
Function createRs(pTable As ADODB.Recordset) As ADODB.Recordset
Set createRs = New ADODB.Recordset
createRs.ActiveConnection = Nothing
createRs.CursorLocation = adUseClient ' Or adUseServer if needed
createRs.LockType = adLockBatchOptimistic ' Or other suitable lock type
For i = 0 To pTable.Fields.Count - 1
    If pTable.Fields(i).Type = 131 Then
        createRs.Fields.Append pTable.Fields(i).Name, adDecimal, pTable.Fields(i).NumericScale, adFldIsNullable
        'createRs.Fields(i).NumericScale = pTable.Fields(i).NumericScale
    Else
        createRs.Fields.Append pTable.Fields(i).Name, pTable.Fields(i).Type, pTable.Fields(i).DefinedSize, adFldIsNullable + adFldNegativeScale
    End If
Next
createRs.Open
Do Until pTable.EOF
    createRs.AddNew ' Add a new record
    For i = 0 To pTable.Fields.Count - 1
        If createRs.Fields(i).Type = adBoolean Then
            createRs.Fields(i).Value = IIf(pTable.Fields(i).Value, 1, 0)
       Else
           createRs.Fields(i).Value = pTable.Fields(i).Value
        End If
    Next
    createRs.Update
    pTable.MoveNext
Loop
End Function
Public Function rsField(rsTable As ADODB.Recordset, pValue As Variant, pReturn As String, Optional pField As String = "code") As Variant
If Trim(pValue & "") = "" Then
    rsField = Null
    Exit Function
End If

rsTable.Find pField & " = " & (pValue & ""), , adSearchForward, adBookmarkFirst
If Not rsTable.EOF Then
     rsField = rsTable(pReturn)
End If
End Function
