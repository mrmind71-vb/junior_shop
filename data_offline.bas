Attribute VB_Name = "data_offline"
Public Function myRs(pString As String, Optional con As ADODB.Connection, Optional ByVal pConString As String, Optional pType As cmType = ADTEXT, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300, Optional nConTimeout As Integer = 3) As ADODB.Recordset
Dim bClose As Boolean

On Error GoTo myerror

If con Is Nothing Then
    bClose = True
    Set con = New ADODB.Connection
    If Not openConEr(con, pConString, nConTimeout) Then
        Exit Function
    End If
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
Set cmd.ActiveConnection = Nothing
If bClose Then closeCon con
Exit Function
myerror:
If Not myRs Is Nothing Then
    If myRs.State = adStateOpen Then rs.Close
    Set myRs = Nothing
End If

' ?????? ?? ????? ??????? ?????? ???????? ??? ??? ???????
If bClose Then
    If Not con Is Nothing Then
        If con.State = adStateOpen Then con.Close
        Set con = Nothing
    End If
End If
' 5. ????? ??? ????? ?????? ??? ?????? ?????????
Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
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

rsFunc = rsValue(cString, con, pConString, , aParam, nTimeout)
End Function
Public Function rsValue(pString As String, Optional con As ADODB.Connection, Optional ByVal pConString As String, Optional pType As cmType = ADTEXT, Optional aParam As Variant = Empty, Optional pDef As Variant = Null, Optional nTimeout As Integer = 300, Optional nConTimeout = 3) As Variant
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
End Function
Public Function rsValues(pString As String, Optional con As ADODB.Connection, Optional ByVal pConString As String, Optional pType As cmType = ADTEXT, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300, Optional nConTimeout As Integer = 3) As Variant
Dim loctable As ADODB.Recordset
Set loctable = myRs(pString, con, pConString, pType, aParam, nTimeout, nConTimeout)
If Not (loctable.BOF And loctable.EOF) Then
    For i = 0 To loctable.Fields.Count - 1
        rsValues = AddFlag(rsValues, LCase(loctable.Fields(i).Name), loctable.Fields(i).Value)
    Next
Else
    rsValues = Null
End If
loctable.Close
Set loctable = Nothing
End Function
Public Function myRsCmd(pString As String, Optional con As ADODB.Connection, Optional ByVal pConString As String, Optional pType As cmType = ADTEXT, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300, Optional nConTimeout As Integer = 3) As Boolean
Dim bClose As Boolean
If con Is Nothing Then
    Set con = New ADODB.Connection
    If Not openConEr(con, pConString, nConTimeout) Then
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

cmd.Execute
cmd.ActiveConnection = Nothing
If bClose Then closeCon con
myRsCmd = True
End Function
Public Function rsEx(pString As String, Optional ByVal pConString As String, Optional pType As cmType = ADTEXT, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300, Optional nConTimeout As Integer = 3) As Integer
Dim bClose As Boolean

On Error GoTo myerror

Dim con As New ADODB.Connection
If Not openConEr(con, pConString, nConTimeout) Then
    Exit Function
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

cmd.Execute rsEx
Set cmd.ActiveConnection = Nothing
If bClose Then closeCon con
Exit Function
myerror:
' ?????? ?? ????? ??????? ?????? ???????? ??? ??? ???????
If Not con Is Nothing Then
    If con.State = adStateOpen Then con.Close
    Set con = Nothing
End If
' 5. ????? ??? ????? ?????? ??? ?????? ?????????
Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

