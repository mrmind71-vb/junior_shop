Attribute VB_Name = "data_offline"
Public Function myRs(pString As String, Optional Con As adodb.Connection, Optional ByVal pConString As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300, Optional nConTimeout As Integer = 3) As adodb.Recordset
Dim bClose As Boolean

On Error GoTo myError

If Con Is Nothing Then
    bClose = True
    Set Con = New adodb.Connection
    If pConString = "" Then
        If Not openConEr(Con, , nConTimeout) Then
            Exit Function
        End If
    ElseIf Not openConEr(Con, pConString, nConTimeout) Then
        Exit Function
    End If
End If

Dim cmd As New adodb.command
cmd.CommandTimeout = nTimeout
cmd.ActiveConnection = Con
cmd.CommandType = pType
cmd.CommandText = pString

If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If

Set myRs = New adodb.Recordset
myRs.CursorLocation = adUseClient
Set myRs = cmd.Execute
Set myRs.ActiveConnection = Nothing
Set cmd.ActiveConnection = Nothing
If bClose Then closeCon Con
Exit Function
myError:
If Not myRs Is Nothing Then
    If myRs.State = adStateOpen Then Rs.Close
    Set myRs = Nothing
End If

' ?????? ?? ????? ??????? ?????? ???????? ??? ??? ???????
If bClose Then
    If Not Con Is Nothing Then
        If Con.State = adStateOpen Then Con.Close
        Set Con = Nothing
    End If
End If
' 5. ????? ??? ????? ?????? ??? ?????? ?????????
err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Function
Public Function rsFunc(pFunction As String, Optional Con As adodb.Connection, Optional pParam1 As String = "", Optional pParam2 As String = "", Optional pParam3 As String = "", Optional pParam4 As String = "", Optional pParam5 As String = "", Optional pParam6 As String = "", Optional pParam7 As String = "", Optional pParam8 As String = "", Optional pParam9 As String = "", Optional pParam10 As String = "") As Variant
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
rsFunc = rsValue(cString, Con, pConString, pType, aParam, nTimeout)
End Function
Public Function rsValue(pString As String, Optional Con As adodb.Connection, Optional ByVal pConString As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional pDef As Variant = Null, Optional nTimeout As Integer = 300, Optional nConTimeout = 3) As Variant
Dim loctable As adodb.Recordset
Set loctable = myRs(pString, Con, pConString, pType, aParam, nTimeout)
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
Public Function rsValues(pString As String, Optional Con As adodb.Connection, Optional ByVal pConString As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300, Optional nConTimeout As Integer = 3) As Variant
Dim loctable As adodb.Recordset
Set loctable = myRs(pString, Con, pConString, pType, aParam, nTimeout, nConTimeout)
If Not (loctable.BOF And loctable.EOF) Then
    For i = 0 To loctable.Fields.Count - 1
        rsValues = AddFlag(rsValues, LCase(loctable.Fields(i).Name), loctable.Fields(i).Value)
    Next
End If
loctable.Close
Set loctable = Nothing
End Function
Public Function myRsCmd(pString As String, Optional Con As adodb.Connection, Optional ByVal pConString As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeout As Integer = 300, Optional nConTimeout As Integer = 3) As Boolean
Dim bClose As Boolean
If Con Is Nothing Then
    Set Con = New adodb.Connection
    If pConString = "" Then
        If Not openConEr(Con, , nConTimeout) Then
            Exit Function
        End If
    ElseIf Not openConEr(Con, pConString, nConTimeout) Then
        Exit Function
    End If
    bClose = True
End If

Dim cmd As New adodb.command
cmd.CommandTimeout = nTimeout
cmd.ActiveConnection = Con
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
If bClose Then closeCon Con
myRsCmd = True
End Function

