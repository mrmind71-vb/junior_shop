VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ADO Connection and Command objects
Private cn As ADODB.Connection
Private sqlStatements() As String
Public ErrNumber As Integer
Public ErrDesc As String
Public ErrSource As String
Public cError As ClsError
' Initialize the connection
Private Sub Class_Initialize()
clearSql
End Sub
' Clear previous errors
Public Function myRs(strSql, Optional pType As Long = 1, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 300, Optional nConTimeout As Integer = 3) As ADODB.Recordset
    On Error GoTo myerror
            
    Set myRs = New ADODB.Recordset
    Dim cmd As New ADODB.command
    cmd.CommandTimeout = nTimeOut
    cmd.ActiveConnection = cn
    cmd.CommandType = pType
    cmd.CommandText = strSql ' Ensure strSql is declared globally or in your module

    If Not IsEmpty(aParam) Then
        Dim i As Long
        For i = 0 To UBound(aParam) Step 2
            cmd.Parameters("@" & aParam(i)).value = aParam(i + 1)
        Next
    End If

    myRs.CursorLocation = 3 ' adUseClient
    Set myRs = cmd.Execute

    Set myRs.ActiveConnection = Nothing
    Set cmd.ActiveConnection = Nothing
    Exit Function

myerror:
    If Not myRs Is Nothing Then
        If myRs.State = 1 Then myRs.Close ' Fixed "rs.Close" typo
        Set myRs = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function OpenCon(Optional ByVal ConnectionString As String = "", _
                               Optional ByVal TimeOut As Long = 15, _
                               Optional ByVal QueryTimeOut As Long = 180) As Boolean
    
    
    On Error GoTo ErrorHandler
    clearSql

    ' ????? ??????? ????? ??????? ??? ?? ??? ?????? ?? ???
    If cn Is Nothing Then
        Set cn = New ADODB.Connection
    Else
        If cn.State = adStateOpen Then cn.Close
    End If

    
    ' ????? ??????? ????? ??? ??? ??????? ????? ???????
    
    ' ??? ??????? ???????
    cn.CursorLocation = adUseClient
    cn.ConnectionTimeout = TimeOut
    cn.CommandTimeout = QueryTimeOut
    
    ' ????? ?? ??????? ????????? ?? ??????
    Dim finalConnectionString As String
    If ConnectionString = "" Then
        finalConnectionString = strCon
    Else
        finalConnectionString = ConnectionString
    End If
    
    ' ?????? ??? ???????
    cn.Open finalConnectionString
    
    ' ????? ???? ???? ???????
    OpenCon = True
    Exit Function

ErrorHandler:
    ' ?????? ????? ?????? ??????? ?????
    OpenCon = False
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
    
    ' ??? ?????? ????? ????????
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ «·« ’«· »ﬁ«⁄œ… «·»Ì«‰« :" & vbCrLf & _
           "—ﬁ„ «·Œÿ√ : " & Err.Number & vbCrLf & _
           "«·Ê’› : " & Err.Description, vbCritical, "Œÿ√ ›Ï «·« ’«·"
           
    Err.Clear
End Function
Public Function closeCon(Optional bIgError As Boolean = True) As Boolean
On Error Resume Next ' ??? ??? ????? ??????
    
    Dim lngErrNum As Long
    Dim strErrDesc As String
    Dim strErrSource As String
    
    ' ??? ?????? ????? ?????
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    strErrSource = Err.Source
    
    ' ????? ???? ??????? ??????? ?????
    If Not cn Is Nothing Then
        ' ?????? ??????? ?? ???????? ??? ??? ??? ??????? ???????
        If cn.State = adStateOpen Then
            cn.RollbackTrans
            cn.Close
        End If
        Set cn = Nothing
    End If
    
    ' ????? ????? ????? ?? ????? ??????
    If lngErrNum <> 0 Then
        ' ????? ???? ????? ?????? ?? ????? ??????? ??????
        If Not bIgError Then
            Error.Raise lngErrNum, strErrSource, strErrDesc
        End If
    Else
        closeCon = True
    End If
    
    ' ????? ????? ??????? ?????? ????
    Err.Clear
    On Error GoTo 0
End Function
Private Sub Class_Terminate()
    closeCon
    Set cError = Nothing
    ' 1. Close any open files or database connections
    ' Close #1
    
    ' 2. Release COM objects or API handles (if applicable)
    ' Set myExternalApp = Nothing
    
    ' 3. Optional: Add a log or message (useful for debugging)
    ' MsgBox "Class has been terminated"
End Sub
Public Function ExecuteTransaction() As Boolean
    On Error GoTo ErrorHandler
    Dim IsTransactionStarted As Boolean
    
    ' «› —«÷ «·‰Ã«Õ ›Ï «·»œ«Ì…
    ExecuteTransaction = True
    
    ' ?????? ?? ?? ??????? ?????? ???????? ?????
    If cn.State = adStateClosed Then
        MsgBox "« ’«· ﬁ«⁄œ… «·»Ì«‰«  „€·ﬁ", vbCritical, "Œÿ√"
        ExecuteTransaction = False
        Exit Function
    End If

    ' ??? ???????? (Transaction)
    cn.BeginTrans

    Dim i As Long
    ' ?????? ??? ???? ????? ?????? ?? ???????? ????????
    For i = 1 To UBound(sqlStatements)
        ' ?????? ?? ?? ?????? ???? ?????
        If Trim$(sqlStatements(i)) <> "" Then
            cn.Execute sqlStatements(i), , adCmdText + adExecuteNoRecords
             IsTransactionStarted = True
        End If
    Next i

    ' ????? ????????? ?? ??? ???? ???? ?????
    cn.CommitTrans
    Exit Function

ErrorHandler:
     
    MsgBox "›‘·  «·⁄„·Ì…  „ «· —«Ã⁄ ⁄‰ ﬂ«›… «· €ÌÌ—« " & vbCrLf & _
           "«·Œÿ√ : " & Err.Description & vbCrLf & _
           "—ﬁ„ «·Ã„·… «·„”»»… ··Œÿ«: " & CStr(i), vbCritical, "Œÿ√ ›Ï ﬁ«⁄œ… «·»Ì«‰« "
     
    On Error Resume Next
    'Error.Raise Err.Number, Err.Source, Err.Description
    If IsTransactionStarted Then
        cn.RollbackTrans
    End If
           
    ' ?? ??? ???? ?? ???? ??? ????? ???? ???????? ???????? ????
    
    ' ??? ????? ????? ?????? ?????? ???? ????? ?? ???????
    ExecuteTransaction = False
    On Error GoTo 0
End Function
Public Function Execute(strSql As String) As Boolean
    On Error GoTo ErrorHandler

    ' «› —«÷ «·‰Ã«Õ ›Ï «·»œ«Ì…
    
    ' ?????? ?? ?? ??????? ?????? ???????? ?????
    If cn.State = adStateClosed Then
        MsgBox "« ’«· ﬁ«⁄œ… «·»Ì«‰«  „€·ﬁ", vbCritical, "Œÿ√"
        Exit Function
    End If


    ' ?????? ??? ???? ????? ?????? ?? ???????? ????????
    cn.Execute strSql, , adCmdText + adExecuteNoRecords


    Execute = True

    ' ????? ????????? ?? ??? ???? ???? ?????
    Exit Function

ErrorHandler:
    MsgBox "›‘·  «·⁄„·Ì… " & vbCrLf & _
           "«·Œÿ√ : " & Err.Description _
           , vbCritical, "Œÿ√ ›Ï ﬁ«⁄œ… «·»Ì«‰« "
    
    On Error Resume Next
    'Error.Raise Err.Number, Err.Source, Err.Description
    
           
    ' ?? ??? ???? ?? ???? ??? ????? ???? ???????? ???????? ????
    
    ' ??? ????? ????? ?????? ?????? ???? ????? ?? ???????
    Execute = False
    On Error GoTo 0
End Function
Public Function addSql(sValue As String)
ReDim Preserve sqlStatements(UBound(sqlStatements) + 1)
sqlStatements(UBound(sqlStatements)) = sValue
End Function
Public Function clearSql()
ReDim sqlStatements(0)
End Function
Function IncData(sTable As String, sField As String, Optional pWhere As String = "", Optional bNum As Boolean = False) As String
Dim locTable As New ADODB.Recordset
Dim cString As String
cString = "Select Max(" & sField & ") as Maxof " & _
       " FROM " & sTable
If pWhere <> "" Then cString = cString & " where " & pWhere
Set locTable = myRs(cString)
If Not (locTable.EOF) Then
    If IsNull(locTable!maxOf) Then
        IncData = "1"
    ElseIf bNum Then
        IncData = Val(locTable!maxOf & "") + 1
    Else
        IncData = IncrementString(locTable!maxOf & "")
    End If
End If
locTable.Close
Set locTable = Nothing
End Function
Public Function rsField(pString As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional pDef As Variant = Null, Optional nTimeOut As Integer = 300) As Variant
Dim locTable As ADODB.Recordset
Set locTable = Me.myRs(pString, pType, aParam, nTimeOut)
If Not locTable.EOF Then
    rsField = locTable.Fields(0).value
ElseIf Not IsEmpty(pDef) Then
    rsField = pDef
Else
    rsField = Null
End If
If locTable.State = adStateOpen Then locTable.Close
Finally:
Set locTable = Nothing
End Function
Public Function rsValues(pString As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 300) As Variant
Dim locTable As ADODB.Recordset
Set locTable = myRs(pString, pType, aParam, nTimeOut)
Dim i As Long
If Not (locTable.BOF And locTable.EOF) Then
    For i = 0 To locTable.Fields.Count - 1
        rsValues = AddFlag(rsValues, LCase(locTable.Fields(i).Name), locTable.Fields(i).value)
    Next
Else
    rsValues = Null
End If
locTable.Close
Set locTable = Nothing
End Function
Public Function cmd(pString As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 1000) As ADODB.command
Set cmd = New ADODB.command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = cn
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).value = aParam(i + 1)
    Next
End If
End Function
