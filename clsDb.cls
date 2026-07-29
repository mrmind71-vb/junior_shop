VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Private member variables to manage connection state
Private conn As ADODB.Connection
Private mConnectionString As String
Private mTimeOut As Long
Private mQueryTimeOut As Long
Private sqlStatements() As String
Private mForm As Form
' Executed automatically when Class is initialized
Private Sub Class_Initialize()
    Set conn = New ADODB.Connection
    mConnectionString = strCon
    mTimeOut = 10
    mQueryTimeOut = 180
End Sub
' Property to safely pass connection string into the class
Public Property Let ConnectionString(ByVal sValue As String)
    mConnectionString = sValue
End Property
Public Property Let TimeOut(ByVal sValue As String)
    mTimeOut = sValue
End Property
Public Property Let QueryTimeOut(ByVal sValue As String)
    mQueryTimeOut = sValue
End Property
Public Property Let myForm(ByRef sValue As Form)
    Set mForm = sValue
End Property

' Open connection strategy
Public Function OpenConnection() As Boolean
    On Error GoTo ErrHandler
    
    If conn.State = adStateClosed Then
        conn.ConnectionString = mConnectionString
        conn.CursorLocation = adUseClient ' Essential for record counts & sorting
        conn.ConnectionTimeout = mTimeOut
        conn.CommandTimeout = mQueryTimeOut
        conn.Open
    End If
    OpenConnection = True
    Exit Function

ErrHandler:
    OpenConnection = False
    MsgBox "Database connection failed: " & Err.Description, vbCritical, "Database Error"
End Function
' Close connection strategy
Public Sub CloseConnection()
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set conn = Nothing
End Sub
Public Function myRs(strSql, Optional pType As Long = 1, Optional aParam As Variant = Empty) As ADODB.Recordset
    On Error GoTo ErrHandler
    Dim rs As ADODB.Recordset
    
    If Not OpenConnection() Then Exit Function
                
    Set myRs = New ADODB.Recordset
    Dim cmd As New ADODB.command
    cmd.CommandTimeout = mTimeOut
    cmd.ActiveConnection = conn
    cmd.CommandType = pType
    cmd.CommandText = strSql ' Ensure strSql is declared globally or in your module

    If Not IsEmpty(aParam) Then
        Dim i As Long
        For i = 0 To UBound(aParam) Step 2
            cmd.Parameters("@" & aParam(i)).value = aParam(i + 1)
        Next
    End If

    myRs.CursorLocation = adUseClient
    Set myRs = cmd.Execute

    Set myRs.ActiveConnection = Nothing
    
CleanUp:
    On Error Resume Next
    If Not cmd Is Nothing Then Set cmd.ActiveConnection = Nothing: Set cmd = Nothing
    Exit Function
ErrHandler:
    Set myRs = Nothing
    MsgBox "Œÿ« ›Ï «· ⁄œÌ· : " & Err.Description, vbCritical, "SQL Error"
    Resume CleanUp
End Function
' Best Practice Method: Executes INSERTS, UPDATES, or DELETES safely
Public Function Execute(ByVal Sql As String) As Long
    On Error GoTo ErrHandler
    Dim recordsAffected As Long
    
    If Not OpenConnection() Then Exit Function
    
    conn.Execute Sql, recordsAffected, adCmdText
    Execute = recordsAffected ' Returns number of rows altered
    Exit Function

ErrHandler:
    Execute = -1
    MsgBox "Execution error: " & Err.Description, vbCritical, "SQL Error"
End Function
' Executed automatically when Class is destroyed (Set Nothing)
Private Sub Class_Terminate()
    Call CloseConnection
End Sub
Public Property Let Sql(ByVal sValue As String)
    
    If IsArrayEmpty(sqlStatements) Then
        ReDim sqlStatements(0)
    Else
        ' ????? ??? ???????? ?????? 1 ?? ?????? ??? ???????? ???????
        ReDim Preserve sqlStatements(UBound(sqlStatements) + 1)
    End If
    
    ' ????? ?????? ??????? ?? ??? ????
    sqlStatements(UBound(sqlStatements)) = sValue
End Property
' 2. ???? ?????? ???? ?? ??? ???? ???????? ????? ???? ?????
Private Function IsArrayEmpty(ByRef arr() As String) As Boolean
    On Error Resume Next
    Dim upperBounds As Long
    
    ' ?????? ??? ???? ?????? ????????
    upperBounds = UBound(arr)
    
    ' ??? ??? ??? ??? 9? ???? ???? ?? ???????? ??? ????? ???
    If Err.Number = 9 Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
    On Error GoTo 0
End Function
Public Function StrList(strSql, Optional pType As cmType = adText, Optional aParam As Variant = Empty) As String
Dim locTable As ADODB.Recordset
Set locTable = myRs(strSql, pType, aParam)

If locTable Is Nothing Then Exit Function

Do Until locTable.EOF
    StrList = StrList & "|#" & locTable.Fields(0) & ";" & locTable.Fields(1)
    locTable.MoveNext
Loop
locTable.Close
Set locTable = Nothing
End Function
Public Sub BeginTran()
    If OpenConnection() Then conn.BeginTrans
End Sub
' ???? ?????? ????????
Public Sub CommitTrans()
    If Not conn Is Nothing Then conn.CommitTrans
End Sub
Public Sub RollbackTran()
    If Not conn Is Nothing Then conn.RollbackTrans
End Sub
Public Function rsValue(strSql As String, ByRef value As Variant, Optional pType As cmType = adText, Optional aParam As Variant = Empty) As Boolean
Dim locTable As ADODB.Recordset
Set locTable = myRs(strSql, pType, aParam)
If locTable Is Nothing Then Exit Function

If Not locTable.EOF Then
    value = locTable.Fields(0).value
Else
    value = Empty
End If
Set locTable = Nothing
rsValue = True
End Function
Public Function rsField(strSql As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty) As Variant
Dim locTable As ADODB.Recordset
Set locTable = myRs(strSql, pType, aParam)
If locTable Is Nothing Then Exit Function

If Not locTable.EOF Then
    rsField = locTable.Fields(0).value
Else
    rsField = Empty
End If
Set locTable = Nothing
End Function
Public Function rsValues(strSql As String, ByRef aValues As Variant, Optional pType As cmType = adText, Optional aParam As Variant = Empty) As Boolean
Dim locTable As ADODB.Recordset
Set locTable = myRs(strSql, pType, aParam)
If locTable Is Nothing Then Exit Function

Dim i As Long
If Not (locTable.BOF And locTable.EOF) Then
    For i = 0 To locTable.Fields.Count - 1
        aValues = AddFlag(aValues, LCase(locTable.Fields(i).Name), locTable.Fields(i).value)
    Next
Else
    aValues = Empty
End If
locTable.Close
Set locTable = Nothing
rsValues = True
End Function
Public Function cmd(strSql As String, Optional pType As cmType = adText, Optional aParam As Variant = Empty) As ADODB.command
Set cmd = New ADODB.command

If Not OpenConnection() Then Exit Function

cmd.CommandTimeout = mTimeOut
cmd.ActiveConnection = conn
cmd.CommandType = pType
cmd.CommandText = strSql
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).value = aParam(i + 1)
    Next
End If
End Function
Public Function ExecuteTransaction() As Boolean
    
    If Not OpenConnection() Then Exit Function
    
    On Error GoTo ErrorHandler
    Dim IsTransactionStarted As Boolean
    ExecuteTransaction = True
    
    conn.BeginTrans
    IsTransactionStarted = True
    
    Dim sCaption As String
    If Not mForm Is Nothing Then
        mForm.prog1.Visible = True
        mForm.prog1.value = 0
        sCaption = mForm.Caption
    End If
        
    Dim i As Long
    For i = 0 To UBound(sqlStatements)
        If Trim$(sqlStatements(i)) <> "" Then
            If Not mForm Is Nothing Then
                mForm.Caption = sCaption & " ”Ã· " & i + 1 & " „‰ " & UBound(sqlStatements) + 1
                mForm.prog1.value = Round(i / (UBound(sqlStatements)), 2) * 100
            End If
            DoEvents
            conn.Execute sqlStatements(i), , adCmdText + adExecuteNoRecords
        End If
    Next i
    
    conn.CommitTrans
    IsTransactionStarted = False
    
CleanUp:
    On Error Resume Next
    
    ' Fix 4: Proper order of Rollback BEFORE checking connection closures
    If IsTransactionStarted Then
        conn.RollbackTrans
    End If
    
    If Not mForm Is Nothing Then
        mForm.prog1.Visible = False
        mForm.prog1.value = 0
        mForm.Caption = sCaption
    End If
    
    Err.Clear
    On Error GoTo 0
    Exit Function
    
ErrorHandler:
    MsgBox "›‘·  «·⁄„·Ì…  „ «· —«Ã⁄ ⁄‰ ﬂ«›… «· €ÌÌ—« " & vbCrLf & _
           "«·Œÿ√ : " & Err.Description & vbCrLf & _
           "—ﬁ„ «·Ã„·… «·„”»»… ··Œÿ« : " & CStr(i + 1), vbCritical, "Œÿ√ ›Ï ﬁ«⁄œ… «·»Ì«‰« "
    
    On Error Resume Next
    If IsTransactionStarted And (Not conn Is Nothing) Then
        If conn.State = adStateOpen Then
            conn.RollbackTrans
            IsTransactionStarted = False
        End If
    End If
    ExecuteTransaction = False
    On Error GoTo 0
    Resume CleanUp
End Function
Function IncData(sTable As String, sField As String, Optional pWhere As String = "", Optional bNum As Boolean = False) As String
Dim locTable As New ADODB.Recordset
Dim cString As String
cString = "Select Max(" & sField & ") as Maxof " & _
       " FROM " & sTable
If pWhere <> "" Then cString = cString & " where " & pWhere
Set locTable = myRs(cString)
If locTable Is Nothing Then Exit Function
If Not (locTable.EOF) Then
    If IsNull(locTable!maxOf) Then
        IncData = ""
    ElseIf bNum Then
        IncData = Val(locTable!maxOf & "") + 1
    Else
        IncData = IncrementString(locTable!maxOf & "")
    End If
End If
locTable.Close
Set locTable = Nothing
End Function
