Attribute VB_Name = "mdlMall"
Public Function getTokenMall(Optional ByRef pError As String) As String

If rsMall.EOF Then
    pError = "„·› ÷»ÿ »Ì«‰«  «· ÊﬂÌ‰ €Ì— „ÊÃÊœ"
    Exit Function
End If

Dim httpB As New ChilkatHttp
Dim reqB As New ChilkatHttpRequest


reqB.httpVerb = "POST"
reqB.Path = rsMall!authUrl & ""

'reqB.AddHeader "Authorization", loctable!auth & ""
reqB.AddHeader "Content-Type", "application/x-www-form-urlencoded"
reqB.AddHeader "Accept", "Application/json"
reqB.AddHeader "contentType", "Application/json"
'reqB.AddHeader "x-api-key", loctable!apiKey & ""
reqB.AddParam "email", rsMall!mail & ""
reqB.AddParam "password", rsMall!PassWord & ""

Dim respB As ChilkatHttpResponse
Set respB = httpB.PostUrlEncoded(rsMall!baseUrl & rsMall!authUrl, reqB)
If (httpB.LastMethodSuccess = 0) Then
    pError = httpB.LastErrorText
    Exit Function
End If

Dim json As New ChilkatJsonObject
success = json.Load(respB.BodyStr)
If (success <> 1) Then
    pError = "Error : " & json.LastErrorText
    Exit Function
End If
'If Not json.ObjectOf("result") Is Nothing Then
    getTokenMall = json.StringOf("access")
'End If
End Function
Public Function SendInvoicesMall(con As ADODB.Connection, pBranch As String, ByRef pError As String, pCount As Long, Optional myform As Form) As Boolean
Dim loctable As New ADODB.Recordset
If rsMall.EOF Then
    SendInvoicesMall = False
    Exit Function
End If

Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")

On Error GoTo myError:

Dim aPrm As Variant
aPrm = AddFlag(Empty, "Branch", pBranch)
aPrm = AddFlag(aPrm, "date_begin", myFormat_sp(rsMall!date_begin))

Set loctable = mycmd("[dbo].[sp_invoices_mall]", con, adStoredProc, aPrm)

If loctable.EOF Then
    SendInvoicesMall = True
    Exit Function
End If


Dim httpB As New ChilkatHttp
Dim reqB As New ChilkatHttpRequest

Dim sToken As String
sToken = getTokenMall(pError)

If pError <> "" Then
    Exit Function
End If

httpB.SetRequestHeader "Authorization", "Bearer " & sToken
httpB.SetRequestHeader "Content-Type", "application/json"
httpB.SetRequestHeader "Acrcept", "Application/json"
'httpB.SetRequestHeader "x-api-key", apiKey
'httpB.SetRequestHeader "x-token", token

Dim json As New ChilkatJsonObject
Dim respB As ChilkatHttpResponse
Dim sid As String

If Not myform Is Nothing Then
    myform.prog1.Visible = True
    myform.prog1.Value = 0
End If

Dim nRecord As Long
Do Until loctable.EOF
    
    If Not myform Is Nothing Then
        nRecord = nRecord + 1
        myform.prog1.Value = Round(nRecord / loctable.RecordCount, 2) * 100
    End If
        
    If createLines(loctable, json) Then
        Set respB = New ChilkatHttpResponse
        Set respB = httpB.PostJson2(rsMall!baseUrl & rsMall!invoiceUrl, "application/json", json.Emit())
        If (respB.LastMethodSuccess <> 1) Then
            pError = respB.LastErrorText & vbCrLf & _
                    respB.BodyStr
            Exit Function
        ElseIf respB.StatusCode >= 400 Then
            pError = respB.StatusCode & vbCrLf & _
                     respB.StatusText & vbCrLf & _
                     respB.LastErrorText & vbCrLf & _
                     respB.BodyStr
        ElseIf respB.StatusCode = 200 Or respB.StatusCode = 201 Then
            pCount = pCount + 1
            
            json.Clear
            success = json.Load(respB.BodyStr)
            sid = json.StringOf("id")
            
            con.Execute "update file6_20h set " & _
                   " posted_mall = 1, " & _
                   "ID_MALL = " & addstring(sid) & _
                   " Where doc_no = " & MyParn(loctable!doc_no)
        End If
        
    End If
    loctable.MoveNext
Loop
SendInvoicesMall = True
Finally:
If Not myform Is Nothing Then
    myform.prog1.Visible = False
    myform.prog1.Value = 0
End If
Exit Function
myError:
pError = Err.Description
Err.Clear
GoTo Finally
End Function
Private Function createLines(loctable As ADODB.Recordset, ByRef json As ChilkatJsonObject) As Boolean
Dim success As Integer
json.Clear
success = json.AddStringAt(-1, "branch", loctable!branch)
success = json.AddStringAt(-1, "date", Format(Date, "YYYY-MM-DDThh:nn:ss.msZ"))
success = json.AddStringAt(-1, "discount", Format(mRound(loctable!discount), "Fixed"))
success = json.AddStringAt(-1, "invoice_number", loctable!doc_no)
success = json.AddStringAt(-1, "subtotal", Format(mRound(loctable!Subtotal), "Fixed"))
success = json.AddStringAt(-1, "tax", Format(mRound(loctable!tax), "Fixed"))
success = json.AddStringAt(-1, "total", Format(mRound(loctable!TOTAL), "Fixed"))
success = json.AddBoolAt(-1, "is_return", loctable!is_return)
createLines = True
End Function
Public Function GetJsonArray(ByVal pUrl As String, pToken As String, ByRef pError As String, Optional pName As String = "", Optional pValue As String) As ChilkatJsonArray
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl
Dim sbJson As New ChilkatStringBuilder
Dim json As New ChilkatJsonObject
Dim Succees As Integer

success = url.ParseUrl(pUrl)
success = rest.AddHeader("Authorization", pToken)
If pName <> "" Then
    success = rest.AddQueryParam(pName, pValue)
End If

Dim bAutoReconnect As Long
bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)

Dim jsonarraystr As String
jsonarraystr = rest.FullRequestNoBody("GET", url.Path)
If (rest.LastMethodSuccess <> 1) Then
    pError = rest.LastErrorText
    Exit Function
ElseIf rest.ResponseStatusCode >= 400 Then
    pError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")"
    Exit Function
End If

success = sbJson.Append("{""array"":")
success = sbJson.Append(jsonarraystr)
success = sbJson.Append("}")

success = json.Load(sbJson.GetAsString())

Set GetJsonArray = json.ArrayAt(0)
End Function
Public Function GetRest(ByVal pUrl As String, pToken As String, ByRef pError As String, Optional pName As String = "", Optional pValue As String) As String
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl

Dim sb As New ChilkatStringBuilder
Dim Succees As Integer

success = url.ParseUrl(pUrl)
success = rest.AddHeader("Authorization", pToken)

If url.Query <> "" Then
    rest.AddQueryParams (url.Query)
End If

Dim bAutoReconnect As Long
Dim json As New ChilkatJsonObject
Dim jsonarraystr As String

bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)
GetRest = rest.FullRequestNoBody("GET", url.Path)

If (rest.LastMethodSuccess <> 1) Then
    pError = rest.LastErrorText
    Exit Function
ElseIf rest.ResponseStatusCode >= 400 Then
    pError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")"
    Exit Function
End If
End Function
Public Function GetJsonSb(ByVal pUrl As String, pToken As String, ByRef pError As String, Optional pName As String = "", Optional pValue As String) As ChilkatStringBuilder
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl

Dim sb As New ChilkatStringBuilder
Dim Succees As Integer

success = url.ParseUrl(pUrl)
success = rest.AddHeader("Authorization", pToken)
If pName <> "" Then
    success = rest.AddQueryParam(pName, pValue)
End If

Dim bAutoReconnect As Long
Dim json As New ChilkatJsonObject
bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)

success = sb.Append(rest.FullRequestNoBody("GET", url.Path))

If (rest.LastMethodSuccess <> 1) Then
    pError = rest.LastErrorText
    Exit Function
ElseIf rest.ResponseStatusCode >= 400 Then
    pError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")"
    Exit Function
End If
Set GetJsonSb = sb
End Function


