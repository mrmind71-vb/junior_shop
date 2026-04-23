Attribute VB_Name = "mdApiStock"
Public rsApiStock As ADODB.Recordset
Public Function SendBranches(con As ADODB.Connection, myForm As Form, Optional bDelete As Boolean = False) As Boolean
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl
Dim Succees As Integer
Dim json As New ChilkatJsonObject
Dim aJson As New ChilkatJsonArray
Dim sError As String
Dim nAdd As Integer

sToken = getToken(sError)
'sToken = "0d324832288cbced07fdd2f5dd73aeedd98dc60c"
If pError <> "" Then
    myForm.AddNote pError, vbRed
    Exit Function
End If

success = url.ParseUrl(rsApiStock!baseUrl)
success = rest.AddHeader("Authorization", "Token " & sToken)
success = rest.AddHeader("Content-Type", "application/json")

Dim bAutoReconnect As Long
bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)

If success <> 1 Then
    MsgBox rest.LastErrorText
    myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «·›—Ê⁄ »‰Ã«Õ", vbRed
    Exit Function
End If


Dim loctable As New ADODB.Recordset
Dim bExit As Boolean

Set loctable = cmd("dbo.sp_api_stock_branches", con).Execute

If Not IsMissing(myForm) Then
    myForm.prog1.Visible = True
    myForm.prog1.Value = 0
End If

Dim i As Long
Dim nRecordCount As Long

nRecordCount = loctable.RecordCount

SetKbLayout Lang_AR
Dim sCaption As String
sCaption = myForm.Caption

If bDelete Then
    myForm.AddNote "Ì „ Õ–› »Ì«‰«  «·›—Ê⁄"
    jsonResponseStr = rest.FullRequestNoBody("GET", rsApiStock!branchesurl)
    If aJson.Load(jsonResponseStr) = 1 Then
        If aJson.Size > 0 Then
            For i = 0 To aJson.Size - 1
                myForm.prog1.Value = Round((i + 1) / (aJson.Size), 2) * 100
                myForm.Caption = sCaption & " - Record " & (i + 1) & " from " & aJson.Size
                jsonResponseStr = rest.FullRequestNoBody("DELETE", rsApiStock!branchesurl & aJson.ObjectAt(i).StringOf("id") & "/")
            Next
            myForm.AddNote "Õ–› " & aJson.Size & " ›—⁄"
        End If
    End If
End If

Dim bAdd As Boolean
i = 0

myForm.AddNote "Ì „ ‰ﬁ· »Ì«‰«  «·›—Ê⁄"
Do Until loctable.EOF
    myForm.prog1.Value = Round(loctable.AbsolutePosition / (loctable.RecordCount), 2) * 100
    myForm.Caption = sCaption & " - Record " & (loctable.AbsolutePosition) & " from " & loctable.RecordCount
    
'    bAdd = False
'    If Not IsNull(loctable!ID) Then
'        jsonResponseStr = rest.FullRequestNoBody("GET", rsApiStock!branchesurl & loctable!ID & "/")
'        If json.Load(jsonResponseStr) = 1 Then
'            If json.StringOf("id") = "" Then
'                bAdd = True
'            End If
'        End If
'    End If
    
       
    success = json.UpdateString("code", loctable!code)
    success = json.UpdateString("name", loctable!Name)
                                           
        
    If IsNull(loctable!ID) Or bDelete Then
        jsonResponseStr = rest.FullRequestString("POST", rsApiStock!branchesurl, json.Emit())
        
        If json.Load(jsonResponseStr) = 1 Then
            If IsNumeric(json.StringOf("id")) Then
                con.Execute "UPDATE FILE0_40 " & _
                            " SET FILE0_40.ID_STOCK = " & json.StringOf("id") & _
                            " WHERE FILE0_40.CODE = " & MyParn(loctable!code)
                nAdd = 1
            End If
        End If
    Else
        jsonResponseStr = rest.FullRequestString("PUT", rsApiStock!branchesurl & loctable!ID & "/", json.Emit())
        nEdit = 1
    End If
            
    If (rest.LastMethodSuccess <> 1) Then
        cError = rest.LastErrorText
        MsgBox cError
        myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «·›—Ê⁄ »‰Ã«Õ", vbRed
        Exit Do
   ElseIf rest.ResponseStatusCode >= 400 Then
        success = json.Load(jsonResponseStr)
        cError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")" & vbCrLf & _
                  jsonResponseStr
        MsgBox cError
        myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «·«ﬁ”«„ »‰Ã«Õ", vbRed
        Exit Do
    Else
        nCountAdd = nCountAdd + nAdd
        nCountEdit = nCountEdit + nEdit
    End If
lastloop:
    json.Clear
    loctable.MoveNext
Loop
If nCountAdd > 0 Then
    myForm.AddNote " „ «÷«›… " & nCountAdd & " ›—⁄"
ElseIf nCountEdit > 0 Then
    myForm.AddNote " „  ⁄œÌ· " & nCountEdit & " ›—⁄"
Else
    myForm.AddNote "·„ Ì „  ⁄œÌ· «Ì ›—Ê⁄"
End If

SendBranches = cError = ""

If Not IsMissing(myForm) Then
    myForm.Caption = sCaption
    myForm.prog1.Visible = False
    myForm.prog1.Value = 0
End If
End Function
Private Function SendModels(con As ADODB.Connection, myForm As Form, Optional bAll As Boolean = False) As Boolean
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl
Dim Succees As Integer
Dim json As ChilkatJsonObject
Dim aJson As Variant

myForm.AddNote "”Ì „ ‰ﬁ· »Ì«‰«  «·„ÊœÌ·«  «·Ì «·„Êﬁ⁄"

success = url.ParseUrl(mainUrl)
success = rest.AddHeader("Authorization", mbToken)
success = rest.AddHeader("Content-Type", "application/json")

Dim bAutoReconnect As Long
bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)

Dim aParam(0) As Variant
aParam(0) = Array("ALL", adBoolean, adParamInput, -1, IIf(bAll, 1, 0))

Dim loctable As New ADODB.Recordset
Set loctable = myRecordProc("GetModels", con, aParam)

If Not IsMissing(myForm) Then
    myForm.prog1.Visible = True
    myForm.prog1.Value = 0
End If

Dim jArray As New ChilkatJsonArray
Dim i As Long
Dim nRecordCount As Long
Dim sCaption As String
Dim CLIST As String

sCaption = myForm.Caption
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    If Not IsMissing(myForm) Then
        myForm.prog1.Value = Round((i + 1) / (nRecordCount), 2) * 100
        myForm.Caption = sCaption & " - Record " & (i + 1) & " from " & nRecordCount
    End If
    nLast = jArray.Size
    success = jArray.AddObjectAt(nLast)
    Set json = jArray.ObjectAt(nLast)
    success = json.UpdateString("code", loctable!code & "")
    success = json.UpdateString("name", loctable!Name & "")
    success = json.UpdateString("age", loctable!age & "")
    success = json.UpdateString("section", loctable!Section & "")
    success = json.UpdateString("barcode", loctable!BARCODE & "")
    success = json.UpdateString("unit", loctable!UNIT & "")
    success = json.UpdateInt("limit", mRound(loctable!Limit))
    success = json.UpdateNumber("repeat", mRound(loctable!Repeat))
    success = json.UpdateBool("show_in_app", IIf(Abs(loctable!App) = 0, 0, 1))
    
    CLIST = CLIST & IIf(CLIST = "", "", ",") & MyParn(loctable!code)
    
    If jArray.Size = mbSize Then
        myForm.Caption = sCaption & " - Sending Data"
        jsonResponseStr = rest.FullRequestString("POST", modelUrl, jArray.Emit)
        If (rest.LastMethodSuccess <> 1) Then
            cError = rest.LastErrorText
            Exit Do
        ElseIf rest.ResponseStatusCode >= 400 Then
            cError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")"
            Exit Do
        Else
            nCount = nCount + jArray.Size
        End If
        jArray.Clear
    End If
    loctable.MoveNext
    i = i + 1
Loop

If cError = "" Then
    If Trim(jArray.Emit()) <> "" Then
        jsonResponseStr = rest.FullRequestString("POST", modelUrl, jArray.Emit)
        If (rest.LastMethodSuccess <> 1) Then
            cError = rest.LastErrorText
        ElseIf rest.ResponseStatusCode >= 400 Then
            cError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")"
        Else
            nCount = nCount + jArray.Size
        End If
    End If
End If

If cError = "" Then
    SendModels = True
    If nCount = 0 Then
        myForm.AddNote IIf(bAll, "·« ÌÊÃœ »Ì«‰«  ··«—”«·", "·« ÌÊÃœ »Ì«‰«  ··«—”«·"), vbGreen
    Else
        myForm.AddNote " „ «—”«· »Ì«‰«  " & nCount & " „ÊœÌ· «·Ì «·„Êﬁ⁄"
    End If
Else
    MsgBox Err.Description
    myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «·„ÊœÌ·«  »‰Ã«Õ", vbRed
End If

If Not IsMissing(myForm) Then
    myForm.Caption = sCaption
    myForm.prog1.Visible = False
    myForm.prog1.Value = 0
End If
End Function
Private Function sendItems(con As ADODB.Connection, myForm As Form, Optional bAll As Boolean = False) As Boolean
Dim httpB As New ChilkatHttp
Dim respB As New ChilkatHttpResponse

Dim Succees As Integer

Dim json As New ChilkatJsonObject
Dim cError As String

myForm.AddNote "”Ì „ ‰ﬁ· »Ì«‰«  «·«’‰«› «·Ì «·„Êﬁ⁄"


Dim sToken As String
Dim pError As String
'sToken = getToken(pError)
sToken = "0d324832288cbced07fdd2f5dd73aeedd98dc60c"
If pError <> "" Then
    myForm.AddNote pError
    Exit Function
End If

httpB.SetRequestHeader "Authorization", "Token " & sToken
httpB.SetRequestHeader "Content-Type", "application/json"
httpB.SetRequestHeader "Accept", "Application/json"

'respb.AddParam "grant_type", "client_credentials"
'httpB.SetRequestHeader "username", rsApiStock!Us   erName
'httpB.SetRequestHeader "password", rsApiStock!PassWord

Dim bAutoReconnect As Long
bAutoReconnect = 1


Dim aPrm As Variant
'aPrm = AddFlag(aPrm, "changed", IIf(bAll, Null, "1"))
aPrm = AddFlag(aPrm, "changed", Null)

Dim loctable As New ADODB.Recordset
Set loctable = cmd("GetItems", con, adStoredProc, aPrm).Execute

If Not IsMissing(myForm) Then
    myForm.prog1.Visible = True
End If


Dim i As Long
Dim nRecordCount As Long

nRecordCount = loctable.RecordCount


SetKbLayout Lang_AR

Dim sCaption As String
Dim CLIST As String
sCaption = myForm.Caption

Do Until loctable.EOF
    myForm.prog1.Value = Round((i + 1) / (nRecordCount), 2) * 100
    myForm.Caption = sCaption & " - Record " & (i + 1) & " from " & nRecordCount
            
    success = json.UpdateString("code", loctable!code & "")
    success = json.UpdateString("vendor_barcode", loctable!vendor_barcode & "")
    success = json.UpdateString("internal_barcode", loctable!internal_barcode & "")
    success = json.UpdateString("description", loctable!Description & "")
    success = json.UpdateString("size", loctable!Size & "")
    success = json.UpdateString("color", loctable!color & "")
    success = json.UpdateString("model_no", loctable!model_no & "")
    success = json.UpdateNumber("GS1", loctable!GS1 & "")
    
    myForm.Caption = sCaption & " - Sending Data"
    
    Set respB = New ChilkatHttpResponse
    Set respB = httpB.PostJson2(rsApiStock!baseUrl & rsApiStock!itemsURL, "application/json", json.Emit())
    
'    jsonResponseStr = rest.FullRequestString("POST", rsApiStock!itemsurl, json.Emit)
'     If (rest.LastMethodSuccess <> 1) Then
'        cError = rest.LastErrorText
'        Exit Do
'    ElseIf rest.ResponseStatusCode >= 400 Then
'        cError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")"
'        Exit Do
'    End If
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
           con.Execute "update file1_10 " & _
                       " set file1_10.changed = 0 where file1_10.item = " & loctable!code
            MsgBox respB.BodyStr
        End If
    loctable.MoveNext
    i = i + 1
Loop


If cError <> "" Then
    MsgBox Err.Description
    myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «·„ÊœÌ·«  »‰Ã«Õ", vbRed
Else
    sendItems = True
    If nCount = 0 Then
        myForm.AddNote IIf(bAll, "·« ÌÊÃœ »Ì«‰«  ··«—”«·", "ﬂ· «·»Ì«‰«  „ÕœÀ…"), vbGreen
    Else
        myForm.AddNote (" „ «—”«· »Ì«‰« " & Space(1) & nCount & Space(1) & "’‰› «·Ì «·„Êﬁ⁄")
        'CountSentError myForm, itemUrlShow, "«·«’‰«›"
    End If
    'con.Execute "update file1_10 set file1_10.changed = 0 where file1_10.changed = 1"
End If

If Not IsMissing(myForm) Then
    myForm.Caption = sCaption
    myForm.prog1.Visible = False
    myForm.prog1.Value = 0
End If
End Function
Public Function SendAllItems(con As ADODB.Connection, myForm As Form, Optional bAll As Boolean = False) As Boolean
If MsgBox("«—”«· »Ì«‰«  «·«’‰«› «·Ì  ÿ»Ìﬁ «·„Õ„Ê· ?", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
'If Not UpdateBalance(con, myForm, bAll) Then Exit Function
'If Not SendSections(con, myForm, bAll) Then Exit Function
'If Not SendModels(con, myForm, bAll) Then Exit Function
If Not sendItems(con, myForm, bAll) Then Exit Function
SendAllItems = True
End Function
Public Function Getlist(myFrom As Form, pUrl As String, ByRef cError As String, Optional pField As String = "code") As String
Dim jArray As New ChilkatJsonArray
Dim json As ChilkatJsonObject
Set jArray = GetJsonArray(mainUrl & pUrl, cError)

If cError <> "" Then Exit Function
For i = 0 To jArray.Size - 1
    Set json = jArray.ObjectAt(i)
    Getlist = Getlist & "[" & json.StringOf(pField) & "]"
Next
End Function
Private Function GetlistNext(myForm As Form, pUrl As String, ByRef cError As String, Optional pField As String = "code") As String
Dim jArray As New ChilkatJsonArray
Dim json As ChilkatJsonObject
Set jArray = GetJsonArrayNext(myForm, mainUrl & pUrl, cError)
If cError <> "" Then Exit Function
For i = 0 To jArray.Size - 1
    Set json = jArray.ObjectAt(i)
    GetlistNext = GetlistNext & "[" & json.StringOf(pField) & "]"
Next
End Function
Private Function GetlistString(myFrom As Form, pUrl As String, ByRef cError As String, Optional pField As String = "code") As String
Dim jArray As New ChilkatJsonArray
Dim json As ChilkatJsonObject

Set jArray = GetJsonArray(mainUrl & pUrl, cError)
If cError <> "" Then Exit Function
For i = 0 To jArray.Size - 1
    Set json = jArray.ObjectAt(i)
    GetlistString = GetlistString & IIf(GetlistString = "", "", ",") & MyParn(json.StringOf(pField))
Next
End Function
Private Function updateList(ByRef CLIST As String, con As ADODB.Connection)
If Trim(CLIST) <> "" Then
    con.Execute "update file1_10 set changed = 0 where file1_10.item in(" & CLIST & ")", nCount
    CLIST = ""
End If
End Function
Private Function UpdateBalance(con As ADODB.Connection, myForm As Form, bAll As Boolean) As Boolean
Dim nCount As Long
Dim nAffect As Long
con.BeginTrans


myForm.AddNote ("”Ì „  ÕœÌÀ «—’œ… «·«’‰«›")

On Error GoTo myError
'con.Execute "exec [dbo].UpdateItemsNoBal " & IIf(bAll, "1", "0"), nAffect
'nCount = nAffect

con.Execute "exec [dbo].sp_updateItemsBal " & IIf(bAll, "1", "0"), nAffect
nCount = nCount + nAffect
con.CommitTrans

UpdateBalance = True

If nCount = 0 Then
    myForm.AddNote "«·«—’œ… „ÕœÀ… »«·ﬂ«„· !! ·„ Ì „ «· ⁄œÌ·", vbGreen
Else
    myForm.AddNote (" „  ⁄œÌ· «—’œ… " & nCount & " ’‰›")
End If
Exit Function
myError:
MsgBox Err.Description
myForm.AddNote "Œÿ√ «À‰«¡ «· ÕœÌÀ ·„ Ì „ «· ÕœÌÀ »‰Ã«Õ", vbRed
Err.Clear
con.RollbackTrans
End Function
Public Function getToken(Optional ByRef pError As String) As String

Dim httpB As New ChilkatHttp
Dim reqB As New ChilkatHttpRequest


reqB.httpVerb = "POST"
reqB.Path = rsApiStock!authUrl & ""

'reqB.AddHeader "Authorization", loctable!auth & ""
reqB.AddHeader "Content-Type", "application/x-www-form-urlencoded"
reqB.AddHeader "Accept", "Application/json"
reqB.AddHeader "contentType", "Application/json"
'reqB.AddHeader "x-api-key", loctable!apiKey & ""
reqB.AddParam "username", rsApiStock!UserName & ""
reqB.AddParam "password", rsApiStock!PassWord & ""

Dim respB As ChilkatHttpResponse
Set respB = httpB.PostUrlEncoded(rsApiStock!baseUrl & rsApiStock!authUrl, reqB)
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
If json.StringOf("token") <> "" Then
    getToken = json.StringOf("token")
Else
    pError = json.Emit
End If
End Function

