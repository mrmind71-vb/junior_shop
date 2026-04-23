Attribute VB_Name = "mdApiStock"
Public rsApiStock As ADODB.Recordset
Enum enStatus
NoMatch = 404
End Enum
Public Function SendBranches(con As ADODB.Connection, myForm As Form, Optional bDelete As Boolean = False) As Boolean
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl
Dim json As New ChilkatJsonObject
Dim json2 As New ChilkatJsonObject
Dim Succees As Integer

Dim sError As String

Dim nCountAdd As Integer
Dim nCountEdit As Integer

sToken = getToken(con, sError)

If sError <> "" Then
    myForm.AddNote sError, vbRed
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

Dim sCaption As String
sCaption = myForm.Caption

Dim aJson As New ChilkatJsonArray
Dim i As Long

If bDelete Then
    myForm.AddNote "Ì „ Õ–› »Ì«‰«  «·›—Ê⁄"
    jsonResponseStr = rest.FullRequestNoBody("GET", rsApiStock!branchesUrl)
    If aJson.Load(jsonResponseStr) = 1 Then
        If aJson.Size > 0 Then
            For i = 0 To aJson.Size - 1
                myForm.prog1.Value = Round((i + 1) / (aJson.Size), 2) * 100
                myForm.Caption = sCaption & " - Record " & (i + 1) & " from " & aJson.Size
                jsonResponseStr = rest.FullRequestNoBody("DELETE", rsApiStock!branchesUrl & aJson.ObjectAt(i).StringOf("id") & "/")
            Next
            myForm.AddNote "Õ–› " & aJson.Size & " ›—⁄"
        End If
    End If
End If


Dim loctable As New ADODB.Recordset
Set loctable = cmd("dbo.sp_api_stock_branches", con).Execute
If Not IsMissing(myForm) Then
    myForm.prog1.Visible = True
    myForm.prog1.Value = 0
End If

SetKbLayout Lang_AR

myForm.AddNote "Ì „ ‰ﬁ· »Ì«‰«  «·›—Ê⁄"
Do Until loctable.EOF
    myForm.prog1.Value = Round(loctable.AbsolutePosition / (loctable.RecordCount), 2) * 100
    myForm.Caption = sCaption & " - Record " & (loctable.AbsolutePosition) & " from " & loctable.RecordCount
           
    success = json.UpdateString("code", loctable!code)
    success = json.UpdateString("name", loctable!Name)
                                                   
    If IsNull(loctable!ID) Or bDelete Then
        jsonResponseStr = rest.FullRequestString("POST", rsApiStock!branchesUrl, json.Emit())
        
        If json.Load(jsonResponseStr) = 1 Then
            If IsNumeric(json.StringOf("id")) Then
                con.Execute "UPDATE FILE0_40 " & _
                            " SET FILE0_40.ID_STOCK = " & json.StringOf("id") & _
                            " WHERE FILE0_40.CODE = " & MyParn(loctable!code)
                nAdd = 1
            End If
        End If
    Else
        jsonResponseStr = rest.FullRequestString("PUT", rsApiStock!branchesUrl & loctable!ID & "/", json.Emit())
        nEdit = 1
    End If
            
    If (rest.LastMethodSuccess <> 1) Then
        sError = rest.LastErrorText
        Exit Do
   ElseIf rest.ResponseStatusCode >= 400 Then
        sError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")" & vbCrLf & _
                  jsonResponseStr
        Exit Do
    Else
        nCountAdd = nCountAdd + nAdd
        nCountEdit = nCountEdit + nEdit
    End If
lastloop:
    json.Clear
    loctable.MoveNext
Loop

If sError <> "" Then
    MsgBox sError
    myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «·›—Ê⁄ »‰Ã«Õ", vbRed
ElseIf nCountAdd > 0 Then
    myForm.AddNote " „ «÷«›… " & nCountAdd & " ›—⁄"
ElseIf nCountEdit > 0 Then
    myForm.AddNote " „  ⁄œÌ· " & nCountEdit & " ›—⁄"
    SendBranches = True
End If

myForm.Caption = sCaption
myForm.prog1.Visible = False
myForm.prog1.Value = 0
End Function
Public Function sendItemsStockApi(con As ADODB.Connection, myForm As Form, Optional bRenew As Boolean = False) As Boolean
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl
Dim json As New ChilkatJsonObject
Dim json2 As New ChilkatJsonObject
Dim jsonResponseStr As String

Dim sError As String
Dim sToken As String

Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")

sToken = getToken(con, sError)
If sError <> "" Then
    myForm.AddNote sError, vbRed
    Exit Function
End If

success = url.ParseUrl(rsApiStock!baseUrl)
success = rest.AddHeader("Authorization", "Token " & sToken)
success = rest.AddHeader("Content-Type", "application/json")

Dim bAutoReconnect As Long
bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)
If Not IsMissing(myForm) Then
    myForm.AddNote "‰ﬁ· »Ì«‰«  «·«’‰«› «·Ì «·„Êﬁ⁄"
End If


Dim aPrm As Variant
aPrm = AddFlag(aPrm, "changed", "1")

Dim loctable As New ADODB.Recordset
Set loctable = cmd("dbo.sp_api_stock_items", con, adStoredProc, aPrm).Execute


If Not IsMissing(myForm) Then
    myForm.prog1.Visible = True
End If

Dim sCaption As String
sCaption = myForm.Caption

Dim bAdd As Boolean
Dim nCountAdd As Long
Dim nCountEdit As Long
Do Until loctable.EOF
    If Not IsMissing(myForm) Then
        myForm.prog1.Value = Round(loctable.AbsolutePosition / (loctable.RecordCount), 2) * 100
        myForm.Caption = sCaption & " - Record " & (loctable.AbsolutePosition) & " from " & loctable.RecordCount
    End If
    
    json.Clear
    
    success = json.UpdateString("code", loctable!code & "")
    If Not IsNull(loctable!vendor_barcode) Then
        success = json.UpdateString("vendor_barcode", loctable!vendor_barcode & "")
    End If
    
    If Not IsNull(loctable!internal_barcode) Then
        success = json.UpdateString("internal_barcode", loctable!internal_barcode & "")
    End If
    
    success = json.UpdateString("description", loctable!Description & "")
        
    success = json.UpdateString("size", loctable!Size & "")
    success = json.UpdateString("color", loctable!color & "")
    success = json.UpdateString("model_no", loctable!model_no & "")
    
    If (Not IsNull(loctable!GS1)) And IsNumeric(loctable!GS1 & "") Then
        success = json.UpdateNumber("GS1", loctable!GS1 & "")
    End If
        
    jsonResponseStr = rest.FullRequestString("POST", rsApiStock!itemsURL, json.Emit)
    If rest.ResponseStatusCode = 200 Then
        nCountAdd = nCountAdd + 1
    ElseIf rest.ResponseStatusCode = 400 Then
        sonResponseStr = rest.FullRequestString("PUT", rsApiStock!itemsURL & loctable!code & "/", json.Emit)
        nCountEdit = nCountEdit + 1
    End If
        
    If (rest.LastMethodSuccess <> 1) Then
        sError = rest.LastErrorText
        GoTo Finally
    ElseIf rest.ResponseStatusCode >= 400 Then
        sError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")" & vbCrLf & _
        jsonResponseStr
        GoTo Finally
    Else
        con.Execute "UPDATE FILE1_10 SET " & _
                    "FILE1_10.CHANGED = 0," & _
                    "FILE1_10.SEND_TO_API = 1" & _
                    " WHERE FILE1_10.ITEM = " & MyParn(loctable!code)
    End If
    loctable.MoveNext
Loop
Finally:
If sError <> "" Then
    MsgBox sError
    If Not IsMissing(myForm) Then
        myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «·„ÊœÌ·«  »‰Ã«Õ", vbRed
    End If
Else
    sendItemsStockApi = True
    If nCountAdd + nCountEdit = 0 Then
        If Not IsMissing(mhform) Then
            myForm.AddNote IIf(bAll, "·« ÌÊÃœ »Ì«‰«  ··«—”«·", "ﬂ· «·»Ì«‰«  „ÕœÀ…"), vbGreen
        End If
    Else
        If Not IsMissing(myForm) Then
            If nCountAdd > 0 Then myForm.AddNote (" „  «÷«›… »Ì«‰« " & Space(1) & nCountAdd & Space(1) & "’‰›")
            If nCountEdit > 0 Then myForm.AddNote (" „  ⁄œÌ· »Ì«‰« " & Space(1) & nCountEdit & Space(1) & "’‰›")
        End If
    End If
End If

If Not IsMissing(myForm) Then
    myForm.Caption = sCaption
    myForm.prog1.Visible = False
    myForm.prog1.Value = 0
End If
End Function
Public Function sendItemsStockApi_All(con As ADODB.Connection, myForm As Form) As Boolean
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl
Dim json As New ChilkatJsonObject
Dim jsonSearch As New ChilkatJsonObject

Dim sError As String
Dim sToken As String

Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")

sToken = getToken(con, sError)
If sError <> "" Then
    myForm.AddNote sError, vbRed
    Exit Function
End If

success = url.ParseUrl(rsApiStock!baseUrl)
success = rest.AddHeader("Authorization", "Token " & sToken)
success = rest.AddHeader("Content-Type", "application/json")

Dim bAutoReconnect As Long
bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)
myForm.AddNote "‰ﬁ· »Ì«‰«  «·«’‰«› «·Ì «·„Êﬁ⁄"

Dim aJson As New ChilkatJsonArray
Dim sbList As New ChilkatStringBuilder


Dim loctable As New ADODB.Recordset
Set loctable = cmd("dbo.sp_api_stock_items", con, adStoredProc).Execute

If Not IsMissing(myForm) Then
    myForm.prog1.Visible = True
End If


SetKbLayout Lang_AR

Dim sCaption As String
sCaption = myForm.Caption

Dim nCount As Long
Do Until loctable.EOF
    myForm.prog1.Value = Round(loctable.AbsolutePosition / (loctable.RecordCount), 2) * 100
    myForm.Caption = sCaption & " - Record " & (loctable.AbsolutePosition) & " from " & loctable.RecordCount
        
        
    aJson.AddObjectAt (-1)
    Set json = aJson.ObjectAt(aJson.Size - 1)
    
    success = json.UpdateString("code", loctable!code & "")
    
    If Not IsNull(loctable!vendor_barcode) Then
        success = json.UpdateString("vendor_barcode", loctable!vendor_barcode & "")
    End If
    
    If Not IsNull(loctable!internal_barcode) Then
        success = json.UpdateString("internal_barcode", loctable!internal_barcode & "")
    End If
    
    success = json.UpdateString("description", loctable!Description & "")
        
    success = json.UpdateString("size", loctable!Size & "")
    success = json.UpdateString("color", loctable!color & "")
    success = json.UpdateString("model_no", loctable!model_no & "")
    
    If (Not IsNull(loctable!GS1)) And IsNumeric(loctable!GS1 & "") Then
        success = json.UpdateNumber("GS1", loctable!GS1 & "")
    End If
                  
    
    sbList.Append addstring(loctable!code) & ","
    nCount = nCount + 1
    
        
    If aJson.Size = Val(rsApiStock!Size) Or loctable.AbsolutePosition = loctable.RecordCount Then
        jsonResponseStr = rest.FullRequestString("POST", rsApiStock!itemsURL, aJson.Emit)
        
        If (rest.LastMethodSuccess <> 1) Then
            sError = rest.LastErrorText
            GoTo Finally
        ElseIf rest.ResponseStatusCode >= 400 Then
            sError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")" & vbCrLf & _
            jsonResponseStr
            GoTo Finally
        Else
            sbList.Shorten 1
            con.Execute "UPDATE FILE1_10 SET " & _
                        "FILE1_10.CHANGED = 0," & _
                        "FILE1_10.SEND_TO_API = 1" & _
                        " WHERE FILE1_10.ITEM IN (" & sbList.GetAsString & ")"
            aJson.Clear
            sbList.Clear
        End If
        
        If PopUp_Msg_Delayed(1, "continue ...") = 2 Then
            Exit Do
        End If
    End If
    loctable.MoveNext
Last_Loop:
Loop


Finally:
If sError <> "" Then
    MsgBox sError
    myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «·„ÊœÌ·«  »‰Ã«Õ", vbRed
Else
    sendItems = True
    If nCount = 0 Then
        myForm.AddNote IIf(bAll, "·« ÌÊÃœ »Ì«‰«  ··«—”«·", "ﬂ· «·»Ì«‰«  „ÕœÀ…"), vbGreen
    Else
        myForm.AddNote (" „  √÷«›… " & Space(1) & nCount & Space(1) & "’‰› „⁄œ· «·Ì  ÿ»Ìﬁ «·Ã—œ")
    End If
End If

Dim aPrm As Variant
aPrm = AddFlag(aPrm, "changed", "1")
Set loctable = cmd("dbo.sp_api_stock_items", con, adStoredProc, aPrm).Execute
If loctable.RecordCount > 0 Then
    MsgBox "»«ﬁÌ " & loctable.RecordCount & " «’‰«› ·„ Ì „ «—”«·Â«"
End If

myForm.Caption = sCaption
myForm.prog1.Visible = False
myForm.prog1.Value = 0
End Function
Public Function UpdateItemsStockApi(con As ADODB.Connection, myForm As Form) As Boolean
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl
Dim json As New ChilkatJsonObject

Dim sError As String
Dim sToken As String

Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")

sToken = getToken(con, sError)
If sError <> "" Then
    myForm.AddNote sError, vbRed
    Exit Function
End If

success = url.ParseUrl(rsApiStock!baseUrl)
success = rest.AddHeader("Authorization", "Token " & sToken)
success = rest.AddHeader("Content-Type", "application/json")

Dim bAutoReconnect As Long
bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)
myForm.AddNote "‰ﬁ· »Ì«‰«  «·«’‰«› «·Ì «·„Êﬁ⁄"

Dim JsonResp As New ChilkatJsonObject

aPrm = AddFlag(aPrm, "changed", "1")
Dim loctable As New ADODB.Recordset
Set loctable = cmd("dbo.sp_api_stock_items", con, adStoredProc, aPrm).Execute

If Not IsMissing(myForm) Then
    myForm.prog1.Visible = True
End If


SetKbLayout Lang_AR

Dim sCaption As String
sCaption = myForm.Caption


Dim nCount As Long
Do Until loctable.EOF
    myForm.prog1.Value = Round(loctable.AbsolutePosition / (loctable.RecordCount), 2) * 100
    myForm.Caption = sCaption & " - Record " & (loctable.AbsolutePosition) & " from " & loctable.RecordCount
        
    'aJson.Clear
    'aJson.AddObjectAt (-1)
    'Set json = aJson.ObjectAt(aJson.Size - 1)
    json.Clear
    success = json.UpdateString("code", loctable!code & "")
    
    If Not IsNull(loctable!vendor_barcode) Then
        success = json.UpdateString("vendor_barcode", loctable!vendor_barcode & "")
    End If
    
    If Not IsNull(loctable!internal_barcode) Then
        success = json.UpdateString("internal_barcode", loctable!internal_barcode & "")
    End If
    
    success = json.UpdateString("description", loctable!Description & "")
        
    success = json.UpdateString("size", loctable!Size & "")
    success = json.UpdateString("color", loctable!color & "")
    success = json.UpdateString("model_no", loctable!model_no & "")
    
    If (Not IsNull(loctable!GS1)) And IsNumeric(loctable!GS1 & "") Then
        success = json.UpdateNumber("GS1", loctable!GS1 & "")
    End If
    
    If loctable!send_to_api Then
        jsonResponseStr = rest.FullRequestString("PUT", rsApiStock!itemsURL & loctable!code & "/", json.Emit)
    Else
        jsonResponseStr = rest.FullRequestString("POST", rsApiStock!itemsURL, json.Emit)
    End If
            
    If rest.ResponseStatusCode >= 400 Then
        If Not loctable!send_to_api Then
            JsonResp.Clear
            success = JsonResp.Load(jsonResponseStr)
    
            If success = 1 Then
                If JsonResp.JsonTypeOf("code") = 4 Then
                    If InStr(1, LCase(JsonResp.ArrayOf("code").StringAt(0)), "code already exists") > 0 Then
                        jsonResponseStr = rest.FullRequestString("PUT", rsApiStock!itemsURL & loctable!code & "/", json.Emit)
                    End If
                End If
            End If
        Else
            jsonResponseStr = rest.FullRequestString("POST", rsApiStock!itemsURL, json.Emit)
        End If
    End If
        
    If (rest.LastMethodSuccess <> 1) Then
        sError = rest.LastErrorText
        GoTo Finally
    ElseIf rest.ResponseStatusCode >= 400 Then
        sError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")" & vbCrLf & _
        jsonResponseStr
        GoTo Finally
    Else
        con.Execute "UPDATE FILE1_10 SET " & _
                    "FILE1_10.CHANGED = 0," & _
                    "FILE1_10.SEND_TO_API = 1" & _
                    " WHERE FILE1_10.ITEM = " & loctable!code
    End If
    If loctable.AbsolutePosition Mod 500 = 0 Then
        If PopUp_Msg_Delayed(1, "continue ...") = 2 Then
            Exit Do
        End If
    End If
    loctable.MoveNext
Last_Loop:
Loop


Finally:
If sError <> "" Then
    MsgBox sError
    myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «·„ÊœÌ·«  »‰Ã«Õ", vbRed
Else
    updateitemsstock = True
    If nCount = 0 Then
        myForm.AddNote IIf(bAll, "·« ÌÊÃœ »Ì«‰«  ··«—”«·", "ﬂ· «·»Ì«‰«  „ÕœÀ…"), vbGreen
    Else
        myForm.AddNote (" „  √÷«›… " & Space(1) & nCount & Space(1) & "’‰› „⁄œ· «·Ì  ÿ»Ìﬁ «·Ã—œ")
    End If
End If

Set loctable = cmd("dbo.sp_api_stock_items", con, adStoredProc, aPrm).Execute
If loctable.RecordCount > 0 Then
    MsgBox "»«ﬁÌ " & loctable.RecordCount & " «’‰«› ·„ Ì „ «—”«·Â«"
End If

myForm.Caption = sCaption
myForm.prog1.Visible = False
myForm.prog1.Value = 0
End Function
Public Function SendAllItems(con As ADODB.Connection, myForm As Form, Optional bRenew As Boolean = False) As Boolean
'If MsgBox("«—”«· »Ì«‰«  «·«’‰«› «·Ì  ÿ»Ìﬁ «·„Õ„Ê· ?", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
'If Not UpdateBalance(con, myForm, bAll) Then Exit Function
'If Not SendSections(con, myForm, bAll) Then Exit Function
'If Not SendModels(con, myForm, bAll) Then Exit Function
'If Not sendItems(con, myForm, bRenew) Then Exit Function
'SendAllItems = True
End Function
Private Function getToken(con As ADODB.Connection, Optional ByRef pError As String) As String
Dim httpB As New ChilkatHttp
Dim reqB As New ChilkatHttpRequest

If Not IsNull(rsApiStock!token) Then
    getToken = rsApiStock!token & ""
    Exit Function
End If

reqB.httpVerb = "POST"
reqB.Path = rsApiStock!authUrl & ""

reqB.AddHeader "Content-Type", "application/x-www-form-urlencoded"
reqB.AddHeader "Accept", "Application/json"
reqB.AddHeader "contentType", "Application/json"
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
    con.Execute "update SettingApiStock set token = " & MyParn(getToken)
    Set rsApiStock = createRs(cmd("SettingApiStock", con, adTable).Execute)
Else
    pError = json.Emit
End If
End Function
Public Function PopUp_Msg_Delayed(AckTime As Integer, sMsg, Optional sCaption As String, Optional defButton As Integer = vbDefaultButton1) As Integer
    
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    
    ' Display the pop-up. It will close automatically after 't' seconds or if a button is clicked.
    ' Case 1 means the user clicked OK, Case -1 means the time expired.
    
     PopUp_Msg_Delayed = InfoBox.Popup(sMsg & vbCrLf & _
                                "This window closes automatically after " & AckTime & " seconds", _
                              AckTime, _
                              sCaption, _
                              vbOKCancel + defButton) ' Use vbOKCancel for OK and Cancel buttons
    Set InfoBox = Nothing
End Function
Public Function PopUp_Msg(AckTime As Integer, sMsg As String, Optional sCaption As String, Optional defButton As Integer = vbDefaultButton1) As Integer
    
 Call MessageBoxTimeout(0, sCaption, sMsg, vbInformation, 0, AckTime)
End Function



