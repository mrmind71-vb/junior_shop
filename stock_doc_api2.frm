VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form stock_doc_api 
   BackColor       =   &H00FFFFFF&
   Caption         =   "„” ‰œ«  Ã—œ «· ÿÌ»ﬁ"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   6570
      Width           =   6360
      Begin VB.CheckBox chkTest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«Œ »«— «·«’‰«›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1500
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   2970
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "stock_doc_api2.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdBranchApi 
         Height          =   510
         Left            =   2970
         TabIndex        =   4
         Top             =   135
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "÷»ÿ «ﬂÊ«œ «·›—Ê⁄"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdItemsUpdate 
         Height          =   510
         Left            =   3330
         TabIndex        =   1
         Top             =   135
         Visible         =   0   'False
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "”Õ» „” ‰œ«  «·Ã—œ"
         ButtonStyle     =   3
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   2
      Top             =   7380
      Visible         =   0   'False
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6495
      Left            =   90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   45
      Width           =   6135
      _cx             =   10821
      _cy             =   11456
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "stock_doc_api"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Stock_appFrm
Dim sId_stock As String
Dim bAct As Boolean
Public con As ADODB.Connection
Private Sub cmdBranchApi_Click()
Dim sError As String
Dim JArray As New ChilkatJsonArray
Dim json As New ChilkatJsonObject
Set JArray = GetJsonArray(rsApiStock!baseUrl & rsApiStock!branchesUrl, "Token " & rsApiStock!token, sError)
If sError <> "" Then
    MsgBox sError
    Exit Sub
End If

If JArray.LastMethodSuccess = False Then
     MsgBox JArray.LastErrorText
End If

prog1.Visible = True
For i = 0 To JArray.Size - 1
    prog1.Value = Round((i + 1) / JArray.Size, 2) * 100
    Set json = JArray.ObjectAt(i)
    con.Execute "UPDATE FILE0_40 " & _
                " SET FILE0_40.ID_STOCK = " & json.StringOf("id") & _
                " WHERE FILE0_40.CODE = " & MyParn(json.StringOf("code")), nAffect
    nAffected = nAffected + nAffect
Next
Finally:
prog1.Visible = False
MsgBox " „  ⁄œÌ· " & nAffected & " „Œ“‰"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
GoTo Finally
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdTest_Click()

End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    GetDocs
End If
End Sub

Private Sub Form_Load()
Set rsApiStock = createRs(cmd("SettingApiStock", con, adTable).Execute)

Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")

Fixgrd

sId_stock = GetField("Select id_stock from file0_40 where code = " & MyParn(myForm.xStore.BoundText), con) & ""

End Sub
Private Sub Fixgrd()
grid1.TextMatrix(0, 0) = "«· «—ÌŒ"
grid1.TextMatrix(0, 1) = "«·›—⁄"
grid1.TextMatrix(0, 2) = "«·›—⁄"
grid1.TextMatrix(0, 3) = "«·»Ì«‰"
grid1.TextMatrix(0, 4) = "id"
grid1.ColWidth(0) = 1300
grid1.ColWidth(1) = 1000
grid1.ColWidth(2) = 1800
grid1.ColWidth(3) = 2800
grid1.ColWidth(4) = 1000
grid1.ColHidden(1) = True
grid1.ColHidden(4) = True
For i = 0 To grid1.Cols - 1
    grid1.ColAlignment(i) = flexAlignRightCenter
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set stock_doc_api = Nothing
End Sub
Private Sub GetDocs()
Dim sError As String
Dim JArray As New ChilkatJsonArray
Set JArray = GetJsonArray(rsApiStock!baseUrl & rsApiStock!StockDocURL, "Token " & rsApiStock!token, sError)
If sError <> "" Then
    MsgBox sError
    Exit Sub
End If
Dim i As Long
Dim nAffect As Integer
Dim json As ChilkatJsonObject
Dim aInsert As Variant
prog1.Visible = True
prog1.Value = 0
On Error GoTo myerror
For i = 0 To JArray.Size - 1
    prog1.Value = Round((i + 1) / JArray.Size, 2) * 100
    Set json = JArray.ObjectAt(i)
    If json.StringOf("branch") = sId_stock Then
        grid1.AddItem ""
        grid1.TextMatrix(grid1.Rows - 1, 0) = myFormat_p(json.StringOf("date"))
        grid1.TextMatrix(grid1.Rows - 1, 1) = json.StringOf("branch")
        grid1.TextMatrix(grid1.Rows - 1, 2) = json.StringOf("branch_name")
        grid1.TextMatrix(grid1.Rows - 1, 3) = json.StringOf("name")
        grid1.TextMatrix(grid1.Rows - 1, 4) = json.StringOf("id")
    End If
Next
Finally:
prog1.Visible = False
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
GoTo Finally
End Sub
Private Function AddDocApp(ByRef pError As String, ByRef nAffected As Long) As Boolean

Dim sError As String

Dim json As New ChilkatJsonObject
Dim JArray As New ChilkatJsonArray
Dim sURL As String

Dim bExit As Boolean
Dim sb As New ChilkatStringBuilder

sURL = rsApiStock!baseUrl & rsApiStock!StockItemsURL & "?doc_id=" & grid1.TextMatrix(grid1.Row, grid1.Cols - 1) & "&page_size=1000"
sCaption = Me.Caption

prog1.Visible = True
Dim i As Long
Do
    success = json.Load(GetRest(sURL, "Token " & rsApiStock!token, sError))
    If json.IsNullOf("previous") Then
        nCount = json.StringOf("count")
    End If
    
    JArray.AppendArrayItems json.ArrayOf("results")
    Me.Caption = sCaption & " - " & "”Ã· " & JArray.Size & " „‰ " & nCount
    prog1.Value = Round(JArray.Size / nCount, 2) * 100
    
    sURL = json.StringOf("next") & ""
Loop Until json.IsNullOf("next") = 1 Or sError <> ""



Me.Caption = sCaption
prog1.Value = 0
con.BeginTrans
On Error GoTo myerror
For i = 0 To JArray.Size - 1
    Me.Caption = sCaption & " - " & "”Ã· " & i + 1 & " „‰ " & JArray.Size
    Set json = JArray.ObjectAt(i)
    If json.StringOf("doc_no") = grid1.TextMatrix(grid1.Row, grid1.Cols - 1) Then
        
        
        i2 = i2 + 1
        sb.Append "(" & _
                    addstring(myForm.xDoc_No.text) & "," & _
                    addvalue(json.StringOf("code")) & "," & _
                    json.IntOf("count") & "," & _
                    addstring(json.StringOf("zone")) & "," & _
                    json.StringOf("doc_no") & "," & _
                    addstring(grid1.TextMatrix(grid1.Row, 3)) & "," & _
                    addDate(ISODate(json.StringOf("created_at"))) & "," & _
                    addstring(json.StringOf("created_by")) & _
                     "),"
        
        
        If i2 = 500 Or i = JArray.Size - 1 Then
            i2 = 0
            sb.Shorten 1
            con.Execute "INSERT INTO FILE0_100(" & _
                        "doc_no," & _
                        " item," & _
                        " RealBal," & _
                        " [ZONE]," & _
                        " [doc_app_id]," & _
                        " [doc_app_name]," & _
                        " [doc_app_date]," & _
                        " [doc_app_create])" & _
                         "VALUES" & _
                         sb.GetAsString
            sb.Clear
        End If
        nAffected = nAffected + 1
    End If
Next
con.CommitTrans
AddDocApp = True
Finally:
prog1.Visible = False
Me.Caption = sCaption
Exit Function
myerror:
pError = Err.Description
Err.Clear
con.RollbackTrans
End Function
Private Function TestData() As Boolean
Dim sError As String

Dim json As New ChilkatJsonObject
Dim JArray As New ChilkatJsonArray
Dim sURL As String

Dim bExit As Boolean
Dim sb As New ChilkatStringBuilder

sURL = rsApiStock!baseUrl & rsApiStock!StockItemsURL & "?doc_id=" & grid1.TextMatrix(grid1.Row, grid1.Cols - 1) & "&page_size=1000"
sCaption = Me.Caption

prog1.Visible = True
Dim i As Long
Do
    success = json.Load(GetRest(sURL, "Token " & rsApiStock!token, sError))
    If json.IsNullOf("previous") Then
        nCount = json.StringOf("count")
    End If
    
    JArray.AppendArrayItems json.ArrayOf("results")
    Me.Caption = sCaption & " - " & "”Ã· " & JArray.Size & " „‰ " & nCount
    prog1.Value = Round(JArray.Size / nCount, 2) * 100
    
    sURL = json.StringOf("next") & ""
Loop Until json.IsNullOf("next") = 1 Or sError <> ""

Me.Caption = sCaption
prog1.Value = 0

Dim oError As New ShowError
oError.sCaption = "«’‰«› €Ì— „”Ã·… ›Ï «·—∆Ì”Ì"

'On Error GoTo myerror
For i = 0 To JArray.Size - 1
    Me.Caption = sCaption & " - " & "”Ã· " & i + 1 & " „‰ " & JArray.Size
    prog1.Value = Round(i / (JArray.Size - 1), 2) * 100
    
    Set json = JArray.ObjectAt(i)
    sitem = json.StringOf("code")
    If json.StringOf("doc_no") = grid1.TextMatrix(grid1.Row, grid1.Cols - 1) Then
        If IsNull(rsFunc("dbo.fn_ret_item", con, MyParn(sitem))) Then
            oError.grid1.AddItem ""
            oError.grid1.TextMatrix(oError.grid1.Rows - 1, 0) = oError.grid1.Rows - 1
            oError.grid1.TextMatrix(oError.grid1.Rows - 1, 1) = sitem
        End If
    End If
Next
If oError.grid1.Rows > 1 Then
    oError.Show 1
Else
    Unload oError
    MsgBox "·«  ÊÃœ «’‰«› €Ì— „”Ã·… ›Ï «·—∆Ì”Ì"
End If
TestData = True
Finally:
prog1.Visible = False
Me.Caption = sCaption
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Private Sub grid1_DblClick()
'If grid1.Row = 0 Then Exit Sub
'If Not IsEmpty(GetField("SELECT TOP 1 ID FROM FILE0_100 WHERE DOC_NO = " & MyParn(myForm.xDoc_No.text) & " AND DOC_APP_ID = " & MyParn(grid1.TextMatrix(grid1.Row, 4)), con)) Then
'    MsgBox "«·„” ‰œ „ÊÃÊœ ›Ï «·„·›"
'    Exit Sub
'End If
'
'Me.MousePointer = vbHourglass
'Dim sError As String
'Dim nAffected As Long
'
'If AddDocApp(sError, nAffected) Then
'    MsgBox " „  «÷«›… " & nAffected & " ”Ã·"
'    myForm.myProcAppAdd
'    Exit Sub
'Else
'    MsgBox IIf(sError <> "", sError, "·„   „ «÷«›… «Ì ”Ã·« ")
'End If
'
'Me.MousePointer = vbNormal

If chkTest.Value = 0 Then
    getStock
Else
    testStock
End If
End Sub
Private Sub getStock()
If grid1.Row = 0 Then Exit Sub
If Not IsEmpty(GetField("SELECT TOP 1 ID FROM FILE0_100 WHERE DOC_NO = " & MyParn(myForm.xDoc_No.text) & " AND DOC_APP_ID = " & MyParn(grid1.TextMatrix(grid1.Row, 4)), con)) Then
    MsgBox "«·„” ‰œ „ÊÃÊœ ›Ï «·„·›"
    Exit Sub
End If

Me.MousePointer = vbHourglass
Dim sError As String
Dim nAffected As Long
If AddDocApp(sError, nAffected) Then
    MsgBox " „  «÷«›… " & nAffected & " ”Ã·"
    myForm.myProcAppAdd
    Exit Sub
Else
    MsgBox IIf(sError <> "", sError, "·„   „ «÷«›… «Ì ”Ã·« ")
End If

Me.MousePointer = vbNormal
End Sub
Private Sub testStock()
If grid1.Row = 0 Then Exit Sub
If Not IsEmpty(GetField("SELECT TOP 1 ID FROM FILE0_100 WHERE DOC_NO = " & MyParn(myForm.xDoc_No.text) & " AND DOC_APP_ID = " & MyParn(grid1.TextMatrix(grid1.Row, 4)), con)) Then
    MsgBox "«·„” ‰œ „ÊÃÊœ ›Ï «·„·›"
    Exit Sub
End If

Me.MousePointer = vbHourglass
Dim sError As String
Dim nAffected As Long
TestData
Me.MousePointer = vbNormal
End Sub
