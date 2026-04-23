VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form balance_factfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "√€·«ﬁ › —…"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   990
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand cmdExit 
      Height          =   555
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   979
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
      Picture         =   "balance_fact.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdApply 
      Height          =   555
      Left            =   1710
      TabIndex        =   1
      Top             =   90
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   979
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "”Õ» «—’œ… «·«’‰«›"
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   795
      Visible         =   0   'False
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   3
      Top             =   330
      Visible         =   0   'False
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   45
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   714
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "balance_factfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New adodb.Connection
Dim conFact As New adodb.Connection
Dim bAct As Boolean
Private Sub CmdApply_Click()
DoEvents
Me.MousePointer = 11
cmdApply.Enabled = False
getData
cmdApply.Enabled = True
Me.MousePointer = 0
End Sub
Private Function getData()
Dim cString As String
'con.Execute "UPDATE FILE1_10 SET FILE1_10.BALANCE_FACT = 0"

cmd("UPDATE FILE1_10 SET FILE1_10.BALANCE_FACT = 0", con).Execute

Dim loctable As New adodb.Recordset
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "STORE", "004")
Set loctable = cmd("[dbo].[sp_balance]", conFact, adStoredProc, aPrm).Execute
Dim sCaption As String
Dim strTab As New ChilkatStringTable
Dim nRecord As Long
prog1.Visible = True
sCaption = Me.Caption
Do Until loctable.EOF
    nRecord = nRecord + 1
    Me.Caption = sCaption & " ”Ã· " & nRecord & " „‰ " & loctable.RecordCount
    prog1.Value = Round(nRecord / loctable.RecordCount, 2) * 100
          
    strTab.Append "update file1_10 " & _
                 " set file1_10.balance_fact = " & mRound(loctable!balance) & _
                   " where BARCODE2 = " & MyParn(loctable!BARCODE2) & ";"
    loctable.MoveNext
Loop

prog1.Visible = True
nCount = 1000
For i = 0 To strTab.Count - 1 Step nCount
    Me.Caption = sCaption & " - " & IIf(i + nCount > strTab.Count, strTab.Count, i + nCount) & " „‰ " & strTab.Count
    prog1.Value = Round((i + 1) / strTab.Count, 2) * 100
    
    On Error Resume Next
    'cmd(strTab.GetStrings(i, nCount, 1), con).Execute
    
    con.Execute strTab.GetStrings(i, nCount, 1), nAffect
    
    If Err.Number <> 0 Then
        MsgBox Err.Description
        MsgBox strTab.GetStrings(i, nCount, 1)
        GoTo Finally
        'Err.Clear
    End If
Next

MsgBox "⁄œœ «·”Ã·«  " & strTab.Count
Unload Me
Finally:
prog1.Visible = False
Me.Caption = sCaption
Set loctable = Nothing
Exit Function
myError:
MsgBox Err.Description
Err.Clear
GoTo Finally
End Function
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    CmdApply_Click
End If
End Sub

Private Sub Form_Load()
openConFACT conFact
openCon con

End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set closefrm = Nothing
End Sub

Private Sub xClosed_Click()
cmdApply.Caption = IIf(xClosed.Value = 0, "› Õ", "«€·«ﬁ")
End Sub

Private Sub xDate1_DblClick()
Set datefrm.oDate = xDate1
datefrm.Show 1
End Sub

Private Sub xdate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xDate1
End Sub

Private Sub xdate2_DblClick()
Set datefrm.oDate = xdate2
datefrm.Show 1
End Sub

Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xdate2
End Sub
Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate xdate2
End Sub

