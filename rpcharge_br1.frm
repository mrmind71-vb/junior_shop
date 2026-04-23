VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpcharge_br1 
   Caption         =   " ﬁ«—Ì— «·‰ﬁœÌ…"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3105
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   2295
      Width           =   1500
   End
   Begin VB.CommandButton cmdClear 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1575
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   2295
      Width           =   1500
   End
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   45
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   2295
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   2130
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   90
      Width           =   5145
      Begin VB.CheckBox chkCash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "‰ﬁœÌ ›ﬁÿ"
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
         Height          =   330
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1710
         Width           =   1320
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2325
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2325
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xbox 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   945
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XBRANCH 
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   1305
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc DATA2 
         Height          =   330
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·Ï  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Œ“‰…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   990
         Width           =   360
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3555
      Top             =   2295
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2925
      Top             =   2070
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "rpcharge_br1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
    doprint1
End Sub
Private Sub doprint1()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
ReDim aHeader(4)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select Charge,Sum(Value) as SumofValue," & _
          "File8_51.MainGroup,File8_51.DescA,File8_52.Desca as MainGroupDesca " & _
          " FROM acc_charge_BR INNER JOIN File8_51 ON acc_charge_BR.CHARGE = File8_51.CODE LEFT JOIN file8_52 ON File8_51.MAINGROUP = file8_52.CODE LEFT JOIN FILE0_50 ON FILE0_50.CODE = acc_charge_BR.BOX WHERE VALUE <> 0"

If lIsBranchStore Then
    cString = cString & " AND FILE0_50.BRANCH = " & MyParn(cBranch)
End If

If xbox.BoundText <> "" Then
    cString = cString & " AND BOX = " & MyParn(xbox.BoundText)
    aHeader(1) = "[ Œ“‰… " & xbox.text & "]"
End If
If XBRANCH.BoundText <> "" Then
    cString = cString & " AND acc_charge_BR.BRANCH = " & MyParn(XBRANCH.BoundText)
    aHeader(2) = "[ ›—⁄ " & XBRANCH.text & "]"
End If
If IsDate(xdate1.text) Then
    cString = cString & " AND acc_charge_BR.date  >= " & DateSq(xdate1.text)
    aHeader(0) = "[" & BetweenString(Format(xdate1.text, "d-m-yyyy"), Format(xdate2.text, "d-m-yyyy")) & "]"
End If

If IsDate(xdate2.text) Then
    cString = cString & " AND acc_charge_BR.date <= " & DateSq(xdate2.text)
     aHeader(0) = "[" & BetweenString(Format(xdate1.text, "d-m-yyyy"), Format(xdate2.text, "d-m-yyyy")) & "]"
End If
    
If chkCash.Value = 1 Then
    cString = cString & " AND BOX IS NOT NULL"
    aHeader(4) = "[‰ﬁœÌ ›ﬁÿ]"
End If
cString = cString & " Group By Charge,File8_51.MainGroup,File8_51.DescA,FILE8_52.DESCA"
          
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !CHARGE
    temptable!str2 = !DESCA
    temptable!str4 = !MainGroup
    temptable!str5 = !MainGroupDesca
    temptable!VAL1 = !Sumofvalue

    temptable!STR7 = TurnValue(retHeader(aHeader, 0, 4))
    temptable.Update
    sourcetable.MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
Main.Report1.ReportFileName = App.Path & "\Reports\charge1.rpt"
Main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.Report1.Action = 1
End Sub

Private Sub cmdClear_Click()
DefineText Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
FixRpImage Me
openCon con
If cBranch = "00" Then
    Set data1.Recordset = myCmd("FILE0_50", con, adTable)
Else
    Set data1.Recordset = myCmd("SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch), con)
End If
Set xbox.RowSource = data1
xbox.BoundColumn = "CODE"
xbox.ListField = "DESCA"

Set data2.Recordset = myCmd("select * from BRANCH where isbranch2 = 1 ", con)
Set XBRANCH.RowSource = data2
XBRANCH.BoundColumn = "CODE"
XBRANCH.ListField = "DESCA"
If cBranch <> "00" Then
   XBRANCH.BoundText = cBranch
   XBRANCH.Enabled = False
End If

End Sub
Private Function RetCharge(pCharge)
ChargeTable.Find "Code = " & MyParn(pCharge), , adSearchForward, adBookmarkFirst
If Not ChargeTable.EOF Then RetCharge = ChargeTable!DESCA
End Function
Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function MYVALID() As Boolean
If (Not IsDate(xdate1.text)) And Trim(xdate1.text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xdate2.text)) And Trim(xdate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub xbox_GotFocus()
myGotFocus xbox
End Sub
Private Sub xbox_LostFocus()
myLostFocus xbox
If Not xbox.MatchedWithList Then xbox.BoundText = ""
End Sub
Private Sub XBRANCH_GotFocus()
myGotFocus XBRANCH
End Sub
Private Sub XBRANCH_LostFocus()
myLostFocus XBRANCH
If Not XBRANCH.MatchedWithList Then XBRANCH.BoundText = ""
End Sub
