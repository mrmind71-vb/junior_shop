VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpcharge_BR2 
   Caption         =   " ﬁ«—Ì— «·‰ﬁœÌ…"
   ClientHeight    =   3630
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
   ScaleHeight     =   3630
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
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
      Left            =   90
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   2925
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
      Left            =   1620
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   2925
      Width           =   1500
   End
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
      Left            =   3150
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   2925
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
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
         TabIndex        =   16
         Top             =   2430
         Width           =   1320
      End
      Begin VB.TextBox xdesca 
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
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1665
         Width           =   3345
      End
      Begin VB.TextBox xdate1 
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
         Left            =   2295
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xbox 
         Height          =   315
         Left            =   315
         TabIndex        =   2
         Top             =   945
         Width           =   3345
         _ExtentX        =   5900
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
      Begin MSDataListLib.DataCombo xCharge 
         Height          =   315
         Left            =   315
         TabIndex        =   3
         Top             =   1305
         Width           =   3345
         _ExtentX        =   5900
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
         Left            =   315
         TabIndex        =   11
         Top             =   2025
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSAdodcLib.Adodc DATA3 
         Height          =   330
         Left            =   -1800
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
      Begin VB.Label Label6 
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
         TabIndex        =   12
         Top             =   2025
         Width           =   405
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "»Ì«‰"
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
         TabIndex        =   10
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label Label4 
         Caption         =   "«·„’—Ê›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1350
         Width           =   1005
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
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
         TabIndex        =   6
         Top             =   585
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   5
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
      Left            =   3330
      Top             =   1800
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
   Begin MSAdodcLib.Adodc data2 
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
End
Attribute VB_Name = "rpcharge_BR2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
    doprint1
End Sub

Private Sub cmdClear_Click()
DefineText Me
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub doprint1()
Dim sourcetable As New ADODB.Recordset
Dim temptable As New ADODB.Recordset
Dim aHeader(4)
If Not MYVALID Then Exit Sub

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select acc_charge_BR.Date,acc_charge_BR.DescA,file0_50.desca AS FILE0_50DESCA,FILE8_51.DESCA AS CHARGEDESCA,Value,acc_charge_BR.DOC_NO,BOX , branch.desca  as branch_desca " & _
          " FROM   branch INNER JOIN acc_charge_BR ON branch.code = acc_charge_BR.BRANCH LEFT OUTER JOIN FILE8_51 ON acc_charge_BR.CHARGE = FILE8_51.CODE LEFT OUTER JOIN FILE0_50 ON acc_charge_BR.BOX = FILE0_50.CODE WHERE VALUE <> 0  "

If lIsBranchStore Then
    cString = cString & " AND FILE0_50.BRANCH = " & MyParn(cBranch)
End If

If XBOX.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  acc_charge_BR.BOX = " & MyParn(XBOX.BoundText)
     aHeader(1) = "[ Œ“‰… " & XBOX.text & "]"
End If

If XBRANCH.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  acc_charge_BR.branch = " & MyParn(XBRANCH.BoundText)
     aHeader(1) = "[ ›—⁄ " & XBRANCH.text & "]"
End If

If IsDate(xDate1.text) Then
    cString = cString & turnFound(cString) & " acc_charge_BR.date >= " & DateSq(xDate1.text)
    aHeader(0) = "[" & BetweenString(xDate1.text, xDate2.text) & "]"
End If

If IsDate(xDate2.text) Then
    cString = cString & turnFound(cString) & " acc_charge_BR.date <= " & DateSq(xDate2.text)
    aHeader(0) = "[" & BetweenString(xDate1.text, xDate2.text) & "]"
End If
    
If XCHARGE.BoundText <> "" Then
    cString = cString & turnFound(cString) & " acc_charge_BR.charge = " & MyParn(XCHARGE.BoundText)
End If

If Trim(xdesca.text) <> "" Then
    cString = cString & turnFound(cString) & MyParnAnd(xdesca.text, "acc_charge_BR.desca")
    aHeader(2) = "[" & "«·»Ì«‰ : " & xdesca.text & "]"
End If

If chkCash.Value = 1 Then
    cString = cString & " AND acc_charge_BR.BOX IS NOT NULL"
    aHeader(3) = "[‰ﬁœÌ ›ﬁÿ]"
End If

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    If Len(!doc_no) > 10 Then
        temptable!str1 = RetZero(DelZero(Mid(!doc_no, 2)), 6)
    Else
        temptable!str1 = !doc_no
    End If
    temptable!str2 = TurnValue(!FILE0_50DESCA)
    'temptable!Str3 = TurnValue(ArbString(!desca & turn(!chargeDesca & "", turn(!desca, "-") & !chargeDesca) & ""))
    temptable!str3 = !DESCA
    
    If Not IsNull(!chargeDesca) Then
        temptable!str3 = TurnValue(ArbString(temptable!str3 & turn(temptable!str3 & "", "-") & !chargeDesca & ""))
    End If
    
    temptable!VAL1 = !Value
    temptable!Date1 = !Date
    temptable!STR21 = "  ›’Ì·Ï „’—Ê› " & XCHARGE.text
    temptable!str22 = TurnValue(retHeader(aHeader, 0, 4))
    temptable!str6 = !branch_desca
    temptable.Update
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
Set temptable = Nothing

contemp.BeginTrans
contemp.CommitTrans

REPORT1.ReportFileName = App.Path & "\Reports\charge2.rpt"
REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
REPORT1.Action = 1
End Sub
Private Sub doprint2()
Dim sourcetable As New ADODB.Recordset
Dim temptable As New ADODB.Recordset
Dim aHeader(3)
If Not MYVALID Then Exit Sub

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select acc_INCOM.Date,acc_INCOM.DescA,file0_50.desca AS FILE0_50DESCA,file8_61.DESCA AS CHARGEDESCA,Value,acc_INCOM.DOC_NO,BOX" & _
          " From acc_INCOM  Left join file8_61 on acc_INCOM.charge = file8_61.code LEFT JOIN FILE0_50 ON acc_INCOM.BOX = FILE0_50.CODE WHERE acc_INCOM.DOC_NO IS NOT NULL "

If lIsBranchStore Then cString = cString & " AND FILE0_50.BRANCH = " & MyParn(cBranch)
If XBOX.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  acc_INCOM.BOX = " & MyParn(XBOX.BoundText)
     aHeader(1) = "[ Œ“‰… " & XBOX.text & "]"
End If
If XBRANCH.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  acc_INCOM.branch = " & MyParn(XBRANCH.BoundText)
     aHeader(1) = "[ ›—⁄ " & XBRANCH.text & "]"
End If

If IsDate(xDate1.text) Then
    cString = cString & turnFound(cString) & " acc_INCOM.date >= " & DateSq(xDate1.text)
    aHeader(0) = "[" & BetweenString(xDate1.text, xDate2.text) & "]"
End If


If IsDate(xDate2.text) Then
    cString = cString & turnFound(cString) & " acc_INCOM.date <= " & DateSq(xDate2.text)
    aHeader(0) = "[" & BetweenString(xDate1.text, xDate2.text) & "]"
End If
    

If XCHARGE.BoundText <> "" Then
    cString = cString & turnFound(cString) & " acc_INCOM.charge = " & MyParn(XCHARGE.BoundText)
End If

If Trim(xdesca.text) <> "" Then
    cString = cString & turnFound(cString) & MyParnAnd(xdesca.text, "acc_INCOM.desca")
    aHeader(2) = "[" & "«·»Ì«‰ : " & xdesca.text & "]"
End If


sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !doc_no
    temptable!str2 = TurnValue(!FILE0_50DESCA)
    temptable!str3 = TurnValue(!DESCA, "", Null)
    If Not IsNull(!chargeDesca) Then
        temptable!str3 = TurnValue(ArbString(temptable!str3 & turn(temptable!str3 & "", "-") & !chargeDesca & ""))
    End If
    temptable!VAL1 = !Value
    temptable!Date1 = !Date
    temptable!STR21 = "  ›’Ì·Ï „’—Ê› " & XCHARGE.text
    temptable!str22 = TurnValue(retHeader(aHeader, 0, 3))
    temptable.Update
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
Set temptable = Nothing

contemp.BeginTrans
contemp.CommitTrans
REPORT1.ReportFileName = App.Path & "\Reports\income2.rpt"
REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
REPORT1.Action = 1
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
FixRpImage Me
openCon con
Set sourcetable = New ADODB.Recordset
Set temptable = New ADODB.Recordset

Set data1.Recordset = mycmd("Select * FROM " & IIf(publicFlag = 2, "file8_51", "file8_61") & " ORDER BY DESCA", con)
Set XCHARGE.RowSource = data1
XCHARGE.BoundColumn = "Code"
XCHARGE.ListField = "DescA"

DATA2.ConnectionString = strCon
If cBranch = "00" Then
    Set DATA2.Recordset = mycmd("FILE0_50", con, adTable)
Else
    Set DATA2.Recordset = mycmd("SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch), con)
End If
Set XBOX.RowSource = DATA2
XBOX.BoundColumn = "CODE"
XBOX.ListField = "DESCA"

Set data3.Recordset = mycmd("select * from BRANCH where isbranch2 = 1 ", con)
Set XBRANCH.RowSource = data3
XBRANCH.BoundColumn = "CODE"
XBRANCH.ListField = "DESCA"
If cBranch <> "00" Then
   XBRANCH.BoundText = cBranch
   XBRANCH.Enabled = False
End If
End Sub
Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function MYVALID() As Boolean
'If xCharge.BoundText = "" Then
'    MsgBox IIf(publicFlag = 2, "»Ì«‰ «·„’—Ê› „ÿ·Ê»", "»Ì«‰ «·«Ì—«œ „ÿ·Ê»")
'    Exit Function
'End If
If (Not IsDate(xDate1.text)) And Trim(xDate1.text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xDate2.text)) And Trim(xDate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
Private Sub xdesca_GotFocus()
myGotFocus xdesca
End Sub
Private Sub xdesca_LostFocus()
myLostFocus xdesca
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub XBOX_GotFocus()
myGotFocus XBOX
End Sub
Private Sub XBOX_LostFocus()
myLostFocus XBOX
If Not XBOX.MatchedWithList Then XBOX.BoundText = ""
End Sub
Private Sub XCHARGE_GotFocus()
myGotFocus XCHARGE
End Sub
Private Sub XCHARGE_LostFocus()
myLostFocus XCHARGE
If Not XCHARGE.MatchedWithList Then XCHARGE.BoundText = ""
End Sub
Private Sub XBRANCH_GotFocus()
myGotFocus XBRANCH
End Sub
Private Sub XBRANCH_LostFocus()
myLostFocus XBRANCH
If Not XBRANCH.MatchedWithList Then XBRANCH.BoundText = ""
End Sub


