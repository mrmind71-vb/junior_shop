VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpCharge2 
   Caption         =   " ﬁ«—Ì— «·‰ﬁœÌ…"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
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
   ScaleHeight     =   4740
   ScaleWidth      =   5865
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
      Left            =   270
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   4050
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
      Left            =   1800
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   4050
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
      Left            =   3330
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   4050
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      Height          =   645
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3330
      Width           =   5550
      Begin VB.CheckBox chkNoBranch 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ ›—⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   1230
      End
      Begin VB.CheckBox XISCASH 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "‰ﬁœÏ ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2655
      Width           =   5550
      Begin VB.OptionButton XALL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«’Ê· ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   1185
      End
      Begin VB.OptionButton XALL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„’«—Ì› ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1365
      End
      Begin VB.OptionButton XALL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "ﬂ· «·„’«—Ì›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2490
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   180
      Width           =   5505
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
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   4
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
         Left            =   2700
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
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xbox 
         Height          =   330
         Left            =   720
         TabIndex        =   2
         Top             =   945
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xCharge 
         Height          =   330
         Left            =   720
         TabIndex        =   3
         Top             =   1305
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XBRANCH 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   2025
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
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
         Left            =   4245
         RightToLeft     =   -1  'True
         TabIndex        =   20
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
         Left            =   4245
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1620
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
         Left            =   4245
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1305
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
         Left            =   4245
         RightToLeft     =   -1  'True
         TabIndex        =   17
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
         Left            =   4245
         RightToLeft     =   -1  'True
         TabIndex        =   16
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
         Left            =   4245
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   945
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
      Left            =   495
      Top             =   1890
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
   Begin MSAdodcLib.Adodc DATA3 
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
Attribute VB_Name = "rpCharge2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New adodb.Connection
Private Sub CmdApply_Click()
If publicFlag = 2 Then
    doprint1
Else
    doprint2
End If
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub doprint1()
Dim sourcetable As New adodb.Recordset
Dim temptable As New adodb.Recordset
Dim aHeader(4)
If Not myValid Then Exit Sub


contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

Dim cString As String

cString = "Select acc_charge.Date," & _
          "acc_charge.DescA," & _
          "file0_50.desca AS FILE0_50DESCA," & _
          "FILE8_51.DESCA AS CHARGEDESCA," & _
          "Value," & _
          "acc_charge.DOC_NO," & _
          "BOX," & _
          "branch.desca  as branch_desca " & _
          " FROM   acc_charge " & _
          " left join branch ON branch.code = acc_charge.BRANCH" & _
          " LEFT OUTER JOIN FILE8_51 ON acc_charge.CHARGE = FILE8_51.CODE" & _
          " LEFT JOIN FILE0_50 ON acc_CHARGE.BOX = FILE0_50.CODE"

Dim cWhere As String
If lIsBranchStore Then
    cString = cString & turnFound(cString) & " FILE0_50.BRANCH = " & MyParn(cBranch)
End If
If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  acc_charge.BOX = " & MyParn(xbox.BoundText)
     aHeader(1) = "[ Œ“‰… " & xbox.text & "]"
End If

If XBRANCH.MatchedWithList Then
    cString = cString & turnFound(cString) & "  acc_charge.branch = " & MyParn(XBRANCH.BoundText)
     aHeader(1) = "[ ›—⁄ " & XBRANCH.text & "]"
End If

If IsDate(xdate1.text) Then
    cString = cString & turnFound(cString) & " acc_charge.date >= " & DateSq(xdate1.text)
    aHeader(0) = "[" & BetweenString(xdate1.text, xdate2.text) & "]"
End If

If XISCASH.Value <> 0 Then
    cString = cString & " AND BOX IS NOT NULL "
    aHeader(3) = " ‰ﬁœÏ ›ﬁÿ ›ﬁÿ "
End If

If IsDate(xdate2.text) Then
    cString = cString & turnFound(cString) & " acc_charge.date <= " & DateSq(xdate2.text)
    aHeader(0) = "[" & BetweenString(xdate1.text, xdate2.text) & "]"
End If
    
If XALL(1).Value Then
    cString = cString & turnFound(cString) & " File8_51.ISasset = 0 "
    aHeader(4) = " „’«—Ì› ›ﬁÿ "
End If

If XALL(2).Value Then
    cString = cString & turnFound(cString) & " File8_51.ISasset = 1 "
    aHeader(4) = " «’Ê· ›ﬁÿ "
End If
    

If xCharge.MatchedWithList Then
    cString = cString & turnFound(cString) & " acc_charge.charge = " & MyParn(xCharge.BoundText)
End If

If Trim(xdesca.text) <> "" Then
    cString = cString & turnFound(cString) & MyParnAnd(xdesca.text, "acc_charge.desca")
    aHeader(2) = "[" & "«·»Ì«‰ : " & xdesca.text & "]"
End If

If chkNoBranch.Value = 1 Then
    cString = cString & turnFound(cString) & "acc_charge.branch is null"
End If

Set sourcetable = cmd(cString, con).Execute

If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !doc_no
    temptable!str2 = TurnValue(!FILE0_50DESCA)
    'temptable!Str3 = TurnValue(ArbString(!desca & turn(!chargeDesca & "", turn(!desca, "-") & !chargeDesca) & ""))
    temptable!str3 = Mid(!DESCA, 1, 250)
    If Not IsNull(!chargeDesca) Then
        temptable!str3 = Mid(TurnValue(ArbString(temptable!str3 & turn(temptable!str3 & "", "-") & !chargeDesca & "")), 1, 250)
    End If
    temptable!VAL1 = !Value
    temptable!Date1 = !Date
    temptable!STR21 = "  ›’Ì·Ï „’—Ê› " & xCharge.text
    temptable!str22 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!str6 = !branch_desca
    temptable.Update
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
Set temptable = Nothing

contemp.BeginTrans
contemp.CommitTrans
Report1.ReportFileName = App.Path & "\Reports\charge2.rpt"
Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Report1.Action = 1
End Sub
Private Sub doprint2()
Dim sourcetable As New adodb.Recordset
Dim temptable As New adodb.Recordset
Dim aHeader(3)
If Not myValid Then Exit Sub

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select acc_INCOM.Date,acc_INCOM.DescA,file0_50.desca AS FILE0_50DESCA,file8_61.DESCA AS CHARGEDESCA,Value,acc_INCOM.DOC_NO,BOX" & _
          " From acc_INCOM  Left join file8_61 on acc_INCOM.charge = file8_61.code LEFT JOIN FILE0_50 ON acc_INCOM.BOX = FILE0_50.CODE WHERE acc_INCOM.DOC_NO IS NOT NULL "

If lIsBranchStore Then cString = cString & " AND FILE0_50.BRANCH = " & MyParn(cBranch)
If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  acc_INCOM.BOX = " & MyParn(xbox.BoundText)
     aHeader(1) = "[ Œ“‰… " & xbox.text & "]"
End If
If XBRANCH.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  acc_INCOM.branch = " & MyParn(XBRANCH.BoundText)
     aHeader(1) = "[ ›—⁄ " & XBRANCH.text & "]"
End If

If IsDate(xdate1.text) Then
    cString = cString & turnFound(cString) & " acc_INCOM.date >= " & DateSq(xdate1.text)
    aHeader(0) = "[" & BetweenString(xdate1.text, xdate2.text) & "]"
End If


If IsDate(xdate2.text) Then
    cString = cString & turnFound(cString) & " acc_INCOM.date <= " & DateSq(xdate2.text)
    aHeader(0) = "[" & BetweenString(xdate1.text, xdate2.text) & "]"
End If
    

If xCharge.BoundText <> "" Then
    cString = cString & turnFound(cString) & " acc_INCOM.charge = " & MyParn(xCharge.BoundText)
End If

If Trim(xdesca.text) <> "" Then
    cString = cString & turnFound(cString) & MyParnAnd(xdesca.text, "acc_INCOM.desca")
    aHeader(2) = "[" & "«·»Ì«‰ : " & xdesca.text & "]"
End If
If XISCASH.Value <> 0 Then
    cString = cString & " AND BOX IS NOT NULL "
    aHeader(3) = " ‰ﬁœÏ ›ﬁÿ ›ﬁÿ "
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
        temptable!str3 = TurnValue(!DESCA, "", Null) & " - " & TurnValue(!chargeDesca, Null, "")
    End If
    temptable!VAL1 = !Value
    temptable!Date1 = !Date
    temptable!STR21 = "  ›’Ì·Ï „’—Ê› " & xCharge.text
    temptable!str22 = TurnValue(retHeader(aHeader, 0, 3))
    temptable.Update
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
Set temptable = Nothing

contemp.BeginTrans
contemp.CommitTrans
Main.Report1.ReportFileName = App.Path & "\Reports\income2.rpt"
Main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.Report1.Action = 1
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
FixRpImage Me

openCon con

chkNoBranch.Visible = publicFlag = 2

Set sourcetable = New adodb.Recordset
Set temptable = New adodb.Recordset
data1.ConnectionString = strCon

data1.RecordSource = "Select * FROM " & IIf(publicFlag = 2, "file8_51", "file8_61") & " ORDER BY DESCA"
Set xCharge.RowSource = data1
xCharge.BoundColumn = "Code"
xCharge.ListField = "DescA"

data2.ConnectionString = strCon
If cBranch = "00" Or cBranch > "60" Then
    data2.RecordSource = "FILE0_50"
Else
    data2.RecordSource = "SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch)
End If
Set xbox.RowSource = data2
xbox.BoundColumn = "CODE"
xbox.ListField = "DESCA"


If cBranch = "00" Then
DATA3.ConnectionString = strCon
DATA3.RecordSource = "BRANCH"
Set XBRANCH.RowSource = DATA3
XBRANCH.BoundColumn = "CODE"
XBRANCH.ListField = "DESCA"
If cBranch <> "00" Then
   XBRANCH.BoundText = cBranch
   XBRANCH.Enabled = False
End If
End If
If publicFlag <> 2 Then Label4 = "«·«Ì—«œ :"
Frame2.Visible = (publicFlag = 2)
End Sub
Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function myValid() As Boolean
'If xCharge.BoundText = "" Then
'    MsgBox IIf(publicFlag = 2, "»Ì«‰ «·„’—Ê› „ÿ·Ê»", "»Ì«‰ «·«Ì—«œ „ÿ·Ê»")
'    Exit Function
'End If
If (Not IsDate(xdate1.text)) And Trim(xdate1.text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xdate2.text)) And Trim(xdate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
myValid = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xdesca_GotFocus()
myGotFocus xdesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xdesca
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
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
Private Sub XCHARGE_GotFocus()
myGotFocus xCharge
End Sub
Private Sub XCHARGE_LostFocus()
myLostFocus xCharge
If Not xCharge.MatchedWithList Then xCharge.BoundText = ""
End Sub
Private Sub XBRANCH_GotFocus()
myGotFocus XBRANCH
End Sub
Private Sub XBRANCH_LostFocus()
myLostFocus XBRANCH
If Not XBRANCH.MatchedWithList Then XBRANCH.BoundText = ""
End Sub
