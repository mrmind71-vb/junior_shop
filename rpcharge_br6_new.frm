VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpcharge_br6 
   Caption         =   "≈Ã„«·Ï „Êﬁ› Õ”«»  ÊﬂÌ·- ›—⁄ ·› —…"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
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
   ScaleHeight     =   2430
   ScaleWidth      =   6960
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
      Left            =   3285
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   1845
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
      Left            =   1755
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   1845
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
      Left            =   225
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1845
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   6495
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
         Left            =   4170
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   270
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
         Left            =   4170
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo XACC 
         Height          =   315
         Left            =   1710
         TabIndex        =   5
         Top             =   990
         Width           =   3825
         _ExtentX        =   6747
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
         Left            =   -1935
         Top             =   45
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
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   1710
         TabIndex        =   7
         Top             =   1350
         Width           =   3825
         _ExtentX        =   6747
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄ "
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
         Left            =   5625
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1350
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·Õ”«»"
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
         Left            =   5625
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1035
         Width           =   585
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
         Left            =   5625
         RightToLeft     =   -1  'True
         TabIndex        =   4
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
         Left            =   5625
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   675
         Width           =   690
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   7020
      Top             =   135
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
      Left            =   6705
      Top             =   945
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
Attribute VB_Name = "rpcharge_br6"
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
Dim temptable2 As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset, nFBal As Double, cAcc As String
ReDim aHeader(3)
contemp.Execute "DELETE * FROM TEMP"
contemp.Execute "DELETE * FROM TEMP2"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
temptable2.Open "temp2", contemp, adOpenStatic, adLockOptimistic, adCmdTable

'If XACC.BoundText = "" Then Exit Sub
If XACC.MatchedWithList Then
    cAcc = "6" & XACC.BoundText
Else
    cAcc = ""
End If

If IsDate(xDate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
    aHeader(2) = BetweenString(Format(xDate1.text, "d-m-yyyy"), Format(xDate2.text, "d-m-yyyy"))
End If

If IsDate(xDate2.text) Then
    aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xDate2.text))
    aHeader(2) = BetweenString(Format(xDate1.text, "d-m-yyyy"), Format(xDate2.text, "d-m-yyyy"))
End If

          
If XACC.MatchedWithList Then
    aPrm = AddFlag(aPrm, "ACC", cAcc)
    aHeader(0) = "≈Ã„«·Ï Õ”«» " & XACC.text
End If

If xStore.MatchedWithList Then
    aPrm = AddFlag(aPrm, "STORE", xStore.BoundText)
    aHeader(1) = "›—⁄" & xStore.text
End If

          
Set sourcetable = myCmd("dbo.sp_acc_move_store", con, adStoredProc, aPrm)
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    If IsNull(!charge2) Then
        If Trim(!FLAG_DESCA) = "„»Ì⁄«  „’«—Ì›" Then
            temptable!str2 = "„” ·“„«  »Ì⁄"
            temptable!str1 = "„’«—Ì›  ÊﬂÌ·"
        Else
            temptable!str2 = TurnValue(!FLAG_DESCA)
            temptable!str1 = TurnValue(!SUBGR)
        End If
    Else
        temptable!str2 = TurnValue(GetDesca("SELECT FILE8_52.DESCA FROM FILE8_51 INNER JOIN FILE8_52 ON FILE8_51.[MAINGROUP] = FILE8_52.CODE  WHERE FILE8_51.CODE = " & MyParn(!charge2), con))
        temptable!str1 = "„’«—Ì›  ÊﬂÌ·"
    End If
    temptable!VAL1 = !value1
    temptable!val2 = !Value2
    temptable!str6 = TurnValue(retHeader(aHeader, 0, 1))
    temptable!STR7 = TurnValue(retHeader(aHeader, 1, 1))
    If temptable!str1 <> "„’«—Ì›  ÊﬂÌ·" Then
        temptable!str10 = "0"
    End If
    temptable.Update
    sourcetable.MoveNext
Loop
End With

'contemp.Execute " INSERT INTO TEMP ( VAL10, STR1, STR2, VAL1, VAL2, STR5, STR6, STR7, STR10 ) SELECT TEMP2.VAL10, TEMP2.STR1, TEMP2.STR2, SUM(TEMP2.VAL1), SUM(TEMP2.VAL2), TEMP2.STR5, TEMP2.STR6, TEMP2.STR7, TEMP2.STR10 FROM TEMP2 GROUP BY TEMP2.VAL10, TEMP2.STR1, TEMP2.STR2, TEMP2.STR5, TEMP2.STR6 ,TEMP2.STR7 ,TEMP2.STR10"

contemp.BeginTrans
contemp.CommitTrans
REPORT1.ReportFileName = App.Path & "\Reports\charge6.rpt"
REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
REPORT1.Action = 1
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

Set DATA2.Recordset = myCmd("SELECT * FROM acc0_10", con)
Set XACC.RowSource = DATA2
XACC.BoundColumn = "CODE"
XACC.ListField = "DESCA"

Set DATA3.Recordset = myCmd("SELECT * FROM FILE0_40 WHERE ISONEST2 = 1 ORDER BY CODE  ", con)
Set xStore.RowSource = DATA3
xStore.BoundColumn = "CODE"
xStore.ListField = "DESCA"
End Sub
Function myValid() As Boolean
If (Not IsDate(xDate1.text)) And Trim(xDate1.text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xDate2.text)) And Trim(xDate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
myValid = True
End Function
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set rpcharge_br6 = Nothing
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub XACC_GotFocus()
myGotFocus XACC
End Sub
Private Sub XACC_LostFocus()
myLostFocus XACC
If Not XACC.MatchedWithList Then XACC.BoundText = ""
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
