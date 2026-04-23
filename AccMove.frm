VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form AccMove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«” «– ⁄«„ Õ”«»« "
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5850
      RightToLeft     =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   630
      Width           =   1365
   End
   Begin VB.TextBox xlevel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   4725
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   90
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.CommandButton Cmd_print 
      Caption         =   "ÿ»«⁄… «” «– «·Õ”«»"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9450
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   225
      Width           =   1860
   End
   Begin VB.CommandButton Cmdgo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "«” Ã«»…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9450
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   675
      Width           =   1860
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   8010
      Left            =   540
      TabIndex        =   5
      Top             =   1125
      Width           =   19800
      _cx             =   34925
      _cy             =   14129
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
      AutoResize      =   -1  'True
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
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.TextBox xType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3525
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   11340
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Width           =   8985
      Begin VB.TextBox date2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   1410
      End
      Begin VB.TextBox Date1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1815
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5895
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xCent 
         Height          =   315
         Left            =   135
         TabIndex        =   13
         Top             =   585
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·⁄„·…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   630
         Width           =   960
      End
      Begin VB.Label xBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   645
         Width           =   675
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·Õ”«»"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7755
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   300
         Width           =   975
      End
   End
   Begin Threed.SSCheck XPRINTSE 
      Height          =   315
      Left            =   495
      TabIndex        =   12
      Top             =   675
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      _Version        =   196610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ÿ»«⁄… «·„Õœœ ›ﬁÿ"
      Alignment       =   1
   End
   Begin Threed.SSCommand cmd_excel 
      Height          =   870
      Left            =   7290
      TabIndex        =   16
      Top             =   225
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   1535
      _Version        =   196610
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
      Picture         =   "AccMove.frx":0000
      Caption         =   " ÕÊÌ· «ﬂ”Ì· "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   17
      Top             =   10845
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   -630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
Attribute VB_Name = "AccMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADOdb.Connection
Dim AccTable As ADOdb.Recordset, nFirst As Double
Sub fillgrd()
nFirst = 0
Dim GrdTable As New ADOdb.Recordset
If IsDate(Date1.text) Then
    cString = "Select Sum([VAL1])AS TVAL1 ,Sum([VAL2]) AS TVAL2 FROM ACC_MOVE LEFT JOIN ACC1_10 ON ACC_MOVE.ACC = ACC1_10.CODE " & _
              " WHERE  Date < " & DateSq(Date1.text)
    If xcode.text <> "" Then
        cString = cString & " AND ACC = " & MyParn(xcode.text)
    End If
    GrdTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    If Not (GrdTable.BOF And GrdTable.EOF) Then nFirst = TurnValue(GrdTable!TVAL1, Null, 0) - TurnValue(GrdTable!TVAL2, Null, 0)
    GrdTable.Close
End If

cString = "Select ACC_MOVE.*     FROM ACC_MOVE Where ACC IS NOT NULL "
cString = cString & " AND ACC = " & MyParn(xcode.text)

If IsDate(Date1.text) Then cString = cString & " and [date] >= " & DateSq(Date1.text)
If IsDate(date2.text) Then cString = cString & " and [date] <= " & DateSq(date2.text)
cString = cString & " Order by DATE , VAL1 "

GrdTable.Open cString, con, adOpenForwardOnly, adLockReadOnly, adCmdText

Dim nPrevious As Double
Dim nValue As Double
nPrevious = 0
grid1.Rows = 1
With grid1
If nFirst <> 0 Then
    grid1.AddItem ""
    grid1.TextMatrix(.Rows - 1, 0) = "—’Ìœ «Ê· «·„œ…"
    grid1.TextMatrix(.Rows - 1, 1) = Date1.text
    grid1.TextMatrix(.Rows - 1, IIf(nFirst > 0, 3, 4)) = Format(nFirst, "Fixed")
    grid1.TextMatrix(.Rows - 1, 5) = Format(nFirst, "Fixed")
    nPrevious = Format(nFirst, "Fixed")
End If

Do Until GrdTable.EOF
    grid1.AddItem ""
    grid1.TextMatrix(.Rows - 1, 0) = IIf(IsDate(GrdTable!Date), Format(GrdTable!Date, "d-m-yyyy"), "")
    grid1.TextMatrix(.Rows - 1, 1) = GrdTable!DESCA & ""
    grid1.TextMatrix(.Rows - 1, 2) = GrdTable!DOC_ID & ""
    grid1.TextMatrix(.Rows - 1, 3) = mRound(GrdTable!VAL1)
    grid1.TextMatrix(.Rows - 1, 4) = mRound(GrdTable!val2)
    nValue = Format(nPrevious + Val(grid1.TextMatrix(.Rows - 1, 3)) - Val(grid1.TextMatrix(.Rows - 1, 4)), "Fixed")
         
    
    grid1.TextMatrix(.Rows - 1, 5) = Format(nValue, "Fixed")
    nCol3 = nCol3 + Val(.TextMatrix(.Rows - 1, 3))
    nCol4 = nCol4 + Val(.TextMatrix(.Rows - 1, 4))
    nPrevious = nValue
    GrdTable.MoveNext
    I = I + 1
Loop
End With
xBal.Caption = Format(nValue, "Fixed")
End Sub
Sub myProc()
    xcode.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    xdesca.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
    Unload Search3
End Sub

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
fillgrd
End Sub

Private Sub Form_Load()
con.CursorLocation = adUseClient
con.Open strCon
Set AccTable = New ADOdb.Recordset
AccTable.Open "ACC1_10", con, adOpenDynamic, adLockOptimistic, adCmdTableDirect

'data1.ConnectionString = strCon
'data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
'Set xCent.RowSource = data1
'xCent.ListField = "Desca"
'xCent.BoundColumn = "Code"

grid1.FormatString = " «—ÌŒ |" & "«·»Ì«‰ |" & "—ﬁ„ «·„” ‰œ |" & "„œÌ‰|" & "œ«∆‰|" & "—’Ìœ|" & "ÿ»«⁄…|"
grid1.Cols = 7
grid1.ColWidth(0) = 1400
grid1.ColWidth(1) = 9500
grid1.ColWidth(2) = 1600
grid1.ColWidth(3) = 1600
grid1.ColWidth(4) = 1600
grid1.ColWidth(5) = 1600
grid1.ColWidth(6) = 800
For I = 0 To grid1.Cols - 1
    grid1.ColAlignment(I) = flexAlignRightCenter
Next
grid1.Rows = 1
grid1.FixedCols = 0
grid1.FixedRows = 1
grid1.ColDataType(6) = flexDTBoolean
Date1.text = "1-1-" & Year(Date)
date2.text = Format(Date, "DD-MM-YYYY")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_EnterCell()
    With grid1
        If .col = 6 Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        AcntLookup_Sub Me, Search3
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub xcode_LostFocus()
If xcode.text = "" Then Exit Sub
AccTable.Find " code = " & MyParn(xcode.text), , adSearchForward, adBookmarkFirst
If Not AccTable.EOF Then
    xdesca.Caption = AccTable!DESCA & ""
    cmdGo.Enabled = Not AccTable.EOF
End If
End Sub
Private Sub CMD_PRINT_Click()
    Dim temptable As New ADOdb.Recordset
    Dim sourcetable As New ADOdb.Recordset
    contemp.Execute "DELETE * FROM TEMP"
    temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
                        
    With grid1
    For I = 1 To .Rows - 1
        If XPRINTSE.Value <> -1 Or TurnValue(.TextMatrix(I, 6), "", False) Then
            temptable.AddNew
            temptable!STR19 = cComp
            
            temptable!str8 = "ÿ»«⁄… «” «– ⁄«„ ··Õ”«» —ﬁ„ " & xcode.text & "   " & xdesca.Caption
            temptable!STR7 = " „‰  «—ÌŒ " & Date1.text & "  ≈·Ï  «—ÌŒ " & date2.text
            
            temptable!str1 = TurnValue(.TextMatrix(I, 2), "", Null)
            temptable!str3 = TurnValue(.TextMatrix(I, 1), "", Null)
            temptable!Date1 = IIf(IsDate(.TextMatrix(I, 0)), .TextMatrix(I, 0), Null)
            temptable!VAL1 = Val(.TextMatrix(I, 3))
            temptable!val2 = Val(.TextMatrix(I, 4))
            temptable!val3 = Val(.TextMatrix(I, 5))
            temptable!STR30 = TurnValue(cusername)
            
            temptable.Update
        End If
    Next I
    End With

    contemp.BeginTrans
    contemp.CommitTrans
    If XPRINTSE.Value Then
        REPORT1.ReportFileName = App.Path & "\Reports\R_ACC30.rpt"
    Else
        REPORT1.ReportFileName = App.Path & "\Reports\R_ACC3.rpt"
    End If
    REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    REPORT1.Action = 1
End Sub

