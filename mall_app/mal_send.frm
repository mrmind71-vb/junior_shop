VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Mall_Sendfrm 
   Caption         =   "»Ì«‰«  «· ÿ»Ìﬁ ⁄·Ì «·„Êﬁ⁄"
   ClientHeight    =   1785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1785
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   945
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   90
      Width           =   10455
      Begin MSDataListLib.DataCombo xBranch 
         Height          =   330
         Left            =   6120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   225
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   5640
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄"
         Height          =   270
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   405
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   1590
      Visible         =   0   'False
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   4725
      TabIndex        =   0
      Top             =   810
      Width           =   5820
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "mal_send.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
      Begin Threed.SSCommand cmdSendMall 
         Height          =   510
         Left            =   3735
         TabIndex        =   1
         Top             =   180
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
         Enabled         =   0   'False
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
         Caption         =   "√—”«· «·›Ê« Ì—"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMallFix 
         Height          =   510
         Left            =   1620
         TabIndex        =   3
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
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
         Caption         =   "÷»ÿ »Ì«‰«  «·„Êﬁ⁄"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "Mall_Sendfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bSend As Boolean
Dim con As New ADODB.Connection
Dim bAct As Boolean
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdFixMobileSetting_Click()
'invoice_send_settingfrm.Show 1
End Sub
Private Sub cmdItemsUpdate_Click()
'sendinvoicesBranch
End Sub

Private Sub cmdMallFix_Click()
mall_settingfrm.Show 1
End Sub

Private Sub cmdSendMall_Click()
Dim sError As String
Dim nCount As Long
SendInvoicesMall con, XBRANCH.BoundText, sError, nCount, Me
If sError <> "" Then
    MsgBox sError
Else
    MsgBox " „ «—”«· " & nCount
End If
End Sub

Private Sub Command1_Click()
Dim sError As String
Dim sToken As String

sToken = getTokenMall(sError)
If sError <> "" Then
    MsgBox (sError)
Else
    Dim json As New ChilkatJsonArray
    Set json = GetJsonArray(rsMall!baseUrl & rsMall!invoiceUrl, "Bearer " & sToken, sError)
    
    MsgBox json.Size
    MsgBox "done"
End If

End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    If bSend Then
        If sendinvoicesBranch Then Unload Me
    End If
End If
End Sub
Private Sub Form_Load()
openCon con

Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")


Set data1.Recordset = cmd("SELECT * FROM BRANCH WHERE IS_MALL = 1 ", con).Execute
Set XBRANCH.RowSource = data1
XBRANCH.ListField = "Desca"
XBRANCH.BoundColumn = "Code"

End Sub
Public Function AddNote(sNote As String, Optional color As Long = 0)
grid1.AddItem sNote
grid1.Select grid1.Rows - 1, 0
grid1.ShowCell grid1.Rows - 1, 0
If color > 0 Then
    grid1.Cell(flexcpBackColor, grid1.Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = color
End If
DoEvents
End Function
Private Sub FixGrdMissing()
With grdMissing
    .TextMatrix(0, 0) = "«·ÌÊ„"
    .TextMatrix(0, 1) = "«·ÌÊ„"
    .TextMatrix(0, 2) = "«·≈Ã„«·Ì"
    .ColWidth(0) = 1200
    .ColWidth(0) = 1300
    .ColWidth(0) = 1200
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter

    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#0.00", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ì"

End With
End Sub
Private Function SendDay(pDate As String) As Boolean
Dim JArray As New ChilkatJsonArray
Dim sResponse As String
Me.MousePointer = vbHourglass

SendDay = SendInvoices(con, Me, myFormat(pDate), JArray, sResponse)
Me.MousePointer = vbNormal
End Function
Private Sub Form_Unload(Cancel As Integer)
Set Mall_Sendfrm = Nothing
End Sub

Private Sub xbranch_Click(Area As Integer)
cmdSendMall.Enabled = XBRANCH.MatchedWithList
If XBRANCH.MatchedWithList Then
    Dim aPrm As Variant
    aPrm = AddFlag(Empty, "Branch", XBRANCH.BoundText)
    aPrm = AddFlag(aPrm, "date_begin", myFormat_sp(rsMall!date_begin))
    Dim loctable As New ADODB.Recordset
    Set loctable = mycmd("[dbo].[sp_invoices_mall]", con, adStoredProc, aPrm)
    lblCount.Caption = "⁄œœ «·”Ã·«  «·€Ì— „—Õ·… : " & loctable.RecordCount
    Set loctable = Nothing
End If
End Sub
