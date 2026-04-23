VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PriceSalesFix 
   Caption         =   "÷»ÿ „ Ê”ÿ ”⁄— «·„” Â·ﬂ"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4320
      Width           =   5280
      Begin Threed.SSCommand cmdYear 
         Height          =   510
         Left            =   2115
         TabIndex        =   4
         Top             =   135
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
         BackColor       =   16777215
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
         Caption         =   "«Œ Ì«— «·”‰…"
         TagVariant      =   "«Œ Ì«— «·”‰…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdApply 
         Height          =   510
         Left            =   45
         TabIndex        =   5
         Top             =   135
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   192
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
         Caption         =   "÷»ÿ «·«”⁄«—"
         ButtonStyle     =   3
      End
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   45
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PriceSalesFix.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   5085
      UseMaskColor    =   -1  'True
      Width           =   5280
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   5655
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4200
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   5190
      _cx             =   9155
      _cy             =   7408
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
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
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
Attribute VB_Name = "PriceSalesFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty

Private Sub cmdApply_Click()
Me.MousePointer = vbHourglass
grid1.Rows = 0
Dim sErroNumber As Variant
Dim sErrorMsg As Variant
Dim nUpdate As Long
Dim nDelete As Long
Dim nInsert As Long

grid1.Rows = 0
AddNote "Ì „ «÷«›… «”⁄«— ”‰… " & cmdYear.Tag
If create_Item_price(cmdYear.Tag, con, sErrorNumber, sErrorMsg, nUpdate, nDelete, nInsert) Then
    AddNote " „ Õ–› " & nInsert & " ”Ã· «”⁄«—"
    AddNote " „ «÷«›… " & nInsert & " ”Ã· «”⁄«—"
Else
    If Not IsNull(sErrorNumber) Then
        AddNote "Error Number : " & sErrorNumber, vbRed
    End If
    If Not IsNull(sErrorMsg) Then
        AddNote sErrorMsg, vbRed
    End If
End If
Me.MousePointer = vbNormal
End Sub


Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdYear_Click()
'YearsLookup Me, oSearchYear, , cmdYear.Tag <> ""
YearsLookup Me, oSearchYear, , , "CODE <> " & cmdYear.Tag
End Sub

Private Sub Form_Load()
openCon con
cmdYear.Tag = Format(Date, "YYYY")
cmdYear.Caption = Format(Date, "YYYY")
grid1.ColWidth(0) = grid1.Width - 400
grid1.ColAlignment(0) = flexAlignRightCenter
End Sub
Public Function AddNote(sNote As Variant, Optional color As Long = 0)
grid1.AddItem sNote
If color > 0 Then
    grid1.Cell(flexcpBackColor, grid1.Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = color
End If
DoEvents
End Function
Sub myProc()
ActiveControl.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
ActiveControl.Caption = IIf(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0) = "", ActiveControl.TagVariant, oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
oSearchYear.Hide
End Sub
Function create_Item_price(pYear As String, con As ADODB.Connection, Optional ByRef pErrorNumber As Variant, Optional ByRef pErrorMsg As Variant, Optional ByRef pUpdate As Long = 0, Optional ByRef pDelete As Long = 0, Optional ByRef pInsert As Long = 0) As Boolean
Dim cmd As New ADODB.Command
aPrm = AddFlag(Empty, "YEAR", TurnValue(pYear))
Set cmd = mycmdEx("dbo.sp_items_price", con, aPrm)

If Not IsNull(cmd.Parameters("@ERROR_NUMBER").Value) Then
     pErrorNumber = cmd.Parameters("@ERROR_NUMBER").Value
     pErrorMsg = cmd.Parameters("@ERROR_MSG").Value
     Exit Function
End If

pUpdate = mRound(cmd.Parameters("@COUNT_UPDATE"))
pDelete = mRound(cmd.Parameters("@COUNT_DELETE"))
pInsert = mRound(cmd.Parameters("@COUNT_INSERT"))
create_Item_price = True
End Function

