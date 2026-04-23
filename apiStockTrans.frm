VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form apiStockTransfrm 
   Caption         =   "»Ì«‰«  «· ÿ»Ìﬁ ⁄·Ì «·„Êﬁ⁄"
   ClientHeight    =   4680
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1305
      Visible         =   0   'False
      Width           =   8070
      Begin Threed.SSCommand cmdUsers 
         Height          =   510
         Left            =   4140
         TabIndex        =   12
         Top             =   135
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
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
         Caption         =   "”Õ» »Ì«‰«  «·„” Œœ„Ì‰"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame Frame3 
      Height          =   825
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3015
      Visible         =   0   'False
      Width           =   8115
      Begin Threed.SSCommand cmdOrders 
         Height          =   555
         Left            =   4095
         TabIndex        =   7
         Top             =   180
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   979
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
         Caption         =   "”Õ» ÿ·»Ì«  «·„⁄«—÷"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOrdersSend 
         Height          =   555
         Left            =   90
         TabIndex        =   13
         Top             =   180
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   979
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
         Caption         =   "«—”«· ÿ·»Ì«  «·„⁄«—÷"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1230
      Left            =   135
      TabIndex        =   4
      Top             =   45
      Width           =   4020
      Begin Threed.SSCommand cmdItemsUpdate 
         Height          =   960
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   1693
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
         Caption         =   "‰ﬁ· ﬂ· «·«’‰«›"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdUpdateBranch 
         Height          =   420
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   1215
         Visible         =   0   'False
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   741
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
         Caption         =   "Õ–› Ê ÕœÌÀ ﬂ· »Ì«‰«  «·›—Ê⁄"
         ButtonStyle     =   3
      End
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   2205
      MaskColor       =   &H00FFFFFF&
      Picture         =   "apiStockTrans.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   3870
      UseMaskColor    =   -1  'True
      Width           =   6000
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   4185
      TabIndex        =   0
      Top             =   45
      Width           =   4020
      Begin Threed.SSCommand cmdItemsUpdate 
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   741
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
         Caption         =   "‰ﬁ· «·«’‰«› «·„ÕœÀ… ›ﬁÿ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdUpdateBranch 
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   675
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   741
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
         Caption         =   " ÕœÌÀ »Ì«‰«  «·›—Ê⁄"
         ButtonStyle     =   3
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   4485
      Visible         =   0   'False
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4200
      Left            =   8190
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   135
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
   Begin Threed.SSCommand apiStockTransfrm 
      Height          =   510
      Left            =   45
      TabIndex        =   10
      Top             =   3870
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
Attribute VB_Name = "apiStockTransfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub apiStockTransfrm_Click()
ApiStockSettingfrm.Show 1
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdFixMobileSetting_Click()
ApiStockSettingfrm.Show 1
End Sub
Private Sub cmdItemsUpdate_Click(Index As Integer)

If MsgBox("«—”«· »Ì«‰«  «·«’‰«› «·Ì  ÿ»Ìﬁ «·„Õ„Ê· ?", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub

cmdItemsUpdate(Index).Enabled = False
Me.MousePointer = vbHourglass
grid1.Rows = 0

If rsApiStock Is Nothing Then Exit Sub
If rsApiStock.EOF Then Exit Sub

UpdateItemsStockApi con, Me

Me.MousePointer = vbNormal
cmdItemsUpdate(Index).Enabled = True
End Sub
Private Sub cmdUpdateBranch_Click(Index As Integer)
cmdUpdateBranch(Index).Enabled = False
Me.MousePointer = vbHourglass
grid1.Rows = 0
SendBranches con, Me, Index = 1
Me.MousePointer = vbNormal
cmdUpdateBranch(Index).Enabled = False
End Sub
Private Sub Form_Load()
openCon con
Set rsApiStock = createRs(cmd("SettingApiStock", con, adTable).Execute)
grid1.ColWidth(0) = grid1.Width - 400
grid1.ColAlignment(0) = flexAlignRightCenter
End Sub
Public Function AddNote(sNote As String, Optional color As Long = 0)
grid1.AddItem sNote
If color > 0 Then
    grid1.Cell(flexcpBackColor, grid1.Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = color
End If
DoEvents
End Function
