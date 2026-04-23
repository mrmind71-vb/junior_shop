VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form apiStockTransfrm 
   Caption         =   "ÈíÇäÇÊ ÇáÊØÈíÞ Úáí ÇáãæÞÚ"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   645
      Left            =   1485
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2385
      Width           =   2895
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1260
      Width           =   8070
      Begin Threed.SSCommand cmdUsers 
         Height          =   510
         Left            =   4140
         TabIndex        =   13
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
         Caption         =   "ÓÍÈ ÈíÇäÇÊ ÇáãÓÊÎÏãíä"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame Frame3 
      Height          =   825
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3015
      Width           =   8115
      Begin Threed.SSCommand cmdOrders 
         Height          =   555
         Left            =   4095
         TabIndex        =   8
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
         Caption         =   "ÓÍÈ ØáÈíÇÊ ÇáãÚÇÑÖ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOrdersSend 
         Height          =   555
         Left            =   90
         TabIndex        =   14
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
         Caption         =   "ÇÑÓÇá ØáÈíÇÊ ÇáãÚÇÑÖ"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1185
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   4020
      Begin Threed.SSCommand cmdItems 
         Height          =   420
         Left            =   90
         TabIndex        =   5
         Top             =   180
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
         Caption         =   "ÊÍÏíË ÈíÇäÇÊ ßá ÇáÇÕäÇÝ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClients 
         Height          =   420
         Left            =   90
         TabIndex        =   6
         Top             =   675
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
         Caption         =   "ÊÍÏíË ÈíÇäÇÊ ßá  ÇáÚãáÇÁ"
         ButtonStyle     =   3
      End
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   2205
      MaskColor       =   &H00FFFFFF&
      Picture         =   "MobilTransfrm.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "ÎÑæÌ"
      Top             =   3870
      UseMaskColor    =   -1  'True
      Width           =   6000
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   4185
      TabIndex        =   0
      Top             =   90
      Width           =   4020
      Begin Threed.SSCommand cmdItemsUpdate 
         Height          =   420
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
         Caption         =   "ÊÍÏíË ÈíÇäÇÊ ÇáÇÕäÇÝ ÇáãÚÏáÉ ÝÞØ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClientsUpdate 
         Height          =   420
         Left            =   90
         TabIndex        =   3
         Top             =   675
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
         Caption         =   "ÊÍÏíË ÈíÇäÇÊ ÇáÚãáÇÁ ÇáãÍÏËÉ ÝÞØ"
         ButtonStyle     =   3
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   4485
      Visible         =   0   'False
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4200
      Left            =   8280
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   180
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
      TabIndex        =   11
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
      Caption         =   "ÖÈØ ÈíÇäÇÊ ÇáãæÞÚ"
      ButtonStyle     =   3
   End
End
Attribute VB_Name = "apiStockTransfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdClients_Click()
Me.MousePointer = vbHourglass
grid1.Rows = 0
SendAllCustomers con, Me, True
Me.MousePointer = vbNormal
End Sub
Private Sub cmdClientsUpdate_Click()
Me.MousePointer = vbHourglass
grid1.Rows = 0
SendAllCustomers con, Me, False
Me.MousePointer = vbNormal
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdFixMobileSetting_Click()
ApiStockSettingfrm.Show 1
End Sub

Private Sub cmdItems_Click()
Me.MousePointer = vbHourglass
grid1.Rows = 0
SendAllItems con, Me, True
Me.MousePointer = vbNormal
End Sub
Private Sub cmdItemsUpdate_Click()
Me.MousePointer = vbHourglass
grid1.Rows = 0

If rsApiStock = Nothing Then Exit Sub
If rsApiStock.EOF Then Exit Sub

SendAllItems con, Me, False
'subAddOrder Me, com
Me.MousePointer = vbNormal
End Sub
Private Sub cmdOrders_Click()
Me.MousePointer = vbHourglass
grid1.Rows = 0
GetOrders Me, con
SubAddOrder Me, con
Me.MousePointer = vbNormal
End Sub
Private Sub cmdOrdersSend_Click()
Me.MousePointer = vbHourglass
grid1.Rows = 0
SendAllOrders con, Me, True
Me.MousePointer = vbNormal
End Sub
Private Sub cmdUsers_Click()
Me.MousePointer = vbHourglass
grid1.Rows = 0
GetUsers Me, con
Me.MousePointer = vbNormal
End Sub

Private Sub Command1_Click()
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
