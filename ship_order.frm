VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form online_ship_order 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„” ‰œ«  «·‘Õ‰"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   18255
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame fmClose 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   900
      Width           =   2895
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«€·«ﬁ „” ‰œ"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   270
         Width           =   1230
      End
      Begin Threed.SSCommand cmdClosePeriod 
         Height          =   420
         Left            =   90
         TabIndex        =   34
         Top             =   225
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   741
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
         Caption         =   "«€·«ﬁ › —…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   8325
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   135
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "ship_order.frx":0000
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ship_order.frx":21A7
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "ship_order.frx":41EE
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ship_order.frx":62D9
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   135
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "ship_order.frx":82D3
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ship_order.frx":A3E4
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   135
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "ship_order.frx":C3DE
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ship_order.frx":E602
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   12870
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   -45
      Width           =   5325
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4005
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
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
         Picture         =   "ship_order.frx":106D3
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ship_order.frx":12A9E
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   2700
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
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
         Picture         =   "ship_order.frx":14B47
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ship_order.frx":16B4F
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1350
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   900
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
         Picture         =   "ship_order.frx":18B06
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ship_order.frx":1B2A2
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
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
         Picture         =   "ship_order.frx":1D736
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Index           =   4
      Left            =   7605
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   495
      Width           =   1545
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   45
         TabIndex        =   18
         Top             =   135
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   820
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
         Picture         =   "ship_order.frx":1FA59
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ship_order.frx":2237E
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   465
         Left            =   45
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   630
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   820
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "ship_order.frx":24BD2
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ship_order.frx":26D32
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14490
      TabIndex        =   8
      Top             =   8325
      Width           =   3705
      Begin MSDataListLib.DataCombo xYear 
         Height          =   330
         Left            =   1665
         TabIndex        =   9
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
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
      Begin MSDataListLib.DataCombo xMonth 
         Height          =   330
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   870
         _ExtentX        =   1535
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
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·”‰…"
         Height          =   285
         Left            =   2970
         TabIndex        =   12
         Tag             =   "Color"
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·‘Â—"
         Height          =   330
         Left            =   945
         TabIndex        =   11
         Tag             =   "Color"
         Top             =   225
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   9180
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   630
      Width           =   9015
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   585
         Width           =   3750
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6570
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   180
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo xship 
         Height          =   360
         Left            =   5580
         TabIndex        =   2
         Top             =   540
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·ÕÊŸ…"
         Height          =   270
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   585
         Width           =   585
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "‘—ﬂ… «·‘Õ‰"
         Height          =   270
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   630
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ „” ‰œ"
         Height          =   240
         Left            =   7830
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   795
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   465
      Left            =   5085
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6630
      Left            =   180
      TabIndex        =   4
      Top             =   1665
      Width           =   18015
      _cx             =   31776
      _cy             =   11695
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
      BackColorFixed  =   12648384
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      TabBehavior     =   1
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   9720
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   45
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   476
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
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   3240
         TabIndex        =   15
         Top             =   45
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
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
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   7110
         TabIndex        =   16
         Top             =   45
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   582
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
   Begin Threed.SSCommand cmdSup 
      Height          =   420
      Left            =   10755
      TabIndex        =   30
      Top             =   8460
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   741
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«Œ «— «·„Ê—œ"
      TagVariant      =   "«Œ «— «·„Ê—œ"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand cmdProject 
      Height          =   420
      Left            =   6975
      TabIndex        =   31
      Top             =   8460
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   741
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "√Œ «— «·„‘—Ê⁄"
      TagVariant      =   "√Œ «— «·„‘—Ê⁄"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand cmdFilter 
      Height          =   555
      Left            =   2835
      TabIndex        =   36
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Undo Filter"
      ButtonStyle     =   3
   End
End
Attribute VB_Name = "online_ship_order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nType As Byte, bEdit As Boolean, sDoc_no As String
Dim bIg As Boolean
Public bCheck As Boolean
Dim dbm As New DBManager
Dim bEditRecord As Boolean, bAct As Boolean
Dim oSeachCode As New Search, oSearchDoc As New Search, oSearchProject As New Search, oSearchSup As New Search
Dim cFilter As String, cSelect As String
Dim CardTable As ADODB.Recordset
Dim clist1 As String, cList2 As String, cList3 As String, sDef As String
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[SHIP]", addstring(xship.BoundText))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.text))
'aInsert = AddFlag(aInsert, IIf(xDoc_No.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cUserName))
'aInsert = AddFlag(aInsert, IIf(xDoc_No.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")


If dbm.opencn Then
    dbm.ClearState
    If xDoc_No.Tag = DefineMode Then
        xDoc_No.text = dbm.IncData("file6_90SH", "doc_NO", , True)
        aInsert = AddFlag(aInsert, "DOC_NO", addvalue(xDoc_No.text))
        dbm.addState addInsert(aInsert, "FILE6_90SH")
    Else
        dbm.addState addUpdate(aInsert, "FILE6_90SH", "doc_no = " & addvalue(xDoc_No.text))
    End If
    
    'myreplaceGrd Row
    
    myreplace = dbm.ExecuteTransaction
    
    dbm.closeCon
End If
End Function
Sub myProc(Optional sControl As String = "")
If ActiveControl.Name = grid1.Name Then
    grid1.TextMatrix(grid1.Row, 1) = oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0)
    Grid1_AfterEdit grid1.Row, grid1.col
    oSeachCode.Hide
    CellPos 13, grid1.Row, grid1.col
ElseIf ActiveControl.Name = cmdInform.Name Then
    openCardTable tbMode.tbFind, oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
ElseIf ActiveControl.Name = cmdProject.Name Then
'    If oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 0) = "" Then
'        cmdProject.Tag = ""
'        cmdProject.Caption = cmdProject.TagVariant
'    Else
'        cmdProject.Tag = oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 0)
'        cmdProject.Caption = oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 1)
'    End If
'    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
'        If Not openCardTable Then myDefine
'    End If
'    oSearchProject.Hide
End If
End Sub
Private Sub CmdDel_Click()
If MsgBox("Õ–› ?", vbOKCancel + vbDefaultButton2 + vbCritical) <> vbOK Then Exit Sub

Dim con As New ADdbm.Connection
If Not opencn(con) Then Exit Sub

con.BeginTrans
On Error GoTo myError
con.Execute "Delete  From FILE6_90S where Doc_No = " & addvalue(xDoc_No.text)
con.Execute "Delete  From FILE6_90SH where Doc_No = " & addvalue(xDoc_No.text)
con.CommitTrans
closeCon con

If sDoc_no <> "" Then Exit Sub

If Not openCardTable(tbMode.tbPrevious, xDoc_No.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
Finally:
closeCon con
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
If TransCount(con) > 0 Then con.RollbackTrans
Resume Finally
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
cString = "SELECT s.DOC_NO," & _
          " FORMAT(s.DATE,'yyyy/M/d')," & _
          " SHIP.DESCA," & _
          " s.Notes" & _
          " FROM FILE6_90SH s" & _
          " INNER JOIN SHIP ON s.SHIP = SHIP.CODE"
If pFilter <> "" Then cString = cString & " WHERE " & pFilter
Generalarray(1) = cString
Generalarray(2) = "order by sDATE ,sDOC_NO "
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·ÿ·»Ì…"
listarray(0, 1) = "DOC_NO IN (SELECT DOC_NO FROM FILE6_90S WHERE %%SHIP_NO%%)"

listarray(1, 0) = "—ﬁ„ «·›« Ê—…"
listarray(1, 1) = "DOC_NO IN (SELECT DOC_NO FROM FILE6_90S WHERE %%INV_NO%%)"

listarray(1, 0) = " «—ÌŒ «·„” ‰œ"
listarray(1, 1) = "(##[Date]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1350

GrdArray(2, 0) = "‘—ﬂ… «·‘Õ‰"
GrdArray(2, 1) = 2500

GrdArray(3, 0) = "„·ÕÊŸ…"
GrdArray(3, 1) = 3000

Dim aFilter As Variant
aFilter = AddFlag(aFilter, "FILTER", True)
aFilter = AddFlag(aFilter, "FIELD", "DOC_NO")

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.sCaption = "„” ‰œ«  «·‘Õ‰"
oSearchDoc.aFilter = aFilter
oSearchDoc.Show 1
End Sub
Private Sub CmdInform_Click()
CardLookup cFilter
End Sub
Private Sub CmdFirst_Click()
If Not openCardTable(tbMode.tbFirst) Then
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
If Not openCardTable Then
    myDefine
End If
End Sub
Private Sub CmdNext_Click()
If Not openCardTable(tbMode.tbNext, xDoc_No.text) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, xDoc_No.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdNewInv_Click()
myDefine
End Sub
Private Sub cmdProject_Click()
'ProjectLookup Me, oSearchProject, , , cmdProject.Tag <> ""
End Sub
Private Sub cmdSave_Click()
If Not myValid Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If sDoc_no <> "" Then
    Unload Me
    Exit Sub
End If
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub cmdSup_Click()
suplookup Me, oSearchSup, , , cmdSup.Tag <> ""
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub

Private Sub Command1_Click()
End Sub
Private Sub cmdYear_Click()
Set oSearchYear.myForm = Me
oSearchYear.bEmpty = cmdYear.Tag <> ""
oSearchYear.Show 1
End Sub
Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xDoc_No.Tag = LoadMode Then
        grid1.SetFocus
    Else
        xDate.SetFocus
    End If
    Err.Clear
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then KeyAscii = 0
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End If
End Sub
Private Sub Form_Load()
bEdit = True
If dbm.openCon Then
    Set xship.RowSource = dbm.myRs("Select Code,DescA From ship ORDER BY CODE")
    xship.ListField = "Desca"
    xship.BoundColumn = "Code"
End If
If Not openCardTable Then myDefine
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Set online_ship_order = Nothing
Err.Clear
End Sub
Private Function myValid(Optional bIgMsg As Boolean = True) As Boolean

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Not xship.MatchedWithList Then
    MsgBox "‘—ﬂ… «·‘Õ‰ €Ì— „”Ã·…"
    Exit Function
End If
myValid = True
End Function
Private Sub myload()
xDoc_No.text = CardTable!doc_no
xDate.text = myFormat_p(CardTable!Date)
xship.BoundText = CardTable!SHIP & ""
xNotes.text = CardTable!NOTES & ""
bIg = True
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
bIg = False

panel1(0).Caption = CardTable!UserName & " " & myFormat_p(CardTable!Time, True)
panel1(1).Caption = CardTable!USERNAME2 & " " & myFormat_p(CardTable!TIME2, True)
Handlecontrols LoadMode
myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
xDoc_No.text = ""
xDate.text = myFormat_p(Date)
xship.BoundText = ""
xNotes.text = ""
bIg = True
xClosed.Value = 0
bIg = False

Fixgrd

grid1.Rows = 1

myAddItem

Handlecontrols DefineMode

CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit And xClosed.Value = 0
xClosed.Enabled = bopt1 And nMode = LoadMode
cmdClosePeriod.Enabled = bopt1

cmdFilter.Visible = cmdFilter.Tag <> ""
cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmddel.Enabled = nMode = LoadMode And bEditRecord

cmdSave.Enabled = bEditRecord
cmddel.Enabled = nMode = LoadMode And bEditRecord

Dim nRecord As Long, nRecords As Long
retRecords xDoc_No.text, nRecords, nRecord

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
     panel1(0).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""

xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If Not ValidNum(xDoc_No.text) Then
     If xDoc_No.Tag = LoadMode Then
        myDefine
    Else
        xDoc_No.text = ""
    End If
Else
    If (Not (CardTable.EOF)) And xDoc_No.Tag = LoadMode Then
        If CardTable!doc_no = xDoc_No.text Then
            Exit Sub
        End If
    End If
    
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If

End If
End Sub
Private Function CalcTotals(Optional Row As Long = 0)
Dim nTotal As Double
With grid1
Dim i As Long
For i = 1 To .Rows - 2
    nTotal = nTotal + mRound(.TextMatrix(i, 2))
Next

If Row = 0 Then Exit Function

.TextMatrix(Row, 5) = ""
.TextMatrix(Row, 6) = ""

If .TextMatrix(Row, 1) = "" Then Exit Function
If .TextMatrix(Row, 0) <> "" Then
End If
End With
End Function
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_no As String = "", Optional bDefine As Boolean = True) As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_no) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

'On Error GoTo myError
cFilter = ""
If cmdFilter.Tag <> "" Then cFilter = cFilter & Tr(cFilter) & "DOC_NO IN (" & cmdFilter.Tag & ")"
If xYear.text <> "" Then cFilter = cFilter & Tr(cFilter) & "YEAR(DATE) = " & xYear.text
If xMonth.text <> "" Then cFilter = cFilter & Tr(cFilter) & "MONTH(DATE) = " & xMonth.text

If sDoc_no <> "" Then cFilter = "DOC_NO = " & sDoc_no
cString = "SELECT TOP 1 * " & _
          " FROM file6_90SH"

If pMode = tbMode.tbFirst Then
    cOrder = "Order by Doc_no"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by Doc_no DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "Doc_no = " & pDoc_no
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "Doc_no < " & pDoc_no
    cOrder = "Order by Doc_no desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "Doc_no > " & pDoc_no
    cOrder = "Order by Doc_no"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

If cFilter <> "" Then
    cString = cString & IIf(cWhere = "", " WHERE ", " AND ") & cFilter
End If

cString = cString & " " & cOrder
If dbm.openCon Then
    Set CardTable = dbm.myRs(cString)
    If (Not CardTable.EOF) Then
        myload
        openCardTable = True
    End If
    Me.MousePointer = 0
End If
Exit Function
myError:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function
Private Sub myUndo()
If xDoc_No.Tag = DefineMode Then
    If Not openCardTable Then
        CmdNewInv_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xDate_DblClick()
Set datefrm.oDate = xDate
datefrm.Show 1
End Sub
Private Sub chkDay_Click()
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub chkMonth_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Sub chkOpen_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Sub chkYear_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_no <> "" Then
    cString = "SELECT Count(*) AS records,COUNT(CASE WHEN DOC_NO <= " & pDoc_no & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM FILE6_90SH"
If cFilter <> "" Then cString = cString & " WHERE " & cFilter

Set loctable = myRs(cString)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Sub cmdFilter_Click()
cmdFilter.Tag = ""
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Sub myproc2(pFilter As String)
oSearchDoc.Hide
cmdFilter.Tag = pFilter
If Not openCardTable(tbMode.tbFirst, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub myreplaceGrd(Row As Long)
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addvalue(xDoc_No.text))
        aInsert = AddFlag(aInsert, "ORDER_NO", addvalue(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "ORDER_NO_SUP", addstring(grid1.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "INV_NO", addstring(grid1.TextMatrix(i, 3)))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            dbm.addState addInsert(aInsert, "FILE6_90S")
        Else
            dbm.addState addUpdate(aInsert, "FILE6_90S", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
If Not myValid(True) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myLoadGrd
    If Row < grid1.Rows - 1 Then
        grid1.Select Row, col
    Else
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If


If Not validRow(Row, col) Then
    CalcTotals Row
    Exit Sub
End If

With grid1
If Row = grid1.Rows - 1 Then
    myAddItem
ElseIf Row = grid1.Rows - 2 And (col = 0) Then
    MyEditItem grid1, Row, col
End If

If myreplace(Row) Then
    If xDoc_No.Tag = DefineMode Then
        Handlecontrols LoadMode
        myLoadGrd
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myLoadGrd
    Else
        CalcTotals Row
    End If
Else
    myLoadGrd
End If
End With
End Sub
Private Sub grid1_EnterCell()
If Not bEditRecord Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub myLoadGrd()
Dim cString As String
Dim aFields(10)
aFields(0) = "s.SHIP_NO"
aFields(1) = "s.ORDER_NO"
aFields(2) = "s.ORDER_NO_SUP"
aFields(3) = "v.TYPE_DESCA"
aFields(4) = "s.INV_NO"
aFields(5) = "v.DATE"
aFields(6) = "v.NAME"
aFields(7) = "v.PHONE"
aFields(8) = "v.ADDRESS"
aFields(9) = "v.TOTAL_QUANT"
aFields(10) = "s.ID"

cString = "SELECT " & arString(aFields, ",") & _
         " FROM FILE6_90S s " & _
         " INNER JOIN vw_online_orders v ON s.order_no = v.order_no AND s.ORDER_NO_SUP = v.ORDER_NO_SUP"
cString = cString & " WHERE  s.DOC_NO = " & addvalue(xDoc_No.text)
If dbm.openCon Then
    Set grid1.DataSource = dbm.myRs(cString)
End If
myAddItem
Fixgrd
End Sub
Private Sub grid1_GotFocus()
grid1_EnterCell
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
ElseIf Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 112 And grid1.col = 1 Then
    'SupLookupAll Me, oSeachCode
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
        On Error GoTo myError
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from FILE6_90S where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        myRemove grid1.Row
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 0 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
        MsgBox "«·„‘—Ê⁄ €Ì— „”Ã·"
    End If
ElseIf col = 1 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
        MsgBox "«·Œ“‰… €Ì— „”Ã·…"
    End If
End If
End Sub
Private Sub Fixgrd()
With grid1
Dim aFields(10)
aFields(0) = "s.SHIP_NO"
aFields(1) = "s.ORDER_NO"
aFields(2) = "s.ORDER_SUP"
aFields(3) = "v.ONLINE_TYPES.DESCA"
aFields(4) = "s.INV_NO"
aFields(5) = "v.DATE"
aFields(6) = "v.NAME"
aFields(7) = "v.PHONE"
aFields(8) = "v.ADDRESS"
aFields(9) = "v.TOTAL_QUANT"
aFields(10) = "s.ID"

.FormatString = "„|" & "»Ê·Ì’… «·‘Õ‰|" & "»Ê·Ì’… «÷«›Ì…|" & "‰Ê⁄ «·»Ê·Ì’…|" & "—ﬁ„ «·›« Ê—…|" & "«· «—ÌŒ|" & "«·«”„|" & "«· ·Ì›Ê‰|" & "«·⁄‰Ê«‰|" & "«·ﬂ„Ì…|" & "«·ﬁÌ„…|"
.ColWidth(0) = 800
.ColWidth(1) = 1500
.ColWidth(2) = 800
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.ColWidth(5) = 1500
.ColWidth(6) = 1500
.ColWidth(7) = 1500
.ColWidth(8) = 2000
.ColWidth(10) = 800
.ColHidden(.Cols - 1) = True
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub myAddItem()
With grid1
.AddItem ""
.TextMatrix(.Rows - 1, 0) = .Rows - 1
End With
End Sub
Private Function validRow(Row As Long, Optional col As Long = -1) As Boolean
With grid1
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If col < grid1.Cols - 4 Then
    grid1.col = col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 2)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, col
End If
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
CalcTotals
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
'    If col = 0 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
'    If col = 1 And (grid1.TextMatrix(Row, col) = "" Or grid1.TextMatrix(Row, grid1.Cols - 1) = "") Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'    If grid1.col = 0 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
'    If grid1.col = 1 And (grid1.TextMatrix(grid1.Row, grid1.col) = "" Or grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub xClosed_Click()
If bIg Then Exit Sub
On Error GoTo myError
con.Execute "UPDATE FILE6_90H" & _
            " SET CLOSED = " & xClosed.Value & _
            " WHERE DOC_NO = " & MyParn(xDoc_No.text)
Inform " „ " & IIf(xClosed.Value = 0, "› Õ «·„” ‰œ", "«€·«ﬁ «·„” ‰œ") & " »‰Ã«Õ"
myUndo
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdClosePeriod_Click()
closefrm.sFile = cFileHeader
closefrm.Show 1
myUndo
End Sub
Private Function GrdDesc(pShip_no As String, ByRef Row As Long, Optional ByRef sMsg As String) As Boolean
With grid1
If Trim(pShip_no) = "" Or Len(pShip_no) < 3 Or Len(pShip_no) > 30 Then
    sMsg = "—ﬁ„ »Ê·Ì’… €Ì— ’«·Õ"
    Exit Function
End If

Dim loctable As New ADODB.Recordset

Set loctable = myRs("dbo.sp_online_order_ship", , , adStoredProc, AddFlag(Empty, "SHIP_NO", pShip_no))
If loctable.EOF Then
    MsgBox "—ﬁ„ »Ê·Ì’… €Ì— ’ÕÌÕ"
    GoTo Finally
End If

Dim cmDiscount As New ADODB.command
aPrm = AddFlag(Empty, "ITEM", cm.Parameters("@ITEM").Value)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
Set cmDiscount = cmd("dbo.sp_offer_price", con, adStoredProc, aPrm)
cmDiscount.Execute

grid1.TextMatrix(Row, 1) = cm.Parameters("@ITEM").Value
grid1.TextMatrix(Row, 2) = cm.Parameters("@MOSM").Value
grid1.TextMatrix(Row, 3) = cm.Parameters("@FACT").Value
grid1.TextMatrix(Row, 4) = cm.Parameters("@SUPP").Value & ""
grid1.TextMatrix(Row, 5) = cm.Parameters("@MODELFACT0").Value
grid1.TextMatrix(Row, 6) = cm.Parameters("@DESCA").Value
grid1.TextMatrix(Row, 7) = cm.Parameters("@SCAL").Value
grid1.TextMatrix(Row, 8) = cm.Parameters("@COLOR").Value
grid1.TextMatrix(Row, 9) = cm.Parameters("@PRICE").Value
grid1.TextMatrix(Row, 10) = "1"

If IsNull(cmDiscount.Parameters("@ID").Value) Then
    grid1.TextMatrix(Row, 11) = cm.Parameters("@PRICE").Value
End If
grid1.TextMatrix(Row, 17) = Val(cm.Parameters("@COSTITEM").Value & "")
If Not IsNull(cmDiscount.Parameters("@ID").Value) Then
    grid1.TextMatrix(Row, 11) = cmDiscount.Parameters("@PRICE").Value
    grid1.TextMatrix(Row, 18) = cmDiscount.Parameters("@DOC_NO").Value
    grid1.TextMatrix(Row, 19) = 0
    grid1.TextMatrix(Row, 20) = 0
End If

GrdDesc = True

End With
End Function

