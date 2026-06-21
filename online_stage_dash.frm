VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form online_stage_dash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "ÿ·»Ì«  «Ê‰ ·«Ì‰"
   ClientHeight    =   10635
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10635
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkEdit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ÿ·»Ì«  „⁄œ·…"
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
      Height          =   420
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   405
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   855
      Width           =   1545
      Begin Threed.SSCommand cmdGet 
         Height          =   870
         Left            =   45
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   135
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1535
         _Version        =   196610
         CaptionStyle    =   1
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
         Picture         =   "online_stage_dash.frx":0000
         Caption         =   " ÃÂÌ“ ÿ·»Ì…"
         TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
         ButtonStyle     =   3
         PictureAlignment=   11
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1890
      Width           =   1500
      Begin Threed.SSCommand cmdGo 
         Height          =   510
         Left            =   90
         TabIndex        =   7
         Top             =   180
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
         Picture         =   "online_stage_dash.frx":2665
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   900
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
         Picture         =   "online_stage_dash.frx":5518
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   -315
      Top             =   -135
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
   Begin MSAdodcLib.Adodc DATA2 
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
   Begin MSAdodcLib.Adodc data5 
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
   Begin MSAdodcLib.Adodc DATA6 
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
   Begin MSAdodcLib.Adodc DATA7 
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
   Begin MSAdodcLib.Adodc Adodc1 
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
      Caption         =   "data7"
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
   Begin MSAdodcLib.Adodc Adodc2 
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA4 
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3120
      Left            =   18000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   2265
      _cx             =   3995
      _cy             =   5503
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   27
      FixedRows       =   2
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
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   10215
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   741
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
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
         Left            =   4095
         TabIndex        =   11
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
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
         Left            =   8100
         TabIndex        =   12
         Top             =   45
         Width           =   4000
         _ExtentX        =   7064
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
         Index           =   3
         Left            =   12150
         TabIndex        =   13
         Top             =   45
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
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
         Index           =   4
         Left            =   16155
         TabIndex        =   14
         Top             =   45
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grdError 
      Height          =   3120
      Left            =   1845
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   16125
      _cx             =   28443
      _cy             =   5503
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   2
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
      ExplorerBar     =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdPrep 
      Height          =   3120
      Left            =   10305
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3285
      Width           =   9960
      _cx             =   17568
      _cy             =   5503
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   9
      FixedRows       =   2
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
      ExplorerBar     =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdDone 
      Height          =   3120
      Left            =   270
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3285
      Width           =   9960
      _cx             =   17568
      _cy             =   5503
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   2
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
      ExplorerBar     =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdEdit 
      Height          =   2985
      Left            =   10305
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6480
      Width           =   9960
      _cx             =   17568
      _cy             =   5265
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   9
      FixedRows       =   2
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
      ExplorerBar     =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdCancel 
      Height          =   2985
      Left            =   270
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6480
      Width           =   9960
      _cx             =   17568
      _cy             =   5265
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   9
      FixedRows       =   2
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
      ExplorerBar     =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "online_stage_dash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oPassword As online_stage_pass
Dim cString As String
Dim CLIST As String
Dim oSearchMan As New Search_abd
Dim cStr1 As String, cStr2 As String
Dim oSearchItem As New Search3
Dim con As New ADODB.Connection
Private Sub cmd_addexel_Click()
Inform "  „ «÷«›… «·ÿ·»Ì«  "
myload
End Sub

Private Sub chkEcit_Click()
myloadGrdDone
End Sub

Private Sub chkEdit_Click()
myloadGrdDone
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGet_Click()
Set oPassword = New online_stage_pass
Set oPassword.myForm = Me
oPassword.Show 1
End Sub
Private Sub cmdGo_Click()
myload
End Sub
Private Sub Form_Load()
    myload
    Exit Sub
myError:
    MsgBox Err.Description
    Err.Clear
End Sub
Public Sub myload()
myLoadGrd
myloadGrdError
myloadgrdPrep
myloadGrdDone
myloadGrdEdit
myloadGrdCancel
End Sub
Public Sub myProc(Optional pValue As String)
If ActiveControl.Name = cmdGet.Name Then
    Unload oPassword
    GetOrder pValue
ElseIf ActiveControl.Name = grdError.Name Then
    Unload oPassword
    GetError pValue
ElseIf ActiveControl.Name = grdDone.Name Then
    Unload oPassword
    GetErrorRp pValue
End If
End Sub
Private Sub GetOrder(sManCode As String)

Dim con As New ADODB.Connection
If Not openCn(con) Then Exit Sub


Dim aRet As Variant
Dim cString As String
cString = "SELECT TOP 1 " & _
        " vw_online_order.ORDER_NO, " & _
        " vw_online_order.DOC_NO" & _
        " FROM vw_online_order " & _
        " WHERE STAGE = 2" & _
        " AND MAN = " & MyParn(sManCode) & _
        "  ORDER BY PAYMOB DESC,DATE "

aRet = rsValues(cString, con)

Dim sStage As String
If Not IsNull(aRet) Then
    If MsgBox("ÿ·»Ì…  Õ  «· ÃÂÌ“ ··„‰œÊ» !! ÕœÌÀ", vbOKCancel) <> vbOK Then GoTo Finally
Else
    cString = "SELECT TOP 1 " & _
            " vw_online_order.ORDER_NO, " & _
            " vw_online_order.DOC_NO" & _
            " FROM vw_online_order " & _
            " WHERE vw_online_order.STAGE = 1" & _
            " AND vw_online_order.PREP = 1" & _
            " ORDER BY PAYMOB DESC,DATE"
    aRet = rsValues(cString, con)
End If

If Not IsNull(aRet) Then
    Set online_Stage_order.myForm = Me
    online_Stage_order.sOrder_No = retFlag(aRet, "order_no")
    online_Stage_order.sDoc_no = retFlag(aRet, "doc_no")
    online_Stage_order.sManCode = sManCode
    
    online_Stage_order.Show 1
End If
Finally:
closeCon con
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
Resume Finally
End Sub
Private Sub GetError(sManCode As String)
Set online_Stage_order.myForm = Me
online_Stage_order.sOrder_No = grdError.TextMatrix(grdError.Row, 0)
online_Stage_order.sDoc_no = Val(grdError.TextMatrix(grdError.Row, 1))
online_Stage_order.sManCode = sManCode
online_Stage_order.Show 1
End Sub
Private Sub GetErrorRp(sManCode As String)
Set online_Stage_order.myForm = Me
online_Stage_order.sOrder_No = grdDone.TextMatrix(grdDone.Row, 0)
online_Stage_order.sDoc_no = Val(grdDone.TextMatrix(grdDone.Row, 1))
online_Stage_order.sManCode = sManCode
online_Stage_order.Show 1
End Sub
Public Sub myLoadGrd()
On Error GoTo myError
With grid1
cString = "SELECT  STAGES_CODES.CODE, " & _
    "        STAGES_CODES.DESCA, " & _
    "        COUNT(*) " & _
    "FROM    vw_online_order INNER JOIN " & _
    "        STAGES_CODES ON vw_online_order.STAGE = STAGES_CODES.CODE " & _
    " AND vw_online_order.PREP = 1" & _
    " GROUP BY STAGES_CODES.CODE, STAGES_CODES.DESCA " & _
    " ORDER BY STAGES_CODES.CODE "
    Set grid1.DataSource = myRs(cString)
End With
Fixgrd
'grid1.Cell(flexcpAlignment, 0, 0, grid1.Rows - 1, grid1.Cols - 1) = 7
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Sub Fixgrd()
With grid1
    '.RowHeight(0) = 400
    .RowHeight(1) = 400
    
    .TextMatrix(0, 0) = "«·„—Õ·…"
    .TextMatrix(0, 1) = "«·„—Õ·…"
    .TextMatrix(0, 2) = "⁄œœ «·ÿ·»Ì« "
    
    .TextMatrix(0, 0) = " ÃÂÌ“ ÿ·»Ì« "
    .TextMatrix(0, 1) = " ÃÂÌ“ ÿ·»Ì« "
    .TextMatrix(0, 2) = " ÃÂÌ“ ÿ·»Ì« "
    
    .TextMatrix(1, 0) = "«·„—Õ·…"
    .TextMatrix(1, 1) = "«·„—Õ·…"
    .TextMatrix(1, 2) = "«·⁄œœ"
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .ColHidden(0) = True
    .ColWidth(1) = 1400
    .ColWidth(2) = 700
    
    '.ExplorerBar = flexExSort
     .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 2, "#", vbYellow, , True, ""
    If grid1.Rows > 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 1) = "«·≈Ã„«·Ì"
    End If
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'closeCon con
    'If cBranch <> "00" Then closeCon con_MyShop
    SaveText Me, , Array(xDate1.Name, xdate2.Name)
End Sub

Private Sub grdDone_Click()
If grdDone.Row > 1 Then
    Set oPassword = New online_stage_pass
    Set oPassword.myForm = Me
    oPassword.Show 1
End If
End Sub

Private Sub grdError_DblClick()
If grdError.Row > 1 Then
    Set oPassword = New online_stage_pass
    Set oPassword.myForm = Me
    oPassword.Show 1
End If
End Sub

Private Sub grid1_DblClick()
If grid1.Row < 1 Or grid1.Row = grid1.Rows - 1 Then Exit Sub
End Sub
Private Sub myloadGrdError()
Dim strSql As String
strSql = "SELECT v.ORDER_NO, " & _
         "NULLIF(v.DOC_NO, 0), " & _
         "t.desca, " & _
         "FORMAT(v.[DATE],'yyyy/M/d'), " & _
         "FORMAT(v.TIME2,'yyyy/M/d HH:mm'), " & _
         "dbo.fn_time_diff(v.TIME2, GETDATE()) AS TimePeriod," & _
         "FILE6_25.DESCA, " & _
         "v.TOTAL_QUANT, " & _
         "e.DESCA, " & _
         " FORMAT(e.[TIME],'yyyy/M/d HH:mm'), " & _
         " dbo.fn_time_diff(e.[TIME], GETDATE()) AS TimePeriod" & _
         " FROM vw_online_order v " & _
         " INNER JOIN ONLINE_TYPES t ON v.TYPE = t.CODE" & _
         " INNER JOIN FILE6_25 ON v.MAN = FILE6_25.CODE " & _
         " LEFT JOIN vw_online_er e ON v.ORDER_NO = e.ORDER_NO AND v.DOC_NO = e.doc_no_sup and v.stage = e.stage" & _
         " WHERE v.STAGE = 4" & _
         " ORDER BY e.ID"
On Error GoTo myError
Set grdError.DataSource = myRs(strSql)
fixGrdError
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrdError()
Dim i As Long
With grdError
    '.RowHeight(0) = 400
    '.RowHeight(1) = 600
    For i = 0 To .Cols - 1
        .TextMatrix(0, i) = "ÿ·»Ì«  »Â« „‘ﬂ·« "
        .ColAlignment(i) = flexAlignCenterCenter
    Next
    
    .TextMatrix(1, 0) = "—ﬁ„ «·ÿ·»Ì…"
    .TextMatrix(1, 1) = "›—⁄Ì"
    .TextMatrix(1, 2) = "‰Ê⁄ «·ÿ·»Ì…"
    .TextMatrix(1, 3) = " «—ÌŒ «·ÿ·»Ì…"
    .TextMatrix(1, 4) = "Êﬁ  «· ÃÂÌ“"
    .TextMatrix(1, 5) = "„œ… «· ÃÂÌ“"
    .TextMatrix(1, 5 + 1) = "«·„‰œÊ»"
    .TextMatrix(1, 6 + 1) = "«·ﬂ„Ì…"
    .TextMatrix(1, 7 + 1) = "»Ì«‰ «·„‘ﬂ·…"
    .TextMatrix(1, 8 + 1) = "Êﬁ  «·„‘ﬂ·…"
    .TextMatrix(1, 9 + 1) = "„œ… «·„‘ﬂ·…"
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .ColHidden(1) = True
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 1250
    .ColWidth(3) = 1200
    .ColWidth(4) = 1700
    .ColWidth(5) = 1000
    .ColWidth(6) = 1500
    .ColWidth(7) = 700
    .ColWidth(8) = 4000
    .ColWidth(9) = 1700
    .ColWidth(10) = 1100
    
     .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub
Private Sub myloadgrdPrep()
Dim strSql As String
strSql = "SELECT v.ORDER_NO, " & _
         "FORMAT(v.[DATE],'yyyy/M/d'), " & _
         "t.desca, " & _
         "FILE6_25.DESCA, " & _
         "FORMAT(v.TIME2,'yyyy/M/d HH:mm'), " & _
         "dbo.fn_time_diff(time2, GETDATE()) AS TimePeriod," & _
         "v.TOTAL_QUANT " & _
         " FROM vw_online_order v " & _
         " INNER JOIN FILE6_25 ON v.MAN = FILE6_25.CODE " & _
         " INNER JOIN ONLINE_TYPES t ON v.TYPE = t.CODE " & _
         " WHERE v.STAGE = 2"
On Error GoTo myError
Set grdPrep.DataSource = myRs(strSql)
fixgrdPrep
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixgrdPrep()
Dim i As Long
With grdPrep
    '.RowHeight(0) = 400
    '.RowHeight(1) = 600
    For i = 0 To .Cols - 1
        .TextMatrix(0, i) = "»Ì«‰ ÿ·»Ì«   Õ  «· ÃÂÌ“"
        .ColAlignment(i) = flexAlignCenterCenter
    Next
    
    .TextMatrix(1, 0) = "—ﬁ„ «·ÿ·»Ì…"
    .TextMatrix(1, 1) = " «—ÌŒ «·ÿ·»Ì…"
    .TextMatrix(1, 2) = "«·‰Ê⁄"
    .TextMatrix(1, 3) = "«·„‰œÊ»"
    .TextMatrix(1, 4) = "Êﬁ  «· ÃÂÌ“"
    .TextMatrix(1, 5) = "„œ… «· ÃÂÌ“"
    .TextMatrix(1, 6) = "«·ﬂ„Ì…"
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    '.ColHidden(0) = True
    '.ColHidden(2) = True
    '.ColHidden(8) = True
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1300
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1700
    .ColWidth(5) = 1000
    .ColWidth(6) = 800
    
'    '.ExplorerBar = flexExSort
     .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
'    .SubtotalPosition = flexSTBelow
'    .Subtotal flexSTCount, -1, 2, "#", vbYellow, , True, ""
'    If grid1.Rows > 1 Then
'        grid1.TextMatrix(grid1.Rows - 1, 1) = "«·≈Ã„«·Ì"
'    End If
    End With
End Sub
Private Sub myloadGrdDone()
Dim strSql As String
strSql = "SELECT v.ORDER_NO, " & _
         "NULLIF(v.DOC_NO, 0), " & _
         "FORMAT(v.[DATE],'yyyy/M/d'), " & _
         "t.desca, " & _
         "FILE6_25.DESCA, " & _
         "FORMAT(v.TIME2,'yyyy/M/d HH:mm'), " & _
         "FORMAT(v.TIME3,'yyyy/M/d HH:mm'), " & _
         "  dbo.fn_time_diff(time2, time3) AS TimePeriod," & _
         " v.TOTAL_QUANT " & _
         " FROM vw_online_order v" & _
         " INNER JOIN ONLINE_TYPES t ON v.TYPE = t.CODE" & _
         " INNER JOIN FILE6_25 ON v.MAN = FILE6_25.CODE " & _
        " WHERE v.STAGE = 3"
If chkEdit.Value = 1 Then
    strSql = strSql & " AND (SELECT COUNT(*)  FROM FILE6_90_ER WHERE SEEN = 0 AND STAGE = 4 AND ORDER_NO = V.ORDER_NO AND DOC_NO = v.DOC_NO) > 0"
End If

On Error GoTo myError
Set grdDone.DataSource = myRs(strSql)
fixGrdDone
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrdDone()
Dim i As Long
With grdDone
    '.RowHeight(0) = 400
    '.RowHeight(1) = 600
    For i = 0 To .Cols - 1
        .TextMatrix(0, i) = "ÿ·»Ì«   „  ÃÂÌ“Â«"
        .ColAlignment(i) = flexAlignCenterCenter
    Next
    
    .TextMatrix(1, 0) = "—ﬁ„ «·ÿ·»Ì…"
    .TextMatrix(1, 1) = "›—⁄Ì"
    .TextMatrix(1, 1 + 1) = " «—ÌŒ «·ÿ·»Ì…"
    .TextMatrix(1, 2 + 1) = "‰Ê⁄ «·ÿ·»Ì…"
    
    .TextMatrix(1, 3 + 1) = "«·„‰œÊ»"
    .TextMatrix(1, 4 + 1) = "»œ¡ «· ÃÂÌ“"
    .TextMatrix(1, 5 + 1) = "«‰ Â«¡ «· ÃÂÌ“"
    .TextMatrix(1, 6 + 1) = "«·„œ…"
    .TextMatrix(1, 7 + 1) = "«·ﬂ„Ì…"
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .ColHidden(1) = True
    
    .ColWidth(0) = 800
    .ColWidth(1 + 1) = 1250
    .ColWidth(2 + 1) = 1000
    .ColWidth(3 + 1) = 1400
    .ColWidth(4 + 1) = 1700
    .ColWidth(5 + 1) = 1700
    .ColWidth(6 + 1) = 1000
    .ColWidth(7 + 1) = 700
    
     .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub
Private Sub myloadGrdEdit()
Dim strSql As String
strSql = "SELECT v.ORDER_NO, " & _
         "v.DOC_NO," & _
         "t.DESCA, " & _
         "FORMAT(v.[DATE],'yy/M/d'), " & _
         "FORMAT(v.TIME2,'yy/M/d HH:mm'), " & _
         "FILE6_25.DESCA, " & _
         "v.TOTAL_QUANT, " & _
         "e.DESCA, " & _
         "FORMAT(e.[TIME],'yyyy/M/d HH:mm'), " & _
         "FORMAT(DATEADD(MINUTE, DATEDIFF(MINUTE, e.[TIME], GETDATE()), 0), 'HH:mm') AS TimePeriod" & _
         " FROM vw_online_order v " & _
         " INNER JOIN FILE6_25 ON v.MAN = FILE6_25.CODE " & _
         " INNER JOIN ONLINE_TYPES t ON v.TYPE = t.CODE " & _
         " INNER JOIN vw_online_er e ON v.ORDER_NO =  e.ORDER_NO AND v.DOC_NO =  e.DOC_NO_SUP AND v.STAGE = e.STAGE  " & _
         " WHERE v.STAGE = 6" & _
         " ORDER BY e.ID"
On Error GoTo myError
Set grdEdit.DataSource = myRs(strSql)
fixGrdEdit
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrdEdit()
Dim i As Long
With grdEdit
    '.RowHeight(0) = 400
    '.RowHeight(1) = 600
    For i = 0 To .Cols - 1
        .TextMatrix(0, i) = "ÿ·»Ì«  ﬁÌœ «· ⁄œÌ·"
        .ColAlignment(i) = flexAlignCenterCenter
    Next
    
    .TextMatrix(1, 0) = "—ﬁ„ «·ÿ·»Ì…"
    .TextMatrix(1, 1) = "›—⁄Ì"
    .TextMatrix(1, 2) = "‰Ê⁄ «·ÿ·»Ì…"
    .TextMatrix(1, 3) = " «—ÌŒ «·ÿ·»Ì…"
    .TextMatrix(1, 4) = "Êﬁ  «· ÃÂÌ“"
    .TextMatrix(1, 5) = "«·„‰œÊ»"
    .TextMatrix(1, 6) = "«·ﬂ„Ì…"
    .TextMatrix(1, 7) = "”»» «· ⁄œÌ·"
    .TextMatrix(1, 8) = "Êﬁ  «·«—”«·"
    .TextMatrix(1, 9) = "«·„œ…"
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .ColHidden(1) = True
    .ColHidden(4) = True
    
    .ColWidth(0) = 900
    .ColWidth(1) = 0
    .ColWidth(2) = 1100
    .ColWidth(3) = 1250
    .ColWidth(4) = 1700
    .ColWidth(5) = 1400
    .ColWidth(6) = 700
    .ColWidth(7) = 1500
    .ColWidth(8) = 1700
    .ColWidth(9) = 1000
    
     .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub
Private Sub myloadGrdCancel()
Dim strSql As String
strSql = "SELECT v.ORDER_NO, " & _
         "v.DOC_NO," & _
         "t.DESCA, " & _
         "FORMAT(v.[DATE],'yy/M/d'), " & _
         "FORMAT(v.TIME2,'yy/M/d HH:mm'), " & _
         "FILE6_25.DESCA, " & _
         "v.TOTAL_QUANT, " & _
         "e.DESCA, " & _
         "FORMAT(e.[TIME],'yyyy/M/d HH:mm'), " & _
         "FORMAT(DATEADD(MINUTE, DATEDIFF(MINUTE, e.[TIME], GETDATE()), 0), 'HH:mm') AS TimePeriod" & _
         " FROM vw_online_order v " & _
         " INNER JOIN FILE6_25 ON v.MAN = FILE6_25.CODE " & _
         " INNER JOIN ONLINE_TYPES t ON v.TYPE = t.CODE " & _
         " INNER JOIN vw_online_er e ON v.ORDER_NO =  e.ORDER_NO AND v.DOC_NO =  e.DOC_NO_SUP AND v.STAGE = e.STAGE  " & _
         " WHERE v.STAGE = 5" & _
         " ORDER BY e.ID"
On Error GoTo myError
Set grdCancel.DataSource = myRs(strSql)
fixGrdcancel
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrdcancel()
Dim i As Long
With grdCancel
    '.RowHeight(0) = 400
    '.RowHeight(1) = 600
    For i = 0 To .Cols - 1
        .TextMatrix(0, i) = "ÿ·»Ì«  ﬁÌœ «·«·€«¡"
        .ColAlignment(i) = flexAlignCenterCenter
    Next
    
    .TextMatrix(1, 0) = "—ﬁ„ «·ÿ·»Ì…"
    .TextMatrix(1, 1) = "›—⁄Ì"
    .TextMatrix(1, 2) = "‰Ê⁄ «·ÿ·»Ì…"
    .TextMatrix(1, 3) = " «—ÌŒ «·ÿ·»Ì…"
    .TextMatrix(1, 4) = "Êﬁ  «· ÃÂÌ“"
    .TextMatrix(1, 5) = "«·„‰œÊ»"
    .TextMatrix(1, 6) = "«·ﬂ„Ì…"
    .TextMatrix(1, 7) = "”»» «· ⁄œÌ·"
    .TextMatrix(1, 8) = "Êﬁ  «·«—”«·"
    .TextMatrix(1, 9) = "«·„œ…"
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .ColHidden(1) = True
    .ColHidden(4) = True
    
    .ColWidth(0) = 900
    .ColWidth(1) = 0
    .ColWidth(2) = 1100
    .ColWidth(3) = 1250
    .ColWidth(4) = 1700
    .ColWidth(5) = 1400
    .ColWidth(6) = 700
    .ColWidth(7) = 1500
    .ColWidth(8) = 1700
    .ColWidth(9) = 1000
    
     .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub




