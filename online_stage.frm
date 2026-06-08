VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form online_stagefrm 
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
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1395
      Width           =   1500
      Begin Threed.SSCommand cmdGo 
         Height          =   510
         Left            =   90
         TabIndex        =   12
         Top             =   675
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
         Picture         =   "online_stage.frx":0000
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   90
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1215
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
         Picture         =   "online_stage.frx":2EB3
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdGet 
         Height          =   510
         Left            =   90
         TabIndex        =   14
         Top             =   135
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
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
         Caption         =   "”Õ» ÿ·»Ì…"
         TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
         ButtonStyle     =   3
         PictureAlignment=   9
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
      Left            =   16425
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   3840
      _cx             =   6773
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
      TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   14460
      _cx             =   25506
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdRepair 
      Height          =   3120
      Left            =   10350
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3285
      Width           =   9915
      _cx             =   17489
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
      TabIndex        =   9
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
      Left            =   10350
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6480
      Width           =   9915
      _cx             =   17489
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2985
      Left            =   270
      TabIndex        =   15
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "online_stagefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oPassword As Password_man
Dim cString As String
Dim CLIST As String
Dim oSearchMan As New Search_abd
Dim cStr1 As String, cStr2 As String
Dim oSearchItem As New Search3
Dim con As New ADODB.Connection
Private Sub cmd_addexel_Click()
AddFromExel
Inform "  „ «÷«›… «·ÿ·»Ì«  "
myload
End Sub
Private Sub CMD_PRINT_Click()
doprint_day
End Sub
Private Sub CMD_SEND_Click()
'Dim cDocSalPost As String
If cBranch = "00" Then Exit Sub

If grid1.TextMatrix(grid1.Row, 16) <> "" Then Exit Sub

Dim oOnlineCheck As New OnlineCheck
Set oOnlineCheck.myForm = Me
oOnlineCheck.sDoc_no = grid1.TextMatrix(grid1.Row, 0)
oOnlineCheck.Show 1

'myreplace_Sales

'If grid1.TextMatrix(grid1.Row, 16) = "" Then
'    If MsgBox(" —ÕÌ· ·›« Ê—… „»Ì⁄«  ", vbYesNo + vbDefaultButton2) = vbYes Then
'        cDocSalPost = myreplace_Sales()
'        If cDocSalPost <> "" Then
'            MsgBox " „  —ÕÌ· ·»Ê‰ „»Ì⁄«  —ﬁ„ " & cDocSalPost
'        Else
'            MsgBox "·„ Ì „ ⁄„· „” ‰œ «·„»Ì⁄«   "
'        End If
'    End If
'    myload
'    grid2.Rows = 2
'End If
End Sub
Private Sub cmdCSV_Click()
Set onlineCSVfrm.myForm = Me
'Set onlineCSVfrm.con = con
onlineCSVfrm.Show 1
End Sub

Private Sub cmdExcel_Click()
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGet_Click()
Set oPassword = New Password_man
Set oPassword.myForm = Me
oPassword.Show 1
End Sub
Private Sub CmdGo_Click()
myload
End Sub

Private Sub Form_Load()
    myload
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Public Sub myload()
myLoadGrd
myloadGrdError
myloadGrdRepair
myloadGrdDone
myloadGrdEdit
End Sub
Public Sub myProc()
If ActiveControl.Name = cmdGet.Name Then
    GetOrder
ElseIf ActiveControl.Name = grdError.Name Then
    GetError
End If
End Sub
Private Sub GetOrder()
Dim sValue As String
Dim sManCode As String

Dim con As New ADODB.Connection
If Not openCn(con) Then Exit Sub

sManCode = Trim(oPassword.xPass.text)
sValue = rsValue("SELECT CODE FROM MAN_ONLINE_cODES WHERE CODE = " & MyParn(sManCode), con) & ""
If sValue = "" Then
    MsgBox "ﬂÊœ €Ì— ’ÕÌÕ"
    GoTo Finally
Else
    Unload oPassword
End If

Dim cString As String
cString = "SELECT TOP 1 FILE6_90H.DOC_NO " & _
          " FROM FILE6_90H " & _
          " INNER JOIN vw_online_orders_open v ON FILE6_90H.DOC_NO = v.ORDER_NO " & _
          " WHERE FILE6_90H.STAGE = 2" & _
          " AND MAN_STAGE = " & sManCode

sValue = rsValue(cString, con) & ""

Dim sStage As String
If sValue <> "" Then
    If MsgBox("ÿ·»Ì…  Õ  «· ÃÂÌ“ ··„‰œÊ» !! ÕœÌÀ", vbOKCancel) <> vbOK Then GoTo Finally
    sStage = "3"
Else
    cString = "SELECT TOP 1 FILE6_90H.DOC_NO " & _
            " FROM FILE6_90H " & _
            " INNER JOIN vw_online_orders_open v ON FILE6_90H.DOC_NO = v.ORDER_NO " & _
            " WHERE FILE6_90H.STAGE = 1" & _
            " ORDER BY FILE6_90H.PAYMOB DESC,FILE6_90H.DATE"
    sValue = rsValue(cString, con) & ""
    sStage = "2"
End If

If sValue <> "" Then
    Set ShowOrderManfrm.myForm = Me
    ShowOrderManfrm.sDoc_no = sValue
    ShowOrderManfrm.sManCode = sManCode
    ShowOrderManfrm.sStage = sStage
    ShowOrderManfrm.Show 1
End If
Finally:
closeCon con
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Resume Finally
End Sub
Private Sub GetError()
Dim sManCode As String

Dim con As New ADODB.Connection
If Not openCn(con) Then Exit Sub

sManCode = Trim(oPassword.xPass.text)
sValue = rsValue("SELECT CODE FROM MAN_ONLINE_cODES WHERE CODE = " & MyParn(sManCode), con) & ""
If sValue = "" Then
    MsgBox "ﬂÊœ €Ì— ’ÕÌÕ"
    GoTo Finally
Else
    Unload oPassword
End If
If sValue <> "" Then
    Set ShowOrderManfrm.myForm = Me
    ShowOrderManfrm.sDoc_no = grdError.TextMatrix(2, 0)
    ShowOrderManfrm.sManCode = sManCode
    ShowOrderManfrm.Show 1
End If
Finally:
closeCon con
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Resume Finally
End Sub
Public Sub myLoadGrd()
On Error GoTo myerror
With grid1
cString = "SELECT  STAGES_CODES.CODE, " & _
    "        STAGES_CODES.DESCA, " & _
    "        COUNT(*) " & _
    "FROM    FILE6_90H INNER JOIN " & _
    "        STAGES_CODES ON FILE6_90H.STAGE = STAGES_CODES.CODE " & _
    "INNER JOIN vw_online_orders_open " & _
    "ON FILE6_90H.DOC_NO = vw_online_orders_open.ORDER_NO " & _
    "GROUP BY  STAGES_CODES.CODE, STAGES_CODES.DESCA " & _
    "ORDER BY STAGES_CODES.CODE "
    Set grid1.DataSource = myRs(cString)
End With
fixGrd
'grid1.Cell(flexcpAlignment, 0, 0, grid1.Rows - 1, grid1.Cols - 1) = 7
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrd()
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
    .ColWidth(1) = 2500
    .ColWidth(2) = 1000
    
    '.ExplorerBar = flexExSort
     .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTCount, -1, 2, "#", vbYellow, , True, ""
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

Private Sub grdError_DblClick()
If grdError.Row > 1 Then
    Set oPassword = New Password_man
    Set oPassword.myForm = Me
    oPassword.Show 1
End If
End Sub

Private Sub grid1_DblClick()
If grid1.Row < 1 Or grid1.Row = grid1.Rows - 1 Then Exit Sub
End Sub
Private Sub myloadGrdError()
Dim strSql As String
strSql = "SELECT FILE6_90H.DOC_NO, " & _
         "FORMAT(FILE6_90H.[DATE],'yyyy/M/d'), " & _
         "FILE6_90H.Payment_Method, " & _
         "FORMAT(FILE6_90H.TIME2,'yyyy/M/d HH:mm'), " & _
         "MAN_ONLINE_CODES.DESCA, " & _
         "FILE6_90H.TOTAL_QUANT, " & _
         "FILE6_90H_ER.DESCA, " & _
         "FORMAT(FILE6_90H_ER.[TIME],'yyyy/M/d HH:mm'), " & _
         "FORMAT(DATEADD(MINUTE, DATEDIFF(MINUTE, FILE6_90H_ER.[TIME], GETDATE()), 0), 'HH:mm') AS TimePeriod" & _
         " FROM FILE6_90H INNER JOIN " & _
         "MAN_ONLINE_CODES ON FILE6_90H.MAN_STAGE = MAN_ONLINE_CODES.CODE " & _
         "INNER JOIN FILE6_90H_ER ON FILE6_90H.DOC_NO = FILE6_90H_ER.DOC_NO " & _
         "INNER JOIN (SELECT MAX(ID) AS ID " & _
         "  FROM FILE6_90H_ER " & _
         "  GROUP BY FILE6_90H_ER.DOC_NO) AS L " & _
         " ON L.ID = FILE6_90H_ER.ID" & _
        " INNER JOIN vw_online_orders_open ON FILE6_90H.DOC_NO = vw_online_orders_open.ORDER_NO " & _
        " WHERE FILE6_90H.STAGE = 4" & _
        " ORDER BY FILE6_90H_ER.ID"
On Error GoTo myerror
Set grdError.DataSource = myRs(strSql)
fixGrdError
Exit Sub
myerror:
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
    .TextMatrix(1, 1) = " «—ÌŒ «·ÿ·»Ì…"
    .TextMatrix(1, 2) = "‰Ê⁄ «·”œ«œ"
    .TextMatrix(1, 3) = "Êﬁ  «· ÃÂÌ“"
    .TextMatrix(1, 4) = "«·„‰œÊ»"
    .TextMatrix(1, 5) = "«·ﬂ„Ì…"
    .TextMatrix(1, 6) = "»Ì«‰ «·„‘ﬂ·…"
    .TextMatrix(1, 7) = "Êﬁ  «·„‘ﬂ·…"
    .TextMatrix(1, 8) = "„œ… «·„‘ﬂ·…"
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .ColHidden(0) = True
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1300
    .ColWidth(2) = 1500
    .ColWidth(3) = 1600
    .ColWidth(4) = 1200
    .ColWidth(5) = 800
    .ColWidth(6) = 5000
    .ColWidth(7) = 1600
    .ColWidth(8) = 1000
    
     .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub
Private Sub myloadGrdRepair()
Dim strSql As String
strSql = "SELECT FILE6_90H.DOC_NO, " & _
         "FORMAT(FILE6_90H.[DATE],'yyyy/M/d'), " & _
         "FILE6_90H.Payment_Method, " & _
         "MAN_ONLINE_CODES.DESCA, " & _
         "FORMAT(FILE6_90H.TIME2,'yyyy/M/d HH:mm'), " & _
         "FORMAT(DATEADD(MINUTE, DATEDIFF(MINUTE, time2, GETDATE()), 0), 'HH:mm') AS TimePeriod," & _
         "FILE6_90H.TOTAL_QUANT, " & _
         "FILE6_90H_ER.DESCA, " & _
         "FORMAT(FILE6_90H_ER.TIME,'yyyy/M/d HH:mm') " & _
         "FROM FILE6_90H INNER JOIN " & _
         "MAN_ONLINE_CODES ON FILE6_90H.MAN_STAGE = MAN_ONLINE_CODES.CODE " & _
         "LEFT  JOIN FILE6_90H_ER ON FILE6_90H.DOC_NO = FILE6_90H_ER.DOC_NO " & _
         "LEFT JOIN (SELECT MAX(ID) AS ID " & _
         "  FROM FILE6_90H_ER WHERE STAGE = 4 " & _
         "  GROUP BY FILE6_90H_ER.DOC_NO) AS L " & _
         " ON L.ID = FILE6_90H_ER.ID" & _
        " INNER JOIN vw_online_orders_open ON FILE6_90H.DOC_NO = vw_online_orders_open.ORDER_NO " & _
        " WHERE FILE6_90H.STAGE = 2" & _
        " ORDER BY FILE6_90H_ER.ID"
On Error GoTo myerror
Set grdRepair.DataSource = myRs(strSql)
fixGrdRepair
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrdRepair()
Dim i As Long
With grdRepair
    '.RowHeight(0) = 400
    '.RowHeight(1) = 600
    For i = 0 To .Cols - 1
        .TextMatrix(0, i) = "»Ì«‰ ÿ·»Ì«   Õ  «· ÃÂÌ“"
        .ColAlignment(i) = flexAlignCenterCenter
    Next
    
    .TextMatrix(1, 0) = "—ﬁ„ «·ÿ·»Ì…"
    .TextMatrix(1, 1) = " «—ÌŒ «·ÿ·»Ì…"
    .TextMatrix(1, 2) = "‰Ê⁄ «·”œ«œ"
    .TextMatrix(1, 3) = "«·„‰œÊ»"
    .TextMatrix(1, 4) = "Êﬁ  «· ÃÂÌ“"
    .TextMatrix(1, 5) = "„œ… «· ÃÂÌ“"
    .TextMatrix(1, 6) = "«·ﬂ„Ì…"
    .TextMatrix(1, 7) = "”»» «·„‘ﬂ·…"
    .TextMatrix(1, 8) = "Êﬁ  «·„‘ﬂ·…"
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .ColHidden(0) = True
    .ColHidden(2) = True
    '.ColHidden(8) = True
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1300
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 800
    .ColWidth(6) = 800
    .ColWidth(7) = 2000
    .ColWidth(8) = 1600
    
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
strSql = "SELECT FILE6_90H.DOC_NO, " & _
         "FORMAT(FILE6_90H.[DATE],'yyyy/M/d'), " & _
         "FILE6_90H.Payment_Method, " & _
         "MAN_ONLINE_CODES.DESCA, " & _
         "FORMAT(FILE6_90H.TIME2,'yyyy/M/d HH:mm'), " & _
         "FORMAT(FILE6_90H.TIME3,'yyyy/M/d HH:mm'), " & _
         "FORMAT(DATEADD(MINUTE, DATEDIFF(MINUTE, time2, time3), 0), 'HH:mm') AS TimePeriod," & _
         "FILE6_90H.TOTAL_QUANT " & _
         "FROM FILE6_90H INNER JOIN " & _
         "MAN_ONLINE_CODES ON FILE6_90H.MAN_STAGE = MAN_ONLINE_CODES.CODE " & _
        " INNER JOIN vw_online_orders_open ON FILE6_90H.DOC_NO = vw_online_orders_open.ORDER_NO " & _
        " WHERE FILE6_90H.STAGE = 3"
On Error GoTo myerror
Set grdDone.DataSource = myRs(strSql)
fixGrdDone
Exit Sub
myerror:
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
    .TextMatrix(1, 1) = " «—ÌŒ «·ÿ·»Ì…"
    .TextMatrix(1, 2) = "‰Ê⁄ «·”œ«œ"
    
    .TextMatrix(1, 3) = "«·„‰œÊ»"
    .TextMatrix(1, 4) = "»œ¡ «· ÃÂÌ“"
    .TextMatrix(1, 5) = "«‰ Â«¡ «· ÃÂÌ“"
    .TextMatrix(1, 6) = "«·„œ…"
    .TextMatrix(1, 7) = "«·ﬂ„Ì…"
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .ColHidden(0) = True
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1300
    .ColWidth(2) = 1500
    .ColWidth(3) = 1800
    .ColWidth(4) = 1600
    .ColWidth(5) = 1600
    .ColWidth(6) = 800
    .ColWidth(7) = 800
    
     .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub
Private Sub myloadGrdEdit()
Dim strSql As String
strSql = "SELECT FILE6_90H.DOC_NO, " & _
         "FORMAT(FILE6_90H.[DATE],'yyyy/M/d'), " & _
         "FILE6_90H.Payment_Method, " & _
         "FORMAT(FILE6_90H.TIME2,'yyyy/M/d HH:mm'), " & _
         "MAN_ONLINE_CODES.DESCA, " & _
         "FILE6_90H.TOTAL_QUANT, " & _
         "FILE6_90H_ER.DESCA, " & _
         "FORMAT(FILE6_90H_ER.[TIME],'yyyy/M/d HH:mm'), " & _
         "FORMAT(DATEADD(MINUTE, DATEDIFF(MINUTE, FILE6_90H_ER.[TIME], GETDATE()), 0), 'HH:mm') AS TimePeriod" & _
         " FROM FILE6_90H INNER JOIN " & _
         " MAN_ONLINE_CODES ON FILE6_90H.MAN_STAGE = MAN_ONLINE_CODES.CODE " & _
         " INNER JOIN FILE6_90H_ER ON FILE6_90H.DOC_NO = FILE6_90H_ER.DOC_NO " & _
         " INNER JOIN (SELECT MAX(ID) AS ID " & _
         "  FROM FILE6_90H_ER WHERE STAGE = 6" & _
         "  GROUP BY FILE6_90H_ER.DOC_NO) AS L " & _
         " ON L.ID = FILE6_90H_ER.ID" & _
        " INNER JOIN vw_online_orders_open ON FILE6_90H.DOC_NO = vw_online_orders_open.ORDER_NO " & _
        " WHERE FILE6_90H.STAGE = 6" & _
        " ORDER BY FILE6_90H_ER.ID"
On Error GoTo myerror
Set grdEdit.DataSource = myRs(strSql)
fixGrdEdit
Exit Sub
myerror:
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
    .TextMatrix(1, 1) = " «—ÌŒ «·ÿ·»Ì…"
    .TextMatrix(1, 2) = "‰Ê⁄ «·”œ«œ"
    .TextMatrix(1, 3) = "Êﬁ  «· ÃÂÌ“"
    .TextMatrix(1, 4) = "«·„‰œÊ»"
    .TextMatrix(1, 5) = "«·ﬂ„Ì…"
    .TextMatrix(1, 6) = "”»» «· ⁄œÌ·"
    .TextMatrix(1, 7) = "Êﬁ  «·«—”«·"
    .TextMatrix(1, 8) = "«·„œ…"
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .ColHidden(0) = True
    .ColHidden(2) = True
    .ColHidden(3) = True
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1300
    .ColWidth(2) = 1500
    .ColWidth(3) = 1600
    .ColWidth(4) = 1200
    .ColWidth(5) = 800
    .ColWidth(6) = 3800
    .ColWidth(7) = 1600
    .ColWidth(8) = 800
    
     .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub


