VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form online_stage_manager 
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
      Height          =   1320
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1620
      Width           =   1500
      Begin Threed.SSCommand cmdGo 
         Height          =   510
         Left            =   90
         TabIndex        =   3
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
         Picture         =   "online_stage_dash_manager.frx":0000
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   90
         TabIndex        =   4
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
         Picture         =   "online_stage_dash_manager.frx":2EB3
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
      Height          =   5505
      Left            =   17910
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   2355
      _cx             =   4154
      _cy             =   9710
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
   Begin VSFlex7Ctl.VSFlexGrid grdRefund2 
      Height          =   2625
      Left            =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2970
      Width           =   8790
      _cx             =   15505
      _cy             =   4630
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
      Height          =   2850
      Left            =   1620
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   16260
      _cx             =   28681
      _cy             =   5027
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
      AutoSearch      =   1
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
   Begin VSFlex7Ctl.VSFlexGrid grdRefund 
      Height          =   2625
      Left            =   8955
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2970
      Width           =   8925
      _cx             =   15743
      _cy             =   4630
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
   Begin VSFlex7Ctl.VSFlexGrid grdInvoices 
      Height          =   4200
      Left            =   135
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5625
      Width           =   8790
      _cx             =   15505
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
   Begin VSFlex7Ctl.VSFlexGrid grdInvoicesReturn 
      Height          =   2220
      Left            =   8955
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5625
      Width           =   11355
      _cx             =   20029
      _cy             =   3916
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
   Begin VSFlex7Ctl.VSFlexGrid grdReplace 
      Height          =   1950
      Left            =   8955
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7875
      Width           =   11355
      _cx             =   20029
      _cy             =   3440
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "online_stage_manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sDateSales As String
Dim dbm As New DBManage
Private Sub cmd_addexel_Click()
Inform "  „ «÷«›… «·ÿ·»Ì«  "
myload
End Sub

Private Sub chkEdit_Click()
myloadGrdDone
End Sub

Private Sub cmdExit_Click()
    Unload Me
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
Private Sub myload()
sDateSales = myFormat(rsDateBranch(sBranchOnline))
myLoadGrd
myloadGrdRefund
myloadGrdRefund2
myloadGrdDone
myloadGrdInvoices
myloadGrdInvoicesReturn
myloadGrdReplace
End Sub
Public Sub myProc(Optional pValue As String)
End Sub
Private Sub GetOrder(sManCode As String)

Dim con As New ADODB.Connection
If Not opencn(con) Then Exit Sub


Dim aRet As Variant
Dim cString As String
cString = "SELECT TOP 1 " & _
        " vw_online_order.ORDER_NO, " & _
        " vw_online_order.DOC_NO" & _
        " FROM vw_online_order " & _
        " WHERE STAGE = 2" & _
        " AND MAN = " & MyParn(sManCode) & _
        "  ORDER BY PR_ORDER, DATE,PAYMOB DESC"

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
            " ORDER BY PR_ORDER, DATE,PAYMOB DESC"
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
myerror:
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
Public Sub myLoadGrd()
On Error GoTo myerror
With grid1
    Dim aPrm As Variant
    aPrm = AddFlag(aPrm, "DATE", sDateSales)
    Set grid1.DataSource = myRs("sp_manager_stage", , , adStoredProc, aPrm)
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
    .ColWidth(1) = 1600
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
    Set online_stage_manager = Nothing
    Set dmb = Nothing
    SaveText Me, , Array(xDate1.Name, xdate2.Name)
End Sub

Private Sub grdDone_DblClick()
If grdDone.Row > 1 Then
    online_stage_invoice_trans.sOrder_No = grdDone.TextMatrix(grdDone.Row, 0)
    online_stage_invoice_trans.sDoc_no = Val(grdDone.TextMatrix(grdDone.Row, 1))
    online_stage_invoice_trans.Show 1
End If
End Sub

Private Sub grdError_DblClick()
If grdError.Row > 1 Then
    Set oPassword = New online_stage_pass
    Set oPassword.myForm = Me
    oPassword.Show 1
End If
End Sub
Private Sub grdPrep_DblClick()
If grdError.Row > 1 Then
    Set oPassword = New online_stage_pass
    Set oPassword.myForm = Me
    oPassword.Show 1
End If
End Sub
Private Sub grdRefund_DblClick()
If grdRefund.Row > 1 Then
    online_stage_invoice_trans.sOrder_No = grdRefund.TextMatrix(grdRefund.Row, 0)
    online_stage_invoice_trans.sDoc_no = Val(grdRefund.TextMatrix(grdRefund.Row, 1))
    online_stage_invoice_trans.Show 1
End If
End Sub

Private Sub grdRefund2_DblClick()
If grdRefund2.Row > 1 Then
    online_stage_invoice_trans.sOrder_No = grdRefund2.TextMatrix(grdRefund2.Row, 0)
    online_stage_invoice_trans.sDoc_no = Val(grdRefund2.TextMatrix(grdRefund2.Row, 1))
    online_stage_invoice_trans.Show 1
End If
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
On Error GoTo myerror
Set grdPrep.DataSource = myRs(strSql)
fixgrdPrep
Exit Sub
myerror:
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
         "NULLIF(V.DOC_NO,0)," & _
         "s.desca, " & _
         "t.desca, " & _
         "FORMAT(v.[DATE],'yyyy/M/d'), " & _
         "FILE6_25.DESCA, " & _
         "FORMAT(v.TIME2,'yyyy/M/d HH:mm'), " & _
         "FORMAT(v.TIME3,'yyyy/M/d HH:mm'), " & _
         "  dbo.fn_time_diff(time2, time3) AS TimePeriod," & _
         " v.TOTAL_QUANT, " & _
         " v.TOTAL_ITEM, " & _
         " v.DISCOUNT, " & _
         " v.TOTAL_ITEM - v.DISCOUNT " & _
         " FROM vw_online_order v" & _
         " INNER JOIN ONLINE_TYPES t ON v.TYPE = t.CODE" & _
         " INNER JOIN STAGES_CODES s ON v.STAGE = s .CODE" & _
         " INNER JOIN FILE6_25 ON v.MAN = FILE6_25.CODE " & _
         " WHERE (v.STAGE = 3 OR v.STAGE = 8)" & _
         " ORDER BY STAGE DESC,DATE ,V.DOC_NO"
         

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
    .TextMatrix(1, 1) = "›—⁄Ì"
    .TextMatrix(1, 1 + 1) = "«·„—Õ·…"
    .TextMatrix(1, 2 + 1) = "‰Ê⁄ «·ÿ·»Ì…"
    .TextMatrix(1, 3 + 1) = " «—ÌŒ «·ÿ·»Ì…"
    
    .TextMatrix(1, 4 + 1) = "«·„‰œÊ»"
    .TextMatrix(1, 5 + 1) = "»œ¡ «· ÃÂÌ“"
    .TextMatrix(1, 6 + 1) = "«‰ Â«¡ «· ÃÂÌ“"
    .TextMatrix(1, 7 + 1) = "«·„œ…"
    .TextMatrix(1, 8 + 1) = "«·ﬂ„Ì…"
    .TextMatrix(1, 9 + 1) = "≈Ã„«·Ì"
    .TextMatrix(1, 10 + 1) = "Œ’„"
    .TextMatrix(1, 11 + 1) = "’«›Ì"
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .ColHidden(1) = True
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 2000
    .ColWidth(3) = 1250
    .ColWidth(4) = 2000
    .ColWidth(5) = 1700
    .ColWidth(6) = 1700
    .ColWidth(7) = 1000
    .ColWidth(8) = 700
    .ColWidth(9) = 1100
    .ColWidth(10) = 1000
    .ColWidth(11) = 1100
    
     .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub
Private Sub myloadGrdRefund()
Dim strSql As String
strSql = "SELECT v.ORDER_NO, " & _
         "V.DOC_NO," & _
         "t.desca, " & _
         "FORMAT(v.[DATE],'yyyy/M/d'), " & _
         "v.[SALES_RET], " & _
         "FORMAT(v.[DATE_MAIL],'yyyy/M/d'), " & _
         "FORMAT(v.TIME,'yyyy/M/d HH:mm'), " & _
         "dbo.fn_time_diff([time], getDate()) AS TimePeriod," & _
         " v.TOTAL_QUANT, " & _
         " v.TOTAL_ITEM, " & _
         " v.DISCOUNT, " & _
         " v.TOTAL_ITEM - v.DISCOUNT " & _
         " FROM vw_online_order v" & _
         " INNER JOIN ONLINE_TYPES t ON v.TYPE = t.CODE" & _
         " WHERE (v.STAGE = 7)" & _
         " AND (TYPE = 2 OR TYPE = 3)"
         
On Error GoTo myerror
Set grdRefund.DataSource = myRs(strSql)
fixGrdRefund
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrdRefund()
Dim i As Long
With grdRefund
For i = 0 To .Cols - 1
    .TextMatrix(0, i) = "ÿ·»Ì«  ﬁÌœ «·«” ·«„ »›« Ê—…"
    .ColAlignment(i) = flexAlignCenterCenter
Next

.TextMatrix(1, 0) = "—ﬁ„ «·ÿ·»Ì…"
.TextMatrix(1, 1) = "›—⁄Ì"
.TextMatrix(1, 2) = "‰Ê⁄ «·ÿ·»Ì…"
.TextMatrix(1, 3) = " «—ÌŒ «·ÿ·»Ì…"
.TextMatrix(1, 4) = "«·›« Ê—… «·„— Ã⁄…"
.TextMatrix(1, 5) = " «—ÌŒ «·Ê’Ê·"
.TextMatrix(1, 6) = "Êﬁ  «·ÿ·»»…"
.TextMatrix(1, 7) = "«·„œ…"
.TextMatrix(1, 8) = "«·ﬂ„Ì…"
.TextMatrix(1, 9) = "≈Ã„«·Ì"
.TextMatrix(1, 10) = "Œ’„"
.TextMatrix(1, 11) = "«·ﬁÌ„…"

.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True

.ColHidden(1) = True
.ColHidden(4) = True
.ColHidden(6) = True
.ColHidden(7) = True
.ColHidden(9) = True
.ColHidden(10) = True

.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 2500
.ColWidth(3) = 1250
.ColWidth(4) = 2000
.ColWidth(5) = 1350
.ColWidth(6) = 1400
.ColWidth(7) = 1200
.ColWidth(8) = 1000
.ColWidth(9) = 1200
.ColWidth(10) = 1000
.ColWidth(11) = 1300
 .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Private Sub myloadGrdRefund2()
Dim strSql As String
strSql = "SELECT v.ORDER_NO, " & _
         "V.DOC_NO," & _
         "t.DESCA," & _
         "FORMAT(v.[DATE],'yyyy/M/d'), " & _
         "v.[SALES_REPLACE], " & _
         "FILE6_25.DESCA, " & _
         "FORMAT(v.TIME,'yyyy/M/d HH:mm'), " & _
         "dbo.fn_time_diff([time], getDate()) AS TimePeriod," & _
         " v.TOTAL_QUANT, " & _
         " v.TOTAL_ITEM, " & _
         " v.DISCOUNT, " & _
         " v.TOTAL_ITEM - v.DISCOUNT " & _
         " FROM vw_online_order v" & _
         " INNER JOIN ONLINE_TYPES t ON v.TYPE = t.CODE" & _
         " INNER JOIN FILE6_25 ON v.MAN = FILE6_25.CODE " & _
         " WHERE (v.STAGE = 7)" & _
         " AND (TYPE = 12)"
         
On Error GoTo myerror
Set grdRefund2.DataSource = myRs(strSql)
fixGrdRefund2
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrdRefund2()
Dim i As Long
With grdRefund2
For i = 0 To .Cols - 1
    .TextMatrix(0, i) = "ÿ·»Ì«  ﬁÌœ «·«” ·«„ »œÊ‰ ›« Ê—…"
    .ColAlignment(i) = flexAlignCenterCenter
Next

.TextMatrix(1, 0) = "—ﬁ„ «·ÿ·»Ì…"
.TextMatrix(1, 1) = "›—⁄Ì"
.TextMatrix(1, 2) = "‰Ê⁄ «·ÿ·»Ì…"
.TextMatrix(1, 3) = " «—ÌŒ «·ÿ·»Ì…"
.TextMatrix(1, 4) = "«·›« Ê—… «·„” »œ·…"
.TextMatrix(1, 5) = "«·„‰œÊ»"
.TextMatrix(1, 6) = "Êﬁ  «·ÿ·»»…"
.TextMatrix(1, 7) = "«·„œ…"
.TextMatrix(1, 8) = "«·ﬂ„Ì…"
.TextMatrix(1, 9) = "≈Ã„«·Ì"
.TextMatrix(1, 10) = "Œ’„"
.TextMatrix(1, 11) = "’«›Ì"

.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True

.ColHidden(1) = True
.ColHidden(2) = True

.ColHidden(4) = True
.ColHidden(9) = True
.ColHidden(10) = True
.ColHidden(11) = True

.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 1500
.ColWidth(3) = 1250
.ColWidth(4) = 1700
.ColWidth(5) = 1800
.ColWidth(6) = 1700
.ColWidth(7) = 1000
.ColWidth(8) = 700
.ColWidth(9) = 1100
.ColWidth(10) = 1000
.ColWidth(11) = 1100

 .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Private Sub myloadGrdInvoices()
Dim strSql As String
strSql = "SELECT FILE6_20H.DOC_NO," & _
         "FILE6_20H.ONLINE_DOC," & _
         "ONLINE_TYPES.DESCA, " & _
         "m.DESCA," & _
         "FILE6_20H.TOTAL_QUANT," & _
         "FILE6_20H.TOTAL_ITEM," & _
         "FILE6_20H.discount, " & _
         "FILE6_20H.TOTAL_ITEM - FILE6_20H.discount " & _
         "FROM FILE6_20H " & _
         "INNER JOIN ONLINE_TYPES ON FILE6_20H.INV_TYPE_ONLINE = ONLINE_TYPES.CODE " & _
         "LEFT JOIN FILE6_25 M ON FILE6_20H.MAN = m.CODE " & _
         "WHERE IS_ONLINE = 1 AND IS_RETURN = 0 " & _
         "AND DATE = " & DateSq(sDateSales)
On Error GoTo myerror
Set grdInvoices.DataSource = myRs(strSql)
fixGrdInvoices
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrdInvoices()
Dim i As Long
With grdInvoices
For i = 0 To .Cols - 1
    .TextMatrix(0, i) = "≈Ã„«·Ì ›Ê« Ì— „»Ì⁄« "
    .ColAlignment(i) = flexAlignCenterCenter
Next

.TextMatrix(1, 0) = "—ﬁ„ «·›« Ê—…"
.TextMatrix(1, 1) = "«·ÿ·»Ì…"
.TextMatrix(1, 2) = "‰Ê⁄ «·ÿ·»Ì…"
.TextMatrix(1, 3) = "«·„‰œÊ»"
.TextMatrix(1, 4) = "«·ﬂ„Ì…"
.TextMatrix(1, 5) = "≈Ã„«·Ì «·«’‰«›"
.TextMatrix(1, 6) = "«·Œ’„"
.TextMatrix(1, 7) = "«·≈Ã„«·Ì"

.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True

.ColWidth(0) = 1800
.ColWidth(1) = 1000
.ColWidth(2) = 1500
.ColWidth(3) = 1800
.ColWidth(4) = 800
.ColWidth(5) = 1100
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColHidden(5) = True
.ColHidden(6) = True


 .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Private Sub myloadGrdInvoicesReturn()
Dim strSql As String
strSql = "SELECT FILE6_20H.DOC_NO," & _
         "FILE6_20H.ONLINE_DOC," & _
         "ONLINE_TYPES.DESCA, " & _
         "m.DESCA," & _
         "FILE6_20H.SALES_RET," & _
         "FILE6_20H.TOTAL_QUANT," & _
         "FILE6_20H.TOTAL_ITEM," & _
         "FILE6_20H.discount, " & _
         "FILE6_20H.TOTAL_ITEM - FILE6_20H.discount " & _
         "FROM FILE6_20H " & _
         "INNER JOIN ONLINE_TYPES ON FILE6_20H.INV_TYPE_ONLINE = ONLINE_TYPES.CODE " & _
         "LEFT JOIN FILE6_25 M ON FILE6_20H.MAN = m.CODE " & _
         "WHERE IS_ONLINE = 1 AND IS_RETURN = 1 " & _
         "AND DATE = " & DateSq(sDateSales)

On Error GoTo myerror
Set grdInvoicesReturn.DataSource = myRs(strSql)
fixGrdInvoicesReturn
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrdInvoicesReturn()
Dim i As Long
With grdInvoicesReturn
For i = 0 To .Cols - 1
    .TextMatrix(0, i) = "≈Ã„«·Ì ›Ê« Ì— „— Ã⁄…"
    .ColAlignment(i) = flexAlignCenterCenter
Next

.TextMatrix(1, 0) = "—ﬁ„ «·›« Ê—…"
.TextMatrix(1, 1) = "«·ÿ·»Ì…"
.TextMatrix(1, 2) = "‰Ê⁄ «·ÿ·»Ì…"
.TextMatrix(1, 3) = "«·„‰œÊ»"
.TextMatrix(1, 4) = "›« Ê—… «·„— Ã⁄"
.TextMatrix(1, 5) = "«·ﬂ„Ì…"
.TextMatrix(1, 6) = "≈Ã„«·Ì «·«’‰«›"
.TextMatrix(1, 7) = "«·Œ’„"
.TextMatrix(1, 8) = "«·≈Ã„«·Ì"

.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True

.ColWidth(0) = 1800
.ColWidth(1) = 1000
.ColWidth(2) = 1800
.ColWidth(3) = 2000
.ColWidth(4) = 1500
.ColWidth(5) = 800
.ColWidth(5 + 1) = 1100
.ColWidth(6 + 1) = 1100
.ColWidth(7 + 1) = 1100
.ColHidden(5 + 1) = True
.ColHidden(6 + 1) = True


 .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Private Sub myloadGrdReplace()
Dim strSql As String
strSql = "SELECT v.ORDER_NO," & _
         "V.DOC_NO," & _
         "t.DESCA," & _
         "FORMAT(v.[DATE],'yyyy/M/d'), " & _
         "v.[SALES_REPLACE], " & _
         "FILE6_25.DESCA, " & _
         "v.TOTAL_QUANT, " & _
         "v.TOTAL_ITEM - v.DISCOUNT " & _
         "FROM FILE6_90BH v " & _
         "INNER JOIN ONLINE_TYPES t ON v.TYPE = t.CODE " & _
         "LEFT JOIN FILE6_25 ON v.MAN = FILE6_25.CODE " & _
         "WHERE (TYPE = 11 OR TYPE = 12) " & _
         "AND v.DATE1 = " & DateSq(sDateSales)
On Error GoTo myerror
Set grdReplace.DataSource = myRs(strSql)
fixGrdReplace
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrdReplace()
Dim i As Long
With grdReplace
For i = 0 To .Cols - 1
    .TextMatrix(0, i) = "ÿ·»Ì… «” »œ«· Ê«” —Ã«⁄ »œÊ‰ ›« Ê—…"
    .ColAlignment(i) = flexAlignCenterCenter
Next

.TextMatrix(1, 0) = "—ﬁ„ «·ÿ·»Ì…"
.TextMatrix(1, 1) = "›—⁄Ì"
.TextMatrix(1, 2) = "‰Ê⁄ «·ÿ·»Ì…"
.TextMatrix(1, 3) = " «—ÌŒ «·ÿ·»Ì…"
.TextMatrix(1, 4) = "›« Ê—… «·«” »œ«·"
.TextMatrix(1, 5) = "«·„‰œÊ»"
.TextMatrix(1, 6) = "«·ﬂ„Ì…"
.TextMatrix(1, 7) = "«·ﬁÌ„…"

.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True

.ColHidden(1) = True

.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 1500
.ColWidth(3) = 1250
.ColWidth(4) = 2000
.ColWidth(5) = 2000
.ColWidth(6) = 700
.ColWidth(7) = 1100
 .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
End With
End Sub

