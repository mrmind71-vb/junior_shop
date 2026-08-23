VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form online_dash_board 
   BackColor       =   &H00FFFFFF&
   Caption         =   "dashboard ONLINE"
   ClientHeight    =   11055
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   22920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   22920
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "«· «—ÌŒ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   4605
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   315
         Width           =   1230
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   315
         Width           =   1230
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   420
         Left            =   45
         TabIndex        =   3
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   741
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
         Picture         =   "online_dash_board.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdGo 
         Default         =   -1  'True
         Height          =   420
         Left            =   990
         TabIndex        =   2
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   741
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
         Picture         =   "online_dash_board.frx":23A0
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1710
      Top             =   -630
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
      Left            =   1710
      Top             =   -1035
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
   Begin MSAdodcLib.Adodc data6 
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
   Begin MSComctlLib.ProgressBar PROG1 
      Align           =   2  'Align Bottom
      Height          =   210
      Left            =   0
      TabIndex        =   5
      Top             =   10845
      Visible         =   0   'False
      Width           =   22920
      _ExtentX        =   40428
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grdTotal 
      Height          =   3345
      Left            =   4725
      TabIndex        =   6
      Top             =   45
      Width           =   18060
      _cx             =   31856
      _cy             =   5900
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   16
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdShip 
      Height          =   3975
      Left            =   45
      TabIndex        =   7
      Top             =   3420
      Width           =   7170
      _cx             =   12647
      _cy             =   7011
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   2580
      Left            =   45
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   810
      Width           =   4650
      _cx             =   8202
      _cy             =   4551
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
      Cols            =   4
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
   Begin VSFlex7Ctl.VSFlexGrid grdModel 
      Height          =   4155
      Left            =   7245
      TabIndex        =   9
      Top             =   7425
      Width           =   7755
      _cx             =   13679
      _cy             =   7329
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdCity 
      Height          =   3975
      Left            =   7245
      TabIndex        =   10
      Top             =   3420
      Width           =   7755
      _cx             =   13679
      _cy             =   7011
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdDays 
      Height          =   4155
      Left            =   45
      TabIndex        =   11
      Top             =   7425
      Width           =   7170
      _cx             =   12647
      _cy             =   7329
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdMan 
      Height          =   3975
      Left            =   15030
      TabIndex        =   12
      Top             =   3420
      Width           =   7755
      _cx             =   13679
      _cy             =   7011
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
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
   Begin VSFlex7Ctl.VSFlexGrid grdMan2 
      Height          =   3975
      Left            =   15030
      TabIndex        =   13
      Top             =   7425
      Width           =   7755
      _cx             =   13679
      _cy             =   7011
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
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
End
Attribute VB_Name = "online_dash_board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_excel_Click()
    ToFileExel2 GridTotal, , , , , 1.1, , , , , , Me
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    If Not myValid Then Exit Sub
    Dim db As New clsDb
    myLoadGrdTotal db
    myloadgrdPrep db
    myLoadGrdShip db
    myloadGrdModel db
    myloadGrdCity db
    myLoadGrdDays db
    myLoadGrdMan db
    myLoadGrdMan2 db
    Set db = Nothing
End Sub
Private Function myValid() As Boolean
If Not IsDate(xDate1.text) Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— „”Ã·"
    Exit Function
End If
If Not IsDate(xdate2.text) Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— „”Ã·"
    Exit Function
End If
myValid = True
End Function
Private Sub MYLOAD1()

End Sub
Private Sub Form_Load()
xDate1.text = myFormat_p(Year(Date) & "-" & Month(Date) & "-" & "01")
xdate2.text = myFormat_p(DateAdd("d", -1, myFormat(DateAdd("m", 1, myFormat(xDate1.text)))))
fixGrdTotal
Fixgrd
FixgrdShip
fixGrdModel
fixGrdCity
fixGrdDays
fixGrdMan
fixGrdMan2
End Sub
Private Sub myLoadGrdTotal(db As clsDb)
Dim aPrm As Variant
If IsDate(xDate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
    aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
End If

Set grdTotal.DataSource = db.myRs("[dbo].[sp_dash_board_total]", adStoredProc, aPrm)

fixGrdTotal
End Sub
Private Sub fixGrdTotal()
With grdTotal
.FixedCols = 0
.FixedRows = 2

.WordWrap = True
.TextMatrix(0, 0) = "‰Ê⁄ «·ÿ·»Ì…"
.ColWidth(0) = 2000
.TextMatrix(1, 0) = "‰Ê⁄ «·ÿ·»Ì…"

Dim i As Long
For i = 1 To 4
    .TextMatrix(0, i) = "ÿ·»Ì« "
Next

.TextMatrix(1, 1) = "⁄œœ"
.ColWidth(1) = 1000
.ColFormat(1) = "#,##0"
.TextMatrix(1, 2) = "ﬁÿ€"
.ColWidth(2) = 1000
.ColFormat(2) = "#,##0"
.TextMatrix(1, 3) = "≈Ã„«·Ï"
.ColWidth(3) = 1500
.ColFormat(3) = "#,##0"
.TextMatrix(1, 4) = "„ Ê”ÿ"

For i = 5 To 9
    .TextMatrix(0, i) = "«·„»«⁄"
Next
.TextMatrix(1, 5) = "⁄œœ"
.ColWidth(5) = 1000
.ColFormat(5) = "#,##0"
.TextMatrix(1, 6) = "ﬁÿ⁄"
.ColWidth(6) = 1000
.ColFormat(6) = "#,##0"
.TextMatrix(1, 7) = "ﬁÌ„…"
.ColWidth(7) = 1500
.ColFormat(7) = "#,##0"
.TextMatrix(1, 8) = "‰”Ì…"
.ColWidth(8) = 1000
.ColFormat(8) = "0.00%"
.TextMatrix(1, 9) = "„ Ê”ÿ"
.ColWidth(9) = 1000

For i = 10 To 14
    .TextMatrix(0, i) = "€Ì— „»«⁄"
Next
.TextMatrix(1, 10) = "⁄œœ"
.ColWidth(10) = 1000
.ColFormat(10) = "#,##0"
.TextMatrix(1, 11) = "ﬁÿ⁄"
.ColWidth(11) = 1000
.ColFormat(11) = "#,##0"
.TextMatrix(1, 12) = "ﬁÌ„…"
.ColWidth(12) = 1500
.ColFormat(12) = "#,##0"
.TextMatrix(1, 13) = "‰”Ì…"
.ColWidth(13) = 1000
.ColFormat(13) = "0.00%"
.TextMatrix(1, 14) = "„ Ê”ÿ"
.ColWidth(14) = 1000

.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True

.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True
.MergeCol(0) = True
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
.Cell(flexcpAlignment, 1, 0, 1, .Cols - 1) = 4


For i = 2 To .Rows - 1
    If .ValueMatrix(i, .Cols - 1) = 2 Or .ValueMatrix(i, .Cols - 1) = 4 Then
        .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE0E0E0
    End If
    If .ValueMatrix(i, .Cols - 2) = 2 Or .ValueMatrix(i, .Cols - 2) = 3 Then
        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
    End If
Next

.WordWrap = True
End With
End Sub
Public Sub myloadgrdPrep(db As clsDb)
With grid1
    Dim aPrm As Variant
    aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
    aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
        
        
    Dim locTable As New ADODB.RecordSet
    Set locTable = db.myRs("[dbo].[sp_dash_stage]", adStoredProc, aPrm)
    
    Dim Col As Long
    Dim Row As Long
    Dim nTotal As Long
    
    .Rows = 2
    Row = 1
    Do Until locTable.EOF
        If Row > 6 And Col = 0 Then
            Col = 2
            Row = 1
        End If
        
        Row = Row + 1
        
        If Row > .Rows - 1 Then
            .AddItem ""
        End If
        
        .TextMatrix(Row, Col) = locTable!DESCA
        .TextMatrix(Row, Col + 1) = locTable!count_order
        nTotal = nTotal + .ValueMatrix(Row, Col + 1)
        
        locTable.MoveNext
    Loop

    .ColHidden(2) = Col = 0
    .ColHidden(3) = Col = 0
    
    .ColWidth(0) = IIf(Col = 0, 2500, 1450)
    .ColWidth(1) = IIf(Col = 0, 1000, 700)
    
    If .Rows > 3 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = "«·≈Ã„«·Ì"
        .TextMatrix(.Rows - 1, Col + 1) = nTotal
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &H8000000F
    End If
End With
End Sub
Sub Fixgrd()
With grid1
    .TextMatrix(0, 0) = " ÃÂÌ“ ÿ·»Ì« "
    .TextMatrix(0, 1) = " ÃÂÌ“ ÿ·»Ì« "
    .TextMatrix(0, 2) = " ÃÂÌ“ ÿ·»Ì« "
    
    .TextMatrix(1, 0) = "«·„—Õ·…"
    .TextMatrix(1, 1) = "«·⁄œœ"
    
    .TextMatrix(1, 2) = "«·„—Õ·…"
    .TextMatrix(1, 3) = "«·⁄œœ"
    
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    .RowHidden(0) = True
    
    .ColWidth(0) = 2500
    .ColWidth(1) = 1000
    
    .ColWidth(2) = 1450
    .ColWidth(3) = 700
    
    .ColHidden(2) = True
    .ColHidden(3) = True
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    .SubtotalPosition = flexSTBelow
    
    Dim i As Long
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    End With
End Sub
Public Sub myLoadGrdShip(db As clsDb)
On Error GoTo myerror
With grid1
    Dim aPrm As Variant
    aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
    aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
    Set grdShip.DataSource = db.myRs("[dbo].[sp_dash_board_ship]", adStoredProc, aPrm)
End With
FixgrdShip
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub FixgrdShip()
With grdShip
    For i = 0 To .Cols - 1
        .TextMatrix(0, i) = "»Ê«·’ «·‘Õ‰"
        .ColWidth(i) = 910
    Next
    .ColWidth(0) = 1300
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .RowHeight(1) = 800
                
    .TextMatrix(1, 0) = "«·ÌÊ„"
    .TextMatrix(1, 1) = "⁄œœ «·»Ê«·’"
    .TextMatrix(1, 2) = "»Ê«·’ „€·ﬁ…"
    .TextMatrix(1, 3) = "»Ê«·’ „› ÊÕ…"
    .TextMatrix(1, 4) = "» «—ÌŒ ‘Õ‰"
    .TextMatrix(1, 5) = "»œÊ‰ ‘Õ‰"
    .TextMatrix(1, 6) = "„€·ﬁ… »œÊ‰ ‘Õ‰"
    
    .ColDataType(0) = flexDTDate
    For i = 1 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    .SubtotalPosition = flexSTBelow
    
    For i = 0 To .Cols - 1
        If .ColDataType(i) = flexDTDate Then
            .ColFormat(i) = "yyyy/m/d"
        ElseIf .ColDataType(i) = flexDTDouble Then
            .Subtotal flexSTSum, -1, i, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
        End If
    Next
        
    If .Rows > 2 Then
        .TextMatrix(2, 0) = "„« ﬁ»·Â"
    End If
End With
End Sub
Public Sub myLoadGrdDays(db As clsDb)
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
Set grdDays.DataSource = db.myRs("[dbo].[sp_dash_board_days]", adStoredProc, aPrm)
fixGrdDays
End Sub
Sub fixGrdDays()
With grdDays
    For i = 0 To .Cols - 1
        .TextMatrix(0, i) = "≈Ã„«·Ì ÿ·»Ì«  ÌÊ„Ì…"
        .ColWidth(i) = 910
    Next
    .ColWidth(0) = 1300
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .RowHeight(1) = 600
                
    .TextMatrix(1, 0) = "«·ÌÊ„"
    .TextMatrix(1, 1) = "⁄œœ «·ÿ·»Ì« "
    .TextMatrix(1, 2) = "ÿ·»Ì«  „€·ﬁ…"
    .TextMatrix(1, 3) = "ÿ·»Ì«  „› ÊÕ…"
    .TextMatrix(1, 4) = "ÿ·»Ì«  „·€Ì…"
    .TextMatrix(1, 5) = "ÿ·»Ì«  „‘ÕÊ‰…"
    .TextMatrix(1, 6) = "€Ì— „‘ÕÊ‰…"
    
    .ColDataType(0) = flexDTDate
    For i = 1 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    .SubtotalPosition = flexSTBelow
    
    For i = 0 To .Cols - 1
        If .ColDataType(i) = flexDTDate Then
            .ColFormat(i) = "yyyy/m/d"
        ElseIf .ColDataType(i) = flexDTDouble Then
            .Subtotal flexSTSum, -1, i, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
        End If
    Next
End With
End Sub
Public Sub myloadGrdModel(db As clsDb)
With grid1
    Dim aPrm As Variant
    aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
    aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
    Set grdModel.DataSource = db.myRs("[dbo].[sp_dash_board_model]", adStoredProc, aPrm)
End With
fixGrdModel
End Sub
Sub fixGrdModel()
With grdModel
    For i = 0 To .Cols - 1
        .TextMatrix(0, i) = "„ÊœÌ·«  «·ÿ·»Ì« "
    Next
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 700
    .ColWidth(2) = 1500
    .ColWidth(3) = 2400
    .ColWidth(4) = 1000
    .ColWidth(5) = 800
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    '.RowHeight(0) = 450
                    
    .TextMatrix(1, 0) = "«·„’‰⁄"
    .TextMatrix(1, 1) = "«·„Ê”„"
    .TextMatrix(1, 2) = "„ÊœÌ· «·„’‰⁄"
    .TextMatrix(1, 3) = "«·’‰›"
    .TextMatrix(1, 4) = "«·ﬂ„Ì…"
    .TextMatrix(1, 5) = "«·‰”»…"
    
    For i = 4 To 4
        .ColDataType(i) = flexDTDouble
    Next
    
    
    .ColFormat(5) = "0.00%"
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    .SubtotalPosition = flexSTBelow
    
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
        If .ColDataType(i) = flexDTDate Then
            .ColFormat(i) = "yyyy/m/d"
        ElseIf .ColDataType(i) = flexDTDouble Then
            .Subtotal flexSTSum, -1, i, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
        End If
    Next
End With
End Sub
Public Sub myloadGrdCity(db As clsDb)
Dim strSql As String
strSql = "SELECT  f.Shipping_City ," & _
          "COUNT(*)," & _
          "SUM(CASE WHEN TYPE <= 1 THEN 1 ELSE -1 END * v.TOTAL_QUANT)," & _
          "SUM(CASE WHEN TYPE <= 1 THEN 1 ELSE -1 END *  (v.TOTAL_ITEM - V.DISCOUNT))" & _
          " FROM vw_online_orders v" & _
          " INNER JOIN FILE6_90H F ON v.ORDER_NO = F.DOC_NO" & _
          " WHERE   v.DATE >= " & DateSq(xDate1.text) & _
          " AND     v.DATE <= " & DateSq(xdate2.text) & _
          " AND f.CANCELED = 0" & _
          " AND v.TYPE <= 3" & _
          " GROUP BY " & _
          " f.Shipping_City " & _
          " ORDER BY COUNT(*) DESC"
Set grdCity.DataSource = db.myRs(strSql)
fixGrdCity
End Sub
Sub fixGrdCity()
With grdCity
For i = 0 To .Cols - 1
    .TextMatrix(0, i) = "„Õ«›Ÿ«  «·ÿ·»Ì« "
Next
.Cols = 5

.TextMatrix(1, 0) = "«·„Õ«›Ÿ…"
.TextMatrix(1, 1) = "⁄œœ"
.TextMatrix(1, 2) = "ﬁÿ€"
.TextMatrix(1, 3) = "≈Ã„«·Ï ﬁÌ„…"
.TextMatrix(1, 4) = "‰”Ì…"

    
.ColFormat(1) = "#,##0"
.ColFormat(2) = "#,##0"
.ColFormat(3) = "#,##0"
    
.ColWidth(0) = 3000
.ColWidth(1) = 1000
.ColWidth(2) = 1000
.ColWidth(3) = 1500
.ColWidth(4) = 900
    
.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True
    
    
.ExplorerBar = flexExSort
.Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
.SubtotalPosition = flexSTBelow
    
.ColDataType(1) = flexDTDouble
.ColDataType(2) = flexDTDouble
.ColDataType(3) = flexDTDouble
.ColFormat(4) = "0.00%"

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
    If .ColDataType(i) = flexDTDouble Then
        .Subtotal flexSTSum, -1, i, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    End If
Next

For i = 2 To .Rows - 2
    If .ValueMatrix(.Rows - 1, 1) = 0 Then Exit For
    .TextMatrix(i, 4) = Round(.ValueMatrix(i, 1) / .ValueMatrix(.Rows - 1, 1), 4)
Next
End With
End Sub

Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Sub fixGrdMan()
With grdMan
    For i = 0 To .Cols - 1
        .TextMatrix(0, i) = "≈‰Ã«“ «·ÿ·»Ì«  ··„ÊŸ›Ì‰"
        .ColWidth(i) = 900
    Next
    .ColWidth(0) = 1300
    .ColWidth(1) = 2000
    .ColHidden(0) = True
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .RowHeight(1) = 600
                
    .TextMatrix(1, 0) = "ﬂÊœ «·„ÊŸ›"
    .TextMatrix(1, 1) = "≈”„ «·„ÊŸ›"
    .TextMatrix(1, 2) = "⁄œœ «·ÿ·»Ì« "
    .TextMatrix(1, 3) = " „ «· ÃÂÌ“"
    .TextMatrix(1, 4) = "»Â« „‘ﬂ·…"
    .TextMatrix(1, 5) = "„⁄·ﬁ…"
    .TextMatrix(1, 6) = " „ «· ÃÂÌ“"
    .TextMatrix(1, 7) = "„·€«…"
    
    .ColDataType(0) = flexDTDate
    For i = 2 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    .SubtotalPosition = flexSTBelow
    
    For i = 0 To .Cols - 1
        If .ColDataType(i) = flexDTDouble Then
            .Subtotal flexSTSum, -1, i, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
        End If
    Next
    If .Rows > 2 Then
        .TextMatrix(.Rows - 1, 1) = "«·≈Ã„«·Ì"
    End If
End With
End Sub
Public Sub myLoadGrdMan(db As clsDb)
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
Set grdMan.DataSource = db.myRs("[dbo].[sp_dash_board_man1]", adStoredProc, aPrm)
fixGrdMan
End Sub
Sub fixGrdMan2()
With grdMan2
    For i = 0 To .Cols - 1
        .TextMatrix(0, i) = "≈‰Ã«“ «·ÿ·»Ì«  ··„ÊŸ›Ì‰  Õ· «·„‘ﬂ·…"
        .ColWidth(i) = 900
    Next
    .ColWidth(0) = 1300
    .ColWidth(1) = 4000
    .ColHidden(0) = True
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .RowHeight(1) = 600
                
    .TextMatrix(1, 0) = "ﬂÊœ «·„ÊŸ›"
    .TextMatrix(1, 1) = "≈”„ «·„ÊŸ›"
    .TextMatrix(1, 2) = " „ «· ÕÂÌ“"
    .TextMatrix(1, 3) = "„·€«…"
    
    .ColDataType(0) = flexDTDate
    For i = 2 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
    .SubtotalPosition = flexSTBelow
    
    For i = 0 To .Cols - 1
        If .ColDataType(i) = flexDTDouble Then
            .Subtotal flexSTSum, -1, i, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
        End If
    Next
    If .Rows > 2 Then
        .TextMatrix(.Rows - 1, 1) = "«·≈Ã„«·Ì"
    End If
End With
End Sub
Public Sub myLoadGrdMan2(db As clsDb)
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
Set grdMan2.DataSource = db.myRs("[dbo].[sp_dash_board_man2]", adStoredProc, aPrm)
fixGrdMan2
End Sub


