VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form ShowStockDiffer 
   Caption         =   " ”ÊÌ… ⁄Ã“"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20145
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
   ScaleHeight     =   9180
   ScaleWidth      =   20145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   -90
      Picture         =   "showStockDiffer.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2430
      Top             =   8055
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   3
      Top             =   9030
      Visible         =   0   'False
      Width           =   20145
      _ExtentX        =   35534
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   -90
      Top             =   0
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Left            =   -90
      Top             =   0
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   -90
      Top             =   0
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   8025
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   19995
      _ExtentX        =   35269
      _ExtentY        =   14155
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "«·„’‰⁄"
      TabPicture(0)   =   "showStockDiffer.frx":242A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "GRID4"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "«·ﬁ”„"
      TabPicture(1)   =   "showStockDiffer.frx":2446
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRID3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "«·„ÊœÌ·« "
      TabPicture(2)   =   "showStockDiffer.frx":2462
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "GRID2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "«·«’‰«›"
      TabPicture(3)   =   "showStockDiffer.frx":247E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grid1"
      Tab(3).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   7575
         Left            =   -74910
         TabIndex        =   5
         Top             =   360
         Width           =   19815
         _cx             =   34951
         _cy             =   13361
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
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
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Height          =   7575
         Left            =   90
         TabIndex        =   6
         Top             =   360
         Width           =   19815
         _cx             =   34951
         _cy             =   13361
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
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
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid GRID3 
         Height          =   7575
         Left            =   -74910
         TabIndex        =   8
         Top             =   360
         Width           =   19815
         _cx             =   34951
         _cy             =   13361
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
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
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid GRID4 
         Height          =   7575
         Left            =   -74910
         TabIndex        =   9
         Top             =   360
         Width           =   19815
         _cx             =   34951
         _cy             =   13361
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
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
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame Frame12 
      Height          =   690
      Left            =   6075
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   8190
      Width           =   14010
      Begin VB.TextBox txtBarcode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   11610
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   225
         Width           =   1500
      End
      Begin VB.TextBox txtModel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8190
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   225
         Width           =   2040
      End
      Begin Threed.SSCommand cmdSection 
         Height          =   375
         Left            =   4365
         TabIndex        =   15
         Top             =   225
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   661
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
         Caption         =   "... «Œ «— «·ﬁ”„"
         TagVariant      =   "... «Œ «— «·ﬁ”„"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFact 
         Height          =   375
         Left            =   90
         TabIndex        =   16
         Top             =   225
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   661
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
         Caption         =   "... «Œ «— «·„’‰⁄"
         TagVariant      =   "... «Œ «— «·„’‰⁄"
         ButtonStyle     =   3
      End
      Begin VB.Label Label6 
         Caption         =   "»«—ﬂÊœ"
         Height          =   375
         Left            =   13275
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label7 
         Caption         =   "—ﬁ„ „ÊœÌ·"
         Height          =   375
         Left            =   10305
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   870
      End
   End
   Begin VB.Frame Frame2 
      Height          =   780
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   8100
      Width           =   2535
      Begin VB.CommandButton cmdExel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1260
         Picture         =   "showStockDiffer.frx":249A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   45
         Picture         =   "showStockDiffer.frx":4C85
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   1185
      End
   End
End
Attribute VB_Name = "ShowStockDiffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String, sDate As String, sStore As String
Public con As adodb.Connection
Dim oSearchSection As New Search_abd
Dim oSearchFact As New Search_abd
Private Sub cmdExel_Click()
Me.MousePointer = 11

Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aRow = AddFlag(aRow, aSub)

If SSTab1.Tab = 3 Then
    If grid1.Rows > 1 Then
        aSub = AddFlag(Empty, "row", 1)
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "word_wrap", False)
        aSub = AddFlag(aSub, "back_color", 40)
        aRow = AddFlag(aRow, aSub)
    End If
    ToFileExelNew grid1, , , aRow, , 0.9, , , , , , Me, Array(Me.Caption, SSTab1.TabCaption(SSTab1.Tab), IIf(cmdSection.Tag <> "", "«·ﬁ”„ : " & cmdSection.Caption, ""), IIf(cmdFact.Tag <> "", "«·„’‰⁄ : " & cmdFact.Caption, ""))
ElseIf SSTab1.Tab = 2 Then
    If grid2.Rows > 1 Then
        aSub = AddFlag(Empty, "row", 1)
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "word_wrap", False)
        aSub = AddFlag(aSub, "back_color", 40)
        aRow = AddFlag(aRow, aSub)
    End If
    ToFileExelNew grid2, , , aRow, , 0.9, , , , , , Me, Array(Me.Caption, SSTab1.TabCaption(SSTab1.Tab), IIf(cmdSection.Tag <> "", "«·ﬁ”„ : " & cmdSection.Caption, ""), IIf(cmdFact.Tag <> "", "«·„’‰⁄ : " & cmdFact.Caption, ""))
ElseIf SSTab1.Tab = 1 Then
    If GRID3.Rows > 1 Then
        aSub = AddFlag(Empty, "row", 1)
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "word_wrap", False)
        aSub = AddFlag(aSub, "back_color", 40)
        aRow = AddFlag(aRow, aSub)
    End If
    ToFileExelNew GRID3, , , aRow, , 0.9, , , , , , Me, Array(Me.Caption, SSTab1.TabCaption(SSTab1.Tab), IIf(cmdSection.Tag <> "", "«·ﬁ”„ : " & cmdSection.Caption, ""), IIf(cmdFact.Tag <> "", "«·„’‰⁄ : " & cmdFact.Caption, ""))
ElseIf SSTab1.Tab = 0 Then
    If GRID4.Rows > 1 Then
        aSub = AddFlag(Empty, "row", 1)
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "word_wrap", False)
        aSub = AddFlag(aSub, "back_color", 40)
        aRow = AddFlag(aRow, aSub)
    End If
    ToFileExelNew GRID4, , , aRow, , 0.9, , , , , , Me, Array(Me.Caption, SSTab1.TabCaption(SSTab1.Tab), IIf(cmdSection.Tag <> "", "«·ﬁ”„ : " & cmdSection.Caption, ""), IIf(cmdFact.Tag <> "", "«·„’‰⁄ : " & cmdFact.Caption, ""))
End If
Me.MousePointer = 0
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdPrint_Click()
'Dim aRow(0) As Variant
'aRow(0) = AddFlag(Empty, "row", grid1.Rows - 1)
'aRow(0) = AddFlag(aRow(0), "col", 0)
'aRow(0) = AddFlag(aRow(0), "cols", 2)
'printGrdNew.doprint Me.grid1, 0.8, -1, "ÿ»«⁄…  ”ÊÌ«  Ã—œ —ﬁ„ : " & Val(sDoc_no), sStore & " » «—ÌŒ : " & myFormat_p(sDate), , , False, False, 10, , aRow
'printGrdNew.Show 1
End Sub
Private Sub cmdSection_Click()
SectionLookup Me, oSearchSection, "CODE IN (SELECT FILE1_10.SECTION FROM FILE0_10 INNER JOIN FILE1_10 ON FILE0_10.ITEM = FILE1_10.ITEM WHERE FILE0_10.DOC_NO = " & MyParn(sDoc_no) & ")", , IIf(cmdSection.Tag <> "", "ﬂ· «·«ﬁ”«„", "")
End Sub
Private Sub cmdfact_Click()
FactLookup Me, oSearchFact, "CODE IN (SELECT FILE1_10.FACT FROM FILE0_10 INNER JOIN FILE1_10 ON FILE0_10.ITEM = FILE1_10.ITEM WHERE FILE0_10.DOC_NO = " & MyParn(sDoc_no) & ")", , IIf(cmdFact.Tag <> "", "ﬂ· «·„’«‰⁄", "")
End Sub
Private Sub Form_Load()
Me.Caption = " ”ÊÌ… „” ‰œ  —ﬁ„ : " & sDoc_no
Set grid1.DataSource = DATA1
Set grid2.DataSource = DATA2
Set GRID3.DataSource = data3
Set GRID4.DataSource = data4
myload
MYLOAD2
MYLOAD3
MYLOAD4
End Sub
Private Sub myload()
Dim cString As String
cString = "SELECT Row_Number() over(order by FILE1_10.MODEL,FILE1_10.C_SCAL)," & _
          "FILE0_10.item ," & _
          "file1_10.mosm ," & _
          "fact.desca," & _
          "FILE1_10SC.DESCA," & _
          "file1_10.modelfact0," & _
          "file1_10.desca," & _
          "file1_10.scal," & _
          "file1_10.color," & _
          "file0_10.ComputerBal," & _
          "file0_10.realBal," & _
          "file0_10.differ," & _
          "ROUND(file0_10.differ * file0_10.cost,2), " & _
          "ROUND(file0_10.differ * file1_10.price,2)" & _
          " from FILE0_10 " & _
          " inner join file1_10 on file1_10.item = FILE0_10.item" & _
          " inner join fact on fact.code = file1_10.fact" & _
          " inner join FILE1_10SC on FILE1_10.SECTION = file1_10SC.CODE" & _
          " WHERE  DOC_NO = " & MyParn(sDoc_no) & _
          " AND FILE0_10.IS_DIFFER = 1"

If cmdSection.Tag <> "" Then
    cString = cString & " AND FILE1_10.[SECTION] = " & cmdSection.Tag
End If

If cmdFact.Tag <> "" Then
    cString = cString & " AND FILE1_10.[FACT] = " & cmdFact.Tag
End If

cString = cString & " order by FILE1_10.MODEL,FILE1_10.C_SCAL"
'Dim con As New ADODB.Connection

Set grid1.DataSource = DATA1
Set DATA1.Recordset = cmd(cString, con).Execute
fixGrd
End Sub
Private Sub fixGrd()
With grid1
.RowHeight(0) = 600
'                   0      1           2         3         4          5            6          7          8             9              10              11              12              13
.FormatString = "„|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "ﬁ”„|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "—’Ìœ ﬂ„»ÌÊ —|" & "—’Ìœ Ã—œ|" & " ”ÊÌ… Ã—œ|" & "»”⁄— «· ﬂ·›…|" & "»”⁄— «·»Ì⁄"

.RowHeight(0) = 600
.ColWidth(0) = 1000
.ColWidth(1) = 1100
.ColWidth(2) = 1000
.ColWidth(3) = 1400
.ColWidth(4) = 1800
.ColWidth(5) = 1500
.ColWidth(6) = 3000
.ColWidth(7) = 1100
.ColWidth(8) = 1100
.ColWidth(9) = 1100
.ColWidth(10) = 1100
.ColWidth(11) = 1100
.ColWidth(12) = 1500
.ColWidth(13) = 1600


.ColDataType(9) = flexDTDouble
.ColDataType(10) = flexDTDouble
.ColDataType(11) = flexDTDouble
.ColDataType(12) = flexDTDouble
.ColDataType(13) = flexDTDouble

'.ColHidden(4) = True
.ColHidden(12) = Not (cBranch = "00" And bOpt10)
.ColHidden(13) = Not (cBranch = "00" And bOpt10)

.ExplorerBar = flexExSortShow
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

.SubtotalPosition = flexSTAbove
.Subtotal flexSTSum, -1, 9, "#0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 10, "#0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 11, "#0", , &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 12, "#0,00", , &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 13, "#0,00", , &HC0FFC0, vbBlack, True, "  "
If .Rows > 1 Then
    .TextMatrix(1, 0) = "«·≈Ã„«·Ï"
End If
End With
End Sub
Private Sub MYLOAD2()
Dim cString As String
cString = "SELECT Row_Number() over(order by file1_10.modelfact0)," & _
          "file1_10.modelfact0 ," & _
          "file1_10.mosm ," & _
          "fact.desca," & _
          "file1_10SC.DESCA," & _
          "file1_10.desca," & _
          "sum(file0_10.ComputerBal)," & _
          "sum(file0_10.realBal)," & _
          "SUM(file0_10.differ)," & _
          "ROUND(SUM(file0_10.differ * file1_10.costitem),2), " & _
          "ROUND(SUM(file0_10.differ * file1_10.price),2)" & _
          " from FILE0_10 " & _
          " inner join file1_10 on file1_10.item = FILE0_10.item" & _
          " inner join fact on fact.code = file1_10.fact" & _
          " inner join FILE1_10SC on FILE1_10.SECTION = file1_10SC.CODE" & _
          " WHERE  DOC_NO = " & MyParn(sDoc_no)

If cmdSection.Tag <> "" Then
    cString = cString & " AND FILE1_10.[SECTION] = " & cmdSection.Tag
End If

If cmdFact.Tag <> "" Then
    cString = cString & " AND FILE1_10.[FACT] = " & cmdFact.Tag
End If

cString = cString & " GROUP BY " & _
          "file1_10.modelfact0 ," & _
          "fact.desca," & _
          "file1_10.mosm ," & _
          "file1_10SC.desca," & _
          "file1_10.modelfact0," & _
          "file1_10.desca"
cString = cString & " HAVING SUM(FILE0_10.DIFFER) <> 0 OR SUM(file0_10.differ * file1_10.costitem) <> 0 OR SUM(file0_10.differ * file1_10.price) <> 0"
cString = cString & " order by FILE1_10.modelfact0"

Set grid2.DataSource = DATA2
Set DATA2.Recordset = cmd(cString, con).Execute
Fixgrd2
End Sub
Private Sub Fixgrd2()
With grid2
.RowHeight(0) = 600
'                   0       1             2         3           4            5              6              7              8
.FormatString = "„|" & "—ﬁ„ „ÊœÌ·|" & "„Ê”„|" & "„’‰⁄|" & "ﬁ”„|" & "«·’‰›|" & "—’Ìœ ﬂ„»ÌÊ —|" & "—’Ìœ Ã—œ|" & " ”ÊÌ… Ã—œ|" & "»”⁄— «· ﬂ·›…|" & "»”⁄— «·»Ì⁄"

.RowHeight(0) = 600
.ColWidth(0) = 1000
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 2000
.ColWidth(4) = 2000
.ColWidth(5) = 4600
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColWidth(6 + 2) = 1100
.ColWidth(7 + 2) = 1500
.ColWidth(8 + 2) = 1700

.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
.ColDataType(6 + 2) = flexDTDouble
.ColDataType(7 + 2) = flexDTDouble
.ColDataType(8 + 2) = flexDTDouble

.ColHidden(7 + 2) = Not (cBranch = "00" And bOpt10)
.ColHidden(8 + 2) = Not (cBranch = "00" And bOpt10)

.ExplorerBar = flexExSortShow
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

.SubtotalPosition = flexSTAbove
.Subtotal flexSTSum, -1, 6, "#0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 7, "#0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 6 + 2, "#0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 7 + 2, "#0,00", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 8 + 2, "#0,00", , &HC0FFC0, vbBlack, True, "  "
If .Rows > 1 Then
    .TextMatrix(1, 0) = "«·≈Ã„«·Ï"
End If
.MergeCells = flexMergeFree
End With
End Sub
Private Sub MYLOAD3()
Dim cString As String
cString = "SELECT Row_Number() over(order by file1_10SC.DESCA)," & _
          "file1_10.[SECTION] ," & _
          "file1_10SC.DESCA ," & _
          "sum(file0_10.ComputerBal)," & _
          "sum(file0_10.realBal)," & _
          "SUM(file0_10.differ)," & _
          "ROUND(SUM(file0_10.differ * file1_10.costitem),2), " & _
          "ROUND(SUM(file0_10.differ * file1_10.price),2)" & _
          " from FILE0_10 " & _
          " inner join file1_10 on file1_10.item = FILE0_10.item" & _
          " inner join FILE1_10SC on FILE1_10.[SECTION] = FILE1_10SC.CODE" & _
          " WHERE  DOC_NO = " & MyParn(sDoc_no)

If cmdSection.Tag <> "" Then
    cString = cString & " AND FILE1_10.[SECTION] = " & cmdSection.Tag
End If

If cmdFact.Tag <> "" Then
    cString = cString & " AND FILE1_10.[FACT] = " & cmdFact.Tag
End If
          

cString = cString & " GROUP BY " & _
          "file1_10.[SECTION] ," & _
          "file1_10SC.DESCA "
cString = cString & " HAVING SUM(FILE0_10.DIFFER) <> 0 OR SUM(file0_10.differ * file1_10.costitem) <> 0 OR SUM(file0_10.differ * file1_10.price) <> 0"
cString = cString & " order by FILE1_10SC.DESCA"

Set GRID3.DataSource = data3
Set data3.Recordset = cmd(cString, con).Execute
Fixgrd3
End Sub
Private Sub Fixgrd3()
With GRID3
.RowHeight(0) = 600
'                   0       1             2         3           4            5              6              7              8
.FormatString = "„|" & "«·ﬁ”„|" & "«·ﬁ”„|" & "—’Ìœ ﬂ„»ÌÊ —|" & "—’Ìœ Ã—œ|" & " ”ÊÌ… Ã—œ|" & "»”⁄— «· ﬂ·›…|" & "»”⁄— «·»Ì⁄"

.RowHeight(0) = 600
.ColWidth(0) = 1000
.ColWidth(1) = 1500
.ColWidth(2) = 8100
.ColWidth(3) = 1600
.ColWidth(4) = 1600
.ColWidth(3 + 2) = 1800
.ColWidth(4 + 2) = 1800
.ColWidth(5 + 2) = 2000

.ColDataType(2) = flexDTDouble
.ColDataType(3) = flexDTDouble
.ColDataType(3 + 2) = flexDTDouble
.ColDataType(4 + 2) = flexDTDouble
.ColDataType(5 + 2) = flexDTDouble

.ColHidden(4 + 2) = Not (cBranch = "00" And bOpt10)
.ColHidden(5 + 2) = Not (cBranch = "00" And bOpt10)

.ExplorerBar = flexExSortShow
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

.SubtotalPosition = flexSTAbove
.Subtotal flexSTSum, -1, 2, "#0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 3, "#0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 3 + 2, "#0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 4 + 2, "#0,00", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 5 + 2, "#0,00", , &HC0FFC0, vbBlack, True, "  "
If .Rows > 1 Then
    .TextMatrix(1, 0) = "«·≈Ã„«·Ï"
End If
.MergeCells = flexMergeFree
End With
End Sub
Private Sub MYLOAD4()
Dim cString As String
cString = "SELECT Row_Number() over(order by FACT.DESCA)," & _
          "file1_10.[FACT] ," & _
          "FACT.DESCA ," & _
          "sum(file0_10.ComputerBal)," & _
          "sum(file0_10.realBal)," & _
          "SUM(file0_10.differ)," & _
          "ROUND(SUM(file0_10.differ * file1_10.costitem),2), " & _
          "ROUND(SUM(file0_10.differ * file1_10.price),2)" & _
          " from FILE0_10 " & _
          " inner join file1_10 on file1_10.item = FILE0_10.item" & _
          " inner join FACT on FILE1_10.[FACT] = FACT.CODE" & _
          " WHERE  DOC_NO = " & MyParn(sDoc_no)

If cmdSection.Tag <> "" Then
    cString = cString & " AND FILE1_10.[SECTION] = " & cmdSection.Tag
End If

If cmdFact.Tag <> "" Then
    cString = cString & " AND FILE1_10.[FACT] = " & cmdFact.Tag
End If

cString = cString & " GROUP BY " & _
          "file1_10.[FACT] ," & _
          "FACT.DESCA "
cString = cString & " HAVING SUM(FILE0_10.DIFFER) <> 0 OR SUM(file0_10.differ * file1_10.costitem) <> 0 OR SUM(file0_10.differ * file1_10.price) <> 0"
cString = cString & " order by FACT.DESCA"

Set GRID4.DataSource = data4
Set data4.Recordset = cmd(cString, con).Execute
Fixgrd4
End Sub
Private Sub Fixgrd4()
With GRID4
.RowHeight(0) = 600
'                   0       1             2         3           4            5              6              7              8
.FormatString = "„|" & "«·„’‰⁄|" & "«·„’‰⁄|" & "—’Ìœ ﬂ„»ÌÊ —|" & "—’Ìœ Ã—œ|" & " ”ÊÌ… Ã—œ|" & "»”⁄— «· ﬂ·›…|" & "»”⁄— «·»Ì⁄"

.RowHeight(0) = 600
.ColWidth(0) = 1000
.ColWidth(1) = 1500
.ColWidth(2) = 8100
.ColWidth(3) = 1600
.ColWidth(4) = 1600
.ColWidth(3 + 2) = 1800
.ColWidth(4 + 2) = 1800
.ColWidth(5 + 2) = 2000

.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(3 + 2) = flexDTDouble
.ColDataType(4 + 2) = flexDTDouble
.ColDataType(5 + 2) = flexDTDouble

.ColHidden(4 + 2) = Not (cBranch = "00" And bOpt10)
.ColHidden(5 + 2) = Not (cBranch = "00" And bOpt10)

.ExplorerBar = flexExSortShow
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

.SubtotalPosition = flexSTAbove
.Subtotal flexSTSum, -1, 3, "#0,0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 4, "#0,0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 3 + 2, "#0,0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 4 + 2, "#0,0", &HC0FFC0, vbBlack, True, "  "
.Subtotal flexSTSum, -1, 5 + 2, "#0,0", , &HC0FFC0, vbBlack, True, "  "
If .Rows > 1 Then
    .TextMatrix(1, 0) = "«·≈Ã„«·Ï"
End If
.MergeCells = flexMergeFree
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set ShowStockDiffer = Nothing
End Sub
Private Sub txtBarcode_Change()
SearchGrd grid1, txtBarcode.text, 1, True
End Sub
Private Sub txtModel_Change()
If SSTab1.Tab = 3 Then
    SearchGrd grid1, txtModel.text, 5
ElseIf SSTab1.Tab = 2 Then
    SearchGrd grid2, txtModel.text, 1
End If
End Sub
Public Sub myProc()
If ActiveControl.Name = cmdSection.Name Then
    If oSearchSection.grid1.TextMatrix(oSearchSection.grid1.Row, 0) = "" Then
        cmdSection.Tag = ""
        cmdSection.Caption = cmdSection.TagVariant
    Else
        cmdSection.Tag = oSearchSection.grid1.TextMatrix(oSearchSection.grid1.Row, 0)
        cmdSection.Caption = oSearchSection.grid1.TextMatrix(oSearchSection.grid1.Row, 1)
    End If
    Unload oSearchSection
ElseIf ActiveControl.Name = cmdFact.Name Then
    If oSearchFact.grid1.TextMatrix(oSearchFact.grid1.Row, 0) = "" Then
        cmdFact.Tag = ""
        cmdFact.Caption = cmdFact.TagVariant
    Else
        cmdFact.Tag = oSearchFact.grid1.TextMatrix(oSearchFact.grid1.Row, 0)
        cmdFact.Caption = oSearchFact.grid1.TextMatrix(oSearchFact.grid1.Row, 1)
    End If
    Unload oSearchFact
End If
SSTab1.Enabled = False
Me.MousePointer = vbHourglass
myload
MYLOAD2
MYLOAD3
MYLOAD4
Me.MousePointer = vbNormal
SSTab1.Enabled = True
End Sub
