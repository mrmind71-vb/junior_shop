VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdOnlineDetailsNewfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   " ﬁ—Ì— „ «»⁄… ÿ·»Ì«  √Ê‰ ·«Ì‰"
   ClientHeight    =   10290
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
   ScaleHeight     =   10290
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8700
      Left            =   90
      TabIndex        =   24
      Top             =   1080
      Width           =   20265
      _ExtentX        =   35745
      _ExtentY        =   15346
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
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
      TabCaption(0)   =   " ›’Ì·Ì „ «»⁄… «·ÿ·»Ì« "
      TabPicture(0)   =   "grdOnlineDetailsNew.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "≈Ã„«·Ì „ «»⁄… «·ÿ·»Ì« "
      TabPicture(1)   =   "grdOnlineDetailsNew.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   8250
         Left            =   -74955
         TabIndex        =   25
         Top             =   360
         Width           =   20130
         _cx             =   35507
         _cy             =   14552
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
         BackColorSel    =   12640511
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   80
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
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   8250
         Left            =   45
         TabIndex        =   26
         Top             =   360
         Width           =   20130
         _cx             =   35507
         _cy             =   14552
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
         BackColorSel    =   12640511
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   17
         FixedRows       =   1
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   45
      Width           =   13470
      Begin VB.TextBox xShip_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   180
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "s"
         Top             =   180
         Width           =   2175
      End
      Begin VB.OptionButton xtype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„”·„…"
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
         Height          =   375
         Index           =   1
         Left            =   2940
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   585
         Width           =   960
      End
      Begin VB.OptionButton xtype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "€Ì— „”·„…"
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
         Height          =   375
         Index           =   2
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   585
         Width           =   1185
      End
      Begin VB.OptionButton xtype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
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
         Height          =   375
         Index           =   0
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   585
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.OptionButton xtype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„·€Ì…"
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
         Height          =   375
         Index           =   3
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   585
         Width           =   780
      End
      Begin VB.TextBox xdoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   5490
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Tag             =   "s"
         Top             =   540
         Width           =   1905
      End
      Begin VB.TextBox xOrder_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   5490
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "s"
         Top             =   180
         Width           =   1905
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   10845
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1320
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   8730
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   330
         Left            =   8730
         TabIndex        =   2
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·»Ê·Ì’…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·›« Ê—…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   585
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·ÿ·»Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Õ Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10125
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„‰  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12285
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   660
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·›—⁄"
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
         Height          =   270
         Left            =   12285
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   585
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1035
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   315
      Width           =   5775
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
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
         Picture         =   "grdOnlineDetailsNew.frx":0038
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1170
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
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
         Picture         =   "grdOnlineDetailsNew.frx":235B
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdOnlineDetailsNew.frx":46D1
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   4635
         TabIndex        =   6
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
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
         Picture         =   "grdOnlineDetailsNew.frx":6854
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2295
         TabIndex        =   8
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
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
         Picture         =   "grdOnlineDetailsNew.frx":9707
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSql 
         Height          =   555
         Left            =   3420
         TabIndex        =   7
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   979
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
         Picture         =   "grdOnlineDetailsNew.frx":BC32
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   4410
      Top             =   -90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   3420
      Top             =   -90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   12
      Top             =   10095
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   540
      Top             =   180
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   3330
      Top             =   270
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   1890
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
End
Attribute VB_Name = "grdOnlineDetailsNewfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim dbm As New DBManage
Dim oSearch As New Search_abd, oSearchDoc As New Search_abd
Dim aHeader()
Private Sub cmdExcel_Click()
If SSTab1.Tab = 0 Then
    toExcel1
Else
    toExcel2
End If
End Sub
Private Sub toExcel2()
Me.MousePointer = 11

Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)

aSub = AddFlag(Empty, "row", 1)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "cols", 2)
aSub = AddFlag(aSub, "text", "«·≈Ã„«·Ì")
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)

ToFileExelNew grid2, , , aRow, , 1.2, , , , , , Me

Me.MousePointer = 0
End Sub
Private Sub toExcel1()
Me.MousePointer = 11

Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "col", 0)
aSub = AddFlag(aSub, "cols", 4)
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)

aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "col", 4)
aSub = AddFlag(aSub, "cols", 9)
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)

aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "col", 13)
aSub = AddFlag(aSub, "cols", 12)
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)

aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "col", 25)
aSub = AddFlag(aSub, "cols", 9)
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)

aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "col", 34)
aSub = AddFlag(aSub, "cols", 9)
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)

aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "col", 43)
aSub = AddFlag(aSub, "cols", 9)
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)

aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "col", 52)
aSub = AddFlag(aSub, "cols", 9)
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)

aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "col", 61)
aSub = AddFlag(aSub, "cols", 6)
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)

aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "col", 67)
aSub = AddFlag(aSub, "cols", 6)
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)

aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "col", 73)
aSub = AddFlag(aSub, "cols", 6)
aSub = AddFlag(aSub, "height", 20)
aRow = AddFlag(aRow, aSub)


aSub = AddFlag(Empty, "row", 1)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aSub = AddFlag(aSub, "back_color", 40)
aSub = AddFlag(aSub, "height", 60)
aRow = AddFlag(aRow, aSub)

If grid1.Rows > 2 Then
    aSub = AddFlag(Empty, "row", 2)
    aSub = AddFlag(aSub, "col", 0)
    aSub = AddFlag(aSub, "cols", 2)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "back_color", 40)
    aSub = AddFlag(aSub, "text", "«·≈Ã„«·Ì")
    aRow = AddFlag(aRow, aSub)
End If

ToFileExelNew grid1, , , aRow, , 1.2, , , , , , Me

Me.MousePointer = 0
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
Me.MousePointer = vbHourglass
grid1.ExplorerBar = flexExSortShow
grid2.ExplorerBar = flexExSortShow

If SSTab1.Tab = 0 Then
    myload
ElseIf SSTab1.Tab = 1 Then
    MYLOAD2
End If
Me.MousePointer = vbNormal
End Sub

Private Sub cmdPdf_Click()
myPrint "1", True
End Sub
Private Sub cmdPrint_Click()
myPrint
End Sub

Private Sub cmdSql_Click()
myload True
End Sub
Private Sub Form_Resize()
'SSTab1.Height = IIf(Me.Height - SSTab1.Top - 1000 < 3000, 3000, Me.Height - SSTab1.Top - 500)
'SSTab1.Width = IIf(Me.Width - SSTab1.Left - 150 < 5000, 5000, Me.Width - SSTab1.Left - 150)
'grid1.Height = IIf(SSTab1.Height - grid1.Top - 1000 < 3000, 3000, SSTab1.Height - grid1.Top - 200)
'grid1.Width = IIf(SSTab1.Width - grid1.Left - 150 < 5000, 5000, SSTab1.Width - grid1.Left - 150)
'grid2.Height = IIf(SSTab1.Height - grid2.Top - 1000 < 3000, 3000, SSTab1.Height - grid2.Top - 200)
'grid2.Width = IIf(SSTab1.Width - grid2.Left - 150 < 5000, 5000, SSTab1.Width - grid2.Left - 150)

'Frame1.Left = SSTab1.Width - Frame1.Width - 150
'Frame2.Left = IIf(300 + Me.Width - Frame2.Width - Frame1.Left < 0, 0, 300 + Me.Width - Frame2.Width - Frame1.Left)
'Frame2.Left = Frame1.Left - Frame2.Width - 50
End Sub
Private Sub Form_Load()
Set xStore.RowSource = dbm.myRs("Select code ,descA From file0_40 ORDER BY CODE")
xStore.ListField = "Desca"
xStore.BoundColumn = "CODE"

fixGrd
Fixgrd2
End Sub
Private Sub myload(Optional bString As Boolean = False)
Dim aPrm As Variant

ReDim aHeader(5)

If xStore.MatchedWithList Then
    aPrm = AddFlag(aPrm, "STORE", addstring(xStore.BoundText))
    aHeader(0) = "«·›—⁄ : " & xStore.text
End If

If IsDate(xdate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", addDate(xdate1.text))
    aHeader(1) = BetweenString(xdate1.text, xdate2.text)
End If

If IsDate(xdate1.text) Then
    aPrm = AddFlag(aPrm, "DATE2", addDate(xdate2.text))
    aHeader(1) = BetweenString(xdate1.text, xdate2.text)
End If

If xtype(1).value Then
    aPrm = AddFlag(aPrm, "CLOSED", "1")
    aHeader(2) = "ÿ·»Ì«  " & xtype(1).Caption
ElseIf xtype(2).value Then
    aPrm = AddFlag(aPrm, "OPENED", "1")
    aHeader(2) = "ÿ·»Ì«  " & xtype(2).Caption
ElseIf xtype(3).value Then
    aPrm = AddFlag(aPrm, "CANCELED", "1")
    aHeader(2) = "ÿ·»Ì«  " & xtype(3).Caption
End If

If Trim(xdoc_no.text) <> "" Then
    aPrm = AddFlag(aPrm, "DOC_NO", addstring(xdoc_no.text))
    aHeader(3) = "›« Ê—… : " & xdoc_no.text
End If

If Trim(xOrder_No.text) <> "" Then
    aPrm = AddFlag(aPrm, "ORDER_NO", addstring(xOrder_No.text))
    aHeader(4) = "ÿ·»Ì… : " & xOrder_No.text
End If

If Trim(xship_no.text) <> "" Then
    aPrm = AddFlag(aPrm, "SHIP_NO", addstring(xship_no.text))
    aHeader(5) = "»Ê·Ì’… ‘Õ‰ : " & xship_no.text
End If

Dim cString As String
cString = myPrcString("dbo.sp_online_pay", aPrm)

If bString Then
    Clipboard.Clear
    Clipboard.SetText cString
    Exit Sub
End If

Set grid1.DataSource = dbm.myRs(cString)
fixGrd
End Sub
Sub fixGrd()
With grid1
    .RowHeight(1) = 1000
    .WordWrap = True
    .FrozenCols = 4
    
    For i = 0 To 3
        .TextMatrix(0, i) = "»Ì«‰«  «”«”Ì…"
    Next
        
    For i = 4 To 12
        .TextMatrix(0, i) = "«·›« Ê—… «·—∆Ì”Ì…"
    Next
    If .Rows > 2 Then
        .Cell(flexcpBackColor, 2, 4, .Rows - 1, 12) = &HE0E0E0
        .Cell(flexcpBackColor, 2, 25, .Rows - 1, 33) = &HE0E0E0
        .Cell(flexcpBackColor, 2, 43, .Rows - 1, 51) = &HE0E0E0
        .Cell(flexcpBackColor, 2, 61, .Rows - 1, 66) = &HE0E0E0
        .Cell(flexcpBackColor, 2, 73, .Rows - 1, 78) = &HE0E0E0
    End If

    
    For i = 13 To 24
        .TextMatrix(0, i) = "›Ê« Ì— «” »œ«· Ê«” —Ã«⁄"
    Next
    
    For i = 25 To 33
        .TextMatrix(0, i) = "»Ì«‰«  «· Õ’Ì·"
    Next
    
    For i = 34 To 42
        .TextMatrix(0, i) = "ÿ·»Ì«  «” »œ«·"
    Next
    
    For i = 43 To 51
        .TextMatrix(0, i) = "ÿ·»Ì«  „— Ã⁄"
    Next
    
    For i = 52 To 60
        .TextMatrix(0, i) = "ÿ·»Ì«  —›÷ «” ·«„"
    Next
    
    For i = 61 To 66
        .TextMatrix(0, i) = "ÿ·»Ì«  «” »œ«· »œÊ‰ ›« Ê—…"
    Next
    
    For i = 67 To 72
        .TextMatrix(0, i) = "ÿ·»Ì«  „— Ã⁄ »œÊ‰ ›« Ê—…"
    Next
    
    For i = 73 To 78
        .TextMatrix(0, i) = "≈Ã„«·Ì «·ÿ·»Ì« "
    Next
    
    .TextMatrix(0, 79) = "«·„‰œÊ»"
    
    .TextMatrix(1, 0) = "«· «—ÌŒ"
    .TextMatrix(1, 1) = "—ﬁ„ ÿ·»Ì…"
    .TextMatrix(1, 2) = "⁄œœ «·ﬁÿ⁄"
    .TextMatrix(1, 3) = "ﬁÌ„… «·ÿ·»Ì…"
    
    .TextMatrix(1, 4) = "—ﬁ„ «·›« Ê—…"
    .TextMatrix(1, 5) = " «—ÌŒ «·›« Ê—…"
    .TextMatrix(1, 6) = "⁄œœ «·ﬁÿ⁄"
    .TextMatrix(1, 7) = "ﬁÌ„… «·›« Ê—…"
    .TextMatrix(1, 8) = "—ﬁ„ »Ê·Ì’… «·‘Õ‰"
    .TextMatrix(1, 9) = "‘—ﬂ… «·‘Õ‰"
    .TextMatrix(1, 10) = " «—ÌŒ «·‘Õ‰"
    .TextMatrix(1, 11) = "ÿ—Ìﬁ… «·”œ«œ"
    .TextMatrix(1, 12) = " «—ÌŒ «·«·€«¡"
        
    .TextMatrix(1, 13) = "⁄œœ ›Ê« Ì— «” »œ«·"
    .TextMatrix(1, 14) = "ﬂ„Ì«  ›Ê« Ì— «” »œ«·"
    .TextMatrix(1, 15) = "≈Ã„«·Ì ›Ê« Ì— «” »œ«·"
    
    .TextMatrix(1, 16) = "⁄œœ ›Ê« Ì— „— Ã⁄"
    .TextMatrix(1, 17) = "ﬂ„Ì«  ›Ê« Ì— „— Ã⁄"
    .TextMatrix(1, 18) = "≈Ã„«·Ì ›Ê« Ì— „— Ã⁄"
    
    .TextMatrix(1, 19) = "⁄œœ ›Ê« Ì— —›÷ «” ·«„"
    .TextMatrix(1, 20) = "ﬂ„Ì«  ›Ê« Ì— —›÷ «” ·«„"
    .TextMatrix(1, 21) = "≈Ã„«·Ì ›Ê« Ì— —›÷ «” ·«„"
            
    .TextMatrix(1, 22) = "⁄œœ ›Ê« Ì— »Ì⁄"
    .TextMatrix(1, 23) = "ﬂ„Ì«  ›Ê« Ì— »Ì⁄"
    .TextMatrix(1, 24) = "ﬁÌ„… ›Ê« Ì— »Ì⁄"
    
    .TextMatrix(1, 25) = "”œ«œ «Ê‰ ·«Ì‰"
    .TextMatrix(1, 26) = "⁄„Ê·…  Õ’Ì·"
    
    .TextMatrix(1, 27) = "≈Ã„«·Ì ›Ê« Ì— »Ì⁄"
    .TextMatrix(1, 28) = "≈Ã„«·Ì ›Ê« Ì— „— Ã⁄"
    .TextMatrix(1, 29) = "’«›Ì «·›Ê« Ì—"
    
    .TextMatrix(1, 30) = "COD"
    .TextMatrix(1, 31) = "„’«—Ì› ‘Õ‰ „Õ’·…"
    .TextMatrix(1, 32) = "„’«—Ì› ÿ·»Ì…"
    .TextMatrix(1, 33) = "⁄œœ „” ‰œ«  «· Õ’Ì·"
    
    .TextMatrix(1, 34) = "⁄œœ «” »œ«· »›« Ê—…"
    .TextMatrix(1, 35) = "⁄œœ «” »œ«· »›« Ê—… „‰›–…"
    .TextMatrix(1, 36) = "⁄œœ «” »œ«· »›« Ê—… €Ì— „‰›–…"
    
    .TextMatrix(1, 37) = "ﬂ„Ì«  «” »œ«· »›« Ê—…"
    .TextMatrix(1, 38) = "ﬂ„Ì«  «” »œ«· »›« Ê—… „‰›–…"
    .TextMatrix(1, 39) = "ﬂ„Ì«  «” »œ«· »›« Ê—… €Ì— „‰›–…"
    
    .TextMatrix(1, 40) = "ﬁÌ„ «” »œ«· »›« Ê—…"
    .TextMatrix(1, 41) = "ﬁÌ„ «” »œ«· »›« Ê—… „‰›–…"
    .TextMatrix(1, 42) = "ﬁÌ„ «” »œ«· »›« Ê—… €Ì— „‰›–…"
    
    .TextMatrix(1, 43) = "⁄œœ „— Ã⁄ »›« Ê—…"
    .TextMatrix(1, 44) = "⁄œœ „— Ã⁄ »›« Ê—… „‰›–…"
    .TextMatrix(1, 45) = "⁄œœ „— Ã⁄ »›« Ê—… €Ì— „‰›–…"
        
    .TextMatrix(1, 46) = "ﬂ„Ì«  „— Ã⁄ »›« Ê—…"
    .TextMatrix(1, 47) = "ﬂ„Ì«  „— Ã⁄ »›« Ê—… „‰›–…"
    .TextMatrix(1, 48) = "ﬂ„Ì«  „— Ã⁄ »›« Ê—… €Ì— „‰›–…"
    
    .TextMatrix(1, 49) = "ﬁÌ„ „— Ã⁄ »›« Ê—…"
    .TextMatrix(1, 50) = "ﬁÌ„ „— Ã⁄ »›« Ê—… „‰›–…"
    .TextMatrix(1, 51) = "ﬁÌ„ „— Ã⁄ »›« Ê—… €Ì— „‰›–…"
    
    .TextMatrix(1, 52) = "⁄œœ —›÷ «” ·«„"
    .TextMatrix(1, 53) = "⁄œœ —›÷ «” ·«„ „‰›–…"
    .TextMatrix(1, 54) = "⁄œœ —›÷ «” ·«„ €Ì— „‰›–…"
        
    .TextMatrix(1, 55) = "ﬂ„Ì«  —›÷ «” ·«„"
    .TextMatrix(1, 56) = "ﬂ„Ì«  —›÷ «” ·«„ „‰›–…"
    .TextMatrix(1, 57) = "ﬂ„Ì«  —›÷ «” ·«„ €Ì— „‰›–…"
    
    .TextMatrix(1, 58) = "ﬁÌ„ —›÷ «” ·«„"
    .TextMatrix(1, 59) = "ﬁÌ„ —›÷ «” ·«„ „‰›–…"
    .TextMatrix(1, 60) = "ﬁÌ„ —›÷ «” ·«„ €Ì— „‰›–…"
        
    .TextMatrix(1, 61) = "⁄œœ «” »œ«· »œÊ‰ ›« Ê—…"
    .TextMatrix(1, 62) = "⁄œœ «” »œ«· »œÊ‰ ›« Ê—… „‰›–…"
    .TextMatrix(1, 63) = "⁄œœ «” »œ«· »œÊ‰ ›« Ê—… €Ì— „‰›–…"
    
    .TextMatrix(1, 64) = "ﬂ„Ì«  «” »œ«· »œÊ‰ ›« Ê—…"
    .TextMatrix(1, 65) = "ﬂ„Ì«  «” »œ«· »œÊ‰ ›« Ê—… „‰›–…"
    .TextMatrix(1, 66) = "ﬂ„Ì«  «” »œ«· »œÊ‰ ›« Ê—… €Ì— „‰›–…"
    
    .TextMatrix(1, 67) = "⁄œœ „— Ã⁄ »œÊ‰ ›« Ê—…"
    .TextMatrix(1, 68) = "⁄œœ „— Ã⁄ »œÊ‰ ›« Ê—… „‰›–…"
    .TextMatrix(1, 69) = "⁄œœ „— Ã⁄ »œÊ‰ ›« Ê—… €Ì— „‰›–…"
    
    .TextMatrix(1, 70) = "ﬂ„Ì«  „— Ã⁄ »œÊ‰ ›« Ê—…"
    .TextMatrix(1, 71) = "ﬂ„Ì«  „— Ã⁄ »œÊ‰ ›« Ê—… „‰›–…"
    .TextMatrix(1, 72) = "ﬂ„Ì«  „— Ã⁄ »œÊ‰ ›« Ê—… €Ì— „‰›–…"
    
    .TextMatrix(1, 73) = "⁄œœ «·ÿ·»Ì« "
    .TextMatrix(1, 74) = "⁄œœ «·ÿ·»Ì«  «·„‰›–…"
    .TextMatrix(1, 75) = "⁄œœ «·ÿ·»Ì«  «·€Ì—«·„‰›–…"
    
    .TextMatrix(1, 76) = "ﬁÌ„ «·ÿ·»Ì« "
    .TextMatrix(1, 77) = "ﬁÌ„ «·ÿ·»Ì«  «·„‰›–…"
    .TextMatrix(1, 78) = "ﬁÌ„ «·ÿ·»Ì«  «·€Ì— „‰›–…"
    
    .TextMatrix(1, 79) = "«·„‰œÊ»"

    
    
    .ColWidth(0) = 1300
    .ColWidth(1) = 1100
    .ColWidth(2) = 1100
    .ColWidth(3) = 1400
    .ColWidth(4) = 1600
    .ColWidth(5) = 1300
    .ColWidth(6) = 1100
    .ColWidth(7) = 1300
    .ColWidth(8) = 1700
    .ColWidth(9) = 1300
    .ColWidth(10) = 1400
    .ColWidth(11) = 2200
    .ColWidth(12) = 1300
    .ColWidth(13) = 1000
    .ColWidth(14) = 1000
    
    .ColDataType(0) = flexDTDate
    .ColDataType(5) = flexDTDate
    .ColDataType(10) = flexDTDate
    .ColDataType(12) = flexDTDate
    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    
    .ColDataType(13) = flexDTDouble
    .ColDataType(14) = flexDTDouble
    
    For i = 15 To .Cols - 2
        .ColDataType(i) = flexDTDouble
        .ColWidth(i) = 1400
    Next
    
    .ColWidth(16) = 1000
    .ColWidth(17) = 1000
    .ColWidth(19) = 1000
    .ColWidth(20) = 1000
    .ColWidth(22) = 1000
    .ColWidth(23) = 1000
    .ColWidth(33) = 1000
    .ColWidth(34) = 1000
    .ColWidth(35) = 1000
    .ColWidth(36) = 1000
    .ColWidth(37) = 1000
    .ColWidth(38) = 1000
    .ColWidth(39) = 1000
    .ColWidth(43) = 1000
    .ColWidth(44) = 1000
    .ColWidth(45) = 1000
    .ColWidth(46) = 1000
    .ColWidth(47) = 1000
    .ColWidth(48) = 1000
    
    .ColWidth(52) = 1000
    .ColWidth(53) = 1000
    .ColWidth(54) = 1000
    .ColWidth(55) = 1000
    .ColWidth(56) = 1000
    .ColWidth(57) = 1000
        
    For i = 61 To 75
        .ColWidth(i) = 1000
    Next
    
    .ColWidth(79) = 2000
    
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignCenterCenter
        If .ColDataType(i) = flexDTDate Then
            .ColFormat(i) = "yyyy/m/d"
        End If
    Next
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    .MergeCol(.Cols - 1) = True
    
    .SubtotalPosition = flexSTAbove

    For i = 2 To .Cols - 2
        If .ColDataType(i) = flexDTDouble Then
            .Subtotal flexSTSum, -1, i, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
        End If
    Next
    If .Rows > 2 Then
       .TextMatrix(2, 0) = "«·≈Ã„«·Ì"
    End If
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set grdOnlineDetailsNewfrm = Nothing
End Sub
Private Function myValid() As Boolean
myValid = True
End Function
Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
If grid1.Rows < 3 Then Exit Sub
Dim aRow As Variant, aSub As Variant

aSub = AddFlag(Empty, "row", grid1.Rows - 1)
aSub = AddFlag(aSub, "col", 1)
aSub = AddFlag(aSub, "cols", 2)
aSub = AddFlag(aSub, "text", "≈Ã„«·Ì")
aRow = AddFlag(aRow, aSub)

Set printGrdNew.myForm = Me
printGrdNew.doprint Me.grid1, 0.75, 0, "ÃÊ‰ÌÊ—", Me.Caption, retHeader(aHeader, 0, 2), , True, True, 8, , aRow, Array(1)

If Not bIgPreview Then
    printGrdNew.Show 1
Else
    Unload printGrdNew
End If
End Sub

Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xModelFACT_GotFocus()
myGotFocus xModelFact
End Sub
Private Sub XModelFACT_LostFocus()
myLostFocus xModelFact
End Sub

Private Sub grid1_DblClick()
If grid1.Row < 2 Then Exit Sub
If grid1.col = 1 Then
    ShowOrderfrm.sDoc_no = grid1.TextMatrix(grid1.Row, 1)
    ShowOrderfrm.Show 1
ElseIf grid1.col = 4 Then
    sales_online_popfrm.sDoc_no = grid1.TextMatrix(grid1.Row, 4)
    sales_online_popfrm.bEdit = False
    sales_online_popfrm.Show 1
    
End If

'If grid1.Col = 0 Then
'    OrdersClientsfrm.sDoc_no = grid1.TextMatrix(grid1.Row, 0)
'    OrdersClientsfrm.Show
'Else
'    grdOrdersDelivery_popfrm.sPO_NO = grid1.TextMatrix(grid1.Row, 0)
'    grdOrdersDelivery_popfrm.Show
'End If
End Sub

Private Sub grid1_EnterCell()
'grid1.Editable = IIf(grid1.Col = 0, flexEDKbdMouse, flexEDNone)
End Sub

Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then CLIENTLOOKUP Me, oSearchCode
End Sub
Private Sub Grid2_dblClick()
If grid2.Row < 2 Then
    Exit Sub
ElseIf grid2.col = 2 Or grid2.col = 3 Or grid2.col = 4 Then
    ShowOrdersfrm.bNoInvoice = False
    ShowOrdersfrm.sOrder_No = grid2.TextMatrix(grid2.Row, 1)
    ShowOrdersfrm.Show 1
ElseIf grid2.col = 5 Or grid2.col = 6 Then
    ShowOrdersfrm.bNoInvoice = True
    ShowOrdersfrm.sOrder_No = grid2.TextMatrix(grid2.Row, 1)
    ShowOrdersfrm.Show 1
ElseIf grid2.col = 7 Then
    ShowOrdersInvoicesfrm.sOrder_No = grid2.TextMatrix(grid2.Row, 1)
    ShowOrdersInvoicesfrm.Show 1
End If
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub XPO_NO_GotFocus()
myGotFocus XPO_NO
End Sub
Private Sub xpo_no_LostFocus()
myLostFocus XPO_NO
End Sub
Private Sub xcode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xcode_LostFocus()
myLostFocus xCode
If Not xCode.MatchedWithList Then xCode.BoundText = ""
End Sub
Public Sub myProc()
If ActiveControl.Name = xdoc_no.Name Then
    xdoc_no.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    oSearchDoc.Hide
End If
End Sub
Private Sub xdoc_no_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then SalesOnlineLookup Me, oSearchDoc
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xdoc_no
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xdoc_no
End Sub
Private Sub xship_no_GotFocus()
myGotFocus xship_no
End Sub
Private Sub xship_no_LostFocus()
myLostFocus xship_no
End Sub
Private Sub xOrder_no_GotFocus()
myGotFocus xOrder_No
End Sub
Private Sub xOrder_no_LostFocus()
myLostFocus xOrder_No
End Sub
Private Sub MYLOAD2(Optional bString As Boolean = False)
Dim aPrm As Variant

ReDim aHeader(5)

If xStore.MatchedWithList Then
    aPrm = AddFlag(aPrm, "STORE", addstring(xStore.BoundText))
    aHeader(0) = "«·›—⁄ : " & xStore.text
End If

If IsDate(xdate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", addDate(xdate1.text))
    aHeader(1) = BetweenString(xdate1.text, xdate2.text)
End If

If IsDate(xdate1.text) Then
    aPrm = AddFlag(aPrm, "DATE2", addDate(xdate2.text))
    aHeader(1) = BetweenString(xdate1.text, xdate2.text)
End If

If xtype(1).value Then
    aPrm = AddFlag(aPrm, "CLOSED", "1")
    aHeader(2) = "ÿ·»Ì«  " & xtype(1).Caption
ElseIf xtype(2).value Then
    aPrm = AddFlag(aPrm, "OPENED", "1")
    aHeader(2) = "ÿ·»Ì«  " & xtype(2).Caption
ElseIf xtype(3).value Then
    aPrm = AddFlag(aPrm, "CANCELED", "1")
    aHeader(2) = "ÿ·»Ì«  " & xtype(3).Caption
End If

If Trim(xdoc_no.text) <> "" Then
    aPrm = AddFlag(aPrm, "DOC_NO", addstring(xdoc_no.text))
    aHeader(3) = "›« Ê—… : " & xdoc_no.text
End If

If Trim(xOrder_No.text) <> "" Then
    aPrm = AddFlag(aPrm, "ORDER_NO", addstring(xOrder_No.text))
    aHeader(4) = "ÿ·»Ì… : " & xOrder_No.text
End If

If Trim(xship_no.text) <> "" Then
    aPrm = AddFlag(aPrm, "SHIP_NO", addstring(xship_no.text))
    aHeader(5) = "»Ê·Ì’… ‘Õ‰ : " & xship_no.text
End If

Dim cString As String
cString = myPrcString("dbo.sp_online_pay2", aPrm)

If bString Then
    Clipboard.Clear
    Clipboard.SetText cString
    Exit Sub
End If

Set grid2.DataSource = dbm.myRs(cString)
Fixgrd2
End Sub
Sub Fixgrd2()
With grid2
.RowHeight(0) = 600
.WordWrap = True
.TextMatrix(0, 0) = "«· «—ÌŒ"
.TextMatrix(0, 1) = "—ﬁ„ ÿ·»Ì…"
.TextMatrix(0, 2) = "⁄œœ «·ÿ·»Ì« "
.TextMatrix(0, 3) = "⁄œœ «·ﬁÿ⁄"
.TextMatrix(0, 4) = "ﬁÌ„… «·ÿ·»Ì…"
.TextMatrix(0, 5) = "⁄œœ «·ÿ·»Ì«  »œÊ‰ ›« Ê—…"
.TextMatrix(0, 6) = "⁄œœ «·ﬁÿ⁄ »œÊ‰ ›« Ê—…"
.TextMatrix(0, 7) = "⁄œœ «·›Ê« Ì—"
.TextMatrix(0, 8) = "≈Ã„«·Ì ⁄œœ «·ﬁÿ⁄"
.TextMatrix(0, 9) = "≈Ã„«·Ì «·›Ê« Ì—"
.TextMatrix(0, 10) = " «—ÌŒ «·«·€«¡"
.TextMatrix(0, 11) = "≈Ã„«·Ì «·”œ«œ"
.TextMatrix(0, 12) = "⁄„Ê·…  Õ’Ì·"
.TextMatrix(0, 13) = "’«›Ì «·›Ê« Ì—"
.TextMatrix(0, 14) = "„’«—Ì› ‘Õ‰ „”œœ…"
.TextMatrix(0, 15) = "’«›Ì „’—Ê› «·ÿ·»Ì…"
.TextMatrix(0, 16) = "«·„‰œÊ»"

.ColWidth(0) = 1300
.ColWidth(1) = 1100
.ColWidth(2) = 1100
.ColWidth(3) = 1100
.ColWidth(4) = 1300
.ColWidth(5) = 1100
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColWidth(8) = 1100
.ColWidth(9) = 1300
.ColWidth(11) = 1200
.ColWidth(12) = 1300
.ColWidth(13) = 1400
.ColWidth(14) = 1400
.ColWidth(15) = 1400
.ColWidth(16) = 2800

.ColDataType(2) = flexDTDouble
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
.ColDataType(8) = flexDTDouble
.ColDataType(9) = flexDTDouble


.ColDataType(11) = flexDTDouble
.ColDataType(12) = flexDTDouble
.ColDataType(13) = flexDTDouble
.ColDataType(14) = flexDTDouble
.ColDataType(15) = flexDTDouble
    
.ColDataType(0) = flexDTDate
.ColDataType(10) = flexDTDate
    
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
    If .ColDataType(i) = flexDTDate Then
        .ColFormat(i) = "yyyy/m/d"
    End If
Next
    
.SubtotalPosition = flexSTAbove

For i = 2 To .Cols - 2
    If .ColDataType(i) = flexDTDouble Then
        .Subtotal flexSTSum, -1, i, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    End If
Next
If .Rows > 2 Then
   .TextMatrix(1, 0) = "«·≈Ã„«·Ì"
End If
End With
End Sub

