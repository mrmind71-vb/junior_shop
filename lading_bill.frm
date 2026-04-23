VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form ladingBillfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "»Ê«·’ ‘Õ‰ „”»ﬁ…"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
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
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2130
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   2130
      ScaleWidth      =   20370
      TabIndex        =   24
      Top             =   8550
      Width           =   20370
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   -45
         Width           =   6810
         Begin VB.TextBox xdate_Delivery2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1620
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Tag             =   "filter"
            Top             =   225
            Width           =   1275
         End
         Begin VB.CheckBox chkDay_delivery 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "«” ·«„ «·ÌÊ„"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   270
            Width           =   1230
         End
         Begin VB.TextBox xdate_Delivery1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4410
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Tag             =   "filter"
            Top             =   225
            Width           =   1185
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Õ Ì"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   2970
            TabIndex        =   43
            Tag             =   "COLOR"
            Top             =   270
            Width           =   375
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   " «—ÌŒ «” ·«„"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   5715
            TabIndex        =   42
            Tag             =   "COLOR"
            Top             =   270
            Width           =   915
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   16650
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   -45
         Width           =   3615
         Begin VB.OptionButton optClose 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   270
            Width           =   780
         End
         Begin VB.OptionButton optClose 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "€Ì— „€·ﬁ…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1305
            RightToLeft     =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   270
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VB.OptionButton optClose 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "„€·ﬁ…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1260
         Width           =   2580
         Begin Threed.SSCommand cmdExit 
            Height          =   465
            Left            =   45
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
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
            Picture         =   "lading_bill.frx":0000
            Caption         =   "///"
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdGo 
            Height          =   465
            Left            =   1305
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
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
            Picture         =   "lading_bill.frx":2323
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   6975
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   -45
         Width           =   9645
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   225
            Width           =   1950
         End
         Begin VB.TextBox xship_no 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6390
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   225
            Width           =   1950
         End
         Begin VB.TextBox xdoc_no 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   225
            Width           =   1950
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·„— Ã⁄"
            Height          =   285
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   225
            Width           =   1005
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "—ﬁ„ «·»Ê·Ì’…"
            Height          =   240
            Index           =   0
            Left            =   8460
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "—ﬁ„ «·›« Ê—…"
            Height          =   240
            Left            =   5310
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   270
            Width           =   1005
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   13410
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   585
         Width           =   6855
         Begin VB.TextBox xdate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4050
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Tag             =   "filter"
            Top             =   225
            Width           =   1410
         End
         Begin VB.CheckBox chkDay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "»Ê«·’ «·ÌÊ„"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   270
            Width           =   1365
         End
         Begin VB.TextBox xdate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1575
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Tag             =   "filter"
            Top             =   225
            Width           =   1365
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   " «—ÌŒ »Ê·Ì’…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   15
            Left            =   5625
            TabIndex        =   30
            Tag             =   "COLOR"
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Õ Ì"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   3105
            TabIndex        =   29
            Tag             =   "COLOR"
            Top             =   270
            Width           =   510
         End
      End
      Begin VB.Frame Frame8 
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
         Height          =   690
         Left            =   9450
         TabIndex        =   25
         Top             =   630
         Width           =   3930
         Begin MSDataListLib.DataCombo xYear 
            Height          =   330
            Left            =   1980
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   225
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
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   225
            Width           =   735
            _ExtentX        =   1296
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
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·‘Â—"
            Height          =   330
            Left            =   855
            TabIndex        =   27
            Tag             =   "Color"
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·”‰…"
            Height          =   285
            Left            =   3285
            TabIndex        =   26
            Tag             =   "Color"
            Top             =   270
            Width           =   510
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   585
         Width           =   9285
         Begin VB.TextBox xName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4770
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   225
            Width           =   3255
         End
         Begin VB.TextBox xPhone 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   225
            Width           =   3255
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "≈”„ «·⁄„Ì·"
            Height          =   240
            Left            =   8145
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   " ·Ì›Ê‰ «·⁄„Ì·"
            Height          =   240
            Left            =   3465
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   270
            Width           =   1005
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1215
         Width           =   5550
         Begin Threed.SSCommand cmdShip 
            Height          =   465
            Left            =   45
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   180
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   820
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
            Caption         =   "‘—ﬂ… «·‘Õ‰"
            TagVariant      =   "‘—ﬂ… «·‘Õ‰"
            ButtonStyle     =   3
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   14985
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1215
         Width           =   5280
         Begin Threed.SSCommand cmdType 
            Height          =   465
            Left            =   45
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   180
            Width           =   5190
            _ExtentX        =   9155
            _ExtentY        =   820
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
            Caption         =   "‰Ê⁄ «·»Ê·Ì’…"
            TagVariant      =   "‰Ê⁄ «·»Ê·Ì’…"
            ButtonStyle     =   3
         End
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8205
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   20175
      _cx             =   35586
      _cy             =   14473
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
      BackColorSel    =   8454143
      ForeColorSel    =   128
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   2
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   -1575
      Top             =   3015
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
      Left            =   4455
      Top             =   1575
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
      Left            =   2475
      Top             =   1125
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   270
      Top             =   1350
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   10680
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
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
         Left            =   4095
         TabIndex        =   20
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
         TabIndex        =   21
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
         Index           =   3
         Left            =   12150
         TabIndex        =   22
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
         Index           =   4
         Left            =   16155
         TabIndex        =   23
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
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2850
      Left            =   1035
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   3825
      Visible         =   0   'False
      Width           =   6225
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   0
         Width           =   4740
         Begin VB.OptionButton optDelivery 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»œÊ‰  «—ÌŒ «” ·«„"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1935
            RightToLeft     =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   270
            Width           =   1725
         End
         Begin VB.OptionButton optDelivery 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "» «—ÌŒ «” ·«„"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   270
            Width           =   1410
         End
         Begin VB.OptionButton optDelivery 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3825
            RightToLeft     =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   270
            Value           =   -1  'True
            Width           =   780
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   855
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1260
         Visible         =   0   'False
         Width           =   3615
         Begin VB.OptionButton optDoc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»›« Ê—…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   270
            Width           =   870
         End
         Begin VB.OptionButton optDoc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»œÊ‰ ›« Ê—…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1305
            RightToLeft     =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   270
            Width           =   1230
         End
         Begin VB.OptionButton optDoc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   270
            Value           =   -1  'True
            Width           =   780
         End
      End
   End
End
Attribute VB_Name = "ladingBillfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Public myform As Form
Public sFilter As String, sSup As String
Dim bStopCell As Boolean
Dim cList1 As String, cList2 As String
Dim oSearchInv As New Search_abd
Dim oSearchship As New Search_abd
Dim oSearchLading As New Search_abd
Dim con As New ADODB.Connection
Private Sub myLoadGrd()
Dim cString As String
Dim cWhere As String

Dim aFields(14)
aFields(0) = "LADING_BILL.SHIP_NO AS [—ﬁ„ «·»Ê·Ì’…]"
aFields(1) = "FORMAT(LADING_BILL.DATE,'yyyy/M/d') AS [ «—ÌŒ «·»Ê·Ì’…]"
aFields(2) = "LADING_BILL.SHIP AS [‘—ﬂ… «·‘Õ‰]"
aFields(3) = "LADING_BILL.TYPE AS [‰Ê⁄ «·»Ê·Ì’…]"
aFields(4) = "LADING_BILL.CHARGE1 AS [„’«—Ì› ‘Õ‰]"
aFields(5) = "LADING_BILL.CHARGE2 AS [„’«—Ì› »Ê·Ì’…]"
aFields(6) = "LADING_BILL.DOC_NO AS [—ﬁ„ «·›« Ê—…]"
aFields(7) = "FORMAT(FILE6_20h.DATE,'yyyy/M/d') AS [ «—ÌŒ «·›« Ê—…]"
aFields(8) = "FILE6_20H.NAME AS [«”„ «·⁄„Ì·]"
aFields(9) = "FILE6_20H.PHONE AS [ ·Ì›Ê‰ «·⁄„Ì·]"
aFields(10) = "FILE6_20H.ADDRESS AS [«·⁄‰Ê«‰]"
aFields(11) = "LADING_BILL.DOC_NO_RET AS [›« Ê—… «·„— Ã⁄]"
aFields(12) = "FORMAT(F.DATE,'yyyy/M/d') AS [ «—ÌŒ «·„— Ã⁄]"
aFields(13) = "FORMAT(LADING_BILL.DATE_DELIVERY,'yyyy/M/d') AS [ «—ÌŒ «·«” ·«„]"
aFields(14) = "LADING_BILL.ID"

cString = "SELECT " & arString(aFields, ",") & _
          " FROM LADING_BILL " & _
          " INNER JOIN FILE6_20H ON LADING_BILL.DOC_NO = FILE6_20H.DOC_NO" & _
          " LEFT JOIN FILE6_20H AS F ON LADING_BILL.DOC_NO_RET = F.DOC_NO"

cWhere = retWhere

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

cString = cString & " ORDER BY LADING_BILL.[DATE],LADING_BILL.ID"
Set DATA11.Recordset = myCmd(cString, con)

myAddItem
fixGrd
CalcTotals
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myload()
myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End Sub

Private Sub chkDay_Click()
myLoadGrd
End Sub
Private Sub chkDuplicate_Click()
myload
End Sub

Private Sub chkNoNum_Click()
myload
End Sub

Private Sub chkNoQuant_Click()
myload
End Sub

Private Sub chkOrderReceipt_Click()
myload
End Sub
Private Sub chkDay_delivery_Click()
myLoadGrd
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdGo_Click()
myload
End Sub

Private Sub cmdShip_Click()
ShipLookup Me, oSearchship, , , IIf(cmdShip.Tag <> "", "ﬂ· «·‘—ﬂ« ", "")
End Sub
Private Sub cmdType_Click()
LadingLookup Me, oSearchLading, , , IIf(cmdType.Tag <> "", "ﬂ· «‰Ê«⁄ «·»Ê«·’", "")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) And ActiveControl.Tag <> "filter" Then
        KeyCode = 0
        grid1.SetFocus
    End If
End If
End Sub
Private Sub Form_Resize()
Dim nHeight As Long
nHeight = Me.Height - (grid1.Top + Picture1.Height + panel1(0).Height + 600)
grid1.Height = IIf(nHeight < 2000, 2000, nHeight)
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveText Me
Set ladingBillfrm = Nothing
Err.Clear
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then Exit Sub

If grid1.Row = grid1.Rows - 1 Then
    myAddItem
ElseIf Row = grid1.Rows - 2 And (Col = 2) Then
    'MyEditItem grid1, Row, Col
End If

Dim bNew As Boolean
bNew = grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = ""

Dim nId As Variant
If myreplace(Row, nId) Then
    If bNew Then
        If IsNull(nId) Then
            myLoadGrd
        Else
             grid1.TextMatrix(Row, grid1.Cols - 1) = nId
             'grid1.TextMatrix(Row, grid1.Cols - 2) = GetComputerName
        End If
        grid1.ShowCell grid1.Rows - 1, 1
    End If
    Grid1_EnterCell
Else
    myLoadGrd
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myLoadGrd
End Sub
Private Sub CalcTotal(Row As Long)
End Sub
Private Sub Grid1_EnterCell()

'aFields(0) = "LADING_BILL.SHIP_NO AS [—ﬁ„ «·»Ê·Ì’…]"
'aFields(1) = "FORMAT(LADING_BILL.DATE,'yyyy/M/d') AS [ «—ÌŒ «·»Ê·Ì’…]"
'aFields(2) = "LADING_BILL.SHIP AS [‘—ﬂ… «·‘Õ‰]"
'aFields(3) = "LADING_BILL.TYPE AS [‰Ê⁄ «·»Ê·Ì’…]"
'aFields(4) = "LADING_BILL.CHARGE1 AS [„’«—Ì› ‘Õ‰]"
'aFields(5) = "LADING_BILL.CHARGE2 AS [„’«—Ì› »Ê·Ì’…]"
'aFields(6) = "LADING_BILL.DOC_NO AS [—ﬁ„ «·›« Ê—…]"
'aFields(7) = "FORMAT(FILE6_20h.DATE,'yyyy/M/d') AS [ «—ÌŒ «·›« Ê—…]"
'aFields(8) = "FILE6_20H.NAME AS [«”„ «·⁄„Ì·]"
'aFields(9) = "FILE6_20H.PHONE AS [ ·Ì›Ê‰ «·⁄„Ì·]"
'aFields(10) = "FILE6_20H.ADDRESS AS [«·⁄‰Ê«‰]"
'aFields(11) = "LADING_BILL.DOC_NO_RET AS [›« Ê—… «·„— Ã⁄]"
'aFields(12) = "FORMAT(F.DATE,'yyyy/M/d') AS [ «—ÌŒ «·„— Ã⁄]"
'aFields(13) = "FORMAT(LADING_BILL.DATE_DELIVERY,'yyyy/M/d') AS [ «—ÌŒ «·«” ·«„]"
'aFields(14) = "LADING_BILL.ID"

With grid1
    
    If (cBranch <> "00") Then
        grid1.Editable = flexEDNone
    ElseIf grid1.TextMatrix(grid1.Row, 11) <> "" Then
        grid1.Editable = flexEDNone
    ElseIf grid1.TextMatrix(grid1.Row, 13) <> "" Then
        grid1.Editable = flexEDNone
    ElseIf grid1.Col = 0 Or grid1.Col = 1 Or grid1.Col = 2 Or grid1.Col = 3 Or grid1.Col = 4 Or grid1.Col = 5 Then
        grid1.Editable = flexEDKbdMouse
    ElseIf grid1.Col = 6 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
        grid1.Editable = flexEDKbdMouse
    Else
        grid1.Editable = flexEDNone
    End If
End With
End Sub
Private Sub Form_Load()
bEdit = True

openCon con

Set data2.Recordset = myCmd(YearsString(Year(Date)), con)
Set xYear.RowSource = data2
xYear.ListField = "CODE"
xYear.BoundColumn = "CODE"

Set DATA3.Recordset = myCmd(MonthString(Month(Date)), con)
Set xMonth.RowSource = DATA3
xMonth.ListField = "CODE"
xMonth.BoundColumn = "CODE"

Set grid1.DataSource = DATA11

With grid1
cList1 = StrList2("select code,desca from SHIP order by desca")
cList2 = StrList2("select code,desca from lading_codes order by desca")
End With

myload
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
LoadText Me
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'If cBranch <> "00" Then Exit Sub
If KeyCode = 13 Then
    If Col = 2 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    If Col = 3 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
'If cBranch <> "00" Then Exit Sub
If KeyAscii = 13 Then
    If grid1.Col = 2 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    If grid1.Col = 3 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "—ﬁ„ «·»Ê·Ì’… €Ì— „”Ã·"
        Cancel = True
    End If
ElseIf Col = 1 Then
    grid1.EditText = myFormat_p(grid1.EditText)
    If (Not IsDate(grid1.EditText)) Then
        MsgBox " «—ÌŒ «·»Ê·Ì’… €Ì— „”Ã·"
        Cancel = True
    End If
ElseIf Col = 2 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "‘—ﬂ…«·‘Õ‰ €Ì— „”Ã·…"
        Cancel = True
    End If
ElseIf Col = 3 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "‰Ê⁄ «·»Ê·Ì’… €Ì— „”Ã·"
        Cancel = True
    End If
ElseIf Col = 4 Or Col = 5 Then
    If Not IsNumeric(grid1.EditText) And Trim(grid1.EditText) <> "" Then
        Cancel = True
    End If
ElseIf Col = 4 + 2 Then
    If IsNumeric(grid1.EditText) Then
        If Not invDesc(grid1.EditText, Row) Then
            Cancel = True
        End If
    ElseIf Trim(grid1.EditText) = "" Then
        grid1.TextMatrix(Row, 5) = ""
        grid1.TextMatrix(Row, 9) = ""
    Else
        Cancel = True
    End If
End If
End Sub
Private Function invDesc(pInv_no As String, Row As Long) As Boolean
Dim cmd As New ADODB.Command
Set cmd = myCommand("dbo.sp_lading_invoice", con, adStoredProc, AddFlag(Empty, "doc_no", pInv_no))

If IsNull(cmd.Parameters("@date").Value) Then
    MsgBox "—ﬁ„ ›« Ê—… €Ì— ’ÕÌÕ"
    Exit Function
ElseIf cmd.Parameters("@online").Value <> 2 Then
    MsgBox "«·›« Ê—… ·Ì”  «Ê‰ ·«Ì‰"
    Exit Function
ElseIf cmd.Parameters("@total_item").Value < 0 Then
    MsgBox "›« Ê—… „— Ã⁄"
    Exit Function
ElseIf Not IsNull(cmd.Parameters("@lading").Value) Then
    If MsgBox("«·›« Ê—… „”Ã· ·ÌÂ« »Ê·Ì’… „”»ﬁ…„‰ ﬁ»·", vbOKCancel + vbDefaultButton2) <> vbOK Then
        Exit Function
    End If
ElseIf IsNull(cmd.Parameters("@SHIP_NO").Value) Then
    MsgBox "«·›« Ê—… ·Ì” ·Â« »Ê·Ì’… ‘Õ‰"
    Exit Function
ElseIf IsNull(cmd.Parameters("@SHIP").Value) Then
    MsgBox "«·›« Ê—… ·Ì” ·Â« ‘—ﬂ… ‘Õ‰"
    Exit Function
ElseIf Not cmd.Parameters("@ISCLOSED").Value Then
    MsgBox "«·›« Ê—… ·Ì”  „€·ﬁ…"
    Exit Function
End If
    grid1.TextMatrix(Row, 5 + 2) = myFormat_p(cmd.Parameters("@DATE").Value)
    
    grid1.TextMatrix(Row, 6 + 2) = cmd.Parameters("@name").Value & ""
    
    grid1.TextMatrix(Row, 7 + 2) = cmd.Parameters("@PHONE").Value & ""
    
    'grid1.TextMatrix(Row, 9 + 2) = cmd.Parameters("@DOC_RETURN").Value & ""

    'grid1.TextMatrix(Row, 10 + 2) = myFormat_p(cmd.Parameters("@DATE_DELIVERY").Value)
    invDesc = True
End Function
Private Sub fixGrd()
With grid1
.RowHeight(0) = 600
.ColWidth(0) = 1600
.ColWidth(1) = 1300
.ColWidth(2) = 2000
.ColWidth(3) = 2000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 1600
.ColWidth(7) = 1300
.ColWidth(8) = 2500
.ColWidth(9) = 2000
.ColWidth(10) = 5000
.ColWidth(11) = 1600
.ColWidth(12) = 1300

.ColComboList(2) = cList1
.ColComboList(3) = cList2

.ColHidden(8) = True
.ColHidden(9) = True
.ColHidden(10) = True
.ColHidden(.Cols - 1) = True

For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.WordWrap = True
End With
End Sub
Private Sub CalcTotals()
panel1(0).Caption = IIf(grid1.Rows > 2, "⁄œœ «·«Ì’«·«  " & grid1.Rows - 2, "·« ÌÊÃœ «Ì’«·« ")
End Sub
Private Function validRow(Row As Long) As Boolean
With grid1
If Trim(grid1.TextMatrix(Row, 0)) = "" Then Exit Function
If Not IsDate(grid1.TextMatrix(Row, 1)) Then Exit Function
If Not IsNumeric(grid1.TextMatrix(Row, 2)) Then Exit Function
If Not IsNumeric(grid1.TextMatrix(Row, 3)) Then Exit Function
If grid1.TextMatrix(Row, 4 + 2) = "" Then Exit Function
'If Trim(.TextMatrix(Row, 7)) = "" Then Exit Function
'If Trim(.TextMatrix(Row, 8)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
'If cBranch <> "00" Then Exit Sub
If KeyCode = 13 Then
     CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 112 And grid1.Col = 4 + 2 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    SalesLadingLookup Me, oSearchInv, , , IIf(grid1.TextMatrix(grid1.Row, 4 + 2) <> "", "»œÊ‰ ›« Ê—…", "")
ElseIf KeyCode = 46 And grid1.Row <> 0 And grid1.Row <> grid1.Rows - 1 Then
    If Trim(grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) <> "" Then
        If MsgBox("Õ–› «·«Ì’«· ø", vbOKCancel + vbDefaultButton2) = vbOK Then
            If grid1.TextMatrix(grid1.Row, 11) <> "" Then
                MsgBox "··»Ê·Ì’… „” ‰œ „— Ã⁄"
                Exit Sub
            ElseIf grid1.TextMatrix(grid1.Row, 13) <> "" Then
                MsgBox "··»Ê·’…  «—ÌŒ «” ·«„"
                Exit Sub
            End If
            
            On Error GoTo myerror
            Dim nAffect As Long
            con.Execute "delete from LADING_BILL " & _
                        " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1) & _
                        " AND DOC_NO_RET IS NULL" & _
                        " AND DATE_DELIVERY IS NULL", nAffect
            If nAffect = 0 Then
                MsgBox "·„ Ì „ Õ–› «·»Ê·Ì’… !! ﬁœ ÌﬂÊ‰ ·Â« ›« Ê—… „— Ã⁄ «Ê  «—ÌŒ «” ·«„"
                Exit Sub
            End If
            
            grid1.RemoveItem grid1.Row
            grid1.SetFocus
            CalcTotals
            bStopCell = True
        End If
    End If
ElseIf KeyCode = 112 And grid1.Col = 3 Then
    'SupLookupAll Me, oSearchSup
End If
Exit Sub
myerror:
If Err.Number <> 0 Then MsgBox Err.Description
myLoadGrd
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col < 4 + 2 Then
    grid1.Select Row, Col + 1
    grid1.ShowCell grid1.Row, grid1.Col
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 4 + 2)
    grid1.ShowCell grid1.Row, grid1.FrozenCols
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub myAddItem()
With grid1
    .AddItem ""
    If cmdType.Tag <> "" Then
        grid1.TextMatrix(grid1.Rows - 1, 3) = cmdType.Tag
'    ElseIf grid1.Rows > 2 Then
'        grid1.TextMatrix(grid1.Rows - 1, 3) = grid1.TextMatrix(grid1.Rows - 2, 3)
    End If
End With
End Sub
Private Function myreplace(Row As Long, ByRef nId As Variant) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[SHIP_NO]", addstring(grid1.TextMatrix(Row, 0)))
aInsert = AddFlag(aInsert, "[DATE]", addDate(grid1.TextMatrix(Row, 1)))
aInsert = AddFlag(aInsert, "[SHIP]", addstring(grid1.TextMatrix(Row, 2)))
aInsert = AddFlag(aInsert, "[TYPE]", addvalue(grid1.TextMatrix(Row, 3)))
aInsert = AddFlag(aInsert, "[CHARGE1]", grid1.ValueMatrix(Row, 4))
aInsert = AddFlag(aInsert, "[CHARGE2]", grid1.ValueMatrix(Row, 5))
aInsert = AddFlag(aInsert, "[doc_no]", addstring(grid1.TextMatrix(Row, 6)))
con.BeginTrans

On Error GoTo myerror
If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
    aInsert = AddFlag(aInsert, "[HOST_NAME]", addstring(GetComputerName))
    con.Execute addInsert(aInsert, "LADING_BILL")
    nId = myField("SELECT SCOPE_IDENTITY()", con)
Else
    con.Execute addUpdate(aInsert, "LADING_BILL", "ID = " & grid1.TextMatrix(Row, grid1.Cols - 1))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function
Public Sub myProc()
If ActiveControl.Name = grid1.Name And grid1.Col = 4 + 2 Then
    If Not invDesc(oSearchInv.grid1.TextMatrix(oSearchInv.grid1.Row, 0), grid1.Row) Then Exit Sub
    grid1.TextMatrix(grid1.Row, 4 + 2) = oSearchInv.grid1.TextMatrix(oSearchInv.grid1.Row, 0)
    Grid1_AfterEdit grid1.Row, 4 + 2
    CellPos 13, grid1.Row, 4 + 2
    oSearchInv.Hide
ElseIf ActiveControl.Name = cmdShip.Name Then
    If oSearchship.grid1.TextMatrix(oSearchship.grid1.Row, 0) = "" Then
        cmdShip.Tag = ""
        cmdShip.Caption = cmdShip.TagVariant
    Else
        cmdShip.Tag = oSearchship.grid1.TextMatrix(oSearchship.grid1.Row, 0)
        cmdShip.Caption = oSearchship.grid1.TextMatrix(oSearchship.grid1.Row, 1)
    End If
    myload
    oSearchship.Hide
ElseIf ActiveControl.Name = cmdType.Name Then
    If oSearchLading.grid1.TextMatrix(oSearchLading.grid1.Row, 0) = "" Then
        cmdType.Tag = ""
        cmdType.Caption = cmdType.TagVariant
    Else
        cmdType.Tag = oSearchLading.grid1.TextMatrix(oSearchLading.grid1.Row, 0)
        cmdType.Caption = oSearchLading.grid1.TextMatrix(oSearchLading.grid1.Row, 1)
    End If
    myload
    oSearchLading.Hide
End If
End Sub
Private Sub xCar_Change()
myLoadGrd
End Sub

Private Sub xDiffer_Click()
'myLoadGrd
End Sub
Private Sub optclose_Click(Index As Integer)
myload
End Sub

Private Sub xDoc_no_Change()
'myLoadGrd
End Sub
Private Sub xMonth_Click(Area As Integer)
myLoadGrd
End Sub
Private Function retWhere() As String
Dim cWhere As New ChilkatStringBuilder

If Trim(xDoc_no.text) <> "" Then
    cWhere.Append " AND " & " LADING_BILL.DOC_NO = " & MyParn(xDoc_no.text)
End If

If Trim(XSHIP_NO.text) <> "" Then
    cWhere.Append " AND " & " LADING_BILL.SHIP_NO = " & MyParn(XSHIP_NO.text)
End If

If IsDate(xDate1.text) Then
    cWhere.Append " AND " & " LADING_BILL.DATE >= " & addDate(xDate1.text)
End If

If IsDate(xdate2.text) Then
    cWhere.Append " AND " & " LADING_BILL.DATE <= " & addDate(xdate2.text)
End If

If chkDay.Value = 1 Then
    cWhere.Append " AND " & " LADING_BILL.DATE = " & addDate(Date)
End If

If optclose(1).Value Then
    'cWhere.Append " AND (LADING_BILL.CLOSED_INV = 0  OR LADING_BILL.CLOSED_NO_INV = 0)"
    cWhere.Append " AND (LADING_BILL.CLOSED = 0)"
ElseIf optclose(2).Value Then
    cWhere.Append " AND (LADING_BILL.CLOSED = 1)"
    'cWhere.Append " AND (NOT (LADING_BILL.CLOSED_INV = 0 OR LADING_BILL.CLOSED_NO_INV = 0))"
End If

'If optDoc(1).Value Then
'    cWhere.Append " AND LADING_BILL.DOC_NO IS NULL"
'ElseIf optDoc(2).Value Then
'    cWhere.Append " AND LADING_BILL.DOC_NO IS NOT NULL"
'End If

'If optDelivery(1).Value Then
'    cWhere.Append " AND LADING_BILL.DATE_DELIVERY IS NULL"
'ElseIf optDelivery(2).Value Then
'    cWhere.Append " AND LADING_BILL.DATE_DELIVERY IS NOT NULL"
'End If

If IsDate(xdate_Delivery1.text) Then
    cWhere.Append " AND " & " FILE6_20H.date_delivery >= " & addDate(xdate_Delivery1.text)
End If

If IsDate(xdate_Delivery2.text) Then
    cWhere.Append " AND " & " FILE6_20H.date_delivery <= " & addDate(xdate_Delivery2.text)
End If
'
If chkDay_delivery.Value = 1 Then
    cWhere.Append " AND " & " FILE6_20H.DATE_DELIVERY = " & addDate(Date)
End If
'
If xYear.BoundText <> "" Then
    cWhere.Append " AND " & " YEAR(FILE6_20H.DATE) = " & xYear.BoundText
End If

If xMonth.BoundText <> "" Then
    cWhere.Append " AND " & " MONTH(FILE6_20H.DATE) = " & xMonth.BoundText
End If

If Trim(xName.text) <> "" Then
    cWhere.Append " AND " & MyParnAnd(xName.text, "FILE6_20H.NAME")
End If

If Trim(xphone.text) <> "" Then
    cWhere.Append " AND " & MyParnAnd(xphone.text, "FILE6_20H.PHONE")
End If

If cmdType.Tag <> "" Then
    cWhere.Append " AND " & "LADING_BILL.TYPE = " & cmdType.Tag
End If

If cmdShip.Tag <> "" Then
    cWhere.Append " AND " & "FILE6_20H.SHIP = " & MyParn(cmdShip.Tag)
End If
retWhere = Mid(cWhere.GetAsString, 6)
End Function
Private Sub xYear_Change()
myLoadGrd
End Sub

Private Sub xship_no_GotFocus()
myGotFocus XSHIP_NO
End Sub
Private Sub xship_no_LostFocus()
myLostFocus XSHIP_NO
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_no
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_no
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub xYear_GotFocus()
myGotFocus xYear
End Sub
Private Sub xYear_LostFocus()
myLostFocus xYear
If Not xYear.MatchedWithList Then xYear.BoundText = ""
End Sub
Private Sub xMonth_GotFocus()
myGotFocus xMonth
End Sub
Private Sub xMonth_LostFocus()
myLostFocus xMonth
If Not xMonth.MatchedWithList Then xMonth.BoundText = ""
End Sub
Private Sub xName_GotFocus()
myGotFocus xName
End Sub
Private Sub xname_LostFocus()
myLostFocus xName
End Sub
Private Sub xPhone_GotFocus()
myGotFocus xphone
End Sub
Private Sub XPHONE_LostFocus()
myLostFocus xphone
End Sub
Private Sub xdate_Delivery2_GotFocus()
myGotFocus xdate_Delivery2
End Sub
Private Sub xdate_Delivery2_LostFocus()
myLostFocus xdate_Delivery2
myValidDate xdate_Delivery2
End Sub
Private Sub xdate_Delivery1_GotFocus()
myGotFocus xdate_Delivery1
End Sub
Private Sub xdate_Delivery1_LostFocus()
myLostFocus xdate_Delivery1
myValidDate xdate_Delivery1
End Sub
