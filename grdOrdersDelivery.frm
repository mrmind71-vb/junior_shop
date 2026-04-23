VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdOrdersDeliveryfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   " ﬁ—Ì— ÿ·»Ì«  - «” ·«„ - „»Ì⁄«  - —’Ìœ - ‰”»…  Ê—Ìœ  «·√’‰«›  ›’Ì·Ï"
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
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   0
      Width           =   13785
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   540
         Width           =   2805
      End
      Begin VB.TextBox xModelFact 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         MaxLength       =   15
         TabIndex        =   10
         Top             =   180
         Width           =   2805
      End
      Begin VB.TextBox xScal 
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
         Left            =   10935
         MaxLength       =   15
         TabIndex        =   2
         Top             =   540
         Width           =   1320
      End
      Begin VB.TextBox xbarcode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8505
         MaxLength       =   15
         TabIndex        =   4
         Top             =   900
         Width           =   3750
      End
      Begin VB.TextBox xColor 
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
         Left            =   8505
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   1545
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
         Left            =   10935
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
         Left            =   8505
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox XPO_NO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8505
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1260
         Width           =   3750
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   330
         Left            =   3735
         TabIndex        =   7
         Top             =   540
         Width           =   3570
         _ExtentX        =   6297
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
      Begin MSDataListLib.DataCombo xFact 
         Height          =   330
         Left            =   3735
         TabIndex        =   6
         Top             =   180
         Width           =   3570
         _ExtentX        =   6297
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
      Begin MSDataListLib.DataCombo xcode 
         Height          =   330
         Left            =   3735
         TabIndex        =   9
         Top             =   1260
         Width           =   3570
         _ExtentX        =   6297
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
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   330
         Left            =   3735
         TabIndex        =   8
         Top             =   900
         Width           =   3570
         _ExtentX        =   6297
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
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·«”„"
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
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   585
         Width           =   405
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "«··Ê‰"
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
         TabIndex        =   31
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "«·„ﬁ«”"
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
         Left            =   12330
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Left            =   10170
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Left            =   12330
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   225
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·„’‰⁄"
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
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ã„Ê⁄…"
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
         Index           =   1
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "SKU"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   225
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ"
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
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Barcode ASIN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   12330
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁ”„"
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
         Index           =   0
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No P.O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12330
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1305
         Width           =   660
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   675
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   945
      Width           =   5775
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   16
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
         Picture         =   "grdOrdersDelivery.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1170
         TabIndex        =   15
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
         Picture         =   "grdOrdersDelivery.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdOrdersDelivery.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   4635
         TabIndex        =   12
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
         Picture         =   "grdOrdersDelivery.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2295
         TabIndex        =   14
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
         Picture         =   "grdOrdersDelivery.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSql 
         Height          =   555
         Left            =   3420
         TabIndex        =   13
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
         Picture         =   "grdOrdersDelivery.frx":BBFA
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6540
      Left            =   90
      TabIndex        =   17
      Top             =   1710
      Width           =   20175
      _cx             =   35586
      _cy             =   11536
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Rows            =   1
      Cols            =   20
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
      RightToLeft     =   0   'False
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   4770
      Top             =   225
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
      Left            =   4365
      Top             =   585
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
      TabIndex        =   19
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
Attribute VB_Name = "grdOrdersDeliveryfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearch As New Search_abd
Dim aHeader()
Private Sub CMD_EXEL_Click()
End Sub
Private Sub CMD_PRINT_Click()
End Sub
Private Sub cmdExcel_Click()
Me.MousePointer = 11

Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aRow = AddFlag(aRow, aSub)


If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 40)
    aRow = AddFlag(aRow, aSub)
End If
ToFileExelNew grid1, , , aRow, , 0.9, , , , , , Me, Array(Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), retHeader(aHeader, 4, 5))

Me.MousePointer = 0
Fixgrd
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
Me.MousePointer = vbHourglass
myload
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
grid1.Height = IIf(Me.Height - grid1.Top - 1000 < 3000, 3000, Me.Height - grid1.Top - 1000)
End Sub
Private Sub Form_Load()
openCon con

Set data1.Recordset = mycmd("Select Code,DescA From File1_50 ORDER BY DESCA", con)
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set DATA2.Recordset = mycmd("Select Code,DescA From File1_10SC ORDER BY DESCA", con)
Set XSECTION.RowSource = DATA2
XSECTION.ListField = "Desca"
XSECTION.BoundColumn = "Code"


Set DATA3.Recordset = cmd("Select code ,descA From file4_10 ORDER BY DESCA", con).Execute
Set xcode.RowSource = DATA3
xcode.ListField = "Desca"
xcode.BoundColumn = "CODE"

Set DATA4.Recordset = cmd("Select code ,desca From fact ORDER BY code ", con).Execute
Set xFact.RowSource = DATA4
xFact.ListField = "Desca"
xFact.BoundColumn = "Code"

Set grid1.DataSource = DATA11

Fixgrd
End Sub
Private Sub myload(Optional bString As Boolean = False)
With grid1

Dim cString As New ChilkatStringBuilder
ReDim aHeader(9)
cString.Append "SELECT FILE1_10.ITEM," & _
          "FACT.DESCA," & _
          "FILE1_10.MODELFACT0," & _
          "COALESCE(FILE1_10.DESCA2,FILE1_10.DESCA) AS DESCA," & _
          "FILE1_10.COLOR," & _
          "FILE1_10.SCAL," & _
          "FILE1_10.COST," & _
          "FILE1_10.costitem," & _
          "FILE1_10.PRICE," & _
          "FILE1_10.PRICE2," & _
          "FILE1_10.BARCODE," & _
          "FILE1_10.ENA," & _
          "FILE4_10.DESCA," & _
          "SUM(vw_orders_delivery.QUANT_ORDERED) AS QUANT," & _
          "SUM(vw_orders_delivery.QUANT_CONFIRMED) AS QUANT_CONFIRMED," & _
          "SUM(vw_orders_delivery.QUANT_RECEIVED) AS RECEVD," & _
          "SUM(vw_orders_delivery.QUANT_SOLD) AS SALES," & _
          "COUNT(DISTINCT vw_orders_delivery.PO_NO) AS COUNT1," & _
          "COUNT(DISTINCT vw_orders_delivery.PO_NO_RCV) AS COUNT2," & _
          "COUNT(DISTINCT vw_orders_delivery.PO_NO_SALES) AS COUNT3"
          
cString.Append " FROM FILE1_10" & _
                " INNER JOIN FACT " & _
                " ON FILE1_10.FACT = FACT.CODE" & _
                " INNER JOIN FILE4_10" & _
                " ON FILE4_10.CODE = FILE1_10.CODE" & _
                " INNER JOIN vw_orders_delivery" & _
                " INNER JOIN FILE6_51H ON vw_orders_delivery.PO_NO = FILE6_51H.DOC_NO" & _
                " ON FILE1_10.ITEM = vw_orders_delivery.ITEM"

Dim cWhere As New ChilkatStringBuilder

If xcode.MatchedWithList Then
    cWhere.Append "FILE1_10.CODE = " & MyParn(xcode.BoundText)
    cWhere.Append " AND "
    aHeader(0) = "«·„Ê—œ : " & xcode.text
End If

If xModelFact.text <> "" Then
    cWhere.Append "FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    cWhere.Append " AND "
    aHeader(1) = "SKU : " & xModelFact.text
End If

If xGroup.MatchedWithList Then
    cWhere.Append " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    cWhere.Append " AND "
    aHeader(2) = "«·„Ã„Ê⁄… : " & xGroup.text
End If

If XSECTION.MatchedWithList Then
    cWhere.Append "FILE1_10.[SECTION] = " & Val(XSECTION.BoundText)
    cWhere.Append " AND "
    aHeader(3) = "«·ﬁ”„ : " & xGroup.text
End If

If xFact.MatchedWithList Then
    cWhere.Append "FILE1_10.FACT = " & MyParn(xFact.BoundText)
    cWhere.Append " AND "
    aHeader(4) = "«·„’‰⁄ : " & xGroup.text
End If

If IsDate(xDate1.text) Then
    cWhere.Append "FILE6_51H.DATE >= " & DateSq(xDate1.text)
    cWhere.Append " AND "
    aHeader(5) = BetweenString(xDate1.text, xDate2.text)
End If

If IsDate(xDate2.text) Then
    cWhere.Append "FILE6_51H.DATE <= " & DateSq(xDate2.text)
    cWhere.Append " AND "
    aHeader(5) = BetweenString(xDate1.text, xDate2.text)
End If

If XPO_NO.text <> "" Then
    cWhere.Append "PO_NO = " & MyParn(XPO_NO.text)
    cWhere.Append " AND "
    aHeader(6) = "—ﬁ„ «·ÿ·»Ì… : " & XPO_NO.text
End If

If xDesca.text <> "" Then
    cWhere.Append MyParnAnd(xDesca.text, "file1_10.desca")
    cWhere.Append " AND "
    aHeader(6) = "«”„ «·’‰› : " & XPO_NO.text
End If

If xbarcode.text <> "" Then
    cWhere.Append MyParnAnd(xbarcode.text, "file1_10.BARCODE+FILE1_10.ENA")
    cWhere.Append " AND "
    aHeader(7) = "BARCODE & ASIN" & xbarcode.text
End If

If Trim(xScal.text) <> "" Then
    cWhere.Append "FILE1_10.SCAL = " & MyParn(xScal.text)
    cWhere.Append " AND "
    aHeader(8) = "«·„ﬁ«”" & xScal.text
End If

If Trim(xColor.text) <> "" Then
    cWhere.Append MyParnAnd(xColor, "FILE1_10.COLOR")
    cWhere.Append " AND "
    aHeader(8) = "«··Ê‰" & xColor.text
End If

If cWhere.Length <> 0 Then
    cWhere.Shorten 5
    cString.Append " WHERE " & cWhere.GetAsString()
End If

cString.Append " GROUP BY FILE1_10.model," & _
               " FACT.DESCA," & _
               " FILE1_10.DESCA," & _
               " FILE1_10.DESCA2," & _
               " FILE1_10.MOSM," & _
               " FILE1_10.MODELFACT0," & _
               " FILE1_10.COLOR," & _
               " FILE1_10.SCAL," & _
               " FILE1_10.COST," & _
               " FILE1_10.costitem," & _
               " FILE1_10.PRICE," & _
               " FILE1_10.PRICE2," & _
               " FILE1_10.BARCODE," & _
               " FILE1_10.ENA," & _
               " FILE4_10.DESCA," & _
               " FILE1_10.ITEM," & _
               " FILE1_10.MOSM," & _
               " FILE1_10.C_SCAL"

cString.Append " ORDER BY FILE1_10.model," & _
               " FACT.DESCA," & _
               " FILE1_10.MOSM," & _
               " FILE1_10.MODELFACT0," & _
               " FILE1_10.COLOR," & _
               " FILE1_10.C_SCAL"

If bString Then
    Clipboard.Clear
    Clipboard.SetText cString.GetAsString
    Exit Sub
End If

Set DATA11.Recordset = cmd(cString.GetAsString, con).Execute
End With
Fixgrd
End Sub
Sub Fixgrd()
With grid1
    .RowHeight(0) = 600
    .WordWrap = True
    
    .TextMatrix(0, 0) = "«·’‰›"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "SKU"
    .TextMatrix(0, 3) = "«·«”„"
    .TextMatrix(0, 4) = "«··Ê‰"
    .TextMatrix(0, 5) = "«·„ﬁ«”"
    
    .TextMatrix(0, 6) = "«·”⁄—"
    .TextMatrix(0, 7) = "Net Purchese PRICE"
    .TextMatrix(0, 8) = "Amazon cost"
    .TextMatrix(0, 9) = "List Price"
    
    .TextMatrix(0, 10) = "BARCODE"
    .TextMatrix(0, 11) = "ASIN"
    
    .TextMatrix(0, 12) = "«·„Ê—œ"
    .TextMatrix(0, 13) = "ﬂ„Ì… ÿ·»Ì…"
    .TextMatrix(0, 13 + 1) = "ﬂ„Ì… „ƒﬂœ…"
    .TextMatrix(0, 14 + 1) = "ﬂ„»… „”·„…"
    .TextMatrix(0, 15 + 1) = "ﬂ„»… „»«⁄…"
    
    .TextMatrix(0, 16 + 1) = "⁄œœ «·ÿ·»Ì« "
    .TextMatrix(0, 17 + 1) = "ÿ·»Ì«  „” ·„…"
    .TextMatrix(0, 18 + 1) = "ÿ·»Ì«  „»«⁄…"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1100
    .ColWidth(2) = 1300
    .ColWidth(3) = 3400
    .ColWidth(4) = 1700
    .ColWidth(5) = 800
    .ColWidth(6) = 1700
    .ColWidth(7) = 1700
    .ColWidth(8) = 1700
    .ColWidth(9) = 1300
    
    .ColWidth(10) = 1700
    .ColWidth(11) = 1800
    
    .ColWidth(12) = 1500
    
    .ColWidth(13) = 1000
    .ColWidth(13 + 1) = 1000
    .ColWidth(14 + 1) = 1000
    .ColWidth(15 + 1) = 1000
    .ColWidth(16 + 1) = 1000
    .ColWidth(17 + 1) = 1000
    .ColWidth(18 + 1) = 1000
    
    .ColHidden(6) = True
    .ColHidden(7) = True
    .ColHidden(8) = True
    .ColHidden(9) = True
    .ColHidden(12) = True

    For i = 13 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next

    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 11
    

    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 13, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 13 + 1, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 14 + 1, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 15 + 1, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdOrdersDeliveryfrm = Nothing
End Sub
Private Function MYVALID() As Boolean
MYVALID = True
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
printGrdNew.DOPRINT Me.grid1, 0.75, 0, "ÃÊ‰ÌÊ—", Me.Caption, retHeader(aHeader, 0, 2), , True, True, 8, , aRow, Array(1)

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

Private Sub xPO_NO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    OrdersLookup Me, oSearch
End If
End Sub

Private Sub xScal_GotFocus()
myGotFocus xScal
End Sub
Private Sub xScal_LostFocus()
myLostFocus xScal
End Sub
Private Sub xbarcode_GotFocus()
myGotFocus xbarcode
End Sub
Private Sub xbarcode_LostFocus()
myLostFocus xbarcode
End Sub
Private Sub xColor_GotFocus()
myGotFocus xColor
End Sub
Private Sub xColor_LostFocus()
myLostFocus xColor
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub XPO_NO_GotFocus()
myGotFocus XPO_NO
End Sub
Private Sub xpo_no_LostFocus()
myLostFocus XPO_NO
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub xFact_GotFocus()
myGotFocus xFact
End Sub
Private Sub xFACT_LostFocus()
myLostFocus xFact
If Not xFact.MatchedWithList Then xFact.BoundText = ""
End Sub
Private Sub xcode_GotFocus()
myGotFocus xcode
End Sub
Private Sub xcode_LostFocus()
myLostFocus xcode
If Not xcode.MatchedWithList Then xcode.BoundText = ""
End Sub
Private Sub xSection_GotFocus()
myGotFocus XSECTION
End Sub
Private Sub xSection_LostFocus()
myLostFocus XSECTION
If Not XSECTION.MatchedWithList Then XSECTION.BoundText = ""
End Sub
Public Sub myproc()
XPO_NO.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
oSearch.Hide
End Sub
