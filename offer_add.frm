VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form offerAddfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "«·⁄—Ê÷"
   ClientHeight    =   8835
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14790
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
   ScaleHeight     =   8835
   ScaleWidth      =   14790
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1410
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   45
      Width           =   3705
      Begin VB.Label xdoc_no 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   810
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label xRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   945
         Width           =   1410
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰”»… «·Œ’„"
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   990
         Width           =   1365
      End
      Begin VB.Label xDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   585
         Width           =   1410
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì «·Œ’„"
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   630
         Width           =   1365
      End
      Begin VB.Label xTotalInv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì ﬁÌ„… «·«’‰«›"
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   270
         Width           =   1905
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1410
      Left            =   7470
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   45
      Width           =   3750
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„” Õﬁ ··⁄—÷"
         Height          =   330
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label xQtyOffer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   945
         Width           =   1410
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬂ„Ì… „ »ﬁÌ…"
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label xRest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   585
         Width           =   1410
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬂ„»… „»«⁄…"
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label xSales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   225
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1410
      Left            =   11250
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   45
      Width           =   3480
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·⁄—÷"
         Height          =   330
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   630
         Width           =   870
      End
      Begin VB.Label xQty1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   585
         Width           =   1320
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄·ÌÂ„"
         Height          =   330
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   990
         Width           =   645
      End
      Begin VB.Label xQty2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì «·⁄—÷"
         Height          =   330
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   -9000
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   6525
      Visible         =   0   'False
      Width           =   8610
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„ »ﬁÌ ··Õ’Ê· ⁄·Ì «·⁄—÷ :"
         Height          =   330
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   585
         Width           =   2220
      End
      Begin VB.Label xRestOffer 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   585
         Width           =   1410
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   9945
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   825
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3795
      Left            =   3735
      TabIndex        =   0
      Top             =   1485
      Width           =   10995
      _cx             =   19394
      _cy             =   6694
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   -765
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   -765
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
   Begin VSFlex7Ctl.VSFlexGrid grdOffer 
      Height          =   4560
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   45
      Width           =   3570
      _cx             =   6297
      _cy             =   8043
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   400
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
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   8370
      Width           =   14790
      _ExtentX        =   26088
      _ExtentY        =   820
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
         TabIndex        =   3
         Top             =   45
         Width           =   3600
         _ExtentX        =   6350
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
         Left            =   3645
         TabIndex        =   4
         Top             =   45
         Width           =   3870
         _ExtentX        =   6826
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
         Left            =   7560
         TabIndex        =   5
         Top             =   45
         Width           =   3555
         _ExtentX        =   6271
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
         Left            =   11160
         TabIndex        =   6
         Top             =   45
         Width           =   3555
         _ExtentX        =   6271
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
         TabIndex        =   7
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   2985
      Left            =   90
      TabIndex        =   11
      Top             =   5310
      Width           =   14640
      _cx             =   25823
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
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
   Begin Threed.SSCommand cmdExit 
      Height          =   600
      Left            =   45
      TabIndex        =   34
      Top             =   4680
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1058
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
      Picture         =   "offer_add.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdDel 
      Height          =   600
      Left            =   1125
      TabIndex        =   35
      Top             =   4680
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1058
      _Version        =   196610
      ForeColor       =   255
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
      Picture         =   "offer_add.frx":2323
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "offer_add.frx":4905
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   600
      Left            =   2385
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1058
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
      Picture         =   "offer_add.frx":6CEF
      Alignment       =   8
      ButtonStyle     =   2
      PictureAlignment=   11
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "offer_add.frx":9614
   End
End
Attribute VB_Name = "offerAddfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As sales_abdfrm
Public bEdit As Boolean
Public sDoc_Offer As String
Dim oSearchItem As New Search_abd
Dim bStopCell As Boolean
Dim con As New ADODB.Connection
Private Sub CmdDel_Click()
grid1.Rows = 1
myload 1
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
cmdSave.Enabled = False
myForm.myProc
myDefine
grid1.SetFocus
End Sub

Private Sub Form_Load()
bEdit = True

openCon con
Set grdOffer.DataSource = data1

myLoadGrdOffer

fixGrd
Fixgrd2
End Sub
Private Sub myLoadGrdOffer()
Dim cString As String
cString = "SELECT FILE0_90H.DOC_NO," & _
          "FILE0_90H.DESCA, " & _
          "CAST(0 AS BIT)," & _
          "FILE0_90H.QTY1," & _
          "FILE0_90H.QTY2," & _
          "FILE0_90H.QTY2 " & _
          " FROM FILE0_90H  " & _
          " WHERE FILE0_90H.DATE1 <= " & DateSq(myForm.xDate.text) & _
          " AND FILE0_90H.DATE2 >= " & DateSq(myForm.xDate.text) & _
          " AND FILE0_90H.IS_OFFER = 1" & _
          " AND FILE0_90H.ISSTOP = 0"
Set data1.Recordset = cmd(cString, con).Execute
fixGrdOffer

If grdOffer.Rows = 2 Then
    grdOffer.TextMatrix(1, 2) = -1
    grdOffer_AfterEdit 1, 2
End If
End Sub
Sub fixGrdOffer()
With grdOffer
.RowHeight(0) = 600
.TextMatrix(0, 0) = "ﬂÊœ «·⁄—÷"
.TextMatrix(0, 1) = "«·⁄—÷"
.TextMatrix(0, 2) = "«Œ Ì«—"
.TextMatrix(0, 3) = "ﬂ„Ì…1"
.TextMatrix(0, 4) = "ﬂ„Ì…2"

.ColWidth(0) = 1500
.ColWidth(1) = 2500
.ColWidth(2) = 700

.ColHidden(0) = True
.ColHidden(3) = True
.ColHidden(4) = True
.ColHidden(5) = True

Dim i As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub myLoadDiffer()
Dim i As Long
With grid2

.Subtotal flexSTClear
If .Rows < 2 Then Exit Sub

For i = 1 To .Rows - 1
    nFound = grdTotal.FindRow(.TextMatrix(i, 0), , 0)
    If nFound = -1 Then
         .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbWhite
         .TextMatrix(i, grid2.Cols - 2) = 0
         .TextMatrix(i, grid2.Cols - 1) = .ValueMatrix(i, 6)
    Else
        .TextMatrix(i, grid2.Cols - 2) = grdTotal.ValueMatrix(nFound, 1)
        .TextMatrix(i, grid2.Cols - 1) = .ValueMatrix(i, 6) - .ValueMatrix(i, grid2.Cols - 2)
        If .ValueMatrix(i, grid2.Cols - 1) = 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0FFC0
        ElseIf .ValueMatrix(i, grid2.Cols - 1) > 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H80FFFF
        Else
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H8080FF
        End If
    End If
Next
End With

With grid1
For i = 1 To .Rows - 2
    nFound = grid2.FindRow(.TextMatrix(i, 0), , 0)
    If nFound = -1 Then
        .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H8080FF
        '.TextMatrix(i, 5) = "e1"
    Else
        If grid2.ValueMatrix(nFound, grid2.Cols - 1) = 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0FFC0
        ElseIf grid2.ValueMatrix(nFound, grid2.Cols - 1) > 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H80FFFF
             '.TextMatrix(i, 8) = "w"
        Else
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H8080FF
            '.TextMatrix(i, 8) = "e2"
        End If
    End If
Next
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set offerAddfrm = Nothing
End Sub
Private Sub grdOffer_AfterEdit(ByVal Row As Long, ByVal col As Long)
With grdOffer
Dim i As Long
For i = 1 To .Rows - 1
    If i <> Row Then
        .TextMatrix(i, 2) = 0
    Else
        myload i
    End If
Next
End With
End Sub
Private Sub myload(Row As Long)
With grdOffer
grid1.Rows = 1
If .ValueMatrix(Row, 2) = 0 Then
    xDoc_no.Caption = ""
    xQty1.Caption = ""
    xQty2.Caption = ""
    xTotal.Caption = ""
Else
    xDoc_no.Caption = .TextMatrix(Row, 0)
    xQty1.Caption = .TextMatrix(Row, 3)
    xQty2.Caption = .TextMatrix(Row, 4)
    xTotal.Caption = Myvalue(.ValueMatrix(Row, 3) + .ValueMatrix(Row, 4))
    grid1.AddItem ""
End If
CalcTotals
cmddel.Enabled = False
End With
End Sub
Private Sub myDefine()
Dim i As Long
Dim Row As Long
For i = 1 To grdOffer.Rows - 1
    If grdOffer.ValueMatrix(i, 2) <> 0 Then
        grdOffer.TextMatrix(i, 2) = 1
        myload i
    End If
Next
End Sub
Private Sub GrdOffer_Click()
If grdOffer.Row > 0 Then
    Dim i As Long
    If grdOffer.ValueMatrix(grdOffer.Row, 2) = 0 Then
        For i = 1 To grdOffer.Rows - 1
            If i <> grdOffer.Row Then
                grdOffer.TextMatrix(i, 2) = 0
            Else
                grdOffer.TextMatrix(i, 2) = -1
                myload i
            End If
        Next
    End If
End If
End Sub

Private Sub GrdOffer_EnterCell()
With grdOffer
    If .Rows <= 2 Then
        .Editable = flexEDNone
    ElseIf .col <> 2 Then
        .Editable = flexEDNone
    ElseIf bEdit Then
        .Editable = flexEDKbdMouse
    End If
End With
End Sub
Private Sub grid1_EnterCell()
With grid1
If Not bEdit Then
    .Editable = flexEDNone
ElseIf .col = 1 Or .col = 10 Then
    If xDoc_no.Caption <> "" Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
Else
    .Editable = flexEDNone
End If
End With
End Sub
Private Function validRow(Row As Long, Optional col As Long = -1) As Boolean
With grid1
    If Not IsDgt(.TextMatrix(Row, 1)) Then Exit Function
    If .ValueMatrix(Row, 10) <= 0 Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If xDoc_no.Caption = "" Then
    Exit Sub
ElseIf KeyCode = 112 Then
    ItemsLook Me, oSearchItem, , xDoc_no.Caption, True
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
        myRemove grid1.Row
        'bStopCell = True
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
End If
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
MakeSerial
CalcTotals
End Sub
Private Sub myLoadGrdTotal()
Dim cString As New ChilkatStringBuilder

Dim i As Long
If grid1.Rows < 3 Then
    grdTotal.Rows = 1
    Exit Sub
End If

cString.Append "WITH tb as ( "
For i = 1 To grid1.Rows - 2
    If i > 1 Then
        cString.Append " UNION ALL "
    End If
    
    cString.Append "SELECT " & _
                grid1.TextMatrix(i, 0) & " AS ITEM " & "," & _
                grid1.TextMatrix(i, 5) & " AS QUANT " & "," & _
                i & " AS ITEM_ORDER"
Next
cString.Append ")"
cString.Append "SELECT ITEM,SUM(QUANT) FROM tb GROUP BY ITEM ORDER BY MIN(ITEM_ORDER)"

Set data2.Recordset = mycmd(cString.GetAsString(), con)
fixGrdTotal

myLoadDiffer

End Sub
Sub fixGrdTotal()
With grdTotal
    .RowHeight(0) = 600
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·ﬂ„Ì…"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 1000
    
    
    Dim i As Long
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next

    .Subtotal flexSTClear
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTCount, -1, 0, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, 1, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"

End With
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
With grid1
    If col = 10 Then
        If Not IsNumeric(.EditText) Then
            Cancel = True
            Exit Sub
        ElseIf Val(.EditText) < 0 Then
            Cancel = True
        ElseIf Val(.EditText) > 0 Then
            If Val(xSales.Caption) + (Val(.EditText) - .ValueMatrix(Row, 10)) > Val(xTotal.Caption) Then
                MsgBox "«·ﬂ„Ì… «ﬂ»— „‰ ﬂ„Ì… «·⁄—÷"
                Cancel = True
                Exit Sub
            End If
            nBalance = fnBalance(grid1.TextMatrix(grid1.Row, 1), con, myForm.xStore.BoundText, myForm.xDate.text) - BalanceFound(grid1.TextMatrix(Row, 1), Row)
            If nBalance < Val(.EditText) Then
                Inform "«·—’Ìœ ·« Ì”„Õ"
                Exit Sub
            End If
        End If
    ElseIf col = 1 Then
        If Not GrdDesc(.EditText, Row) Then
            Cancel = True
        ElseIf .EditText <> .TextMatrix(Row, 1) Then
            .EditText = .TextMatrix(Row, 1)
        End If
    End If
End With
End Sub
Private Function GrdDesc(pItem As String, ByRef Row As Long, Optional ByRef sMsg As String) As Boolean
With grid1
If Trim(pItem) = "" Or Len(pItem) < 3 Or Len(pItem) > 30 Then
    sMsg = "ﬂÊœ €Ì— ’«·Õ"
    Exit Function
End If

Dim cm As New ADODB.command
Set cm = cmd("dbo.sp_item", con, adStoredProc, AddFlag(Empty, "BARCODE_FIND", pItem))
cm.Execute

If IsNull(cm.Parameters("@ITEM").Value) Then
    sMsg = "ﬂÊœ €Ì— ’«·Õ"
    Exit Function
End If

If cm.Parameters("@ISNOITEM").Value Then
    MsgBox "’‰› Œœ„«  ·« ÌÊÃœ ›Ï ⁄—÷"
End If

If cm.Parameters("@PRICE").Value = 0 Then
    MsgBox "”⁄— «·»Ì⁄ ’›— "
    Exit Function
End If


Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "Select doc_no,disc2 " & _
          " from file0_90 " & _
          " where file0_90.modelNo = " & MyParn(cm.Parameters("@MODELNO").Value) & _
          " and file0_90.doc_no = " & MyParn(xDoc_no.Caption)
Set loctable = cmd(cString, con).Execute
If loctable.EOF Then
    MsgBox "’‰› €Ì— „ÊÃÊœ ›Ï «·⁄—÷"
    Exit Function
End If

If fnBalance(cm.Parameters("@ITEM"), con, myForm.xStore.BoundText, myForm.xDate.text) - BalanceFound(cm.Parameters("@ITEM").Value, Row) <= 0 Then
    Inform "·« ÌÊÃœ —’Ìœ «·’‰›"
End If

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
grid1.TextMatrix(Row, 11) = cm.Parameters("@PRICE").Value
grid1.TextMatrix(Row, 12) = Val(loctable!disc2 & "")
grid1.TextMatrix(Row, 13) = Val(cm.Parameters("@COSTITEM").Value & "")
GrdDesc = True
End With
End Function
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'If grid1.Col = 12 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 12 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub fixGrd()
With grid1
'                 0          1          2         3        4          5              6          7          8            9           10          11          12              13
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "”⁄— „” Â·ﬂ|" & "⁄œœ|" & "«·≈Ã„«·Ï|" & "‰”»… «·Œ’„|" & "«· ﬂ·›…"
.RowHeight(0) = 600
.WordWrap = True

.ColWidth(0) = 600
.ColWidth(1) = 1000
.ColWidth(2) = 0
.ColWidth(3) = 1300
.ColWidth(4) = 0
.ColWidth(5) = 1500
.ColWidth(6) = 2800
.ColWidth(7) = 700
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 800
.ColWidth(11) = 1000
.ColWidth(12) = 1000
.ColHidden(2) = True
.ColHidden(3) = True
.ColHidden(4) = True
.ColHidden(12) = True
.ColHidden(13) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Private Sub Fixgrd2()
With grid2
'                 0          1          2         3        4          5              6          7          8            9           10          11          12          13
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "”⁄— „” Â·ﬂ|" & "⁄œœ|" & "«·≈Ã„«·Ï|" & "‰”»… «·Œ’„|" & "«·Œ’„"
.RowHeight(0) = 600
.WordWrap = True

.ColWidth(0) = 600
.ColWidth(1) = 1000
.ColWidth(2) = 0
.ColWidth(3) = 1300
.ColWidth(4) = 0
.ColWidth(5) = 1500
.ColWidth(6) = 2800
.ColWidth(7) = 700
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 800
.ColWidth(11) = 1000
.ColWidth(12) = 1000
.ColWidth(13) = 1400
.ColHidden(2) = True
.ColHidden(3) = True
.ColHidden(4) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Private Sub myAddItem()
grid1.AddItem ""
MakeSerial
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf col = 1 And grid1.TextMatrix(grid1.Row, 1) = "" Then
    grid1.Select Row, col
ElseIf col < 10 Then
    grid1.col = 10
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 10)
    grid1.ShowCell grid1.Row, 1
Else
    grid1.Select Row, col
End If
End Sub
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    If GrdDesc(oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0), grid1.Row) Then
        Grid1_AfterEdit grid1.Row, grid1.col
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Public Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
Dim bNew As Boolean
With grid1

If Not validRow(Row) Then
    CalcTotals
    Exit Sub
End If

If Row = grid1.Rows - 1 Then
    myAddItem
    grid1.Select grid1.Rows - 1, 1
End If

CalcTotals
End With
End Sub
Public Sub CalcTotals()
Dim i As Long
Dim nSales As Double
Dim bDone As Boolean

xRestOffer.Caption = Val(xQty1.Caption)
xRest.Caption = ""
xQtyOffer.Caption = ""

grid2.Rows = 1
xRate.Caption = ""
xDiscount.Caption = ""

cmddel.Enabled = grid1.Rows > 2

For i = 1 To grid1.Rows - 1
    nSales = nSales + grid1.ValueMatrix(i, 10)
    grid1.TextMatrix(i, 11) = Myvalue(grid1.ValueMatrix(i, 10) * grid1.ValueMatrix(i, 9))
    nTotal = nTotal + grid1.ValueMatrix(i, 11)
Next
xSales.Caption = Myvalue(nSales)
xTotalInv.Caption = Myvalue(nTotal)

cmdSave.Enabled = False


If nSales > Val(xTotal.Caption) Then
    MsgBox "«·ﬂ„Ì… «ﬂ»— „‰ ﬂ„Ì… «·⁄—÷"
    Exit Sub
End If

If nSales = 0 Then Exit Sub
    
'If Val(xSales.Caption) Mod Val(xTotal.Caption) = 0 Then
If Val(xSales.Caption) = Val(xTotal.Caption) Then
    xRestOffer.Caption = 0
    xRest.Caption = 0
    xQtyOffer.Caption = (nSales / Val(xTotal.Caption)) * Val(xQty2.Caption)
    bDone = True
Else
    xRestOffer.Caption = Val(xQty1.Caption) - (nSales Mod Val(xQty1.Caption))
    xRest.Caption = Val(xTotal.Caption) - Val(xSales.Caption) Mod Val(xTotal.Caption)
    xQtyOffer.Caption = ((Val(xRest.Caption) + Val(xSales.Caption)) / Val(xTotal.Caption)) * Val(xQty2.Caption)
End If

If Val(xRestOffer.Caption) = 0 And Val(xRest.Caption) = 0 Then
    Addoffer
End If

cmdSave.Enabled = bDone
If bDone Then
    MsgBox " „ «” Ì›«¡ «·⁄—÷"
    If IsFormOpen("search_abd") Then
        Unload oSearchItem
    End If
End If
End Sub
Private Sub Addoffer()
grid2.Redraw = flexRDNone
For Row = 1 To grid1.Rows - 2
    grid2.AddItem ""
    For col = 0 To grid1.Cols - 2
        grid2.TextMatrix(Row, col) = grid1.TextMatrix(Row, col)
    Next
Next
grid2.ColSort(9) = flexSortNumericAscending
grid2.Select 0, 9, 0, 9
grid2.Sort = flexSortUseColSort

Dim nquant As Long
Dim nQuantOffer As Long
Dim nDiscount As Double

nquant = Val(xQtyOffer.Caption)
For i = 1 To grid2.Rows - 1
    nQuantOffer = IIf(nquant > grid2.ValueMatrix(i, 10), grid2.ValueMatrix(i, 10), nquant)
    grid2.TextMatrix(i, 10) = nQuantOffer
    grid2.TextMatrix(i, 11) = grid2.ValueMatrix(i, 10) * grid2.ValueMatrix(i, 9)
    grid2.TextMatrix(i, 13) = mRound((grid2.ValueMatrix(i, 12) * grid2.TextMatrix(i, 11)) / 100, 2)
    nDiscount = nDiscount + grid2.ValueMatrix(i, 13)
    nquant = nquant - nQuantOffer
    If nquant = 0 Then Exit For
Next
xDiscount.Caption = mRound(nDiscount, 2)
If Val(xTotalInv.Caption) <> 0 Then
    xRate.Caption = mRound((Val(xDiscount.Caption) / Val(xTotalInv.Caption)) * 100, 2)
Else
    xRate.Caption = ""
End If
For i = grid2.Rows - 1 To i + 1 Step -1
    grid2.RemoveItem i
Next
grid2.Redraw = flexRDBuffered
End Sub
Private Function BalanceFound(pItem As String, pRow As Long) As Long
For i = 1 To grid1.Rows - 1
    If i <> pRow And grid1.TextMatrix(i, 1) = pItem Then
        BalanceFound = BalanceFound + grid1.ValueMatrix(i, 10)
    End If
Next
End Function

