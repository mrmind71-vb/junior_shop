VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form online_Stage_order 
   BackColor       =   &H00FFFFFF&
   Caption         =   " ÃÂÌ“ ÿ·»Ì…"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   14610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmPr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   45
      Width           =   5595
      Begin VB.Label xtype_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   540
         Width           =   2085
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·‰Ê⁄"
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
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   540
         Width           =   735
      End
      Begin VB.Label xdoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   180
         Width           =   960
      End
      Begin VB.Label lbodoc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·›—⁄Ì"
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
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   225
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   735
      ScaleWidth      =   14610
      TabIndex        =   6
      Top             =   8970
      Width           =   14610
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   -45
         Width           =   3705
         Begin Threed.SSCommand cmdExit 
            Height          =   555
            Left            =   45
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   135
            Width           =   1140
            _ExtentX        =   2011
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
            Picture         =   "online_stage_order.frx":0000
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdPrint 
            Height          =   555
            Left            =   1215
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   135
            Width           =   1185
            _ExtentX        =   2090
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
            Picture         =   "online_stage_order.frx":2323
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "online_stage_order.frx":4699
         End
         Begin Threed.SSCommand cmdExcel 
            Height          =   555
            Left            =   2430
            TabIndex        =   10
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
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
            Picture         =   "online_stage_order.frx":681C
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
      End
      Begin Threed.SSCommand cmdStage 
         Height          =   555
         Index           =   4
         Left            =   9765
         TabIndex        =   13
         Top             =   90
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   979
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
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
         Caption         =   "„‘ﬂ·… ›Ï «·ÿ·»Ì…"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   2
         BevelWidth      =   0
      End
      Begin Threed.SSCommand cmdStage 
         Height          =   555
         Index           =   6
         Left            =   7515
         TabIndex        =   18
         Top             =   90
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   979
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
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
         Caption         =   "«—”«· ·· ⁄œÌ·"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   2
         BevelWidth      =   0
      End
      Begin Threed.SSCommand cmdStage 
         Height          =   555
         Index           =   2
         Left            =   12015
         TabIndex        =   11
         Top             =   90
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   979
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
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
         Caption         =   "»œ« «· ÃÂÌ“"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   2
         BevelWidth      =   0
      End
      Begin Threed.SSCommand cmdStage 
         Height          =   555
         Index           =   5
         Left            =   9765
         TabIndex        =   17
         Top             =   90
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   979
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
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
         Caption         =   "«—”«· ··«·€«¡"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStage 
         Height          =   555
         Index           =   3
         Left            =   12015
         TabIndex        =   12
         Top             =   90
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   979
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
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
         Caption         =   " „ «· ÃÂÌ“"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   2
         BevelWidth      =   0
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   6975
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   7305
      Begin VB.CheckBox xPrinted 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " „  «·ÿ»«⁄…"
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
         Height          =   195
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   675
         Width           =   1320
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„—Õ·…"
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
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label xStage 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   2130
      End
      Begin VB.Label Label2 
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
         Height          =   330
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   540
         Width           =   825
      End
      Begin VB.Label xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   2130
      End
      Begin VB.Label xorder_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   2130
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -8550
      Top             =   5625
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4065
      Left            =   90
      TabIndex        =   0
      Top             =   1080
      Width           =   14235
      _cx             =   25109
      _cy             =   7170
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
      BackColorBkg    =   16777215
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
      Cols            =   10
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdError 
      Height          =   1950
      Left            =   90
      TabIndex        =   14
      Top             =   5175
      Width           =   14235
      _cx             =   25109
      _cy             =   3440
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
      BackColorBkg    =   16777215
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
      Cols            =   10
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdMsg 
      Height          =   1815
      Left            =   90
      TabIndex        =   24
      Top             =   7155
      Width           =   14235
      _cx             =   25109
      _cy             =   3201
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
      BackColorBkg    =   16777215
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
      Rows            =   2
      Cols            =   10
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "online_Stage_order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sOrder_No As String
Public sDoc_no As String
Public sManCode As String
Public bInvoice As Boolean
Public sMsg As String
Public myForm As Form
Dim aHeader(2) As String
Dim bError As Boolean
Dim aRow As Variant, aXl As Variant
Private Sub cmdExel_Click()
ToFileExelNew grid1, , , aXl, , 1
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdPrint_Click()
'Dim aRow(0) As Variant
'aRow(0) = AddFlag(Empty, "row", 1)
'aRow(0) = AddFlag(aRow(0), "col", 0)
'aRow(0) = AddFlag(aRow(0), "cols", IIf(retFlag(aData, "ITEM") = "", 5, 4))
'printGrdNew.sDeviceType = "1"
printGrdNew.bIgPreview = True
printGrdNew.doprint grid1, 0.85, -3, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 4), retHeader(aHeader, 6, 4), False, False, 10, , aRow

Set printGrdNew = Nothing

If xPrinted.Value = 0 Then
    If Val(xdoc_no.Caption) > 0 Then
        nAffect = rsEx("UPDATE FILE6_90BH " & _
             " SET PRINTED = 1" & _
             " WHERE DOC_NO = " & xdoc_no.Caption)
    Else
        nAffect = rsEx("UPDATE FILE6_90H " & _
             " SET PRINTED = 1" & _
             " WHERE DOC_NO = " & MyParn(xorder_no.Caption))
    End If
    xPrinted.Value = nAffect
End If
'printGrdNew.Show 1
End Sub
Private Sub cmdStage_Click(Index As Integer)
If xPrinted.Value = 0 And Index = 2 Then
    If MsgBox("·„   „ ÿ»«⁄… «·„” ‰œ", vbYesNo) <> vbNo Then
        cmdPrint_Click
    End If
End If
If updateStage(Index) Then
    myForm.myload
    Unload Me
End If
End Sub
Function updateStage(Index) As Boolean
If Index = 4 Then
    sMsg = ""
    Set online_error.myForm = Me
    online_error.sCaption = "”»» «·„‘ﬂ·…"
    online_error.Show 1
    If sMsg = "" Then
        Inform_OK "”»» «·„‘ﬂ·… ÷—Ê—Ì"
        Exit Function
    End If
ElseIf Index = 3 And xStage.Tag = "4" Then
    sMsg = ""
    Set online_error.myForm = Me
    online_error.sCaption = "»Ì«‰ Õ· «·„‘ﬂ·…"
    online_error.Show 1
    If sMsg = "" Then sMsg = " „ «·Õ·"
ElseIf Index = 5 Or Index = 6 Then
    sMsg = ""
    Set online_error.myForm = Me
    online_error.sCaption = IIf(Index = 5, "«”»«» «·«·€«¡", "«”»«» «· ⁄œÌ·")
    online_error.Show 1
    If sMsg = "" Then
        Inform_OK "”»» «·«·€«¡ ÷—Ê—Ì"
        Exit Function
    End If
End If

Dim con As New ADODB.Connection
If openCn(con) Then
    If xdoc_no.Caption = "0" Then
        cString = "UPDATE FILE6_90H " & _
                  " SET STAGE = " & (Index) & "," & _
                  " TIME" & Index & " = " & addstring(Format(Now, "YYYY-MM-DD HH:NN")) & "," & _
                  " MAN = " & addstring(sManCode) & "," & _
                  " MAN" & Index & " = " & addstring(sManCode) & _
                  " WHERE DOC_NO = " & MyParn(xorder_no.Caption)
    Else
        cString = "UPDATE FILE6_90BH " & _
                  " SET STAGE = " & (Index) & "," & _
                  " TIME" & Index & " = " & addstring(Format(Now, "YYYY-MM-DD HH:NN")) & "," & _
                  " MAN = " & addstring(sManCode) & "," & _
                  " MAN" & Index & " = " & addstring(sManCode) & _
                  " WHERE DOC_NO = " & xdoc_no.Caption
    End If
    'On Error GoTo myerror
    con.BeginTrans
    con.Execute cString, nRecords
    
    Dim aInsert As Variant
    If Index = 4 Or (Index = 3 And xStage.Tag = "4") Or Index = 5 Or Index = 6 Then
        aInsert = AddFlag(aInsert, "ORDER_NO", addstring(xorder_no.Caption))
        aInsert = AddFlag(aInsert, "DOC_NO_SUP", xdoc_no.Caption)
        aInsert = AddFlag(aInsert, "STAGE", (Index))
        aInsert = AddFlag(aInsert, "DESCA", addstring(sMsg))
        aInsert = AddFlag(aInsert, "TIME", addstring(Format(Now, "YYYY-MM-DD HH:NN")))
        aInsert = AddFlag(aInsert, "MAN", addstring(sManCode))
        con.Execute addInsert(aInsert, "FILE6_90_ER")
    End If
    
    For I = 2 To grdMsg.Rows - 1
        aInsert = AddFlag(Empty, "ORDER_NO", addstring(xorder_no.Caption))
        aInsert = AddFlag(aInsert, "DOC_NO_SUP", xdoc_no.Caption)
        aInsert = AddFlag(aInsert, "STAGE", (Index))
        aInsert = AddFlag(aInsert, "DESCA", addstring(sMsg))
        aInsert = AddFlag(aInsert, "TIME_SEEN", addstring(Format(Now, "YYYY-MM-DD HH:NN")))
        aInsert = AddFlag(aInsert, "MAN", addstring(sManCode))
        con.Execute addUpdate(aInsert, "FILE6_90_MSG", "ID = " & grdMsg.TextMatrix(I, grdMsg.Cols - 1))
    Next
    con.CommitTrans
    If Index = 2 Then
        Inform_OK " „  ÕÊÌ· «·ÿ·»Ì… ·· ÃÂÌ“"
    ElseIf Index = 3 Then
        Inform_OK " „ «·«‰ Â«¡ „‰ «·ÿ·»Ì…"
    ElseIf Index = 4 Then
        Inform_OK " „  ÕÊÌ· «·Ì ÿ·»Ì… »Â« „‘ﬂ·…"
    ElseIf Index = 5 Then
        Inform_OK " „  ÕÊÌ· ··√·€«¡"
    ElseIf Index = 6 Then
        Inform_OK " „  ÕÊÌ· ·· ⁄œÌ·"
    End If
End If
updateStage = True
Finally:
closeCon con
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
Resume Finally
End Function
Private Sub Form_Load()
Dim cString As String
cString = "SELECT v.ORDER_NO," & _
          " v.DOC_NO," & _
          " v.STAGE," & _
          " v.DATE," & _
          " v.type_desca," & _
          " v.printed," & _
          " STAGES_CODES.DESCA AS STAGE_DESCA," & _
          " FILE6_25.DESCA AS MAKE_DESCA" & _
          " FROM vw_online_order v " & _
          " INNER JOIN STAGES_CODES ON v.STAGE = STAGES_CODES.CODE" & _
          " LEFT JOIN FILE6_25 ON v.MAN = FILE6_25.CODE" & _
          " WHERE v.ORDER_NO = " & MyParn(sOrder_No) & _
          " AND v.doc_no = " & Val(sDoc_no)

On Error GoTo myerror
Dim loctable As New ADODB.Recordset
Set loctable = myRs(cString)

xorder_no.Caption = loctable!ORDER_NO & ""
xdoc_no.Caption = loctable!DOC_NO & ""
xStage.Caption = loctable!stage_Desca
xStage.Tag = loctable!Stage
xdate.Caption = myFormat_p(loctable!Date)
xPrinted.Value = IIf(loctable!printed, 1, 0)
xtype_desca.Caption = loctable!TYPE_dESCA & ""
fmPr.Visible = loctable!DOC_NO <> "0"
If loctable!Stage = 4 Then
    cmdStage(3).Caption = " „ Õ· «·„‘ﬂ·… Ê«· ÃÂÌ“"
End If
cmdStage(2).Visible = loctable!Stage = 1
cmdStage(3).Visible = loctable!Stage = 2 Or loctable!Stage = 4
cmdStage(4).Visible = loctable!Stage = 3 Or loctable!Stage = 2
cmdStage(5).Visible = loctable!Stage = 4
cmdStage(6).Visible = loctable!Stage = 4

myload
myloadGrdError
myloadGrdMsg
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub Fixgrd()
With grid1
        
     .RowHeight(0) = 400
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "„ÊœÌ·"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "«·’‰›"
    .TextMatrix(0, 4) = "«··Ê‰"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«·ﬂ„Ì…"
    .TextMatrix(0, 7) = "«·”⁄—"
    .TextMatrix(0, 8) = "«·≈Ã„«·Ì"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1500
    .ColWidth(2) = 2000
    .ColWidth(3) = 4500
    .ColWidth(4) = 2500
    .ColWidth(5) = 1000
    .ColWidth(6) = 800
    .ColWidth(7) = 1000
    .ColWidth(8) = 1200
    
    .ColHidden(.Cols - 3) = True
    .ColHidden(.Cols - 2) = True
    .ColHidden(.Cols - 1) = True
        
                
    .ExplorerBar = flexExSort
    
    .SubtotalPosition = flexSTBelow
    
    
    .Subtotal flexSTSum, -1, 6, "#", &HC0FFC0, , True, "≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, 8, "#", &HC0FFC0, , True, "≈Ã„«·Ì"
    
    
    If grid1.Rows > 1 Then
        aSub = AddFlag(Empty, "row", grid1.Rows - 1)
        aSub = AddFlag(aSub, "col", 0)
        aSub = AddFlag(aSub, "cols", 6)
        aSub = AddFlag(aSub, "text", "«·≈Ã„«·Ì")
        aRow = AddFlag(aRow, aSub)
    
        aSub = AddFlag(Empty, "row", grid1.Rows - 1)
        aSub = AddFlag(aSub, "col", 0)
        aSub = AddFlag(aSub, "cols", 6)
        aSub = AddFlag(aSub, "text", "«·≈Ã„«·Ì")
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "back_color", 19)
        aXl = AddFlag(aXl, aSub)
        
        
'        If grid1.ValueMatrix(1, 9) <> 0 Then
'            grid1.AddItem ""
'            grid1.AddItem ""
'
'            aSub = AddFlag(Empty, "row", grid1.Rows - 2)
'            aSub = AddFlag(aSub, "col", 0)
'            aSub = AddFlag(aSub, "cols", 8)
'            aRow = AddFlag(aRow, aSub)
'
'            aSub = AddFlag(Empty, "row", grid1.Rows - 2)
'            aSub = AddFlag(aSub, "col", 0)
'            aSub = AddFlag(aSub, "cols", 8)
'            aSub = AddFlag(aSub, "bold", True)
'            aSub = AddFlag(aSub, "back_color", 19)
'            aXl = AddFlag(aXl, aSub)
'
'            aSub = AddFlag(Empty, "row", grid1.Rows - 1)
'            aSub = AddFlag(aSub, "col", 0)
'            aSub = AddFlag(aSub, "cols", 8)
'            aRow = AddFlag(aRow, aSub)
'
'            aSub = AddFlag(Empty, "row", grid1.Rows - 1)
'            aSub = AddFlag(aSub, "col", 0)
'            aSub = AddFlag(aSub, "cols", 8)
'            aSub = AddFlag(aSub, "bold", True)
'            aSub = AddFlag(aSub, "back_color", 19)
'            aXl = AddFlag(aXl, aSub)
'
'
'            For i = 0 To 7
'                If i <= 5 Then grid1.TextMatrix(grid1.Rows - 3, i) = "≈Ã„«·Ì"
'                grid1.TextMatrix(grid1.Rows - 2, i) = "«·Œ’„"
'                grid1.TextMatrix(grid1.Rows - 1, i) = "«·’«›Ì"
'            Next
'
'            grid1.TextMatrix(grid1.Rows - 2, 8) = grid1.ValueMatrix(1, 9)
'            grid1.TextMatrix(grid1.Rows - 1, 8) = Round(grid1.ValueMatrix(grid1.Rows - 3, 8) - grid1.ValueMatrix(grid1.Rows - 2, 8), 2)
'
'            grid1.MergeCells = flexMergeRestrictRows
'
'            grid1.MergeRow(grid1.Rows - 3) = True
'            grid1.MergeRow(grid1.Rows - 2) = True
'            grid1.MergeRow(grid1.Rows - 1) = True
'
'            .Cell(flexcpBackColor, .Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = &HC0FFC0
'            .Cell(flexcpBackColor, .Rows - 2, 0, grid1.Rows - 1, grid1.Cols - 1) = &HC0FFC0
'
'
'        End If
    End If
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .Select .Rows - 1, 0
    .ShowCell .Rows - 1, 0
End With
End Sub
Private Sub myload()
Dim cString As String
cString = " SELECT v.item," & _
          " F1_10.MODELFACT0," & _
          " FACT.DESCA," & _
          " F1_10.desca," & _
          " F1_10.COLOR," & _
          " F1_10.SCAL," & _
          " v.QUANT," & _
          " v.PRICE," & _
          " v.TOTAL," & _
          " v.DISCOUNT " & _
          " FROM  vw_online_items v" & _
          " INNER  JOIN FILE1_10 F1_10 ON v.ITEM = F1_10.ITEM" & _
          " INNER  JOIN  FACT ON F1_10.code = FACT.CODE" & _
          " WHERE v.ORDER_NO =  " & MyParn(sOrder_No) & _
          " AND v.DOC_NO = " & sDoc_no

Dim con As New ADODB.Connection
On Error GoTo myerror
Set grid1.DataSource = myRs(cString)
Fixgrd
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set ShowOrderManfrm = Nothing
End Sub
Private Sub myLoadGrd()
If Not openCn(con) Then On Error GoTo Finally

Dim cString As String
cString = " SELECT v.DOC_NO," & _
          " F1_10.MODELFACT012," & _
          " FACT.DESCA," & _
          " F1_10.desca," & _
          " F1_10.COLOR," & _
          " F1_10.SCAL," & _
          " v.QUANT," & _
          " v.PRICE," & _
          " v.TOTAL" & _
          " FROM  vw_online_item v" & _
          " INNER  JOIN FILE1_10 F1_10 ON v.ITEM = F1_10.ITEM" & _
          " INNER  JOIN  FACT ON F1_10.code = FACT.CODE" & _
          " WHERE v.ORDER_NO =  " & MyParn(xorder_no.Caption) & _
          " AND   v.DOC_NO =  " & xdoc_no.Caption & _
          " ORDER BY F1_10.ITEM"

Dim con As New ADODB.Connection
If openCn(con) Then
    On Error GoTo Finally
    Set grid1.DataSource = myRs(cString, con)
End If '
Fixgrd
Finally:
closeCon con
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myloadGrdError()
Dim strSql As String
strSql = "SELECT STAGES_CODES.DESCA AS [«·„—Õ·…], " & _
         " FILE6_25.DESCA AS [«·„‰œÊ»], " & _
         " e.DESCA AS [«·»Ì«‰], " & _
         " FORMAT(e.TIME, 'yyyy/M/d HH:mm') AS [«·Êﬁ ]" & _
         " FROM STAGES_CODES " & _
         "    INNER JOIN FILE6_90_ER e " & _
         "        ON STAGES_CODES.CODE = e.STAGE " & _
         "    INNER JOIN FILE6_25 " & _
         "        ON e.MAN = FILE6_25.CODE" & _
         " WHERE e.ORDER_NO = " & MyParn(xorder_no.Caption) & _
         " AND e.DOC_NO_SUP = " & xdoc_no.Caption

On Error GoTo myerror

With grdError
Set .DataSource = myRs(strSql)

.ColWidth(0) = 1300
.ColWidth(1) = 1500
.ColWidth(2) = 5000
.ColWidth(3) = 1600

For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next

End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub myloadGrdMsg()
Dim strSql As String
strSql = "SELECT FORMAT(m.TIME, 'yyyy/M/d HH:mm') AS [«·Êﬁ ]," & _
         " m.USERNAME AS [«·„—”·], " & _
         " m.DESCA AS [«·„·ÕÊŸ…], " & _
         " FORMAT(m.TIME_SEEN, 'yyyy/M/d HH:mm') AS [Êﬁ  «·ﬁ—«¡…], " & _
         " F25.DESCA AS [«·„‰œÊ»], " & _
         " s.DESCA AS [«·„—Õ·…], " & _
         " m.ID " & _
         " FROM FILE6_90_MSG M" & _
         " LEFT JOIN FILE6_25 F25 ON m.MAN = F25.CODE" & _
         " LEFT JOIN STAGES_CODES S ON m.STAGE = s.CODE" & _
         " WHERE m.ORDER_NO = " & MyParn(xorder_no.Caption) & _
         " AND m.DOC_NO_SUP = " & xdoc_no.Caption & _
         " ORDER BY TIME DESC"
         

On Error GoTo myerror
With grdMsg
Set .DataSource = myRs(strSql)
FixGrdMsg
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub FixGrdMsg()
With grdMsg
Dim I As Long
For I = 0 To .Cols - 1
    .TextMatrix(0, I) = "„·«ÕŸ«  «·›—⁄ «·—∆Ì”Ì"
    .ColAlignment(I) = flexAlignCenterCenter
Next
.RowHidden(0) = False
.ColWidth(0) = 1700
.ColWidth(1) = 2200
.ColWidth(2) = 10000
.ColWidth(3) = 1700
.ColWidth(4) = 2000
.ColWidth(5) = 1500

.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True

.ColHidden(.Cols - 1) = True

.TextMatrix(1, 0) = "«·Êﬁ "
.TextMatrix(1, 1) = "«·„—”·"
.TextMatrix(1, 2) = "«·„·ÕÊŸ…"
.TextMatrix(1, 3) = "Êﬁ  «·ﬁ—«¡…"
.TextMatrix(1, 4) = "«·„‰œÊ»"
.TextMatrix(1, 5) = "«·„—Õ·…"

.ColHidden(.Cols - 4) = True
.ColHidden(.Cols - 3) = True
.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True
     
.Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter

For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
