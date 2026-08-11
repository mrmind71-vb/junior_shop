VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form dashBord_Online 
   BackColor       =   &H00E0E0E0&
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
   Begin VB.Frame Frame7 
      Caption         =   "«·√Ì«„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7800
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3195
      Width           =   5955
      Begin VSFlex7Ctl.VSFlexGrid GRID5 
         Height          =   7440
         Left            =   90
         TabIndex        =   18
         Top             =   270
         Width           =   5775
         _cx             =   10186
         _cy             =   13123
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
         BackColorSel    =   16761024
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   6
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
   End
   Begin VB.Frame Frame6 
      Caption         =   "—’Ìœ ÿ·»»«  €Ì— „‰›–… ··›—Ê⁄"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Left            =   6075
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   6345
      Width           =   3885
      Begin VSFlex7Ctl.VSFlexGrid GRID4 
         Height          =   4065
         Left            =   90
         TabIndex        =   16
         Top             =   270
         Width           =   3705
         _cx             =   6535
         _cy             =   7170
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
         BackColorSel    =   16761024
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
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
   End
   Begin VB.Frame Frame5 
      Caption         =   "«ﬁ”«„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3150
      Width           =   3930
      Begin VSFlex7Ctl.VSFlexGrid GRID3 
         Height          =   2850
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   3750
         _cx             =   6615
         _cy             =   5027
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
         BackColorSel    =   16761024
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
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
   End
   Begin VB.Frame Frame4 
      Caption         =   "„ÊœÌ·« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7755
      Left            =   9990
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3195
      Width           =   6675
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Height          =   7395
         Left            =   45
         TabIndex        =   11
         Top             =   270
         Width           =   6585
         _cx             =   11615
         _cy             =   13044
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
         Rows            =   50
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "„Õ«›Ÿ« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7755
      Left            =   16785
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3150
      Width           =   6090
      Begin VSFlex7Ctl.VSFlexGrid GRID1 
         Bindings        =   "DashBord_Online.frx":0000
         Height          =   7395
         Left            =   90
         TabIndex        =   9
         Top             =   225
         Width           =   5595
         _cx             =   9869
         _cy             =   13044
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
         BackColorSel    =   16761024
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "«·› —…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   14535
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   90
      Width           =   8340
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   90
         Picture         =   "DashBord_Online.frx":0014
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1050
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   2475
         Picture         =   "DashBord_Online.frx":2480
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   1275
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   1230
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   555
         Left            =   1170
         TabIndex        =   14
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   979
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "DashBord_Online.frx":4972
         Alignment       =   4
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6975
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2220
      Left            =   2835
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   945
      Width           =   19995
      Begin VSFlex7Ctl.VSFlexGrid GridTotal 
         Bindings        =   "DashBord_Online.frx":6D86
         Height          =   1905
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   19770
         _cx             =   34872
         _cy             =   3360
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
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
         TabBehavior     =   0
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   180
      Top             =   135
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   90
      Top             =   585
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
   Begin MSAdodcLib.Adodc DATA5 
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
      TabIndex        =   19
      Top             =   10845
      Visible         =   0   'False
      Width           =   22920
      _ExtentX        =   40428
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "DashBord_Online"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conShop As New ADODB.Connection
Dim con As New ADODB.Connection
Dim oSearch_Total As New Search_total
Private Sub cmd_excel_Click()
    ToFileExel2 GridTotal, , , , , 1.1, , , , , , Me
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    If Not myValid Then Exit Sub
    MyLOadTotal
    MYLOAD1
    MYLOAD2
    myload3
    MYLOAD4
    myload5
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
Private Sub Form_Load()
    OpenCon con
        
    xDate1.text = myFormat_p(Format(DateValue("1-" & Month(Date) & "-" & Year(Date)), "DD-MM-YYYY"))
    xdate2.text = myFormat_p(DateAdd("D", -1, DateAdd("M", 1, DateValue(xDate1.text))))
    
    Set GridTotal.DataSource = data1
    Set GRID1.DataSource = DATA2
    Set GRID2.DataSource = DATA3
    Set GRID3.DataSource = DATA4
    Set GRID4.DataSource = DATA5
    Set GRID5.DataSource = data6
    
    
    GridTotal.Rows = 2
    GRID1.Rows = 2
    GRID2.Rows = 2
    GRID3.Rows = 2
    GRID4.Rows = 2
    GRID5.Rows = 2
    
    fixGrdTotal
    FixGrid1
    FixGrid2
    FixGrid3
    FIXGRID4
    FIXGRID5
    
End Sub
Private Sub fixGrdTotal()
With GridTotal
.FixedCols = 0
.FixedRows = 2
.Cols = 21
.RowHeight(0) = 500
.RowHeight(1) = 500

.WordWrap = True
.TextMatrix(0, 0) = "ÿ·»Ì« "
.TextMatrix(0, 1) = "ÿ·»Ì« "
.TextMatrix(0, 2) = "ÿ·»Ì« "
.TextMatrix(0, 3) = "ÿ·»Ì« "

.TextMatrix(1, 0) = "⁄œœ"
.TextMatrix(1, 1) = "ﬁÿ€"
.TextMatrix(1, 2) = "≈Ã„«·Ï"
.TextMatrix(1, 3) = "„ Ê”ÿ"

.TextMatrix(0, 4) = "«·„»«⁄"
.TextMatrix(0, 5) = "«·„»«⁄"
.TextMatrix(0, 6) = "«·„»«⁄"
.TextMatrix(0, 7) = "«·„»«⁄"
.TextMatrix(0, 8) = "«·„»«⁄"

.TextMatrix(1, 4) = "⁄œœ"
.TextMatrix(1, 5) = "ﬁÿ⁄"
.TextMatrix(1, 6) = "ﬁÌ„…"
.TextMatrix(1, 7) = "‰”Ì…"
.TextMatrix(1, 8) = "„ Ê”ÿ"

.TextMatrix(0, 9) = "„·€Ï"
.TextMatrix(0, 10) = "„·€Ï"
.TextMatrix(0, 11) = "„·€Ï"
.TextMatrix(0, 12) = "„·€Ï"
.TextMatrix(0, 13) = "„·€Ï"

.TextMatrix(1, 9) = "⁄œœ"
.TextMatrix(1, 10) = "ﬁÿ⁄"
.TextMatrix(1, 11) = "ﬁÌ„…"
.TextMatrix(1, 12) = "‰”Ì…"
.TextMatrix(1, 13) = "„ Ê”ÿ"


.TextMatrix(0, 14) = "„  »ﬁÏ"
.TextMatrix(0, 15) = "„  »ﬁÏ"
.TextMatrix(0, 16) = "„  »ﬁÏ"
.TextMatrix(0, 17) = "„  »ﬁÏ"
.TextMatrix(0, 18) = "„  »ﬁÏ"

.TextMatrix(1, 14) = "⁄œœ"
.TextMatrix(1, 15) = "ﬁÿ⁄"
.TextMatrix(1, 16) = "ﬁÌ„…"
.TextMatrix(1, 17) = "‰”Ì…"
.TextMatrix(1, 18) = "„ Ê”ÿ"

.TextMatrix(0, 19) = "„— Ã⁄"
.TextMatrix(1, 19) = "⁄œœ"

.TextMatrix(0, 20) = "„— Ã⁄"
.TextMatrix(1, 20) = "‰”»…"

.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True

.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
.Cell(flexcpAlignment, 1, 0, 1, .Cols - 1) = 7

.ColWidth(0) = 800
.ColWidth(1) = 1000
.ColWidth(2) = 1100
.ColWidth(3) = 800

.ColWidth(4) = 800
.ColWidth(5) = 1100
.ColWidth(6) = 1200
.ColWidth(7) = 800
.ColWidth(8) = 800

.ColWidth(9) = 800
.ColWidth(10) = 1100
.ColWidth(11) = 1200
.ColWidth(12) = 800
.ColWidth(13) = 800

.ColWidth(14) = 800
.ColWidth(15) = 1100
.ColWidth(16) = 1200
.ColWidth(17) = 800
.ColWidth(18) = 800

.ColWidth(19) = 800
.ColWidth(20) = 800
.WordWrap = True
For nRow = 2 To .Rows - 1
    If Val(.TextMatrix(nRow, 1)) > 0 Then
        nPrice = Round(Val(.TextMatrix(nRow, 2)) / Val(.TextMatrix(nRow, 0)), 2)
        .TextMatrix(nRow, 3) = nPrice
    End If
        
    If Val(.TextMatrix(nRow, 5)) > 0 Then
        nPrice = Round(Val(.TextMatrix(nRow, 6)) / Val(.TextMatrix(nRow, 4)), 2)
        .TextMatrix(nRow, 8) = nPrice
    End If
    
    If Val(.TextMatrix(nRow, 9)) > 0 Then
        nPrice = Round(Val(.TextMatrix(nRow, 11)) / Val(.TextMatrix(nRow, 9)), 2)
        .TextMatrix(nRow, 13) = nPrice
    End If
    
        
    If Val(.TextMatrix(nRow, 15)) > 0 Then
        nPrice = Round(Val(.TextMatrix(nRow, 16)) / Val(.TextMatrix(nRow, 14)), 2)
        .TextMatrix(nRow, 18) = nPrice
    End If

    .TextMatrix(nRow, 2) = Round(Val(.TextMatrix(nRow, 2)), 0)
    .TextMatrix(nRow, 6) = Round(Val(.TextMatrix(nRow, 6)), 0)
    .TextMatrix(nRow, 11) = Round(Val(.TextMatrix(nRow, 11)), 0)
    .TextMatrix(nRow, 16) = Round(Val(.TextMatrix(nRow, 16)), 0)
    
    If Val(.TextMatrix(nRow, 0)) > 0 Then
        nRate = Round((Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 0))) * 100, 2)
        .TextMatrix(nRow, 7) = nRate
    
        nRate = Round((Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 0))) * 100, 2)
        .TextMatrix(nRow, 12) = nRate
        
        nRate = Round((Val(.TextMatrix(nRow, 14)) / Val(.TextMatrix(nRow, 0))) * 100, 2)
        .TextMatrix(nRow, 17) = nRate
    End If
    If Val(.TextMatrix(nRow, 4)) > 0 Then
        nRate = Round((Val(.TextMatrix(nRow, 19)) / Val(.TextMatrix(nRow, 4))) * 100, 2)
        .TextMatrix(nRow, 20) = nRate
    End If
Next nRow
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTBelow
End With
End Sub
Sub MyLOadTotal()
Dim cString As String
Dim cWhere As String
If IsDate(xDate1.text) Then cWhere = " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xdate2.text) Then cWhere = cWhere & " AND DATE <= " & DateSq(xdate2.text)
cString = " SELECT  COUNT(DISTINCT DOC_NO) AS COUNT1 , SUM(QUANT) AS QTY1 , SUM(TOTALITEM) AS TOTAL1 " & _
            " , 0 " & _
            " ,( SELECT COUNT(DISTINCT DOC_NO ) FROM QFILE6_90 AS QFILE6_90_2 WHERE SALES_DOC IS NOT NULL " & cWhere & ") AS COUNT2 " & _
            " ,( SELECT SUM(QUANT )  FROM QFILE6_90 AS QFILE6_90_2 WHERE SALES_DOC IS NOT NULL " & cWhere & ") AS QTY2 " & _
            " ,( SELECT SUM(TOTALITEM )  FROM QFILE6_90 AS QFILE6_90_2 WHERE SALES_DOC IS NOT NULL " & cWhere & ") AS TOTAL2 " & _
            " , 0 " & _
            " , 0 " & _
            " ,( SELECT COUNT(DISTINCT DOC_NO) FROM QFILE6_90 AS QFILE6_90_2 WHERE DelOrder_Date IS NOT NULL " & cWhere & " ) AS COUNT3 " & _
            " ,( SELECT SUM(QUANT )  FROM QFILE6_90 AS QFILE6_90_2 WHERE DelOrder_Date IS NOT NULL " & cWhere & ") AS QTY3 " & _
            " ,( SELECT SUM(TOTALITEM )  FROM QFILE6_90 AS QFILE6_90_2 WHERE DelOrder_Date IS NOT NULL " & cWhere & ") AS TOTAL3 " & _
            " , 0 " & _
            " , 0 " & _
            " ,( SELECT COUNT(DISTINCT DOC_NO) FROM QFILE6_90 AS QFILE6_90_2 WHERE DelOrder_Date IS NULL AND SALES_DOC IS NULL " & cWhere & ") AS COUNT4 " & _
            " ,( SELECT SUM(QUANT )  FROM QFILE6_90 AS QFILE6_90_2 WHERE DelOrder_Date IS NULL AND SALES_DOC IS NULL " & cWhere & ") AS QTY4 " & _
            " ,( SELECT SUM(TOTALITEM )  FROM QFILE6_90 AS QFILE6_90_2 WHERE DelOrder_Date IS NULL AND SALES_DOC IS NULL " & cWhere & ") AS TOTAL4 " & _
            " , 0 " & _
            " , 0 " & _
            " ,(SELECT COUNT(*) FROM QFILE6_20H WHERE SALES_RET IN (SELECT SALES_DOC FROM QFILE6_90 WHERE SALES_DOC IS NOT NULL " & cWhere & " ) )" & _
            " , 0 " & _
            " FROM QFILE6_90 WHERE DOC_NO IS NOT NULL " & cWhere

Set data1.RecordSet = myRecordSet(cString, con)
fixGrdTotal
End Sub
Sub MYLOAD1()
Dim cString As String
Dim cWhere As String
If IsDate(xDate1.text) Then cWhere = " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xdate2.text) Then cWhere = cWhere & " AND DATE <= " & DateSq(xdate2.text)
cString = " SELECT  Shipping_City ,   COUNT(DOC_NO) , SUM(QUANT) , SUM(TOTALITEM) FROM            QFILE6_90 " & _
            " WHERE DOC_NO IS NOT NULL " & cWhere & _
            " GROUP BY Shipping_City ORDER BY COUNT(DOC_NO) DESC"
Set DATA2.RecordSet = myRecordSet(cString, con)
FixGrid1
End Sub
Private Sub FixGrid1()
With GRID1
.Cols = 5
.RowHeight(0) = 500

.WordWrap = True
.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True
.TextMatrix(0, 0) = "«·„Õ«›Ÿ…"
.TextMatrix(0, 1) = "⁄œœ "
.TextMatrix(0, 2) = "ﬁÿ€"
.TextMatrix(0, 3) = "≈Ã„«·Ï ﬁÌ„…"
.TextMatrix(0, 4) = "‰”Ì…"
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
If .Rows > 1 Then .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
.ColWidth(0) = 1800
.ColWidth(1) = 800
.ColWidth(2) = 900
.ColWidth(3) = 1200
.ColWidth(4) = 800
.SubtotalPosition = flexSTBelow
.Subtotal flexSTCount, -1, 0, "#0", vbYellow, , , " "
.Subtotal flexSTSum, -1, 1, "#0", vbYellow, , , " "
.Subtotal flexSTSum, -1, 2, "#0", vbYellow, , , " "
.Subtotal flexSTSum, -1, 3, "#0.00", vbYellow, , , " "

For nRow = 1 To .Rows - 2
    If Val(.TextMatrix(.Rows - 1, 1)) <> 0 Then
       nRate = Round(Val(.TextMatrix(nRow, 1)) / Val(.TextMatrix(.Rows - 1, 1)) * 100, 2)
        .TextMatrix(nRow, 4) = nRate
    End If
Next nRow
End With
End Sub
Sub MYLOAD2()
Dim cString As String
Dim cWhere As String
If IsDate(xDate1.text) Then cWhere = " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xdate2.text) Then cWhere = cWhere & " AND DATE <= " & DateSq(xdate2.text)
cString = " SELECT        FACT.DESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca, SUM(FILE6_90.QUANT) FROM            FACT INNER JOIN FILE1_10 ON FACT.CODE = FILE1_10.code INNER JOIN FILE6_90 ON FILE1_10.ITEM = FILE6_90.ITEM INNER JOIN FILE6_90H ON FILE6_90.DOC_NO = FILE6_90H.DOC_NO WHERE MODELFACT0 IS NOT NULL " & cWhere & _
            " GROUP BY FACT.DESCA, FILE1_10.MODELFACT0, FILE1_10.desca, FILE1_10.MOSM  ORDER BY SUM(FILE6_90.QUANT ) DESC"
Set DATA3.RecordSet = myRecordSet(cString, con)
FixGrid2
End Sub
Private Sub FixGrid2()
With GRID2
.Cols = 6
.RowHeight(0) = 500

.WordWrap = True
.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True
.TextMatrix(0, 0) = "„’‰⁄"
.TextMatrix(0, 1) = "„Ê”„"
.TextMatrix(0, 2) = "—ﬁ„"
.TextMatrix(0, 3) = "„ÊœÌ·"
.TextMatrix(0, 4) = "ﬁÿ€"
.TextMatrix(0, 5) = "‰”Ì…"
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
If .Rows > 1 Then .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7

.ColWidth(0) = 1300
.ColWidth(1) = 700
.ColWidth(2) = 1000
.ColWidth(3) = 1500
.ColWidth(4) = 800
.ColWidth(5) = 800
.SubtotalPosition = flexSTBelow
.Subtotal flexSTCount, -1, 2, "#0", vbYellow, , , " "
.Subtotal flexSTSum, -1, 4, "#0", vbYellow, , , " "
If Val(.TextMatrix(.Rows - 1, 4)) <> 0 Then
    For nRow = 1 To .Rows - 2
        nRate = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(.Rows - 1, 4)) * 100, 2)
        .TextMatrix(nRow, 5) = nRate
    Next nRow
End If
End With
End Sub

Sub myload3()
Dim cString As String
Dim cWhere As String
If IsDate(xDate1.text) Then cWhere = " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xdate2.text) Then cWhere = cWhere & " AND DATE <= " & DateSq(xdate2.text)
cString = " SELECT  FILE1_10SC.DESCA, SUM(FILE6_90.QUANT) FROM  FILE1_10SC INNER JOIN FILE1_10 ON FILE1_10SC.[CODE] = FILE1_10.[SECTION] INNER JOIN FILE6_90 ON FILE1_10.ITEM = FILE6_90.ITEM INNER JOIN FILE6_90H ON FILE6_90.DOC_NO = FILE6_90H.DOC_NO WHERE MODELFACT0 IS NOT NULL " & cWhere & _
            " GROUP BY FILE1_10SC.DESCA  ORDER BY SUM(FILE6_90.QUANT) DESC"
Set DATA4.RecordSet = myRecordSet(cString, con)
FixGrid3
End Sub
Private Sub FixGrid3()
With GRID3
.Cols = 3
.RowHeight(0) = 500

.WordWrap = True
.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True
.TextMatrix(0, 0) = "«·ﬁ”„"
.TextMatrix(0, 1) = "ﬁÿ€"
.TextMatrix(0, 2) = "‰”Ì…"
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
If .Rows > 1 Then .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7

.ColWidth(0) = 1500
.ColWidth(1) = 800
.ColWidth(2) = 800
.SubtotalPosition = flexSTBelow
.Subtotal flexSTCount, -1, 0, "#0", vbYellow, , , " "
.Subtotal flexSTSum, -1, 1, "#0", vbYellow, , , " "

If Val(.TextMatrix(.Rows - 1, 1)) <> 0 Then
    For nRow = 1 To .Rows - 2
        nRate = Round(Val(.TextMatrix(nRow, 1)) / Val(.TextMatrix(.Rows - 1, 1)) * 100, 2)
        .TextMatrix(nRow, 2) = nRate
    Next nRow
End If
End With
End Sub
Sub MYLOAD4()
Dim cString As String
    cString = " select file0_40.desca , COUNT(DOC_NO) , sum(TOTALITEM) " & _
            " FROM QFILE6_90 LEFT JOIN FILE0_40 ON FILE0_40.CODE = QFILE6_90.STORE " & _
            " Where DelOrder_Date Is Null And SALES_DOC Is Null " & _
            " GROUP BY file0_40.desca , file0_40.CODE  ORDER BY file0_40.CODE"
    Set DATA5.RecordSet = myRecordSet(cString, con)
    FIXGRID4
End Sub
Private Sub FIXGRID4()
With GRID4
.Cols = 3
.RowHeight(0) = 500

.WordWrap = True
.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True
.TextMatrix(0, 0) = "«·›—⁄"
.TextMatrix(0, 1) = "⁄œœ"
.TextMatrix(0, 2) = "ﬁÌ„…"
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
If .Rows > 1 Then .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7

.ColWidth(0) = 1500
.ColWidth(1) = 800
.ColWidth(2) = 1000
.SubtotalPosition = flexSTBelow
.Subtotal flexSTSum, -1, 1, "#0", vbYellow, , , " "
.Subtotal flexSTSum, -1, 2, "#0", vbYellow, , , " "
End With
End Sub
Sub myload5_OLD2()


' select Q_DATE_ORDER.date
' , (SELECT COUNT(FILE6_90H_2.DOC_NO) FROM FILE6_90H AS FILE6_90H_2 WHERE (FILE6_90H_2.SALES_DATE >= Q_DATE_ORDER.DATE OR FILE6_90H_2.SALES_DATE IS NULL ) AND (FILE6_90H_2.DELORDER_DATE IS NULL OR CONVERT(VARCHAR(10), FILE6_90H_2.DELORDER_DATE, 111) >= Q_DATE_ORDER.DATE)  AND FILE6_90H_2.DATE < Q_DATE_ORDER.DATE  )   AS OLDORDER
' , (SELECT COUNT(FILE6_90H_2.DOC_NO) FROM FILE6_90H AS FILE6_90H_2 WHERE  FILE6_90H_2.DATE = Q_DATE_ORDER.DATE  )   AS newRDER
' , (SELECT COUNT(FILE6_90H_2.DOC_NO) FROM FILE6_90H AS FILE6_90H_2 WHERE FILE6_90H_2.SALES_DATE = Q_DATE_ORDER.DATE )  AS SALES
' , (SELECT COUNT(FILE6_90H_2.DOC_NO) FROM FILE6_90H AS FILE6_90H_2 WHERE CONVERT(VARCHAR(10), FILE6_90H_2.DELORDER_DATE, 111) = Q_DATE_ORDER.DATE )  AS delorder
' , (SELECT COUNT(FILE6_90H_2.DOC_NO) FROM FILE6_90H AS FILE6_90H_2 WHERE (FILE6_90H_2.SALES_DATE >  Q_DATE_ORDER.DATE OR FILE6_90H_2.SALES_DATE IS NULL ) AND (FILE6_90H_2.DELORDER_DATE IS NULL OR CONVERT(VARCHAR(10), FILE6_90H_2.DELORDER_DATE, 111)  >  Q_DATE_ORDER.DATE)  AND FILE6_90H_2.DATE <= Q_DATE_ORDER.DATE )
' From Q_DATE_ORDER
' where DATE >= '2023-06-20' AND DATE <= '2023-06-30'
' GROUP BY Q_DATE_ORDER.date
' ORDER BY Q_DATE_ORDER.date



Dim cWhere As String, cF1 As String, cF2 As String
If IsDate(xDate1.text) Then cWhere = " where DATE >= " & DateSq(xDate1.text)
If IsDate(xdate2.text) Then cWhere = cWhere & " AND DATE <= " & DateSq(xdate2.text)

cF1 = " , (SELECT COUNT(FILE6_90H_2.DOC_NO) FROM FILE6_90H AS FILE6_90H_2 WHERE (FILE6_90H_2.SALES_DATE >= Q_DATE_ORDER.DATE OR FILE6_90H_2.SALES_DATE IS NULL ) AND (FILE6_90H_2.DELORDER_DATE IS NULL OR CONVERT(VARCHAR(10), FILE6_90H_2.DELORDER_DATE, 111) >= Q_DATE_ORDER.DATE)  AND FILE6_90H_2.DATE < Q_DATE_ORDER.DATE  )   AS OLDORDER "
cF2 = " , (SELECT COUNT(FILE6_90H_2.DOC_NO) FROM FILE6_90H AS FILE6_90H_2 WHERE  FILE6_90H_2.DATE = Q_DATE_ORDER.DATE  )   AS newRDER"
cF3 = " , (SELECT COUNT(FILE6_90H_2.DOC_NO) FROM FILE6_90H AS FILE6_90H_2 WHERE FILE6_90H_2.SALES_DATE = Q_DATE_ORDER.DATE )  AS SALES"
cF4 = " , (SELECT COUNT(FILE6_90H_2.DOC_NO) FROM FILE6_90H AS FILE6_90H_2 WHERE CONVERT(VARCHAR(10), FILE6_90H_2.DELORDER_DATE, 111) = Q_DATE_ORDER.DATE )  AS delorder"
cF5 = " , (SELECT COUNT(FILE6_90H_2.DOC_NO) FROM FILE6_90H AS FILE6_90H_2 WHERE (FILE6_90H_2.SALES_DATE >  Q_DATE_ORDER.DATE OR FILE6_90H_2.SALES_DATE IS NULL ) AND (FILE6_90H_2.DELORDER_DATE IS NULL OR CONVERT(VARCHAR(10), FILE6_90H_2.DELORDER_DATE, 111)  >  Q_DATE_ORDER.DATE)  AND FILE6_90H_2.DATE <= Q_DATE_ORDER.DATE )"


 
Dim cString As String
    cString = " select Q_DATE_ORDER.date  " & cF1 & cF2 & cF3 & cF4 & cF5 & _
            "   FROM Q_DATE_ORDER " & _
            "   where DATE >= " & DateSq(xDate1.text) & _
            "   AND DATE <= " & DateSq(xdate2.text) & _
            "   GROUP BY Q_DATE_ORDER.date " & _
            "   ORDER BY Q_DATE_ORDER.date "
    Set data6.RecordSet = myRecordSet(cString, con)
    FIXGRID5
End Sub
Private Sub FIXGRID5_old()
    With GRID5
    .Cols = 6
    .RowHeight(0) = 500
    .WordWrap = True
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    .TextMatrix(0, 0) = "«·ÌÊ„"
    .TextMatrix(0, 1) = "”«»ﬁ"
    .TextMatrix(0, 2) = "⁄œœ ÿ·»»« "
    .TextMatrix(0, 3) = "⁄œœ „‰ﬁ–"
    .TextMatrix(0, 4) = "⁄œœ „·€Ï"
    .TextMatrix(0, 5) = "„ »ﬁÏ"
    .RowHeight(0) = 700
    .WordWrap = True
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
    If .Rows > 1 Then .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    .ColWidth(0) = 1100
    .ColWidth(1) = 500
    .ColWidth(2) = 500
    .ColWidth(3) = 500
    .ColWidth(4) = 500
    .ColWidth(5) = 500
    .SubtotalPosition = flexSTBelow
    
    .Subtotal flexSTSum, -1, 2, "#0", vbYellow, , , " „ Ê”ÿ"
    .Subtotal flexSTSum, -1, 3, "#0", vbYellow, , , " „ Ê”ÿ"
    
    .Subtotal flexSTAverage, -1, 2, "#0", vbYellow, , , " „ Ê”ÿ"
    .Subtotal flexSTAverage, -1, 3, "#0", vbYellow, , , " „ Ê”ÿ"
    End With
End Sub

Private Sub FIXGRID5()
    With GRID5
    .RowHeight(0) = 500
    .WordWrap = True
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    .TextMatrix(0, 0) = "«·ÌÊ„"
    .TextMatrix(0, 1) = "⁄œœ ÿ·»»« "
    .TextMatrix(0, 2) = "⁄œœ „‰ﬁ–"
    .TextMatrix(0, 3) = "„ »ﬁÏ"
    .TextMatrix(0, 4) = "⁄œœ ‘Õ‰« "
    .TextMatrix(0, 5) = "‘Õ‰«  „·€Ì…"
    
    
    .RowHeight(0) = 700
    .WordWrap = True
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
    If .Rows > 1 Then .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    .ColWidth(0) = 1200
    .ColWidth(1) = 700
    .ColWidth(2) = 700
    .ColWidth(3) = 700
    .ColWidth(4) = 700
    .ColWidth(5) = 700
    .SubtotalPosition = flexSTBelow
    
    
    .Subtotal flexSTAverage, -1, 1, "#0", vbYellow, , , " „ Ê”ÿ"
    .Subtotal flexSTAverage, -1, 2, "#0", vbYellow, , , " „ Ê”ÿ"
    
    .Subtotal flexSTSum, -1, 4, "#0", vbYellow, , , "≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, 5, "#0", vbYellow, , , "≈Ã„«·Ì"
    
    .AddItem "", 1
     cStr1 = "SELECT COUNT(FILE6_90H.DOC_NO) FROM FILE6_90H WHERE FILE6_90H.SALES_DATE IS NULL AND FILE6_90H.DELORDER_DATE IS NULL AND FILE6_90H.DATE < " & DateSq(xDate1.text)
    .TextMatrix(1, 0) = "—’Ìœ ”«»ﬁ"
    .TextMatrix(1, 3) = GetDesca(cStr1, con)
    
    End With
End Sub
Sub myload5()
'Dim cWhere As String, cF1 As String, cF2 As String
'If IsDate(xDate1.text) Then cWhere = " where DATE >= " & DateSq(xDate1.text)
'If IsDate(xdate2.text) Then cWhere = cWhere & " AND DATE <= " & DateSq(xdate2.text)
'
'cF1 = " , (SELECT COUNT(FILE6_90H.DOC_NO) FROM FILE6_90H WHERE FILE6_90H.DATE = Q_DATE_ORDER.DATE  )   AS newRDER"
'cF2 = " , (SELECT COUNT(FILE6_90H.DOC_NO) FROM FILE6_90H WHERE FILE6_90H.SALES_DATE = Q_DATE_ORDER.DATE )  AS SALES"
'cF3 = " , (SELECT COUNT(FILE6_90H.DOC_NO) FROM FILE6_90H WHERE FILE6_90H.DATE = Q_DATE_ORDER.DATE  AND  FILE6_90H.SALES_DATE IS NULL AND FILE6_90H.DELORDER_DATE IS NULL)  AS BALORDER"
'
'
'
'    Dim cString As String
'    cString = " select Q_DATE_ORDER.date  " & cF1 & cF2 & cF3 & _
'            "   FROM Q_DATE_ORDER " & _
'            "   where DATE >= " & DateSq(xDate1.text) & _
'            "   AND DATE <= " & DateSq(xdate2.text) & _
'            "   GROUP BY Q_DATE_ORDER.date " & _
'            "   ORDER BY Q_DATE_ORDER.date "
'    Set data6.Recordset = myRecordSet(cString, con)
'    FIXGRID5

Dim cString As New ChilkatStringBuilder
'
'    .TextMatrix(0, 0) = "«·ÌÊ„"
'    .TextMatrix(0, 1) = "⁄œœ ÿ·»»« "
'    .TextMatrix(0, 2) = "⁄œœ „‰ﬁ–"
'    .TextMatrix(0, 3) = "„ »ﬁÏ"
'    .TextMatrix(0, 4) = "⁄œœ ‘Õ‰« "
'    .TextMatrix(0, 5) = "‘Õ‰«  „·€Ì…"

cString.Append "SELECT Q_DATE_ORDER.DATE,"
cString.Append "(SELECT COUNT(FILE6_90H.DOC_NO) FROM FILE6_90H WHERE FILE6_90H.DATE = Q_DATE_ORDER.DATE) AS newOrder,"
cString.Append "(SELECT COUNT(FILE6_90H.DOC_NO) FROM FILE6_90H WHERE FILE6_90H.SALES_DATE = Q_DATE_ORDER.DATE)  AS SALES,"
cString.Append "(SELECT COUNT(FILE6_90H.DOC_NO) FROM FILE6_90H WHERE FILE6_90H.DATE = Q_DATE_ORDER.DATE  AND FILE6_90H.SALES_DATE IS NULL AND FILE6_90H.DELORDER_DATE IS NULL),"
cString.Append "(SELECT COUNT(FILE6_20H.DOC_NO) FROM FILE6_20H WHERE FILE6_20H.DATE_SHIP = Q_DATE_ORDER.DATE),"
cString.Append "(SELECT COUNT(FILE6_90H.DOC_NO) FROM FILE6_90H WHERE FILE6_90H.DELORDER_DATE = Q_DATE_ORDER.DATE)"
cString.Append " FROM Q_DATE_ORDER "
cString.Append " WHERE DATE >= " & DateSq(xDate1.text)
cString.Append " AND DATE <= " & DateSq(xdate2.text)
cString.Append " GROUP BY Q_DATE_ORDER.DATE"
cString.Append " ORDER BY Q_DATE_ORDER.DATE"

Set data6.RecordSet = mycmd(cString.GetAsString, con)
FIXGRID5
End Sub
Private Sub SearchManOrder(pWhere As String, pCaption As String, sFieldName As String)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 4)
Set Generalarray(0) = Me

Generalarray(1) = "SELECT FILE6_25.DESCA," & _
                  "COUNT(DISTINCT FILE6_90H.DOC_NO)," & _
                  "SUM(FILE6_90.QUANT)," & _
                  "SUM(FILE6_90.TOTAL)" & _
                  " FROM FILE6_90H " & _
                  " INNER JOIN FILE6_90 ON FILE6_90H.DOC_NO = FILE6_90.DOC_NO " & _
                  " INNER JOIN FILE6_25  ON FILE6_90H.MAN = FILE6_25.CODE"
                                    
If pWhere <> "" Then
    Generalarray(1) = Generalarray(1) & " AND " & pWhere
End If

Generalarray(2) = " GROUP BY FILE6_25.CODE,FILE6_25.DESCA " & _
                  " ORDER BY  COUNT(*)"

Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·„‰œÊ»"
listarray(0, 1) = "(%%FILE6_25.DESCA%%)"

GrdArray(0, 0) = "«·„‰œÊ»"
GrdArray(0, 1) = 6000

GrdArray(1, 0) = sFieldName
GrdArray(1, 1) = 2000
GrdArray(1, 3) = True
GrdArray(1, 4) = "#0"

GrdArray(2, 0) = "«·ﬂ„Ì…"
GrdArray(2, 1) = 1100
GrdArray(2, 3) = True
GrdArray(2, 4) = "#0"

GrdArray(3, 0) = "«·≈Ã„«·Ì"
GrdArray(3, 1) = 1500
GrdArray(3, 3) = True
GrdArray(3, 4) = "#0.00"


searchArray = Array(Generalarray, listarray, GrdArray)
'oSearch_total.sFlag = "charge"
oSearch_Total.sCaption = pCaption
oSearch_Total.bPrint = True
oSearch_Total.Show 1
End Sub
Sub myProc()
End Sub
Sub myPrint(pFlag As String)
Dim sCaption As String
If GRID5.Row = 1 Then
    sCaption = "ÿ·»Ì«  „‰œÊ»Ì‰ „‰›–… „‰ " & BetweenString(xDate1.text, xdate2.text)
ElseIf GRID5.Row > 1 Then
    sCaption = "ÿ·»Ì«  „‰œÊ»Ì‰ „‰›–… ›Ì " & myFormat_p(GRID5.TextMatrix(GRID5.Row, 0))
End If

ReDim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 1)
aRow(0) = AddFlag(aRow(0), "text", "«·≈Ã„«·Ì")
printGrd_abd.doprint oSearch_Total.GRID1, 1, -2, sCaption, , , , False, False, 10, , aRow
printGrd_abd.Show 1
End Sub
Private Sub grid5_DblClick()
Dim cWhere As String
Dim sCaption As String
If GRID5.Row = 1 Then
    If IsDate(xDate1.text) Then
        cWhere = "FILE6_90H.SALES_dATE >= " & DateSq(xDate1.text)
    End If

    If IsDate(xdate2.text) Then
        cWhere = cWhere & Tr(cWhere) & "FILE6_90H.SALES_DATE <= " & DateSq(xdate2.text)
    End If
    sCaption = "ÿ·»Ì«  „‰œÊ»Ì‰ „‰›–… „‰ " & BetweenString(xDate1.text, xdate2.text)
ElseIf GRID5.Row > 1 Then
    cWhere = "FILE6_90H.SALES_DATE = " & DateSq(GRID5.TextMatrix(GRID5.Row, 0))
    sCaption = "ÿ·»Ì«  „‰œÊ»Ì‰ „‰›–… ›Ì " & myFormat_p(GRID5.TextMatrix(GRID5.Row, 0))
End If

SearchManOrder cWhere, sCaption, "«·ÿ·»Ì«  «·„‰›–…"
End Sub
