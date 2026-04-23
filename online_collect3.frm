VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form online_collectfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„œ›Ê⁄«  online"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   18255
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6180
      Left            =   45
      TabIndex        =   33
      Top             =   2115
      Width           =   18060
      _ExtentX        =   31856
      _ExtentY        =   10901
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
      TabCaption(0)   =   "”œ«œ ⁄‰ ÿ—Ìﬁ ‘—ﬂ… «·‘Õ‰"
      TabPicture(0)   =   "online_collect3.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "”œ«œ „‰ Œ·«· «·„Êﬁ⁄"
      TabPicture(1)   =   "online_collect3.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   5730
         Left            =   -74955
         TabIndex        =   34
         Top             =   360
         Width           =   17925
         _cx             =   31618
         _cy             =   10107
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
         Rows            =   1
         Cols            =   8
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
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   5730
         Left            =   90
         TabIndex        =   35
         Top             =   360
         Width           =   17880
         _cx             =   31538
         _cy             =   10107
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
         Rows            =   1
         Cols            =   9
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
   End
   Begin VB.Frame Frame3 
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
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   8325
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   135
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
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
         Picture         =   "online_collect3.frx":0038
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect3.frx":21DF
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   135
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   741
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
         Picture         =   "online_collect3.frx":4226
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect3.frx":6311
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
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
         Picture         =   "online_collect3.frx":830B
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect3.frx":A41C
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   135
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   741
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
         Picture         =   "online_collect3.frx":C416
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect3.frx":E63A
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   12870
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   45
      Width           =   5325
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4005
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "online_collect3.frx":1070B
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect3.frx":12AD6
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   2700
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "online_collect3.frx":14B7F
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect3.frx":16B87
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1350
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   135
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
         Picture         =   "online_collect3.frx":18B3E
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect3.frx":1B2DA
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "online_collect3.frx":1D76E
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      Height          =   1140
      Index           =   4
      Left            =   8685
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   945
      Width           =   1455
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   45
         TabIndex        =   17
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
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
         Picture         =   "online_collect3.frx":1FA91
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect3.frx":223B6
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   465
         Left            =   45
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   630
         Width           =   1365
         _ExtentX        =   2408
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
         Picture         =   "online_collect3.frx":24C0A
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect3.frx":26D6A
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   600
      Left            =   14400
      TabIndex        =   7
      Top             =   8325
      Width           =   3705
      Begin MSDataListLib.DataCombo xYear 
         Height          =   330
         Left            =   1665
         TabIndex        =   8
         Top             =   180
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
         TabIndex        =   9
         Top             =   180
         Width           =   870
         _ExtentX        =   1535
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
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·”‰…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2970
         TabIndex        =   11
         Tag             =   "Color"
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·‘Â—"
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
         Left            =   945
         TabIndex        =   10
         Tag             =   "Color"
         Top             =   225
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   10170
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   765
      Width           =   7935
      Begin VB.TextBox Text1 
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
         Left            =   135
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   900
         Width           =   6630
      End
      Begin VB.TextBox xDoc_No 
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
         Left            =   5400
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox xDate 
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
         Height          =   345
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "D"
         Top             =   180
         Width           =   2220
      End
      Begin MSDataListLib.DataCombo xBank 
         Height          =   330
         Left            =   4545
         TabIndex        =   1
         Top             =   540
         Width           =   2220
         _ExtentX        =   3916
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Left            =   135
         TabIndex        =   36
         Top             =   540
         Width           =   2220
         _ExtentX        =   3916
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "‘—ﬂ… «·‘Õ‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
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
         TabIndex        =   37
         Top             =   585
         Width           =   945
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·«ÕŸ« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   990
         Width           =   795
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·»‰ﬂ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6885
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   585
         Width           =   360
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   270
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ „” ‰œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6885
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   795
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   465
      Left            =   4995
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
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
   Begin Threed.SSCommand cmdFilter 
      Height          =   375
      Left            =   6975
      TabIndex        =   6
      Top             =   1665
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Undo Filter"
      ButtonStyle     =   3
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   9720
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   45
         Width           =   3195
         _ExtentX        =   5636
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
         Left            =   3240
         TabIndex        =   14
         Top             =   45
         Width           =   3855
         _ExtentX        =   6800
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
         Left            =   7110
         TabIndex        =   15
         Top             =   45
         Width           =   4365
         _ExtentX        =   7699
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
   End
   Begin Threed.SSCommand cmdBank 
      Height          =   420
      Left            =   10665
      TabIndex        =   29
      Top             =   8460
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   741
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
      Caption         =   "«Œ «— «·»‰ﬂ"
      TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
      ButtonStyle     =   3
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   465
      Left            =   855
      Top             =   855
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
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
   Begin MSAdodcLib.Adodc data11 
      Height          =   465
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
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
Attribute VB_Name = "online_collectfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nType As Byte, bEdit As Boolean, sDoc_no As String
Public bCheck As Boolean
Dim bEditRecord As Boolean, bAct As Boolean
Dim bStopCell As Boolean
Dim oSeachCode As New Search_abd, oSearchDoc As New Search, oSearchProject As New Search, oSearchSup As New Search
Dim cFilter As String, cSelect As String
Dim con As New ADODB.Connection, CardTable As ADODB.Recordset
Dim cFile As String, cFileHeader As String, aBoxs As Variant
Dim cList1 As String, cList2 As String, cList3 As String, sDef As String
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1, Optional Row2 As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[BANK]", addstring(xBank.BoundText))
aInsert = AddFlag(aInsert, IIf(xDoc_No.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(GetComputerName))
aInsert = AddFlag(aInsert, IIf(xDoc_No.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")
con.BeginTrans
On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    xDoc_No.text = Newflag(cFileHeader, "DOC_NO", con)
    aInsert = AddFlag(aInsert, "DOC_NO", addvalue(xDoc_No.text))
    con.Execute addInsert(aInsert, cFileHeader)
Else
    con.Execute addUpdate(aInsert, cFileHeader, "doc_no = " & addvalue(xDoc_No.text))
End If

If (Row = -1 And Row2 = -1) Or Row <> -1 Then
    myreplaceGrd Row
End If

If (Row = -1 And Row2 = -1) Or Row2 <> -1 Then
    myreplaceGrd2 Row2
End If

con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc(Optional sControl As String = "")
If ActiveControl.Name = grid1.Name Then
    If Not GrdDesc(oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0), "", grid1, grid1.Row) Then Exit Sub
    'grid1.TextMatrix(grid1.Row, 0) = oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0)
    Grid1_AfterEdit grid1.Row, grid1.Col
    oSeachCode.Hide
    CellPos 13, grid1.Row, grid1.Col
ElseIf ActiveControl.Name = grid2.Name Then
    If Not GrdDesc(oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0), "", grid2, grid2.Row) Then Exit Sub
    grid2_AfterEdit grid2.Row, grid2.Col
    oSeachCode.Hide
    CellPos2 13, grid2.Row, grid2.Col
ElseIf ActiveControl.Name = cmdInform.Name Then
    openCardTable tbMode.tbFind, oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
ElseIf ActiveControl.Name = cmdCont.Name Then
    If oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0) = "" Then
        cmdCont.Tag = ""
        cmdCont.Caption = cmdCont.TagVariant
    Else
        cmdCont.Tag = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
        cmdCont.Caption = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 1)
    End If
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
    oSearchSup.Hide
ElseIf ActiveControl.Name = cmdProject.Name Then
    If oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 0) = "" Then
        cmdProject.Tag = ""
        cmdProject.Caption = cmdProject.TagVariant
    Else
        cmdProject.Tag = oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 0)
        cmdProject.Caption = oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 1)
    End If
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
    oSearchProject.Hide
End If
End Sub
Private Sub CmdDel_Click()
If MsgBox("Õ–› ?", vbOKCancel + vbDefaultButton2 + vbCritical) <> vbOK Then Exit Sub

con.BeginTrans
On Error GoTo myerror
con.Execute "Delete  From " & cFile & " where Doc_No = " & addvalue(xDoc_No.text)
con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & addvalue(xDoc_No.text)
con.CommitTrans

If sDoc_no <> "" Then Exit Sub

If Not openCardTable(tbMode.tbPrevious, xDoc_No.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
cString = "SELECT FILE9_10H.Doc_No, Convert(Varchar(10),FILE9_10H.DATE,111),FILE9_10.DESCA " & _
          " FROM FILE9_10H INNER JOIN FILE9_10 ON FILE9_10H.CODE = FILE9_10.CODE"
If pFilter <> "" Then cString = cString & " WHERE " & pFilter
Generalarray(1) = cString
Generalarray(2) = "order by FILE9_10H.DATE ,FILE9_10H.DOC_NO "
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·„ﬁ«Ê·"
listarray(0, 1) = "(%%FILE9_10.DESCA%%)"

listarray(1, 0) = " «—ÌŒ «·„” ‰œ"
listarray(1, 1) = "(##[Date]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1350

GrdArray(2, 0) = "«·≈”„"
GrdArray(2, 1) = 9000

Dim aFilter As Variant
aFilter = AddFlag(aFilter, "FILTER", True)
aFilter = AddFlag(aFilter, "FIELD", "DOC_NO")

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.sCaption = IIf(nType = 0, "«” ⁄·«„ „œ›Ê⁄«  ‰ﬁœÌ… «·Ì „ﬁ«Ê·Ì‰", "«” ⁄·«„ „ﬁ»Ê÷«  ‰ﬁœÌ… „‰ «·„ﬁ«Ê·Ì‰")
oSearchDoc.aFilter = aFilter
oSearchDoc.Show 1
End Sub
Private Sub CmdInform_Click()
CardLookup cFilter
End Sub
Private Sub CmdFirst_Click()
If Not openCardTable(tbMode.tbFirst) Then
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
If Not openCardTable Then
    myDefine
End If
End Sub
Private Sub CmdNext_Click()
If Not openCardTable(tbMode.tbNext, xDoc_No.text) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, xDoc_No.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdNewInv_Click()
myDefine
End Sub
Private Sub cmdProject_Click()
ProjectLookup Me, oSearchProject, , , cmdProject.Tag <> ""
End Sub
Private Sub cmdSave_Click()
If Not myValid Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If sDoc_no <> "" Then
    Unload Me
    Exit Sub
End If
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub cmdCont_Click()
contlookup Me, oSearchSup, , , cmdCont.Tag <> ""
End Sub
Private Sub CmdUndo_Click()
'openCardTable
myUndo
End Sub
Private Sub cmdYear_Click()
Set oSearchYear.myform = Me
oSearchYear.bEmpty = cmdYear.Tag <> ""
oSearchYear.Show 1
End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xDoc_No.Tag = LoadMode Then
        grid1.SetFocus
    Else
        xBank.SetFocus
    End If
    Err.Clear
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then KeyAscii = 0
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End If
End Sub
Private Sub Form_Load()
bEdit = True
cFile = "FILE9_10"
cFileHeader = "FILE9_10H"

openCon con

Set DATA1.Recordset = myCmd("SELECT * FROM FILE5_10 ORDER BY DESCA", con)
Set xBank.RowSource = DATA1
xBank.ListField = "Desca"
xBank.BoundColumn = "Code"

Set grid1.DataSource = data10
Set grid2.DataSource = data11

If Not openCardTable Then myDefine
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveText Me, , Array(chkYear.Name, chkMonth.Name, chkYear.Name)
CardTable.Close
Set CardTable = Nothing
closeCon con
Set online_collectfrm = Nothing
Err.Clear
End Sub
Private Function myValid(Optional bIgMsg As Boolean = True) As Boolean
If Trim(xDoc_No.text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Not xBank.MatchedWithList Then
    MsgBox "«·„ﬁ«Ê· €Ì— „”Ã·"
    Exit Function
End If

myValid = True
End Function
Private Sub myload()
xDoc_No.text = CardTable!doc_no
xDate.text = myFormat_p(CardTable!Date)
xBank.BoundText = CardTable!BANK & ""
panel1(0).Caption = CardTable!UserName & " " & myFormat_p(CardTable!Time, True)
panel1(1).Caption = CardTable!USERNAME2 & " " & myFormat_p(CardTable!TIME2, True)
Handlecontrols LoadMode
myLoadGrd
myloadgrd2
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
CellPos2 13, grid2.Rows - 2, grid2.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
xDoc_No.text = Newflag(cFileHeader, "DOC_NO", con)
xDate.text = myFormat_p(Date)
xBank.BoundText = ""

fixGrd
grid1.Rows = 1
myAddItem

Fixgrd2
grid2.Rows = 1
myAddItem2

Handlecontrols DefineMode
'StatusBar1.Panels(2).text = "·« ÌÊÃœ ”œ«œ"
'StatusBar1.Panels(3).text = ""
'StatusBar1.Panels(4).text = ""

CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit
cmdFilter.Visible = cmdFilter.Tag <> ""
cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmddel.Enabled = nMode = LoadMode And bEditRecord

cmdSave.Enabled = bEditRecord
cmddel.Enabled = nMode = LoadMode And bEditRecord

Dim nRecord As Long, nRecords As Long
retRecords xDoc_No.text, nRecords, nRecord

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
     panel1(0).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""

xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub

Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If Not ValidNum(xDoc_No.text) Then
     If xDoc_No.Tag = LoadMode Then
        myDefine
    Else
        xDoc_No.text = ""
    End If
Else
    If (Not (CardTable.EOF)) And xDoc_No.Tag = LoadMode Then
        If CardTable!doc_no = xDoc_No.text Then
            Exit Sub
        End If
    End If
    
    openCardTable xDoc_No.text
    If Not CardTable.EOF Then
        myload
    ElseIf xDoc_No.Tag = LoadMode Then
        myDefine
    Else
'        xDoc_No.Text = ""
    End If
End If
End Sub
Private Function CalcTotals(Optional Row As Long = -1)
Dim nTotal As Double
With grid1
For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
    .TextMatrix(i, 8) = .ValueMatrix(i, 7) - .ValueMatrix(i, 4)
    nTotal = mRound(nTotal + .ValueMatrix(i, 8))
Next

'If Row = 0 Then Exit Function
'.TextMatrix(Row, 6) = ""
'.TextMatrix(Row, 7) = ""
End With
End Function
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_No As String = "", Optional bDefine As Boolean = True) As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_No) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

'On Error GoTo myerror
cFilter = ""
If cmdFilter.Tag <> "" Then cFilter = cFilter & Tr(cFilter) & "DOC_NO IN (" & cmdFilter.Tag & ")"
If cmdBank.Tag <> "" Then cFilter = cFilter & Tr(cFilter) & "BANK = " & MyParn(cmdBank.Tag)
If xYear.text <> "" Then cFilter = cFilter & Tr(cFilter) & "YEAR(DATE) = " & xYear.text
If xMonth.text <> "" Then cFilter = cFilter & Tr(cFilter) & "MONTH(DATE) = " & xMonth.text

If sDoc_no <> "" Then cFilter = "DOC_NO = " & sDoc_no
cString = "SELECT TOP 1 * " & _
          " FROM " & cFileHeader

If pMode = tbMode.tbFirst Then
    cOrder = "Order by Doc_no"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by Doc_no DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "Doc_no = " & pDoc_No
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "Doc_no < " & pDoc_No
    cOrder = "Order by Doc_no desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "Doc_no > " & pDoc_No
    cOrder = "Order by Doc_no"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

If cFilter <> "" Then
    cString = cString & IIf(cWhere = "", " WHERE ", " AND ") & cFilter
End If

cString = cString & " " & cOrder

Set CardTable = myCmd(cString, con, adText)

If (Not CardTable.EOF) Then
    myload
    openCardTable = True
End If
Me.MousePointer = 0
Exit Function
myerror:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function
Private Sub myUndo()
If xDoc_No.Tag = DefineMode Then
    If Not openCardTable Then
        CmdNewInv_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xDate_DblClick()
Set datefrm.oDate = xDate
datefrm.Show 1
End Sub
Private Sub chkDay_Click()
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub chkMonth_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Sub chkOpen_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Sub chkYear_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Function retRecords(pDoc_No, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_No <> "" Then
    cString = "SELECT Count(*) AS records,SUM(CASE WHEN DOC_NO <= " & pDoc_No & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM " & cFileHeader & Tr(cFilter, " WHERE ") & cFilter

Set loctable = myCmd(cString, con)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Sub cmdFilter_Click()
cmdFilter.Tag = ""
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Sub myproc2(pFilter As String)
oSearchDoc.Hide
cmdFilter.Tag = pFilter
If Not openCardTable(tbMode.tbFirst, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub myreplaceGrd(Row As Long)
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "INV_NO", addstring(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "[VALUE]", grid1.ValueMatrix(i, 7))
        aInsert = AddFlag(aInsert, "[TYPE]", "1")
        aInsert = AddFlag(aInsert, "[FLAG]", "1")
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, cFile)
        Else
            con.Execute addUpdate(aInsert, cFile, "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not myValid(True) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myLoadGrd
    If Row < grid1.Rows - 1 Then
        grid1.Select Row, Col
    Else
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If

If Not validRow(Row, Col) Then
    CalcTotals Row
    Exit Sub
End If

With grid1
If Row = grid1.Rows - 1 Then
    myAddItem
'ElseIf Row = grid1.Rows - 2 And (Col = 0) Then
'    MyEditItem grid1, Row, Col
End If

If myreplace(Row) Then
    If xDoc_No.Tag = DefineMode Then
        Handlecontrols LoadMode
        myLoadGrd
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myLoadGrd
    Else
        CalcTotals Row
    End If
Else
    myLoadGrd
End If
End With
End Sub
Private Sub grid1_EnterCell()
If Not bEditRecord Then
    grid1.Editable = flexEDNone
ElseIf (grid1.Col = 0 Or grid1.Col = 1 Or grid1.Col = 7) Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub myLoadGrd()
Dim cString As String
Dim aFields(9)
aFields(0) = "FILE9_10.INV_NO"
aFields(1) = "FILE6_20H.SHIP_NO"
aFields(2) = "FILE6_20H.DATE"
aFields(3) = "FILE6_20H.[NAME]"
aFields(4) = "FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT"
aFields(5) = "FILE6_20H.CHARGE1 + FILE6_20H.CHARGE2"
aFields(6) = "(FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT) + FILE6_20H.CHARGE1 + FILE6_20H.CHARGE2"
aFields(7) = "FILE9_10.[VALUE]"
aFields(8) = "FILE9_10.[VALUE] - (FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT)"
aFields(9) = "FILE9_10.ID"

cString = "SELECT " & arString(aFields, ",") & _
         " FROM FILE9_10 INNER JOIN FILE9_10H ON FILE9_10.DOC_NO = FILE9_10H.DOC_NO" & _
         " INNER JOIN FILE6_20H ON FILE9_10.INV_NO = FILE6_20H.DOC_NO" & _
         " WHERE FILE9_10.DOC_NO = " & addvalue(xDoc_No.text) & _
         " AND FILE9_10.TYPE = 1"

Set data10.Recordset = myCmd(cString, con)

myAddItem

CalcTotals
fixGrd
End Sub
Private Sub grid1_GotFocus()
grid1_EnterCell
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 112 And grid1.Col = 0 Then
    invLookup Me, oSeachCode
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from " & cFile & " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        myRemove grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
        MsgBox "«·›« Ê—… €Ì— „”Ã·…"
    ElseIf Not GrdDesc(grid1.EditText, "", grid1, Row) Then
        Cancel = True
    End If
ElseIf Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
        MsgBox "—ﬁ„ »Ê·Ì’… «·‘Õ‰ €Ì— „”Ã·"
    ElseIf Not GrdDesc("", grid1.EditText, grid1, Row) Then
        Cancel = True
    End If
End If
End Sub
Private Sub fixGrd()
With grid1
.TextMatrix(0, 0) = "—ﬁ„ «·›« Ê—…"
.TextMatrix(0, 1) = "»Ê·Ì’… «·‘Õ‰"
.TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
.TextMatrix(0, 3) = "«·«”„"
.TextMatrix(0, 4) = "≈Ã„«·Ì «·›« Ê—…"
.TextMatrix(0, 5) = "‘Õ‰ „”Ã·"
.TextMatrix(0, 6) = "≈Ã„«·Ì «·»Ê·Ì’…"
.TextMatrix(0, 7) = "≈Ã„«·Ì «·„œ›Ê⁄"
.TextMatrix(0, 8) = "‘Õ‰ „œ›Ê⁄"

.RowHeight(0) = 600
.ColWidth(0) = 1800
.ColWidth(1) = 1800
.ColWidth(2) = 1300
.ColWidth(3) = 4000
.ColWidth(4) = 1300
.ColWidth(5) = 1300
.ColWidth(6) = 1300
.ColWidth(7) = 1300
.ColWidth(8) = 1300

.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True

For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub myAddItem()
With grid1
.AddItem ""
If grid1.Rows > 2 Then
    '.TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
End If
End With
End Sub
Private Function validRow(Row As Long, Optional Col As Long = -1) As Boolean
With grid1
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If .ValueMatrix(Row, 7) = 0 Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col < 1 And grid1.TextMatrix(Row, 2) = "" Then
    grid1.Col = 1
ElseIf Col < 7 Then
    grid1.Col = 7
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 7)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
CalcTotals
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If Col = 0 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    If Col = 1 And (grid1.TextMatrix(Row, Col) = "" Or grid1.TextMatrix(Row, grid1.Cols - 1) = "") Then Exit Sub
    If Col = 2 And (grid1.TextMatrix(Row, Col) = "" Or grid1.TextMatrix(Row, grid1.Cols - 1) = "") Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid1.Col = 0 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    If grid1.Col = 1 And (grid1.TextMatrix(grid1.Row, grid1.Col) = "" Or grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Then Exit Sub
    If grid1.Col = 2 And (grid1.TextMatrix(grid1.Row, grid1.Col) = "" Or grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Then Exit Sub
    KeyAscii = 0
End If
End Sub
Public Sub invLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(6, 1)
Dim cWhere As String
Set Generalarray(0) = oForm
'                       0                   1
cString = "SELECT FILE6_20H.DOC_NO," & _
          "FILE6_20H.DOC_NO2," & _
          "CONVERT(VARCHAR(10),FILE6_20H.DATE,111)," & _
          "FILE6_20H.SHIP_NO, " & _
          "FILE6_20H.ONLINE_DOC," & _
          "FILE6_20H.[NAME]," & _
          "FILE6_20H.PHONE " & _
          " FROM  FILE6_20H " & _
          " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE AND FILE0_40.ONLINE = 1" & _
          " LEFT JOIN FILE9_10 ON FILE6_20H.DOC_NO = FILE9_10.INV_NO" & _
          " WHERE FILE6_20H.ONLINE = 2" & _
          " AND FILE9_10.INV_NO IS NULL AND FILE6_20H.DATE_PAY IS NULL"


'If cBranch <> "" Then
'    cString = cString & " AND FILE6_20H.BRANCH = " & MyParn(cBranch)
'End If

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " AND " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by FILE6_20H.DATE DESC,FILE6_20H.DOC_NO"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "(FILE6_20H.DOC_NO = 'cFilter' OR FILE6_20H.DOC_NO2 = 'cFilter')"

listarray(1, 0) = "«·«”„-«·⁄‰Ê«‰-—ﬁ„ «· ·Ì›Ê‰"
listarray(1, 1) = "(%%FILE6_20H.NAME%% OR %%FILE6_20H.ADDRESS%% OR PHONE LIKE 'cFilter%')"

listarray(2, 0) = "—ﬁ„ «·»Ê·Ì’…"
listarray(2, 1) = "(SHIP_NO LIKE 'cFilter%')"

listarray(3, 0) = "—ﬁ„ «·ÿ·»Ì…"
listarray(3, 1) = "(ONLINE_DOC LIKE 'cFilter%')"

GrdArray(0, 0) = "—›„ «·›« Ê—…"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "—›„ «·›« Ê—…"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = " «—ÌŒ «·›« Ê—…"
GrdArray(2, 1) = 1300

GrdArray(3, 0) = "—ﬁ„ «·»Ê·Ì’…"
GrdArray(3, 1) = 1500

GrdArray(4, 0) = "—ﬁ„ «·ÿ·»Ì…"
GrdArray(4, 1) = 1500

GrdArray(5, 0) = "«·«”„"
GrdArray(5, 1) = 3000

GrdArray(6, 0) = "«· ·Ì›Ê‰"
GrdArray(6, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE6_20H.DOC_NO")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ ›Ê« Ì— «Ê‰ ·«Ì‰"
oSearch.Show 1
End Sub
Private Function GrdLading(sLading As String, Row As Long, Optional pType As String = "1") As Boolean
Dim loctable As New ADODB.Recordset
Dim cString As String, sFound As String
cString = "SELECT FILE6_20H.* FROM FILE6_20H whe"

If sLading <> "" Then
    cString = cString & " WHERE SHIP_NO = " & MyParn(sLading)
Else
    cString = cString & " WHERE DOC_NO = " & MyParn(sinv_no)
End If

Set loctable = myCmd(cString, con)

If loctable.EOF Then
    MsgBox "›« Ê—… €Ì— „ÊÃÊœ…"
    Exit Function
End If

If Val(loctable!online & "") <> 2 Then
    MsgBox "›« Ê—… ·Ì”  «Ê‰·«Ì‰"
    Exit Function
End If

If Not loctable!PRINTED Then
    MsgBox "›« Ê—… €Ì— „€·ﬁ… »«·‰Ÿ«„"
    Exit Function
End If


If Not IsNull(loctable!DATE_PAY) Then
    MsgBox "«·›« Ê—… „”œœ… » «—ÌŒ " & loctable!DATE_PAY
    Exit Function
End If
    
    sFound = myField("SELECT DOC_NO FROM FILE9_10 WHERE INV_NO = " & MyParn(loctable!doc_no), con) & ""
    
    If sFound <> "" Then
         MsgBox "«·›« Ê—… „”œœ ›Ì „” ‰œ —ﬁ„ " & sFound
         Exit Function
    End If

    pGrid.TextMatrix(Row, 0) = loctable!doc_no
    pGrid.TextMatrix(Row, 1) = loctable!SHIP_NO & ""
    pGrid.TextMatrix(Row, 2) = myFormat_p(loctable!Date)
    pGrid.TextMatrix(Row, 3) = loctable!Name & ""
    pGrid.TextMatrix(Row, 4) = loctable!TOTAL_ITEM - loctable!discount
    pGrid.TextMatrix(Row, 5) = loctable!CHARGE1 + loctable!charge2
    pGrid.TextMatrix(Row, 6) = (loctable!TOTAL_ITEM - loctable!discount) + loctable!CHARGE1 + loctable!charge2
End If
GrdDesc = True
End Function

Private Function GrdDesc(ByVal sinv_no, ByVal sLading, pGrid As Object, Row As Long, Optional pType As String = "1") As Boolean
Dim loctable As New ADODB.Recordset
Dim cString As String, sFound As String
cString = "SELECT FILE6_20H.* FROM FILE6_20H"

If sLading <> "" Then
    cString = cString & " WHERE SHIP_NO = " & MyParn(sLading)
Else
    cString = cString & " WHERE DOC_NO = " & MyParn(sinv_no)
End If

Set loctable = myCmd(cString, con)

If loctable.EOF Then
    MsgBox "›« Ê—… €Ì— „ÊÃÊœ…"
    Exit Function
End If

If loctable.RecordCount > 1 Then
    GrdDesc = True
    Exit Function
End If

If Val(loctable!online & "") <> 2 Then
    MsgBox "›« Ê—… ·Ì”  «Ê‰·«Ì‰"
    Exit Function
End If

If Not loctable!PRINTED Then
    MsgBox "›« Ê—… €Ì— „€·ﬁ… »«·‰Ÿ«„"
    Exit Function
End If


If Not IsNull(loctable!DATE_PAY) Then
    MsgBox "«·›« Ê—… „”œœ… » «—ÌŒ " & loctable!DATE_PAY
    Exit Function
End If
    
    sFound = myField("SELECT DOC_NO FROM FILE9_10 WHERE INV_NO = " & MyParn(loctable!doc_no), con) & ""
    
    If sFound <> "" Then
         MsgBox "«·›« Ê—… „”œœ ›Ì „” ‰œ —ﬁ„ " & sFound
         Exit Function
    End If

    pGrid.TextMatrix(Row, 0) = loctable!doc_no
    pGrid.TextMatrix(Row, 1) = loctable!SHIP_NO & ""
    pGrid.TextMatrix(Row, 2) = myFormat_p(loctable!Date)
    pGrid.TextMatrix(Row, 3) = loctable!Name & ""
    pGrid.TextMatrix(Row, 4) = loctable!TOTAL_ITEM - loctable!discount
    pGrid.TextMatrix(Row, 5) = loctable!CHARGE1 + loctable!charge2
    pGrid.TextMatrix(Row, 6) = (loctable!TOTAL_ITEM - loctable!discount) + loctable!CHARGE1 + loctable!charge2
End If
GrdDesc = True
End Function


Private Sub myreplaceGrd2(Row As Long)
Dim aInsert As Variant
With grid2
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid2.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "INV_NO", addstring(grid2.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "[VALUE]", grid2.ValueMatrix(i, 7))
        aInsert = AddFlag(aInsert, "[CHARGE]", grid2.ValueMatrix(i, 8))
        aInsert = AddFlag(aInsert, "[TYPE]", "2")
        aInsert = AddFlag(aInsert, "[FLAG]", "1")
        If grid2.TextMatrix(i, grid2.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, cFile)
        Else
            con.Execute addUpdate(aInsert, cFile, "ID = " & grid2.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not myValid(True) Then
    On Error Resume Next
    grid2.SetFocus
    Err.Clear
    myloadgrd2
    If Row < grid2.Rows - 1 Then
        grid2.Select Row, Col
    Else
        CellPos2 13, grid2.Rows - 2, grid2.Cols - 1
    End If
    Exit Sub
End If

If Not validrow2(Row, Col) Then
    CalcTotals2 Row
    Exit Sub
End If

With grid2
If Row = grid2.Rows - 1 Then
    myAddItem2
'ElseIf Row = grid2.Rows - 2 And (Col = 0) Then
'    MyEditItem grid2, Row, Col
End If

If myreplace(, Row) Then
    If xDoc_No.Tag = DefineMode Then
        Handlecontrols LoadMode
        myloadgrd2
    ElseIf grid2.TextMatrix(Row, grid2.Cols - 1) = "" Then
        myloadgrd2
    Else
        CalcTotals2 Row
    End If
Else
    myloadgrd2
End If
End With
End Sub
Private Sub grid2_EnterCell()
If Not bEditRecord Then
    grid2.Editable = flexEDNone
ElseIf (grid2.Col = 0 Or grid2.Col = 1 Or grid2.Col = 7 Or grid2.Col = 8) Then
    grid2.Editable = flexEDKbdMouse
Else
    grid2.Editable = flexEDNone
End If
End Sub
Private Sub myloadgrd2()
Dim cString As String
Dim aFields(9)
aFields(0) = "FILE9_10.INV_NO"
aFields(1) = "FILE6_20H.SHIP_NO"
aFields(2) = "FILE6_20H.DATE"
aFields(3) = "FILE6_20H.[NAME]"
aFields(4) = "FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT"
aFields(5) = "FILE6_20H.CHARGE1 + FILE6_20H.CHARGE2"
aFields(6) = "(FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT) + FILE6_20H.CHARGE1 + FILE6_20H.CHARGE2"
aFields(7) = "FILE9_10.[VALUE]"
aFields(8) = "FILE9_10.[CHARGE]"
aFields(9) = "FILE9_10.ID"

cString = "SELECT " & arString(aFields, ",") & _
         " FROM FILE9_10 INNER JOIN FILE9_10H ON FILE9_10.DOC_NO = FILE9_10H.DOC_NO" & _
         " INNER JOIN FILE6_20H ON FILE9_10.INV_NO = FILE6_20H.DOC_NO" & _
         " WHERE FILE9_10.DOC_NO = " & addvalue(xDoc_No.text) & _
         " AND FILE9_10.TYPE = 2"

Set data11.Recordset = myCmd(cString, con)

myAddItem2

CalcTotals2
Fixgrd2
End Sub
Private Sub grid2_GotFocus()
grid2_EnterCell
End Sub
Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos2 KeyCode, grid2.Row, grid2.Col
ElseIf Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 112 And grid2.Col = 0 Then
    invLookup Me, oSeachCode
ElseIf KeyCode = 46 And grid2.Row <> grid2.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid2.TextMatrix(grid2.Row, grid2.Cols - 1) <> "" Then
            con.Execute "Delete from " & cFile & " where ID = " & grid2.TextMatrix(grid2.Row, grid2.Cols - 1)
        End If
        con.CommitTrans
        myRemove2 grid2.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid2_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(grid2.EditText) = "" Then
        Cancel = True
        MsgBox "«·›« Ê—… €Ì— „”Ã·…"
    ElseIf Not GrdDesc(grid2.EditText, "", grid2, Row) Then
        Cancel = True
    End If
ElseIf Col = 1 Then
    If Trim(grid2.EditText) = "" Then
        Cancel = True
        MsgBox "—ﬁ„ »Ê·Ì’… «·‘Õ‰ €Ì— „”Ã·"
    ElseIf Not GrdDesc("", grid2.EditText, grid2, Row) Then
        Cancel = True
    End If
End If
End Sub
Private Sub Fixgrd2()
With grid2
.TextMatrix(0, 0) = "—ﬁ„ «·›« Ê—…"
.TextMatrix(0, 1) = "»Ê·Ì’… «·‘Õ‰"
.TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
.TextMatrix(0, 3) = "«·«”„"
.TextMatrix(0, 4) = "≈Ã„«·Ì «·›« Ê—…"
.TextMatrix(0, 5) = "‘Õ‰ „”Ã·"
.TextMatrix(0, 6) = "≈Ã„«·Ì «·»Ê·Ì’…"
.TextMatrix(0, 7) = "≈Ã„«·Ì «·„œ›Ê⁄"
.TextMatrix(0, 8) = "‘Õ‰ «·‘—ﬂ…"

.RowHeight(0) = 600
.ColWidth(0) = 1800
.ColWidth(1) = 1800
.ColWidth(2) = 1300
.ColWidth(3) = 4000
.ColWidth(4) = 1300
.ColWidth(5) = 1300
.ColWidth(6) = 1300
.ColWidth(7) = 1300
.ColWidth(8) = 1300

.ColHidden(6) = True
.ColHidden(.Cols - 1) = True

For i = 1 To grid2.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub myAddItem2()
With grid2
.AddItem ""
If grid2.Rows > 2 Then
    '.TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
End If
End With
End Sub
Private Function validrow2(Row As Long, Optional Col As Long = -1) As Boolean
With grid2
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If .ValueMatrix(Row, 7) = 0 Then Exit Function
End With
validrow2 = True
End Function
Private Sub CellPos2(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col < 1 And grid2.TextMatrix(Row, 2) = "" Then
    grid2.Col = 1
ElseIf Col < 7 Then
    grid2.Col = 7
ElseIf Col < 8 Then
    grid2.Col = 8
ElseIf Row < grid2.Rows - 1 Then
    grid2.Select Row + 1, NextEmpty(grid2, Row + 1, 0, 7)
    grid2.ShowCell grid2.Row, 0
Else
    grid2.Select Row, Col
End If
End Sub
Private Sub myRemove2(Row As Long)
grid2.RemoveItem Row
CalcTotals2
End Sub
Private Sub grid2_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos2 KeyCode, Row, Col
End If
End Sub
Private Sub grid2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
End If
End Sub
Private Function CalcTotals2(Optional Row As Long = -1)
Dim nTotal As Double
With grid2
'For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
'    .TextMatrix(i, 8) = .ValueMatrix(i, 7) - .ValueMatrix(i, 4)
'    nTotal = mRound(nTotal + .ValueMatrix(i, 8))
'Next

End With
End Function

