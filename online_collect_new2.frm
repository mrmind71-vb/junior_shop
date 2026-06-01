VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form online_collect_newfrm2 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„œ›Ê⁄«  online"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20400
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
   ScaleWidth      =   20400
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   8325
      Width           =   9915
      Begin VB.TextBox xInvoice_Search 
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
         Left            =   90
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   180
         Width           =   1950
      End
      Begin VB.TextBox xShip_No_search 
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
         Left            =   6570
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   180
         Width           =   1950
      End
      Begin VB.Label Label7 
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
         Height          =   330
         Left            =   2160
         TabIndex        =   52
         Tag             =   "Color"
         Top             =   225
         Width           =   1050
      End
      Begin VB.Label Label6 
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
         Height          =   330
         Left            =   8640
         TabIndex        =   50
         Tag             =   "Color"
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   4455
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   1305
      Width           =   1950
      Begin Threed.SSCommand cmdDelRows 
         Height          =   600
         Left            =   45
         TabIndex        =   46
         Top             =   135
         Width           =   1860
         _ExtentX        =   3281
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
         Picture         =   "online_collect_new2.frx":0000
         Caption         =   "Õ–› «·”œ«œ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   6435
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   1305
      Width           =   2535
      Begin Threed.SSCommand cmdScv 
         Height          =   600
         Left            =   45
         TabIndex        =   42
         Top             =   135
         Width           =   2445
         _ExtentX        =   4313
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
         Picture         =   "online_collect_new2.frx":2434
         Caption         =   "› Õ „” ‰œ CSV"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6180
      Left            =   45
      TabIndex        =   38
      Top             =   2115
      Width           =   20310
      _ExtentX        =   35825
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
      TabCaption(0)   =   "»Ê«·’ «” —Ã«⁄ «Ê  »œÌ· »œÊ‰ ›« Ê—…"
      TabPicture(0)   =   "online_collect_new2.frx":4A07
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "”œ«œ ‘—ﬂ… «·‘Õ‰"
      TabPicture(1)   =   "online_collect_new2.frx":4A23
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   5730
         Left            =   -74955
         TabIndex        =   53
         Top             =   360
         Width           =   20175
         _cx             =   35586
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
         Cols            =   12
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
         Left            =   45
         TabIndex        =   54
         Top             =   360
         Width           =   20175
         _cx             =   35586
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
      TabIndex        =   30
      Top             =   8325
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   31
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
         Picture         =   "online_collect_new2.frx":4A3F
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect_new2.frx":6BE6
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   32
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
         Picture         =   "online_collect_new2.frx":8C2D
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect_new2.frx":AD18
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   33
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
         Picture         =   "online_collect_new2.frx":CD12
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect_new2.frx":EE23
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   34
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
         Picture         =   "online_collect_new2.frx":10E1D
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect_new2.frx":13041
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   13635
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   45
      Width           =   6630
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   5310
         TabIndex        =   26
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
         Picture         =   "online_collect_new2.frx":15112
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect_new2.frx":174DD
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   4005
         TabIndex        =   27
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
         Picture         =   "online_collect_new2.frx":19586
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect_new2.frx":1B58E
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   2655
         TabIndex        =   28
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
         Picture         =   "online_collect_new2.frx":1D545
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect_new2.frx":1FCE1
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   29
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
         Picture         =   "online_collect_new2.frx":22175
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   510
         Left            =   1350
         TabIndex        =   44
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
         Picture         =   "online_collect_new2.frx":24498
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
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   945
      Width           =   1455
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   45
         TabIndex        =   23
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
         Picture         =   "online_collect_new2.frx":268AC
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect_new2.frx":291D1
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   465
         Left            =   45
         TabIndex        =   24
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
         Picture         =   "online_collect_new2.frx":2BA25
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect_new2.frx":2DB85
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
      Left            =   16650
      TabIndex        =   13
      Top             =   8325
      Width           =   3705
      Begin MSDataListLib.DataCombo xYear 
         Height          =   330
         Left            =   1665
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   17
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
         TabIndex        =   16
         Tag             =   "Color"
         Top             =   225
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   10485
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   765
      Width           =   9780
      Begin VB.CommandButton cmdPayment_type 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   330
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   900
         Width           =   330
      End
      Begin VB.CommandButton cmdCompHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   330
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   540
         Width           =   330
      End
      Begin VB.CommandButton cmdBankHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   330
         Left            =   4770
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   540
         Width           =   330
      End
      Begin VB.TextBox xDesca 
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
         Left            =   4770
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   900
         Width           =   3930
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
         Left            =   6750
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1950
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
         Left            =   585
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   180
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo xBank 
         Height          =   330
         Left            =   5130
         TabIndex        =   2
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
      Begin MSDataListLib.DataCombo xShip 
         Height          =   330
         Left            =   540
         TabIndex        =   4
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
      Begin MSDataListLib.DataCombo xPayment_type 
         Height          =   330
         Left            =   540
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   900
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
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
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   945
         Width           =   960
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
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   39
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
         Left            =   8775
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   990
         Width           =   795
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Õ”«» »‰ﬂ"
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
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   585
         Width           =   765
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
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   11
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
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   795
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   375
      Left            =   4995
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
      Height          =   465
      Left            =   2700
      TabIndex        =   12
      Top             =   1620
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   820
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
      TabIndex        =   18
      Top             =   9720
      Width           =   20400
      _ExtentX        =   35983
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   21
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
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   3
         Left            =   11520
         TabIndex        =   43
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
      Height          =   510
      Left            =   13635
      TabIndex        =   35
      Top             =   8415
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   900
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
      Caption         =   "«Œ «— Õ”«» «·»‰ﬂ"
      TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
      ButtonStyle     =   3
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   4050
      Top             =   270
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   4455
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   40
      Top             =   9525
      Visible         =   0   'False
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   225
      Top             =   900
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc DATA13 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc data12 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
Attribute VB_Name = "online_collect_newfrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nType As Byte, bEdit As Boolean, sDoc_no As String
Public bCheck As Boolean
Public bDone As Boolean
Dim bEditRecord As Boolean, bAct As Boolean
Dim bStopCell As Boolean
Dim oSeachCode As New Search_abd, oSearchDoc As New Search_abd, oSearchProject As New Search, oSearchSup As New Search
Dim cFilter As String, cSelect As String
Dim con As New ADODB.Connection, CardTable As ADODB.Recordset
Dim cFile As String, cFileHeader As String, aBoxs As Variant
Dim clist1 As String, cList2 As String, cList3 As String, sDef As String
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1, Optional Row2 As Long = -1, Optional Row3 As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[BANK]", addstring(xBank.BoundText))
aInsert = AddFlag(aInsert, "[SHIP]", addstring(xShip.BoundText))
aInsert = AddFlag(aInsert, "[PAYMENT_TYPE]", addvalue(xPayment_type.BoundText))
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

If Row = -1 Then prog1.Visible = True

If (Row = -1 And Row2 = -1 And Row3 = -1) Or Row <> -1 Then
    myreplaceGrd Row
End If

If (Row = -1 And Row2 = -1 And Row3 = -1) Or Row2 <> -1 Then
    myreplaceGrd2 Row2
End If

con.CommitTrans
myreplace = True
Finally:
If Row = -1 Then prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
GoTo Finally:
End Function
Sub myProc(Optional sControl As String = "")
If ActiveControl.Name = grid1.Name Then
    Dim sMsg As String
'    If Not GrdDesc(oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0), "", "", grid1.Row, sMsg, IIf(grid1.ValueMatrix(grid1.Row, 8) = 0, Null, grid1.TextMatrix(grid1.Row, 8))) Then
'        If sMsg <> "" Then MsgBox sMsg
'        Exit Sub
'    End If
    
    Unload oSeachCode
    Grid1_AfterEdit grid1.Row, grid1.col
    CellPos 13, grid1.Row, grid1.col
ElseIf ActiveControl.Name = grid2.Name Then
    If Not grdDesc2(oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0), "", grid2.Row, sMsg, IIf(grid2.ValueMatrix(grid2.Row, 8) = 0, Null, grid2.TextMatrix(grid2.Row, 8))) Then
        If sMsg <> "" Then MsgBox sMsg
        Exit Sub
    End If
    Unload oSeachCode
    grid2_AfterEdit grid2.Row, grid2.col
    CellPos2 13, grid2.Row, grid2.col
ElseIf ActiveControl.Name = GRID3.Name Then
    If Not grdDesc3(oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0), GRID3.Row, sMsg) Then
        If sMsg <> "" Then MsgBox sMsg
        Exit Sub
    End If
    Unload oSeachCode
    grid3_AfterEdit GRID3.Row, GRID3.col
    CellPos3 13, GRID3.Row, GRID3.col
ElseIf ActiveControl.Name = cmdInform.Name Then
    openCardTable tbMode.tbFind, oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
ElseIf ActiveControl.Name = cmdBank.Name Then
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

Private Sub cmdBankEdit_Click()

End Sub

Private Sub cmdBankHeader_Click()
If Not xBank.MatchedWithList Then
    MsgBox "«·»‰ﬂ €Ì— „”Ã·"
    Exit Sub
End If
online_bank_headerfrm.sBank = xBank.BoundText
online_bank_headerfrm.Show 1
End Sub

Private Sub cmdCompEdit_Click()

End Sub

Private Sub cmdCompHeader_Click()
If Not xShip.MatchedWithList Then
    MsgBox "‘—ﬂ… «·‘Õ‰ €Ì— „”Ã·…"
    Exit Sub
End If
online_ship_headerfrm.sShip = xShip.BoundText
online_ship_headerfrm.Show 1
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
Private Sub cmdDelRows_Click()
If MsgBox("Õ–› ?", vbOKCancel + vbDefaultButton2 + vbCritical) <> vbOK Then Exit Sub

con.BeginTrans
On Error GoTo myerror
con.Execute "Delete  From " & cFile & " where Doc_No = " & addvalue(xDoc_No.text)
con.CommitTrans

If sDoc_no <> "" Then Exit Sub

If Not openCardTable(tbMode.tbFind, , xDoc_No.text) Then
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
Private Sub cmdExcel_Click()
Me.MousePointer = 11
Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aRow = AddFlag(aRow, aSub)

For i = 1 To grid1.Rows - 2
    If grid1.TextMatrix(i, 2) = "" Then
        aSub = AddFlag(Empty, "row", i)
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "word_wrap", False)
        aSub = AddFlag(aSub, "back_color", 19)
        aRow = AddFlag(aRow, aSub)
    End If
Next
ToFileExelNew grid2, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, "„” ‰œ ÿ·»Ì… —ﬁ„ : " & xDoc_No.text, "» «—ÌŒ : " & xDate.text)
Me.MousePointer = 0
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
cString = "SELECT ONLINE_PAY.Doc_No, Convert(Varchar(10),ONLINE_PAY.DATE,111),ONLINE_PAY.DESCA " & _
          " FROM ONLINE_PAY "
If pFilter <> "" Then cString = cString & " WHERE " & pFilter
Generalarray(1) = cString
Generalarray(2) = "order by ONLINE_PAY.DATE ,ONLINE_PAY.DOC_NO "
Generalarray(3) = 7000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%ONLINE_PAY.DESCA%%)"

listarray(1, 0) = " «—ÌŒ «·„” ‰œ"
listarray(1, 1) = "(##[Date]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1350

GrdArray(2, 0) = "«·»Ì«‰"
GrdArray(2, 1) = 8000

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

Private Sub cmdPayment_type_Click()
Dim oFlagfrm As New flag_mainfrm
Dim sBound As String
sBound = xPayment_type.BoundText
oFlagfrm.sTable = "payment_type"
oFlagfrm.sCaption = "‰Ê⁄ «·”œ«œ"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
DATA3.Recordset.Requery
xPayment_type.BoundText = sBound
If Not xPayment_type.MatchedWithList Then xPayment_type.BoundText = ""
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
Private Sub cmdSave_Click()
If Not myValid(True) Then Exit Sub
'If Not validRows Then Exit Sub
Me.MousePointer = vbHourglass
If Not myreplace Then Exit Sub
Me.MousePointer = vbNormal
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
'contlookup Me, oSearchSup, , , cmdCont.Tag <> ""
End Sub
Private Sub cmdScv_Click()
If Not myValid Then Exit Sub
If grid1.Rows > 2 Or grid2.Rows > 2 Then
    MsgBox "ÌÊÃœ »Ì«‰«  ›Ï «·„” ‰œ"
    Exit Sub
End If

onlineCollectCSVnewfrm.sDoc_no = xDoc_No.text
onlineCollectCSVnewfrm.sBank = xBank.BoundText
onlineCollectCSVnewfrm.sShip = xShip.BoundText
onlineCollectCSVnewfrm.Show 1

myLoadGrd
myloadgrd2

If grid1.Rows > 2 Then
    SSTab1.Tab = 1
ElseIf grid2.Rows > 2 Then
    SSTab1.Tab = 0
End If
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub cmdYear_Click()
Set oSearchYear.myForm = Me
oSearchYear.bEmpty = cmdYear.Tag <> ""
oSearchYear.Show 1
End Sub

Private Sub cmdWeb_Click()
If grid1.Rows > 2 Or grid2.Rows > 2 Then
    MsgBox "”œ«œ „”Ã· ›Ï «·„” ‰œ"
    Exit Sub
End If
If Not myValid Then Exit Sub

cmdWeb.Enabled = False
getDataWeb
SSTab1.Tab = 1
cmdWeb.Enabled = True
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xDoc_No.Tag = LoadMode Then
        If grid2.Rows > 2 Then
            SSTab1.Tab = 1
        Else
            SSTab1.Tab = 2
            grid1.SetFocus
        End If
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
cFile = "ONLINE_PAY_SUP"
cFileHeader = "ONLINE_PAY"

openCon con

Set DATA1.Recordset = mycmd("SELECT * FROM FILE5_10 WHERE ISSHIP = 1 ORDER BY DESCA", con)
Set xBank.RowSource = DATA1
xBank.ListField = "Desca"
xBank.BoundColumn = "Code"

Set DATA2.Recordset = mycmd("SELECT * FROM SHIP WHERE STOPED = 0 ORDER BY DESCA ", con)
Set xShip.RowSource = DATA2
xShip.ListField = "Desca"
xShip.BoundColumn = "Code"

Set DATA3.Recordset = mycmd("SELECT * FROM PAYMENT_TYPE ORDER BY DESCA ", con)
Set xPayment_type.RowSource = DATA3
xPayment_type.ListField = "Desca"
xPayment_type.BoundColumn = "Code"

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
Private Function myValid(Optional bMsg As Boolean = False) As Boolean
If Trim(xDoc_No.text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Not xBank.MatchedWithList Then
    MsgBox "«·»‰ﬂ €Ì— „”Ã·"
    Exit Function
End If

myValid = True
End Function
Private Function validRows(Optional bMsg As Boolean = True) As Boolean
Dim i As Long
For i = 1 To grid1.Rows - 2
    If Not validRow(i, bMsg) Then
        Exit Function
    End If
Next
validRows = True
End Function
Private Sub myload()
xDoc_No.text = CardTable!doc_no
xDate.text = myFormat_p(CardTable!Date)
xBank.BoundText = CardTable!BANK & ""
xShip.BoundText = CardTable!SHIP & ""
xPayment_type.BoundText = CardTable!payment_type & ""
panel1(2).Caption = CardTable!UserName & " " & myFormat_p(CardTable!Time, True)
panel1(3).Caption = CardTable!USERNAME2 & " " & myFormat_p(CardTable!TIME2, True)

Handlecontrols LoadMode
myLoadGrd
myloadgrd2
CalcTotals
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
'CellPos2 13, GRID2.Rows - 2, GRID2.Cols - 1

On Error Resume Next
If grid2.Rows > 2 Then
    SSTab1.Tab = 0
    grid2.SetFocus
Else
    SSTab1.Tab = 1
    grid1.SetFocus
End If
Err.Clear
End Sub
Private Sub myDefine()
xDoc_No.text = Newflag(cFileHeader, "DOC_NO", con)
xDate.text = myFormat_p(Date)
xBank.BoundText = ""
xShip.BoundText = ""
xPayment_type.BoundText = ""
xDesca.text = ""

grid1.Rows = 1
myAddItem
Fixgrd

grid2.Rows = 1
myAddItem2
Fixgrd2

Handlecontrols DefineMode
'StatusBar1.Panels(2).text = "·« ÌÊÃœ ”œ«œ"
'StatusBar1.Panels(3).text = ""
'StatusBar1.Panels(4).text = ""

CellPos 13, grid1.Rows - 2, grid1.Cols - 1

On Error Resume Next
If SSTab1.Tab = 0 Then
    GRID3.SetFocus
ElseIf SSTab1.Tab = 1 Then
    grid2.SetFocus
Else
    grid1.SetFocus
End If
Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit
cmdFilter.Visible = cmdFilter.Tag <> ""
cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmddel.Enabled = nMode = LoadMode And bEditRecord
cmdSave.Enabled = bEditRecord
cmddel.Enabled = nMode = LoadMode And bEditRecord
cmdDelRows.Enabled = nMode = LoadMode And bEditRecord
cmdScv.Enabled = nMode = LoadMode And bEditRecord

Dim nRecord As Long, nRecords As Long
retRecords xDoc_No.text, nRecords, nRecord

If nMode = LoadMode Then
    panel1(1).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
     panel1(1).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""

xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub

Private Sub Text1_Change()

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
Private Function CalcTotals()
Dim nTotal As Double
Dim i As Long
With grid1
For i = 1 To grid1.Rows - 2
    .TextMatrix(i, 7) = .ValueMatrix(i, 5) - .ValueMatrix(i, 6)
    nTotal = mRound(nTotal + .ValueMatrix(i, 7))
Next
End With

With grid2
For i = 1 To .Rows - 2
    .TextMatrix(i, 7) = .ValueMatrix(i, 5) - .ValueMatrix(i, 6)
    nTotal = mRound(nTotal + .ValueMatrix(i, 7))
Next
End With


panel1(0).Caption = "≈Ã„«·Ì «·„” ‰œ : " & nTotal
End Function
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_no As String = "", Optional bDefine As Boolean = True) As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_no) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

'On Error GoTo myError
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
    cWhere = "Doc_no = " & pDoc_no
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "Doc_no < " & pDoc_no
    cOrder = "Order by Doc_no desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "Doc_no > " & pDoc_no
    cOrder = "Order by Doc_no"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

If cFilter <> "" Then
    cString = cString & IIf(cWhere = "", " WHERE ", " AND ") & cFilter
End If

cString = cString & " " & cOrder

Set CardTable = mycmd(cString, con, adText)

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
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_no <> "" Then
    cString = "SELECT Count(*) AS records,SUM(CASE WHEN DOC_NO <= " & pDoc_no & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM " & cFileHeader & Tr(cFilter, " WHERE ") & cFilter

Set loctable = mycmd(cString, con)
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
        If Row = -1 Then prog1.Value = Round(i / (.Rows - 1), 2) * 100
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "PAYMENT_ID", addstring(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "ORDER_NO", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "[TOTAL]", grid1.ValueMatrix(i, 5))
        aInsert = AddFlag(aInsert, "[COMMISSION]", grid1.ValueMatrix(i, 6))
        aInsert = AddFlag(aInsert, "[TYPE]", "1")
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, cFile)
        Else
            con.Execute addUpdate(aInsert, cFile, "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
Dim i As Long
'For I = 1 To grid1.Rows - 2
'    If I <> Row Then
'        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
'            cmdSave_Click
'            Exit Sub
'        End If
'    End If
'Next

If Not myValid(True) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myLoadGrd
    If Row < grid1.Rows - 1 Then
        grid1.Select Row, col
    Else
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If

If Not validRow(Row, False) Then
    CalcTotals
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
        grid1.ShowCell grid1.Rows - 1, grid1.Cols - 1
    Else
        CalcTotals
    End If
Else
    myLoadGrd
End If
End With
End Sub
Private Sub grid1_EnterCell()
If Not bEditRecord Then
    grid1.Editable = flexEDNone
ElseIf grid2.Row > 2 Then
    grid1.Editable = flexEDNone
ElseIf ((grid1.col = 0 Or grid1.col = 1 Or grid1.col = 2) And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Or grid1.col = 5 Or grid1.col = 6 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub myLoadGrd()
Dim aFields(9)
aFields(0) = "FILE6_90H.PAYMENT_ID"
aFields(1) = "ONLINE_PAY_SUP.ORDER_NO"
aFields(2) = "FILE6_90H.[NAME]"
aFields(3) = "FILE6_90H.[PHONE]"
aFields(4) = "vw_online_pay_invoices.[DOC_NO]"
aFields(5) = "vw_online_pay_invoices.[TOTAL]"

aFields(6) = "ONLINE_PAY_SUP.[TOTAL]"
aFields(7) = "ONLINE_PAY_SUP.[COMMISSION]"
aFields(8) = "ONLINE_PAY_SUP.[VALUE]"
aFields(9) = "ONLINE_PAY_SUP.ID"

Dim cString As String
cString = "SELECT " & arString(aFields, ",") & _
         " FROM ONLINE_PAY_SUP " & _
         " INNER JOIN FILE6_90H ON ONLINE_PAY_SUP.ORDER_NO = FILE6_90H.DOC_NO" & _
         " LEFT JOIN vw_online_pay_invoices ON ONLINE_PAY_SUP.ID = vw_online_pay_invoices.ID" & _
         " WHERE ONLINE_PAY_SUP.DOC_NO = " & addvalue(xDoc_No.text) & _
         " AND ONLINE_PAY_SUP.TYPE = 1"

Set grid1.DataSource = myRs(cString, con)
myAddItem
CalcTotals
Fixgrd
End Sub
Private Sub grid1_GotFocus()
grid1_EnterCell
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
ElseIf Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
        myRemove grid1.Row
    ElseIf MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
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
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
Dim sMsg As String
If col = 0 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
    ElseIf Not GrdDesc(grid1.EditText, "", Row, sMsg) Then
        'MsgBox sMsg
        Cancel = True
    End If
ElseIf col = 1 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
    ElseIf Not GrdDesc(grid1.EditText, "", Row, sMsg) Then
        'MsgBox sMsg
        Cancel = True
    End If
End If
End Sub
Private Sub Fixgrd()
With grid1
.RowHeight(0) = 700
.TextMatrix(0, 0) = "„⁄—› «·”œ«œ"
.TextMatrix(0, 1) = "—ﬁ„ «·ÿ·»Ì…"
.TextMatrix(0, 2) = "«·«”„"
.TextMatrix(0, 3) = "«· ·Ì›Ê‰"
.TextMatrix(0, 4) = "«·›« Ê—…"
.TextMatrix(0, 5) = "≈Ã„«·Ì «·›« Ê—…"
.TextMatrix(0, 6) = "«·„”œœ"
.TextMatrix(0, 7) = "«·⁄„Ê·…"
.TextMatrix(0, 8) = "«·’«›Ì"
.TextMatrix(0, 9) = "ID"

.RowHeight(0) = 600
.ColWidth(0) = 4000
.ColWidth(1) = 1500
.ColWidth(2) = 4000
.ColWidth(3) = 2000
.ColWidth(4) = 2000
.ColWidth(5) = 1400
.ColWidth(6) = 1300
.ColWidth(7) = 1300
.ColWidth(8) = 1300

.ColHidden(.Cols - 1) = True

'.ColComboList(2) = "..."
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
If grid1.Rows > 1 Then
    .Cell(flexcpFontSize, 1, 0, .Rows - 1, 0) = 10
End If
End With
End Sub
Private Sub myAddItem()
With grid1
.AddItem ""
.Cell(flexcpFontSize, 1, 0, .Rows - 1, 0) = 9
If grid1.Rows > 2 Then
    '.TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
End If
End With
End Sub
Private Function validRow(Row As Long, Optional bMsg As Boolean = False) As Boolean
With grid1
If Trim(.TextMatrix(Row, 0)) = "" Then
    If bMsg Then
        MsgBox "—ﬁ„ »Ê·Ì’… «·‘Õ‰ €Ì— „”Ã·"
    End If
    Exit Function
End If

If .ValueMatrix(Row, 5) = 0 And .ValueMatrix(Row, 6) = 0 Then
    If bMsg Then
        MsgBox "·« ÌÊÃœ ⁄„Ê·… «Ê ”œ«œ"
    End If
    Exit Function
End If
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf col = 0 Then
    If grid1.TextMatrix(Row, col) = "" Then
        grid1.Select Row, 1
    Else
        grid1.Select Row, 5
    End If
ElseIf col = 1 Then
    grid1.Select Row, 5
ElseIf col < grid1.Cols - 3 Then
    grid1.col = col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 5)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, col
End If
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
CalcTotals
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If col = 2 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    'If Col = 2 And (grid1.TextMatrix(Row, Col) = "" Or grid1.TextMatrix(Row, grid1.Cols - 1) = "") Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid1.col = 2 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    'If grid1.Col = 1 And (grid1.TextMatrix(grid1.Row, grid1.Col) = "" Or grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Then Exit Sub
    'If grid1.Col = 2 And (grid1.TextMatrix(grid1.Row, grid1.Col) = "" Or grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Function GrdDesc(sPayment_id As String, ByVal sOrder_No As String, Row As Long, ByRef pMsg As String) As Boolean
If sOrder_No = "" Then
    sOrder_No = rsFunc("[dbo].[fn_order_collect]", con, addstring(sPayment_id), "default")
    If sOrder_No = "" Then
        MsgBox "—ﬁ„ " & "Payment id" & " €Ì— ’ÕÌÕ"
        Exit Function
    End If
End If


Dim loctable As New ADODB.Recordset
Set loctable = myRs("SELECT * " & _
                    " FROM FILE6_90H " & _
                    " WHERE DOC_NO = " & MyParn(sOrder_No), con)

If loctable.EOF Then
    MsgBox "—ﬁ„ ÿ·»»… €Ì— ’ÕÌÕ"
    Exit Function
End If

grid1.TextMatrix(Row, 0) = loctable!PAYMENT_ID & ""
grid1.TextMatrix(Row, 1) = loctable!doc_no & ""
grid1.TextMatrix(Row, 2) = loctable!Name & ""
grid1.TextMatrix(Row, 3) = loctable!phone & ""

Set loctable = New ADODB.Recordset
Set loctable = myRs("sp_online_data", con, , adStoredProc, AddFlag(aPrm, "PAYMENT_ID", grid1.TextMatrix(grid1.Row, 0)))
If Not loctable.EOF Then
    grid1.TextMatrix(Row, 4) = loctable!doc_no
    grid1.TextMatrix(Row, 5) = loctable!total
End If
Set loctable = Nothing
GrdDesc = True
End Function
Private Function grdDesc2(ByVal sOrder_No As String, sShip_no As String, Row As Long, ByRef pMsg As String) As Boolean
If sOrder_No = "" Then
    sOrder_No = rsFunc("[dbo].[fn_order_collect]", con, "default", addstring(sShip_no))
    If sOrder_No = "" Then
        MsgBox "—ﬁ„ »Ê·Ì’… «·‘Õ‰ €Ì— ’ÕÌÕ"
    End If
End If

Dim loctable As New ADODB.Recordset
Set loctable = myRs("SELECT * " & _
                    " FROM FILE6_90H WHERE DOC_NO = " & MyParn(sOrder_No), con)
       
grid2.TextMatrix(Row, 0) = loctable!PAYMENT_ID
grid2.TextMatrix(Row, 1) = loctable!doc_no & ""
grid2.TextMatrix(Row, 2) = loctable!Name
grid2.TextMatrix(Row, 3) = loctable!phone & ""
grdDesc2 = True
End Function
Private Sub myreplaceGrd2(Row As Long)
Dim aInsert As Variant
With grid2
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, .Rows - 2, Row)
        If Row = -1 Then prog1.Value = Round(i / (.Rows - 1), 2) * 100
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "SHIP_NO", addstring(.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "ORDER_NO", addstring(.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "[COD]", .ValueMatrix(i, 5))
        aInsert = AddFlag(aInsert, "[FEE]", .ValueMatrix(i, 6))
        aInsert = AddFlag(aInsert, "[TYPE]", "2")
        If .TextMatrix(i, .Cols - 1) = "" Then
            con.Execute addInsert(aInsert, cFile)
        Else
            con.Execute addUpdate(aInsert, cFile, "ID = " & .TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal col As Long)
Dim i As Long
'For I = 1 To grid2.Rows - 2
'    If I <> Row Then
'        If grid2.TextMatrix(I, grid2.Cols - 1) = "" Then
'            cmdSave_Click
'            Exit Sub
'        End If
'    End If
'Next

If Not myValid Then
    On Error Resume Next
    grid2.SetFocus
    Err.Clear
    myloadgrd2
    If Row < grid2.Rows - 1 Then
        grid2.Select Row, col
    Else
        CellPos2 13, grid2.Rows - 2, grid2.Cols - 1
    End If
    Exit Sub
End If

If Not validrow2(Row, False) Then
    CalcTotals
    Exit Sub
End If

With grid2
If Row = grid2.Rows - 1 Then
    myAddItem2
End If

If myreplace(, Row) Then
    If xDoc_No.Tag = DefineMode Then
        Handlecontrols LoadMode
        myloadgrd2
    ElseIf grid2.TextMatrix(Row, grid2.Cols - 1) = "" Then
        myloadgrd2
    Else
        CalcTotals
    End If
Else
    myloadgrd2
End If
End With
End Sub
Private Sub grid2_EnterCell()
If Not bEditRecord Then
    grid2.Editable = flexEDNone
ElseIf grid1.Row > 2 Then
    grid2.Editable = flexEDNone
ElseIf ((grid2.col = 0 Or grid2.col = 1) And grid2.TextMatrix(grid2.Row, grid2.Cols - 1) = "") Or grid2.col = 5 Or grid2.col = 6 Then
    grid2.Editable = flexEDKbdMouse
Else
    grid2.Editable = flexEDNone
End If
End Sub
Private Sub myloadgrd2()
Dim cString As String
Dim aFields(9)
aFields(0) = "ONLINE_PAY_SUP.ORDER_NO"
aFields(1) = "ONLINE_PAY_SUP.SHIP_NO"
aFields(2) = "FILE6_90H.[NAME]"
aFields(3) = "FILE6_90H.[PHONE]"
aFields(4) = "vw_online_pay_invoices.[DOC_NO]"
aFields(5) = "vw_online_pay_invoices.[TOTAL]"
aFields(6) = "ONLINE_PAY_SUP.COD"
aFields(7) = "ONLINE_PAY_SUP.FEE"
aFields(8) = "ONLINE_PAY_SUP.[VALUE]"
aFields(9) = "ONLINE_PAY_SUP.ID"

cString = "SELECT " & arString(aFields, ",") & _
         " FROM ONLINE_PAY_SUP " & _
         " INNER JOIN FILE6_90H ON ONLINE_PAY_SUP.ORDER_NO = FILE6_90H.DOC_NO" & _
         " LEFT JOIN vw_online_pay_invoices ON ONLINE_PAY_SUP.ID = vw_online_pay_invoices.ID" & _
         " WHERE ONLINE_PAY_SUP.DOC_NO = " & addvalue(xDoc_No.text) & _
         " AND ONLINE_PAY_SUP.TYPE = 2"

Set data11.Recordset = mycmd(cString, con)

myAddItem2

CalcTotals
Fixgrd2
End Sub
Private Sub GRID2_GotFocus()
grid2_EnterCell
End Sub
Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos2 KeyCode, grid2.Row, grid2.col
ElseIf Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 46 And grid2.Row <> grid2.Rows - 1 Then
    If grid2.TextMatrix(grid2.Row, grid2.Cols - 1) = "" Then
        myRemove2 grid2.Row
    ElseIf MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
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
Private Sub GRID2_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
With grid2
Dim sMsg As String
If col = 1 Then
    If Trim(grid2.EditText) = "" Then
        Cancel = True
    ElseIf Not grdDesc2(grid2.EditText, "", Row, sMsg) Then
        MsgBox sMsg
        Cancel = True
    End If
End If
End With
End Sub
Private Sub Fixgrd2()
With grid2
.TextMatrix(0, 0) = "»Ê·Ì’… «·‘Õ‰"
.TextMatrix(0, 1) = "—ﬁ„ «·ÿ·»Ì…"
.TextMatrix(0, 2) = "«·«”„"
.TextMatrix(0, 3) = "«· ·Ì›Ê‰"
.TextMatrix(0, 4) = "«·›« Ê—…"
.TextMatrix(0, 5) = "≈Ã„«·Ì «·›« Ê—…"
.TextMatrix(0, 6) = "COD"
.TextMatrix(0, 7) = "ﬁÌ„… «·‘Õ‰"
.TextMatrix(0, 8) = "’«›Ì"
.TextMatrix(0, 9) = "ID"

.RowHeight(0) = 600
.ColWidth(0) = 2000
.ColWidth(1) = 2000
.ColWidth(2) = 5000
.ColWidth(3) = 3500
.ColWidth(4) = 1500
.ColWidth(5) = 1500
.ColWidth(6) = 1500
.ColWidth(7) = 1500
.ColHidden(.Cols - 1) = True
For i = 1 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
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
Private Sub myRemove2(Row As Long)
grid2.RemoveItem Row
CalcTotals
End Sub
Private Function validrow2(Row As Long, Optional bMsg As Boolean = False) As Boolean
With grid2
If Trim(.TextMatrix(Row, 1)) = "" Then
    If bMsg Then
        MsgBox "—ﬁ„ «·ÿ·»Ì… €Ì— „”Ã·"
    End If
    Exit Function
End If

If Trim(.TextMatrix(Row, 2)) = "" Then
    If bMsg Then
        MsgBox "—ﬁ„ «·›« Ê—… €Ì— „”Ã·"
    End If
    Exit Function
End If

If .ValueMatrix(Row, 5) = 0 And .ValueMatrix(Row, 6) = 0 Then
    If bMsg Then
        MsgBox "·« ÌÊÃœ COD «Ê ﬁÌ„… ‘Õ‰"
    End If
    Exit Function
End If
End With
validrow2 = True
End Function
Private Sub CellPos2(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf col = 0 Then
    If NextEmpty(grid2, Row, col, 8) <= 2 Then
        If grid2.TextMatrix(Row, col) = "" Then
            grid2.Select Row, 1
        Else
            grid2.Select Row, 2
        End If
    Else
        grid2.Select Row, NextEmpty(grid2, Row, col, 8)
    End If
ElseIf col = 1 Then
    grid2.Select Row, NextEmpty(grid2, Row, col + 1, 8)
ElseIf col < 8 Then
    grid2.col = 8
ElseIf col < grid2.Cols - 2 Then
    grid2.col = col + 1
ElseIf Row < grid2.Rows - 1 Then
    grid2.Select Row + 1, NextEmpty(grid2, Row + 1, 1, 7)
    grid2.ShowCell grid2.Row, 0
Else
    grid2.Select Row, col
End If
End Sub
Private Sub GRID2_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos2 KeyCode, Row, col
End If
End Sub
Private Sub GRID2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
End If
End Sub
Private Sub xInvoice_Search_Change()
SearchGrd grid1, LCase(Trim(xInvoice_Search.text)), 2, False
End Sub
Private Sub xShip_No_search_Change()
SearchGrd grid1, LCase(Trim(xShip_No_search.text)), 1, False
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
Private Sub xdesca_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xbank_GotFocus()
myGotFocus xBank
End Sub
Private Sub xbank_LostFocus()
myLostFocus xBank
If Not xBank.MatchedWithList Then xBank.BoundText = ""
End Sub
Private Sub xship_GotFocus()
myGotFocus xShip
End Sub
Private Sub xship_LostFocus()
myLostFocus xShip
If Not xShip.MatchedWithList Then xShip.BoundText = ""
End Sub
Private Sub xPayment_type_GotFocus()
myGotFocus xPayment_type
End Sub
Private Sub xPayment_type_LostFocus()
myLostFocus xPayment_type
If Not xPayment_type.MatchedWithList Then xPayment_type.BoundText = ""
End Sub

