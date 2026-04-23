VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form online_collectfrm 
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
   Begin VB.Frame Frame4 
      Height          =   780
      Left            =   765
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   1260
      Width           =   6180
      Begin Threed.SSCommand cmdWeb 
         Height          =   600
         Left            =   3285
         TabIndex        =   37
         Top             =   135
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   1058
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "”Õ» ”œ«œ „‰ Œ·«· «·„Êﬁ⁄"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect8.frx":0000
      End
      Begin Threed.SSCommand cmdShipping 
         Height          =   600
         Left            =   45
         TabIndex        =   38
         Top             =   135
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   1058
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "ﬂ‘› Õ”«» ‘—ﬂ… «·‘Õ‰"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect8.frx":22F6
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6180
      Left            =   45
      TabIndex        =   34
      Top             =   2115
      Width           =   20310
      _ExtentX        =   35825
      _ExtentY        =   10901
      _Version        =   393216
      Tab             =   1
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
      TabPicture(0)   =   "online_collect8.frx":45EC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "GRID3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "”œ«œ ‘—ﬂ… «·‘Õ‰"
      TabPicture(1)   =   "online_collect8.frx":4608
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "”œ«œ „‰ Œ·«· «·„Êﬁ⁄"
      TabPicture(2)   =   "online_collect8.frx":4624
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grid1"
      Tab(2).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   5730
         Left            =   -74910
         TabIndex        =   40
         Top             =   360
         Width           =   20130
         _cx             =   35507
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
      Begin VSFlex7Ctl.VSFlexGrid GRID3 
         Height          =   5730
         Left            =   -74910
         TabIndex        =   41
         Top             =   360
         Width           =   20130
         _cx             =   35507
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
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   5730
         Left            =   90
         TabIndex        =   42
         Top             =   360
         Width           =   20130
         _cx             =   35507
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
      TabIndex        =   26
      Top             =   8325
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
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
         Picture         =   "online_collect8.frx":4640
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect8.frx":67E7
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
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
         Picture         =   "online_collect8.frx":882E
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect8.frx":A919
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   29
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
         Picture         =   "online_collect8.frx":C913
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect8.frx":EA24
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   30
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
         Picture         =   "online_collect8.frx":10A1E
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_collect8.frx":12C42
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   14940
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   45
      Width           =   5325
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4005
         TabIndex        =   22
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
         Picture         =   "online_collect8.frx":14D13
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect8.frx":170DE
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   2700
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
         Picture         =   "online_collect8.frx":19187
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect8.frx":1B18F
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1350
         TabIndex        =   24
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
         Picture         =   "online_collect8.frx":1D146
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect8.frx":1F8E2
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   25
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
         Picture         =   "online_collect8.frx":21D76
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
      Left            =   10845
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   945
      Width           =   1455
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   45
         TabIndex        =   19
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
         Picture         =   "online_collect8.frx":24099
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect8.frx":269BE
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   465
         Left            =   45
         TabIndex        =   20
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
         Picture         =   "online_collect8.frx":29212
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_collect8.frx":2B372
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
      TabIndex        =   9
      Top             =   8325
      Width           =   3705
      Begin MSDataListLib.DataCombo xYear 
         Height          =   330
         Left            =   1665
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   13
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
         TabIndex        =   12
         Tag             =   "Color"
         Top             =   225
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   12330
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   765
      Width           =   7935
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
         Left            =   135
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   3
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
         TabIndex        =   4
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
      Begin MSDataListLib.DataCombo xShip 
         Height          =   330
         Left            =   135
         TabIndex        =   2
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
         TabIndex        =   35
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
         TabIndex        =   33
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
         Left            =   6885
         RightToLeft     =   -1  'True
         TabIndex        =   32
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
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   7
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
         TabIndex        =   6
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
      Height          =   375
      Left            =   13095
      TabIndex        =   8
      Top             =   315
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
      TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   17
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
      Left            =   11790
      TabIndex        =   31
      Top             =   8415
      Width           =   4830
      _ExtentX        =   8520
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
      Left            =   1665
      Top             =   1215
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
      TabIndex        =   39
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
      Left            =   270
      Top             =   540
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
   Begin Threed.SSCommand cmdScv 
      Height          =   555
      Left            =   8775
      TabIndex        =   43
      Top             =   1485
      Width           =   2040
      _ExtentX        =   3598
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
      Picture         =   "online_collect8.frx":2D65F
      Caption         =   "› Õ „” ‰œ CSV"
      ButtonStyle     =   3
      PictureAlignment=   10
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "online_collect8.frx":2FC32
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
Dim oSeachCode As New Search_abd, oSearchDoc As New Search_abd, oSearchProject As New Search, oSearchSup As New Search
Dim cFilter As String, cSelect As String
Dim con As New ADODB.Connection, CardTable As ADODB.Recordset
Dim cFile As String, cFileHeader As String, aBoxs As Variant
Dim cList1 As String, cList2 As String, cList3 As String, sDef As String
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1, Optional Row2 As Long = -1, Optional Row3 As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[BANK]", addstring(xBank.BoundText))
aInsert = AddFlag(aInsert, "[SHIP]", addstring(xShip.BoundText))
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

If (Row = -1 And Row2 = -1 And Row3 = -1) Or Row <> -1 Then
    myreplaceGrd Row
End If

If (Row = -1 And Row2 = -1 And Row3 = -1) Or Row2 <> -1 Then
    myreplaceGrd2 Row2
End If

If (Row = -1 And Row2 = -1 And Row3 = -1) Or Row3 <> -1 Then
    myreplaceGrd3 Row3
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
    Dim sMsg As String
    If Not GrdDesc(oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0), "", "", grid1.Row, sMsg, IIf(grid1.ValueMatrix(grid1.Row, 8) = 0, Null, grid1.TextMatrix(grid1.Row, 8))) Then
        If sMsg <> "" Then MsgBox sMsg
        Exit Sub
    End If
    
    Unload oSeachCode
    Grid1_AfterEdit grid1.Row, grid1.Col
    CellPos 13, grid1.Row, grid1.Col
ElseIf ActiveControl.Name = grid2.Name Then
    If Not grdDesc2(oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0), "", grid2.Row, sMsg, IIf(grid2.ValueMatrix(grid2.Row, 8) = 0, Null, grid2.TextMatrix(grid2.Row, 8))) Then
        If sMsg <> "" Then MsgBox sMsg
        Exit Sub
    End If
    Unload oSeachCode
    grid2_AfterEdit grid2.Row, grid2.Col
    CellPos2 13, grid2.Row, grid2.Col
ElseIf ActiveControl.Name = GRID3.Name Then
    If Not grdDesc3(oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0), GRID3.Row, sMsg) Then
        If sMsg <> "" Then MsgBox sMsg
        Exit Sub
    End If
    Unload oSeachCode
    grid3_AfterEdit GRID3.Row, GRID3.Col
    CellPos3 13, GRID3.Row, GRID3.Col
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
cString = "SELECT FILE9_10H.Doc_No, Convert(Varchar(10),FILE9_10H.DATE,111),FILE9_10H.DESCA " & _
          " FROM FILE9_10H "
If pFilter <> "" Then cString = cString & " WHERE " & pFilter
Generalarray(1) = cString
Generalarray(2) = "order by FILE9_10H.DATE ,FILE9_10H.DOC_NO "
Generalarray(3) = 7000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%FILE9_10H.DESCA%%)"

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
'contlookup Me, oSearchSup, , , cmdCont.Tag <> ""
End Sub
Private Sub cmdScv_Click()
If Not myValid Then Exit Sub
If grid1.Rows > 2 Or grid2.Rows > 2 Or GRID3.Rows > 2 Then
    MsgBox "ÌÊÃœ »Ì«‰«  ›Ï «·”Ã·"
    Exit Sub
End If
onlinecollectCSVfrm.Show 1
End Sub

Private Sub cmdShipping_Click()
If grid1.Rows > 2 Or grid2.Rows > 2 Then
    MsgBox "”œ«œ „”Ã· ›Ï «·„” ‰œ"
    Exit Sub
End If

If Not xShip.MatchedWithList Then
    MsgBox "«·‘—ﬂ… €Ì— „”Ã·…"
    Exit Sub
End If

If Not myValid Then Exit Sub

cmdShipping.Enabled = False
getDataShip
cmdShipping.Enabled = True
SSTab1.Tab = 0
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub cmdYear_Click()
Set oSearchYear.myform = Me
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
Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xDoc_No.Tag = LoadMode Then
        If GRID3.Rows > 2 Then
            SSTab1.Tab = 0
            GRID3.SetFocus
        ElseIf grid2.Rows > 2 Then
            SSTab1.Tab = 1
            grid2.SetFocus
        Else
            SSTab1.Tab = 0
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
cFile = "FILE9_10"
cFileHeader = "FILE9_10H"

openCon con

Set DATA1.Recordset = myCmd("SELECT * FROM FILE5_10 WHERE ISSHIP = 1 ORDER BY DESCA", con)
Set xBank.RowSource = DATA1
xBank.ListField = "Desca"
xBank.BoundColumn = "Code"

Set DATA2.Recordset = myCmd("SELECT * FROM SHIP WHERE STOPED = 0 ORDER BY DESCA ", con)
Set xShip.RowSource = DATA2
xShip.ListField = "Desca"
xShip.BoundColumn = "Code"

Set grid1.DataSource = data10
Set grid2.DataSource = data11
Set GRID3.DataSource = data12

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

If grid2.Rows > 2 Then
    If Not xShip.MatchedWithList Then
        MsgBox "«·‘—ﬂ… €Ì— „”Ã·…"
        Exit Function
    End If
End If

Dim i As Long
If bMsg Then
    For i = 1 To grid1.Rows - 2
        If Not validRow(i, True) Then
            Exit Function
        End If
    Next
End If

If bMsg Then
    For i = 1 To grid2.Rows - 2
        If Not validrow2(i) Then
            Exit Function
        End If
    Next
End If

If bMsg Then
    For i = 1 To GRID3.Rows - 2
        If Not validrow3(i) Then
            Exit Function
        End If
    Next
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
panel1(0).Caption = CardTable!UserName & " " & myFormat_p(CardTable!Time, True)
panel1(1).Caption = CardTable!USERNAME2 & " " & myFormat_p(CardTable!TIME2, True)
Handlecontrols LoadMode
myLoadGrd
myloadgrd2
myloadgrd3
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
CellPos2 13, grid2.Rows - 2, grid2.Cols - 1
CellPos3 13, GRID3.Rows - 2, GRID3.Cols - 1

On Error Resume Next
If GRID3.Rows > 2 Then
    SSTab1.Tab = 0
    GRID3.SetFocus
ElseIf grid2.Rows > 2 Then
    SSTab1.Tab = 1
    grid2.SetFocus
Else
    SSTab1.Tab = 2
    grid1.SetFocus
End If
Err.Clear
End Sub
Private Sub myDefine()
xDoc_No.text = Newflag(cFileHeader, "DOC_NO", con)
xDate.text = myFormat_p(Date)
xBank.BoundText = ""
xShip.BoundText = ""
xDesca.text = ""

grid1.Rows = 1
myAddItem
fixGrd

grid2.Rows = 1
myAddItem2
Fixgrd2

GRID3.Rows = 1
myAddItem3
Fixgrd3

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
Private Function CalcTotals(Optional Row As Long = -1)
Dim nTotal As Double
With grid1

For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
    .TextMatrix(i, 10) = .ValueMatrix(i, 8) - .ValueMatrix(i, 9)
    nTotal = mRound(nTotal + .ValueMatrix(i, 10))
Next

'If Row = 0 Then Exit Function
'.TextMatrix(Row, 6) = ""
'.TextMatrix(Row, 7) = ""
End With
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
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_no <> "" Then
    cString = "SELECT Count(*) AS records,SUM(CASE WHEN DOC_NO <= " & pDoc_no & " THEN 1 ELSE 0 END) AS record"
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
        aInsert = AddFlag(aInsert, "INV_NO", addstring(grid1.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "[TOTAL]", grid1.ValueMatrix(i, 8))
        aInsert = AddFlag(aInsert, "[COMMISSION]", grid1.ValueMatrix(i, 9))
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
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim cWhere As String
If grid1.TextMatrix(Row, 0) <> "" Then
    cWhere = cWhere & " AND FILE6_20H.PAYMENT_ID = " & MyParn(grid1.TextMatrix(Row, 0))
End If
    
If grid1.TextMatrix(Row, 1) <> "" Then
    cWhere = " AND FILE6_20H.SHIP_NO = " & MyParn(grid1.TextMatrix(Row, 1))
End If
    
invLookup Me, oSeachCode, Mid(cWhere, 6)
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i As Long
For i = 1 To grid1.Rows - 2
    If i <> Row Then
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            cmdSave_Click
            Exit Sub
        End If
    End If
Next

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

If Not validRow(Row, False) Then
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
ElseIf grid2.Row > 2 Then
    grid1.Editable = flexEDNone
ElseIf ((grid1.Col = 0 Or grid1.Col = 1 Or grid1.Col = 2) And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Or grid1.Col = 8 Or grid1.Col = 9 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub myLoadGrd()
Dim cString As String
Dim aFields(11)
aFields(0) = "FILE6_20H.PAYMENT_ID"
aFields(1) = "FILE6_20H.SHIP_NO"
aFields(2) = "FILE9_10.INV_NO"
aFields(3) = "FILE6_20H.DATE"
aFields(4) = "FILE6_20H.[NAME]"
aFields(5) = "FILE6_20H.[PHONE]"
aFields(6) = "(FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT)"
aFields(7) = "(FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT) + FILE6_20H.CHARGE1 + FILE6_20H.CHARGE2"
aFields(8) = "FILE9_10.[TOTAL]"
aFields(9) = "FILE9_10.[COMMISSION]"
aFields(10) = "FILE9_10.[VALUE]"
aFields(11) = "FILE9_10.ID"

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
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim sMsg As String
If Col = 0 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
    ElseIf Not GrdDesc("", "", grid1.EditText, Row, sMsg) Then
        MsgBox sMsg
        Cancel = True
    End If
ElseIf Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
    ElseIf Not GrdDesc("", grid1.EditText, "", Row, sMsg) Then
        MsgBox sMsg
        Cancel = True
    End If
End If
End Sub
Private Sub fixGrd()
With grid1
.TextMatrix(0, 0) = "Payment id"
.TextMatrix(0, 1) = "»Ê·Ì’… «·‘Õ‰"
.TextMatrix(0, 2) = "—ﬁ„ «·›« Ê—…"
.TextMatrix(0, 3) = " «—ÌŒ «·›« Ê—…"
.TextMatrix(0, 4) = "«·«”„"
.TextMatrix(0, 5) = "«· ·Ì›Ê‰"
.TextMatrix(0, 6) = "≈Ã„«·Ì «·›« Ê—…"
.TextMatrix(0, 7) = "≈Ã„«·Ì «·»Ê·Ì’…"
.TextMatrix(0, 8) = "≈Ã„«·Ì «·„”œœ"
.TextMatrix(0, 9) = "«·⁄„Ê·…"
.TextMatrix(0, 10) = "«·’«›Ì"
.TextMatrix(0, 11) = "ID"

.RowHeight(0) = 600
.ColWidth(0) = 3000
.ColWidth(1) = 1400
.ColWidth(2) = 1800
.ColWidth(3) = 1300
.ColWidth(4) = 4500
.ColWidth(5) = 1600
.ColWidth(6) = 1300
.ColWidth(7) = 1200
.ColWidth(8) = 1200
.ColWidth(9) = 1200
.ColWidth(10) = 1200
.ColWidth(11) = 1200

.ColHidden(.Cols - 1) = True

.ColComboList(2) = "..."
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ColFormat(3) = "YYYY/M/D"
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
If grid1.Rows > 1 Then
    .Cell(flexcpFontSize, 1, 0, .Rows - 1, 0) = 9
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
If Trim(.TextMatrix(Row, 1)) = "" Then
    If bMsg Then
        MsgBox "—ﬁ„ »Ê·Ì’… «·‘Õ‰ €Ì— „”Ã·"
    End If
    Exit Function
End If

If Trim(.TextMatrix(Row, 2)) = "" Then
    If bMsg Then
        MsgBox "—ﬁ„ «·›« Ê—… €Ì— „”Ã·"
    End If
    Exit Function
End If

If .ValueMatrix(Row, 8) = 0 And .ValueMatrix(Row, 9) = 0 Then
    If bMsg Then
        MsgBox "·« ÌÊÃœ ⁄„Ê·… «Ê ”œ«œ"
    End If
    Exit Function
End If
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col = 0 Then
    If NextEmpty(grid1, Row, Col, 8) <= 2 Then
        If grid1.TextMatrix(Row, Col) = "" Then
            grid1.Select Row, 1
        Else
            grid1.Select Row, 2
        End If
    Else
        grid1.Select Row, NextEmpty(grid1, Row, Col, 8)
    End If
ElseIf Col = 1 Then
    grid1.Select Row, NextEmpty(grid1, Row, Col + 1, 8)
ElseIf Col < 8 Then
    grid1.Col = 8
ElseIf Col < grid1.Cols - 2 Then
    grid1.Col = Col + 1
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
    If Col = 2 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    'If Col = 2 And (grid1.TextMatrix(Row, Col) = "" Or grid1.TextMatrix(Row, grid1.Cols - 1) = "") Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid1.Col = 2 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    'If grid1.Col = 1 And (grid1.TextMatrix(grid1.Row, grid1.Col) = "" Or grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Then Exit Sub
    'If grid1.Col = 2 And (grid1.TextMatrix(grid1.Row, grid1.Col) = "" Or grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Then Exit Sub
    KeyAscii = 0
End If
End Sub
Public Sub invLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "", Optional sType As String = "1")
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(8, 1)
Dim cWhere As String
Set Generalarray(0) = oForm
'                       0                   1
If sType = 1 Then
    cString = "SELECT FILE6_20H.DOC_NO," & _
              "FILE6_20H.DOC_NO2," & _
              "CONVERT(VARCHAR(10),FILE6_20H.DATE,111)," & _
              "FILE6_20H.SHIP_NO, " & _
              "FILE6_20H.ONLINE_DOC," & _
              "FILE6_20H.[NAME]," & _
              "FILE6_20H.PHONE ," & _
              "FILE6_20H.PAYMENT_ID, " & _
              "FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT" & _
              " FROM  FILE6_20H " & _
              " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE AND FILE0_40.ONLINE = 1" & _
              " WHERE FILE6_20H.ONLINE = 2" & _
              " AND FILE6_20H.DATE_PAY IS NULL" & _
              " AND FILE6_20H.SHIP_NO IS NOT NULL" & _
              " AND FILE6_20H.PRINTED = 1"
Else
    cString = "SELECT FILE6_20H.DOC_NO," & _
              "FILE6_20H.DOC_NO2," & _
              "CONVERT(VARCHAR(10),FILE6_20H.DATE,111)," & _
              "FILE6_20H.SHIP_NO, " & _
              "FILE6_20H.ONLINE_DOC," & _
              "FILE6_20H.[NAME]," & _
              "FILE6_20H.PHONE, " & _
              "FILE6_20H.PAYMENT_ID, " & _
              "FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT" & _
              " FROM  FILE6_20H " & _
              " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE AND FILE0_40.ONLINE = 1" & _
              " WHERE FILE6_20H.ONLINE = 2" & _
              " AND FILE6_20H.SHIP_NO IS NOT NULL" & _
              " AND FILE6_20H.PRINTED = 1"
End If

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

listarray(3, 0) = "—ﬁ„ «·ÿ·»Ì…-Payment Id"
listarray(3, 1) = "(ONLINE_DOC LIKE 'cFilter%' OR PAYMENT_ID LIKE 'cFilter%')"

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

GrdArray(7, 0) = "Payment Id"
GrdArray(7, 1) = IIf(sType = 1, 3000, 0)

GrdArray(8, 0) = "≈Ã„«·Ì «·›« Ê—…"
GrdArray(8, 1) = 1100

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
oSearch.bUnLoad = False
oSearch.Caption = "≈” ⁄·«„ ›Ê« Ì— «Ê‰ ·«Ì‰"
oSearch.Show 1
End Sub
Private Function GrdDesc(ByVal sinv_no As String, ByVal sShip_no As String, ByVal sPayment_id As String, Row As Long, ByRef pMsg As String, Optional nTotal As Variant = Null) As Boolean
Dim loctable As New ADODB.Recordset
Dim cString As String, sFound As String
Dim sDup As String
Dim sFilter As String

sDup = IIf(sShip_no = "", "ﬂÊœ ”œ«œ", "»Ê·Ì’… «·‘Õ‰")

cString = "SELECT FILE6_20H.* FROM FILE6_20H " & _
         " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE AND FILE0_40.ONLINE = 1" & _
         " WHERE FILE6_20H.ONLINE = 2"

If sShip_no <> "" Then
    cString = cString & " AND SHIP_NO = " & MyParn(sShip_no)
ElseIf sPayment_id <> "" Then
    cString = cString & " AND PAYMENT_ID = " & MyParn(sPayment_id)
Else
    cString = cString & " AND DOC_NO = " & MyParn(sinv_no)
End If

If IsNull(nTotal) Then
ElseIf nTotal < 0 Then
    sFilter = "IS_RETURN = 1 AND PRINTED = 1"
ElseIf nTotal >= 0 Then
    sFilter = "IS_RETURN = 0 AND PRINTED = 1"
End If
cString = cString & " ORDER BY TOTAL_ITEM"

Set loctable = myCmd(cString, con)
loctable.Filter = sFilter

If loctable.EOF Then
    If IsNull(nTotal) Then
    ElseIf nTotal < 0 Then
        loctable.Filter = "IS_RETURN = 1"
    ElseIf nTotal >= 0 Then
        loctable.Filter = "IS_RETURN = 0"
    End If
        
    If Not loctable.EOF Then
        If IsNull(loctable!ship_no) Then
            pMsg = "»œÊ‰ »Ê·Ì’… ‘Õ‰"
        End If
        pMsg = pMsg & Tr(pMsg, "-") & "·Ì”  „€·›…"
        If sDoc_no = "" Then myLoadNoShip loctable, grid1, Row, pMsg
        Exit Function
    End If
        
    loctable.Filter = "PRINTED = 1"
    If Not loctable.EOF Then
        If IsNull(loctable!ship_no) Then
            pMsg = "»œÊ‰ »Ê·Ì’… ‘Õ‰"
        End If
        If IsNull(nTotal) Then
        ElseIf nTotal < 0 Then
            pMsg = pMsg & Tr(pMsg) & "”œ«œ ”«·» Ê·Ì” „— Ã⁄"
        ElseIf nTotal >= 0 Then
            pMsg = pMsg & Tr(pMsg) & "”œ«œ „ÊÃ» ·„— Ã⁄"
        End If
        If sDoc_no = "" Then myLoadNoShip loctable, grid1, Row, pMsg
        Exit Function
    End If
    
    loctable.Filter = ""
    If Not loctable.EOF Then
        If IsNull(loctable!ship_no) Then
            pMsg = "»œÊ‰ »Ê·Ì’… ‘Õ‰"
        End If
        pMsg = "”œ«œ €Ì— ’ÕÌÕ Ê«·›« Ê—… ·Ì”  „€·ﬁ…"
        If sDoc_no = "" Then myLoadNoShip loctable, grid1, Row, pMsg
        Exit Function
    End If
    
    pMsg = sDup & " €Ì— ’«·Õ"
ElseIf loctable.RecordCount > 1 Then
    pMsg = "«ﬂÀ— „‰ ›« Ê—… »‰›” " & sDup
    myLoadNoShip loctable, grid1, Row, pMsg
    Exit Function
ElseIf (Not IsNull(loctable!DATE_PAY)) And grid1.FindRow(loctable!doc_no, , 2) = -1 Then
    pMsg = "„”œœ… „‰ ﬁ»·"
    If sDoc_no = "" Then myLoadNoShip loctable, grid1, Row, pMsg
ElseIf IsNull(loctable!ship_no) Then
    pMsg = "»œÊ‰ —ﬁ„ »Ê·Ì’…"
    If sDoc_no = "" Then myLoadNoShip loctable, grid1, Row, pMsg
Else
    grid1.TextMatrix(Row, 0) = loctable!PAYMENT_ID & ""
    grid1.TextMatrix(Row, 1) = loctable!ship_no & ""
    grid1.TextMatrix(Row, 2) = loctable!doc_no
    grid1.TextMatrix(Row, 3) = myFormat_p(loctable!Date)
    grid1.TextMatrix(Row, 4) = loctable!Name & "" & IIf(IsNull(loctable!DATE_PAY), "", " („”œœ…)")
    grid1.TextMatrix(Row, 5) = loctable!phone & ""
    grid1.TextMatrix(Row, 6) = loctable!TOTAL_ITEM - loctable!discount
    grid1.TextMatrix(Row, 7) = (loctable!TOTAL_ITEM - loctable!discount) + loctable!CHARGE1 + loctable!charge2
    GrdDesc = True
End If
End Function
Private Sub GRID2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim cWhere As String
If grid2.TextMatrix(Row, 0) <> "" Then
    cWhere = cWhere & " AND FILE6_20H.PAYMENT_ID = " & MyParn(grid2.TextMatrix(Row, 0))
End If
    
If grid2.TextMatrix(Row, 1) <> "" Then
    cWhere = " AND FILE6_20H.SHIP_NO = " & MyParn(grid2.TextMatrix(Row, 1))
    If grid2.ValueMatrix(Row, 8) <> 0 Then
        cWhere = cWhere & " AND FILE6_20H.DATE_PAY IS NULL"
    End If
End If

    
invLookup Me, oSeachCode, Mid(cWhere, 6), , , "2"
End Sub
Private Function grdDesc2(ByVal sinv_no As String, ByVal sShip_no As String, Row As Long, Optional ByRef pMsg As String, Optional nCOD As Variant = Null, Optional nCommission As Variant = Null) As Boolean
Dim loctable As New ADODB.Recordset
Dim cString As String, sFound As String
Dim sFilter As String

sDup = "»Ê·Ì’… «·‘Õ‰"

cString = "SELECT FILE6_20H.* FROM FILE6_20H " & _
         " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE AND FILE0_40.ONLINE = 1" & _
         " WHERE FILE6_20H.ONLINE = 2"

If sShip_no <> "" Then
    cString = cString & " AND SHIP_NO = " & MyParn(sShip_no)
Else
    cString = cString & " AND DOC_NO = " & MyParn(sinv_no)
End If

If IsNull(nCOD) Then
ElseIf nCOD < 0 Then
    sFilter = "IS_RETURN = 1 AND PRINTED = 1"
ElseIf nCOD >= 0 Then
    sFilter = "IS_RETURN = 0 AND PRINTED = 1"
End If

cString = cString & " ORDER BY TOTAL_ITEM"

Set loctable = myCmd(cString, con)
If sFilter <> "" Then loctable.Filter = sFilter

If loctable.EOF Then
    If IsNull(nCOD) Then
    ElseIf nCOD < 0 Then
        loctable.Filter = "IS_RETURN = 1"
    ElseIf nCOD >= 0 Then
        loctable.Filter = "IS_RETURN = 0"
    End If
    
    If Not loctable.EOF Then
        If IsNull(loctable!ship_no) Then
            pMsg = "»œÊ‰ »Ê·Ì’… ‘Õ‰"
        End If
        pMsg = pMsg & Tr(pMsg, "-") & "·Ì”  „€·›…"
        If sDoc_no = "" Then myLoadNoShip loctable, grid2, Row, pMsg
        Exit Function
    End If
        
    loctable.Filter = "PRINTED = 1"
    If Not loctable.EOF Then
        If IsNull(loctable!ship_no) Then
            pMsg = "»œÊ‰ »Ê·Ì’… ‘Õ‰"
        End If
        If IsNull(nTotal) Then
        ElseIf nTotal < 0 Then
            pMsg = pMsg & Tr(pMsg) & "”œ«œ ”«·» Ê·Ì” „— Ã⁄"
        ElseIf nTotal >= 0 Then
            pMsg = pMsg & Tr(pMsg) & "”œ«œ „ÊÃ» ·„— Ã⁄"
        End If
        If sDoc_no = "" Then myLoadNoShip loctable, grid2, Row, pMsg
        Exit Function
    End If
        
    loctable.Filter = ""
    If Not loctable.EOF Then
        If IsNull(loctable!ship_no) Then
            pMsg = "»œÊ‰ »Ê·Ì’… ‘Õ‰"
        End If
        pMsg = "”œ«œ €Ì— ’ÕÌÕ Ê«·›« Ê—… ·Ì”  „€·ﬁ…"
        If sDoc_no = "" Then myLoadNoShip loctable, grid2, Row, pMsg
        Exit Function
    End If
        
    pMsg = "»Ê·Ì’… ‘Õ‰ ·Ì”  „ÊÃÊœ…"
ElseIf loctable.RecordCount > 1 Then
    pMsg = "«ﬂÀ— „‰ ›« Ê—… »‰›” " & sDup
    myLoadNoShip loctable, grid2, Row, pMsg
    Exit Function
ElseIf (Not IsNull(loctable!DATE_PAY)) And grid2.FindRow(loctable!doc_no, , 2) = -1 And nCOD <> 0 Then
    pMsg = "„”œœ… „‰ ﬁ»·"
    If sDoc_no = "" Then myLoadNoShip loctable, grid2, Row, pMsg
ElseIf IsNull(loctable!ship_no) Then
    pMsg = "»œÊ‰ —ﬁ„ »Ê·Ì’…"
    If sDoc_no = "" Then myLoadNoShip loctable, grid2, Row, pMsg
Else
    grid2.TextMatrix(Row, 0) = loctable!PAYMENT_ID & ""
    grid2.TextMatrix(Row, 1) = loctable!ship_no & ""
    grid2.TextMatrix(Row, 2) = loctable!doc_no
    grid2.TextMatrix(Row, 3) = myFormat_p(loctable!Date)
    grid2.TextMatrix(Row, 4) = loctable!Name & "" & IIf(IsNull(loctable!DATE_PAY), "", " („”œœ…)")
    grid2.TextMatrix(Row, 5) = loctable!phone & ""
    grid2.TextMatrix(Row, 6) = loctable!TOTAL_ITEM - loctable!discount
    grid2.TextMatrix(Row, 7) = (loctable!TOTAL_ITEM - loctable!discount) + loctable!CHARGE1 + loctable!charge2
    grdDesc2 = True
End If
End Function
Private Sub myreplaceGrd2(Row As Long)
Dim aInsert As Variant
With grid2
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, .Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "INV_NO", addstring(.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "[COD]", .ValueMatrix(i, 8))
        aInsert = AddFlag(aInsert, "[FEE]", .ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "[TYPE]", "2")
        aInsert = AddFlag(aInsert, "[FLAG]", "1")
        If .TextMatrix(i, .Cols - 1) = "" Then
            con.Execute addInsert(aInsert, cFile)
        Else
            con.Execute addUpdate(aInsert, cFile, "ID = " & .TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i As Long
Dim bAll As Boolean
For i = 1 To grid2.Rows - 2
    If i <> Row Then
        If grid2.TextMatrix(i, grid2.Cols - 1) = "" Then
            cmdSave_Click
            Exit Sub
        End If
    End If
Next

If Not myValid Then
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

If Not validrow2(Row, False) Then
    CalcTotals2 Row
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
ElseIf grid1.Row > 2 Then
    grid2.Editable = flexEDNone
ElseIf ((grid2.Col = 0 Or grid2.Col = 1 Or grid2.Col = 2) And grid2.TextMatrix(grid2.Row, grid2.Cols - 1) = "") Or grid2.Col = 8 Or grid2.Col = 9 Then
    grid2.Editable = flexEDKbdMouse
Else
    grid2.Editable = flexEDNone
End If
End Sub
Private Sub myloadgrd2()
Dim cString As String
Dim aFields(11)
aFields(0) = "FILE6_20H.PAYMENT_ID"
aFields(1) = "FILE6_20H.SHIP_NO"
aFields(2) = "FILE9_10.INV_NO"
aFields(3) = "FILE6_20H.DATE"
aFields(4) = "FILE6_20H.[NAME]"
aFields(5) = "FILE6_20H.[PHONE]"
aFields(6) = "(FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT)"
aFields(7) = "(FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT) + FILE6_20H.CHARGE1 + FILE6_20H.CHARGE2"
aFields(8) = "FILE9_10.COD"
aFields(9) = "FILE9_10.FEE"
aFields(10) = "FILE9_10.[VALUE]"
aFields(11) = "FILE9_10.ID"

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
Private Sub GRID2_GotFocus()
grid2_EnterCell
End Sub
Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos2 KeyCode, grid2.Row, grid2.Col
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
Private Sub GRID2_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid2
Dim sMsg As String
If Col = 0 Then
    If Trim(grid2.EditText) = "" Then
        Cancel = True
    ElseIf Not GrdDesc("", "", grid2.EditText, Row, sMsg) Then
        MsgBox sMsg
        Cancel = True
    End If
ElseIf Col = 1 Then
    If Trim(grid2.EditText) = "" Then
        Cancel = True
    ElseIf Not GrdDesc("", grid2.EditText, "", Row, sMsg) Then
        MsgBox sMsg
        Cancel = True
    End If
End If
End With
End Sub
Private Sub Fixgrd2()
With grid2
.TextMatrix(0, 0) = "Payment id"
.TextMatrix(0, 1) = "»Ê·Ì’… «·‘Õ‰"
.TextMatrix(0, 2) = "—ﬁ„ «·›« Ê—…"
.TextMatrix(0, 3) = " «—ÌŒ «·›« Ê—…"
.TextMatrix(0, 4) = "«·«”„"
.TextMatrix(0, 5) = "«· ·Ì›Ê‰"
.TextMatrix(0, 6) = "≈Ã„«·Ì «·›« Ê—…"
.TextMatrix(0, 7) = "≈Ã„«·Ì «·»Ê·Ì’…"
.TextMatrix(0, 8) = "COD"
.TextMatrix(0, 9) = "ﬁÌ„… «·‘Õ‰"
.TextMatrix(0, 10) = "’«›Ì"
.TextMatrix(0, 11) = "ID"

.RowHeight(0) = 600
.ColWidth(0) = 4000
.ColWidth(1) = 1800
.ColWidth(2) = 1300
.ColWidth(3) = 1300
.ColWidth(4) = 3500
.ColWidth(5) = 1600
.ColWidth(6) = 1300
.ColWidth(7) = 1200
.ColWidth(8) = 1200
.ColWidth(9) = 1200
.ColWidth(10) = 1200
.ColWidth(11) = 1200

.ColHidden(.Cols - 1) = True

.ColComboList(2) = "..."
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 10
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

If .ValueMatrix(Row, 10) = 0 And .ValueMatrix(Row, 8) = 0 Then
    If bMsg Then
        MsgBox "·« ÌÊÃœ COD «Ê ﬁÌ„… ‘Õ‰"
    End If
    Exit Function
End If
End With
validrow2 = True
End Function
Private Sub CellPos2(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col = 0 Then
    If NextEmpty(grid2, Row, Col, 8) <= 2 Then
        If grid2.TextMatrix(Row, Col) = "" Then
            grid2.Select Row, 1
        Else
            grid2.Select Row, 2
        End If
    Else
        grid2.Select Row, NextEmpty(grid2, Row, Col, 8)
    End If
ElseIf Col = 1 Then
    grid2.Select Row, NextEmpty(grid2, Row, Col + 1, 8)
ElseIf Col < 8 Then
    grid2.Col = 8
ElseIf Col < grid2.Cols - 2 Then
    grid2.Col = Col + 1
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
Private Sub GRID2_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos2 KeyCode, Row, Col
End If
End Sub
Private Sub GRID2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
End If
End Sub
Private Function CalcTotals2(Optional Row As Long = -1)
Dim nTotal As Double
With grid2
For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, .Rows - 2, Row)
    .TextMatrix(i, 10) = .ValueMatrix(i, 8) - .ValueMatrix(i, 9)
    nTotal = mRound(nTotal + .ValueMatrix(i, 10))
Next
End With
End Function
Private Function getDataWeb()
Dim cFileName As String
Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

cFileName = Common1.FileName

If cFileName = "" Then Exit Function

Dim cSv As New ChilkatCsv
Dim i As Long

cSv.HasColumnNames = 1
nAccess = cSv.LoadFile(cFileName)
If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If


If cSv.NumRows < 1 Then Exit Function

Dim Tb As New ChilkatStringBuilder
Dim cString As New ChilkatStringBuilder
Dim loctable As New ADODB.Recordset

prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption

Dim sMsg As String, bDone As Boolean
Dim nTotal As Double

grid1.Rows = 1
myAddItem
For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    
    sid = cSv.GetCellByName(i, "id")
    If sid <> "" Then
        sType = cSv.GetCellByName(i, "type")
        nTotal = Val(cSv.GetCellByName(i, "total"))
        nCommission = Val(cSv.GetCellByName(i, "Commission"))
        nValue = Val(cSv.GetCellByName(i, "value"))
    
        If LCase(Mid(sType, 1, 1)) = "r" Then
            nTotal = nTotal * -1
            nValue = nValue * -1
        End If
            
        grid1.TextMatrix(grid1.Rows - 1, 0) = sid
        sMsg = ""
        bDone = GrdDesc("", "", sid, grid1.Rows - 1, sMsg, nTotal)
        If sMsg <> "" Then
            grid1.TextMatrix(grid1.Rows - 1, 4) = sMsg
        End If
        grid1.TextMatrix(grid1.Rows - 1, 8) = nTotal
        grid1.TextMatrix(grid1.Rows - 1, 9) = nCommission
        grid1.TextMatrix(grid1.Rows - 1, 10) = nValue
        myAddItem
    End If
Next
myAddItem
prog1.Visible = False
Me.Caption = sCaption
End Function
Private Function getDataShip()
Dim cFileName As String
Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

cFileName = Common1.FileName

If cFileName = "" Then Exit Function

Dim cSv As New ChilkatCsv
Dim i As Long

cSv.HasColumnNames = 1
nAccess = cSv.LoadFile(cFileName)
If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If


If cSv.NumRows < 1 Then Exit Function

Dim Tb As New ChilkatStringBuilder
Dim cString As New ChilkatStringBuilder
Dim loctable As New ADODB.Recordset

prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption

'grid2.Rows = 1
'GRID3.Rows = 1
'myAddItem2
'myAddItem3


Dim sMsg As String

Dim ship_no As String
Dim bDone As Boolean
Dim cod As Double
Dim fee As Double
Dim Value As Double

For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    
    ship_no = cSv.GetCellByName(i, "ship_no")
    If ship_no <> "" Then
        cod = Val(cSv.GetCellByName(i, "cod"))
        fee = Val(cSv.GetCellByName(i, "fee"))
        Value = Val(cSv.GetCellByName(i, "value"))
        If Not IsNull(myField("SELECT dbo.fn_ship_adv_id(" & MyParn(ship_no) & ")", con)) Then
            GRID3.TextMatrix(GRID3.Rows - 1, 0) = ship_no
            sMsg = ""
            bDone = grdDesc3(ship_no, GRID3.Rows - 1, sMsg)
            If sMsg <> "" Then
                GRID3.TextMatrix(GRID3.Rows - 1, 4) = sMsg
            End If
            GRID3.TextMatrix(GRID3.Rows - 1, 7) = Value
            GRID3.AddItem ""
        Else
            grid2.TextMatrix(grid2.Rows - 1, 1) = ship_no
            sMsg = ""
            bDone = grdDesc2("", ship_no, grid2.Rows - 1, sMsg, cod)
            If sMsg <> "" Then
                grid2.TextMatrix(grid2.Rows - 1, 4) = sMsg
            End If
            grid2.TextMatrix(grid2.Rows - 1, 8) = cod
            grid2.TextMatrix(grid2.Rows - 1, 9) = fee
            grid2.TextMatrix(grid2.Rows - 1, 10) = Value
            grid2.AddItem ""
        End If
    End If
Next
prog1.Visible = False
Me.Caption = sCaption
End Function
Private Function myLoadNoShip(loctable, pGrid, Row As Long, pMsg As String)
With pGrid
pGrid.TextMatrix(Row, 0) = loctable!PAYMENT_ID & ""
'pGrid.TextMatrix(Row, 1) = loctable!ship_no & ""
pGrid.TextMatrix(Row, 2) = loctable!doc_no
pGrid.TextMatrix(Row, 3) = myFormat_p(loctable!Date)
pGrid.TextMatrix(Row, 4) = pMsg
pGrid.TextMatrix(Row, 5) = loctable!phone & ""
pGrid.TextMatrix(Row, 6) = loctable!TOTAL_ITEM - loctable!discount
pGrid.TextMatrix(Row, 7) = (loctable!TOTAL_ITEM - loctable!discount) + loctable!CHARGE1 + loctable!charge2
End With
End Function


Private Sub GRID3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim cWhere As String
If GRID3.TextMatrix(Row, 1) <> "" Then
    cWhere = " AND LADING_BILL.SHIP_NO = " & MyParn(GRID3.TextMatrix(Row, 0))
End If
ShipLookup Me, oSeachCode, Mid(cWhere, 6), , , "3"
End Sub
Private Function grdDesc3(ByVal sShip_no As String, Row As Long, Optional ByRef pMsg As String) As Boolean
Dim loctable As New ADODB.Recordset
Dim cString As String, sFound As String
Dim sFilter As String
sDup = "»Ê·Ì’… «·‘Õ‰"
cString = "SELECT TOP 1 LADING_BILL.DOC_NO," & _
          "LADING_BILL.SHIP_NO," & _
          "FILE6_20H.DATE," & _
          "FILE6_20H.[NAME]," & _
          "FILE6_20H.PHONE," & _
          "LADING_BILL.CHARGE1 + LADING_BILL.CHARGE2 AS CHARGE," & _
          "LADING_BILL.DATE_DELIVERY," & _
          "[dbo].[fn_date_pay_ship](LADING_BILL.SHIP_NO) as DATE_PAY," & _
          " LADING_BILL.CLOSED" & _
          " FROM  LADING_BILL INNER JOIN FILE6_20H ON LADING_BILL.DOC_NO = FILE6_20H.DOC_NO" & _
          " WHERE TYPE > 2" & _
          " AND LADING_BILL.SHIP_NO = " & MyParn(sShip_no)

cString = cString & " ORDER BY LADING_BILL.DATE_DELIVERY,LADING_BILL.ID"
Set loctable = myCmd(cString, con)

loctable.Filter = "CLOSED = 1 AND DATE_PAY = NULL"
If loctable.EOF Then
    loctable.Filter = "DATE_PAY = NULL"
    If Not loctable.EOF Then
        pMsg = "»Ê·Ì’… €Ì— „€·ﬁ…"
        myLoadShip loctable, GRID3, Row, pMsg
        Exit Function
    End If
    
    loctable.Filter = "CLOSED = 1"
    If Not loctable.EOF Then
        pMsg = "»Ê·Ì’… „”œœ… „‰ ﬁ»·"
        myLoadShip loctable, GRID3, Row, pMsg
        Exit Function
    End If
     
    loctable.Filter = ""
    If Not loctable.EOF Then
        pMsg = "»Ê·Ì’… €Ì— „€·ﬁ… Ê»Ê·Ì’… „”œœ… „‰ ﬁ»·"
        myLoadShip loctable, GRID3, Row, pMsg
        Exit Function
    End If
     
    pMsg = "»Ê·Ì’… ‘Õ‰ ·Ì”  „ÊÃÊœ…"
ElseIf (Not IsNull(loctable!DATE_PAY)) Then
    pMsg = "„”œœ… „‰ ﬁ»·"
    myLoadShip loctable, GRID3, Row, pMsg
Else
    myLoadShip loctable, GRID3, Row
    grdDesc3 = True
End If
End Function
Private Sub myreplaceGrd3(Row As Long)
Dim aInsert As Variant
With GRID3
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, .Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "SHIP_NO", addstring(.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "[VALUE]", .ValueMatrix(i, 7))
        aInsert = AddFlag(aInsert, "[TYPE]", "3")
        aInsert = AddFlag(aInsert, "[FLAG]", "2")
        If .TextMatrix(i, .Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE9_10A")
        Else
            con.Execute addUpdate(aInsert, "FILE9_10A", "ID = " & .TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub grid3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i As Long

For i = 1 To GRID3.Rows - 2
    If i <> Row Then
        If GRID3.TextMatrix(i, GRID3.Cols - 1) = "" Then
            cmdSave_Click
            Exit Sub
        End If
    End If
Next
If Not myValid() Then
    On Error Resume Next
    GRID3.SetFocus
    Err.Clear
    myloadgrd3
    If Row < GRID3.Rows - 1 Then
        GRID3.Select Row, Col
    Else
        CellPos3 13, GRID3.Rows - 2, GRID3.Cols - 1
    End If
    Exit Sub
End If

If Not validrow3(Row, False) Then
    CalcTotals3 Row
    Exit Sub
End If

With GRID3
If Row = GRID3.Rows - 1 Then
    myAddItem3
End If

If myreplace(, , Row) Then
    If xDoc_No.Tag = DefineMode Then
        Handlecontrols LoadMode
        myloadgrd3
    ElseIf GRID3.TextMatrix(Row, GRID3.Cols - 1) = "" Then
        myloadgrd3
    Else
        CalcTotals3 Row
    End If
Else
    myloadgrd3
End If
End With
End Sub
Private Sub GRID3_EnterCell()
If Not bEditRecord Then
    GRID3.Editable = flexEDNone
ElseIf grid1.Rows > 2 Then
    GRID3.Editable = flexEDNone
ElseIf ((GRID3.Col = 0 Or GRID3.Col = 1) And GRID3.TextMatrix(GRID3.Row, GRID3.Cols - 1) = "") Or GRID3.Col = 7 Then
    GRID3.Editable = flexEDKbdMouse
Else
    GRID3.Editable = flexEDNone
End If
End Sub
Private Sub myloadgrd3()
Dim cString As String
Dim aFields(8)
aFields(0) = "FILE9_10A.SHIP_NO"
aFields(1) = "FILE6_20H.DOC_NO"
aFields(2) = "FILE6_20H.DATE"
aFields(3) = "FILE6_20H.[NAME]"
aFields(4) = "FILE6_20H.[PHONE]"
aFields(5) = "LADING_BILL.DATE_DELIVERY"
aFields(6) = "LADING_BILL.CHARGE1 + LADING_BILL.CHARGE2 AS CHARGE"
aFields(7) = "FILE9_10A.[VALUE]"
aFields(8) = "FILE9_10A.ID"
cString = "SELECT " & arString(aFields, ",") & _
         " FROM FILE9_10A INNER JOIN LADING_BILL ON LADING_BILL.ID = [dbo].[fn_ship_adv_id](FILE9_10A.SHIP_NO) " & _
         " INNER JOIN FILE6_20H ON LADING_BILL.DOC_NO = FILE6_20H.DOC_NO" & _
         " WHERE FILE9_10A.DOC_NO = " & addvalue(xDoc_No.text)
Set data12.Recordset = myCmd(cString, con)

myAddItem3

CalcTotals3
Fixgrd3
End Sub
Private Sub GRID3_GotFocus()
GRID3_EnterCell
End Sub
Private Sub GRID3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos3 KeyCode, GRID3.Row, GRID3.Col
ElseIf Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 46 And GRID3.Row <> GRID3.Rows - 1 Then
    If GRID3.TextMatrix(GRID3.Row, GRID3.Cols - 1) = "" Then
        myRemove3 GRID3.Row
    ElseIf MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If GRID3.TextMatrix(GRID3.Row, GRID3.Cols - 1) <> "" Then
            con.Execute "Delete from FILE9_10A where ID = " & GRID3.TextMatrix(GRID3.Row, GRID3.Cols - 1)
        End If
        con.CommitTrans
        myRemove3 GRID3.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub GRID3_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid2
Dim sMsg As String
If Col = 0 Then
    If Trim(grid2.EditText) = "" Then
        Cancel = True
    ElseIf Not grdDesc3(.EditText, Row, sMsg) Then
        MsgBox sMsg
        Cancel = True
    End If
ElseIf Col = 1 Then
    If Trim(.EditText) = "" Then
        Cancel = True
    ElseIf Not grdDesc3(.EditText, Row, sMsg) Then
        MsgBox sMsg
        Cancel = True
    End If
End If
End With
End Sub
Private Sub Fixgrd3()
With GRID3
.TextMatrix(0, 0) = "»Ê·Ì’… «·‘Õ‰"
.TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
.TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
.TextMatrix(0, 3) = "«·«”„"
.TextMatrix(0, 4) = "«· ·Ì›Ê‰"
.TextMatrix(0, 5) = " «—ÌŒ «· ”·Ì„"
.TextMatrix(0, 6) = "ﬁÌ„… «·»Ê·Ì’…"
.TextMatrix(0, 7) = "„”œœ"
.TextMatrix(0, 8) = "ID"

.RowHeight(0) = 600
.ColWidth(0) = 2000
.ColWidth(1) = 2000
.ColWidth(2) = 1350
.ColWidth(3) = 5000
.ColWidth(4) = 3000
.ColWidth(6) = 1300
.ColWidth(7) = 1300
.ColWidth(8) = 1300

.ColHidden(.Cols - 1) = True
.ColComboList(1) = "..."
For i = 1 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Private Sub myAddItem3()
With GRID3
.AddItem ""
If GRID3.Rows > 2 Then
    '.TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
End If
End With
End Sub
Private Function validrow3(Row As Long, Optional bMsg As Boolean = False) As Boolean
With GRID3
If Trim(.TextMatrix(Row, 0)) = "" Then
    If bMsg Then
        MsgBox "—ﬁ„ »Ê·Ì’… ‘Õ‰ „”»ﬁ €Ì— „”Ã·"
    End If
    Exit Function
End If

If Trim(.TextMatrix(Row, 1)) = "" Then
    If bMsg Then
        MsgBox "—ﬁ„ «·›« Ê—… €Ì— „”Ã·"
    End If
    Exit Function
End If

If .ValueMatrix(Row, 7) = 0 Then
    If bMsg Then
        MsgBox "·„ Ì „ «œŒ«· ﬁÌ„… »Ê·Ì÷… ‘Õ‰ „”»ﬁ"
    End If
    Exit Function
End If
End With
validrow3 = True
End Function
Private Sub CellPos3(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col = 0 Then
    If NextEmpty(GRID3, Row, Col, 7) <= 1 Then
        GRID3.Select Row, 1
    Else
        GRID3.Select Row, NextEmpty(GRID3, Row, Col, 7)
    End If
ElseIf Col = 1 Then
    GRID3.Select Row, NextEmpty(GRID3, Row, Col + 1, 7)
ElseIf Col < 7 Then
    GRID3.Col = 7
ElseIf Col < GRID3.Cols - 2 Then
    GRID3.Col = Col + 1
ElseIf Row < GRID3.Rows - 1 Then
    GRID3.Select Row + 1, NextEmpty(GRID3, Row + 1, 0, 7)
    GRID3.ShowCell GRID3.Row, 0
Else
    GRID3.Select Row, Col
End If
End Sub
Private Sub myRemove3(Row As Long)
GRID3.RemoveItem Row
CalcTotals3
End Sub
Private Sub GRID3_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos3 KeyCode, Row, Col
End If
End Sub
Private Sub GRID3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
End If
End Sub
Private Function CalcTotals3(Optional Row As Long = -1)
Dim nTotal As Double
With GRID3
For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, .Rows - 2, Row)
    nTotal = mRound(nTotal + .ValueMatrix(i, 3))
Next
End With
End Function
Private Function myLoadShip(loctable, pGrid, Row As Long, Optional pMsg)
With pGrid
pGrid.TextMatrix(Row, 0) = loctable!ship_no & ""
pGrid.TextMatrix(Row, 1) = loctable!doc_no
pGrid.TextMatrix(Row, 2) = myFormat_p(loctable!Date)
If IsMissing(pMsg) Then
    pGrid.TextMatrix(Row, 3) = loctable!Name & ""
Else
    pGrid.TextMatrix(Row, 3) = pMsg
End If
pGrid.TextMatrix(Row, 4) = loctable!phone & ""
pGrid.TextMatrix(Row, 5) = myFormat_p(loctable!Date_delivery)
pGrid.TextMatrix(Row, 6) = loctable!CHARGE
End With
End Function
Public Sub ShipLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "", Optional sType As String = "1")
Dim Generalarray(5)
Dim listarray(2, 5)
Dim GrdArray(8, 1)
Dim cWhere As String
Set Generalarray(0) = oForm
'                       0                   1
cString = "SELECT LADING_BILL.SHIP_NO," & _
          "LADING_BILL.DOC_NO," & _
          "CONVERT(VARCHAR(10),FILE6_20H.DATE,111)," & _
          "FILE6_20H.[NAME]," & _
          "FILE6_20H.PHONE ," & _
          "LADING_BILL.DATE_DELIVERY ," & _
          "LADING_BILL.CHARGE1 + LADING_BILL.CHARGE2" & _
          " FROM  LADING_BILL INNER JOIN FILE6_20H ON LADING_BILL.DOC_NO = FILE6_20H.DOC_NO" & _
          " WHERE LADING_BILL.TYPE > 2" & _
          " AND LADING_BILL.CLOSED = 1" & _
          " AND [dbo].[fn_date_pay_ship](LADING_BILL.SHIP_NO) IS NULL"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " AND " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by LADING_BILL.DATE_DELIVERY DESC,FILE6_20H.DATE"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "(FILE6_20H.DOC_NO = 'cFilter' OR FILE6_20H.DOC_NO2 = 'cFilter')"

listarray(1, 0) = "«·«”„-«·⁄‰Ê«‰-—ﬁ„ «· ·Ì›Ê‰"
listarray(1, 1) = "(%%FILE6_20H.NAME%% OR %%FILE6_20H.ADDRESS%% OR PHONE LIKE 'cFilter%')"

listarray(2, 0) = "—ﬁ„ «·»Ê·Ì’…"
listarray(2, 1) = "(SHIP_NO LIKE 'cFilter%')"

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

GrdArray(7, 0) = "Payment Id"
GrdArray(7, 1) = IIf(sType = 1, 3000, 0)

GrdArray(8, 0) = "≈Ã„«·Ì «·›« Ê—…"
GrdArray(8, 1) = 1100

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
oSearch.bUnLoad = True
oSearch.Caption = "≈” ⁄·«„ ›Ê« Ì— «Ê‰ ·«Ì‰"
oSearch.Show 1
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
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xBank_GotFocus()
myGotFocus xBank
End Sub
Private Sub xBank_LostFocus()
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
