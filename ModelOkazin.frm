VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ModelOkazin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ﬁ«∆„… ⁄—Ê÷ Œ’„ "
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15240
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
   ScaleHeight     =   9855
   ScaleWidth      =   15240
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Height          =   1815
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   315
      Width           =   2040
      Begin VB.CommandButton cmdExcel 
         Caption         =   "”Õ» „‰ «ﬂ”Ì·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1395
         Width           =   1950
      End
      Begin VB.CommandButton CMD_DELSELECT 
         BackColor       =   &H000080FF&
         Caption         =   "Õ–› ”Ã·«  „Õœœ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   585
         Width           =   1950
      End
      Begin VB.CommandButton CD_ADDMODEL 
         Caption         =   "≈÷«›… „ÊœÌ·«  „Ã„⁄…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   180
         Width           =   1950
      End
      Begin VB.CommandButton CMD_AGE 
         BackColor       =   &H0080FFFF&
         Caption         =   "«ŸÂ«— ⁄„— «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   990
         Width           =   1950
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6630
      Left            =   -45
      TabIndex        =   38
      Top             =   2160
      Width           =   17955
      _ExtentX        =   31671
      _ExtentY        =   11695
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   " ›’Ì·Ï „ÊœÌ·«  - »«—ﬂÊœ "
      TabPicture(0)   =   "ModelOkazin.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "GRID2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "≈Ã„«·Ï „ÊœÌ·«  "
      TabPicture(1)   =   "ModelOkazin.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grid3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grid1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Height          =   6060
         Left            =   -74910
         TabIndex        =   39
         Top             =   360
         Width           =   17760
         _cx             =   31327
         _cy             =   10689
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
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
         SelectionMode   =   3
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   6045
         Left            =   90
         TabIndex        =   41
         Top             =   405
         Width           =   11490
         _cx             =   20267
         _cy             =   10663
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
         AllowBigSelection=   -1  'True
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
      Begin VSFlex7Ctl.VSFlexGrid grid3 
         Height          =   6045
         Left            =   11610
         TabIndex        =   43
         Top             =   405
         Width           =   6225
         _cx             =   10980
         _cy             =   10663
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
         AllowBigSelection=   -1  'True
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
      Height          =   1275
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   855
      Width           =   2940
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   675
         Width           =   1590
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ "
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
         Left            =   1860
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   675
         Width           =   810
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ : "
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
         Left            =   1830
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   225
         Width           =   840
      End
   End
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   8820
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   20
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "ModelOkazin.frx":0038
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "ModelOkazin.frx":2208
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   21
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "ModelOkazin.frx":4350
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "ModelOkazin.frx":6518
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   22
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "ModelOkazin.frx":8667
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "ModelOkazin.frx":A847
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   23
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "ModelOkazin.frx":C9A2
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "ModelOkazin.frx":EB5E
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   16
      Top             =   9390
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   820
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   12347
            MinWidth        =   12347
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1080
      Width           =   1410
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ModelOkazin.frx":10CAD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ModelOkazin.frx":13226
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   8820
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   6360
      Begin VB.CommandButton CmdDelInv 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   1620
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ModelOkazin.frx":15589
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
      Begin VB.CommandButton cmdNewInv 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   3195
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ModelOkazin.frx":17E23
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   4770
         Picture         =   "ModelOkazin.frx":1A3CF
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   135
         Width           =   1545
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ModelOkazin.frx":1CBA2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
   End
   Begin VB.Frame Frame6 
      Height          =   780
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   4875
      Begin Threed.SSCommand cmd_excel 
         Height          =   600
         Left            =   45
         TabIndex        =   17
         Top             =   135
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1058
         _Version        =   196610
         BackColor       =   13822956
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
         Picture         =   "ModelOkazin.frx":1F00E
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin Threed.SSCommand CMD_PRINT2 
         Height          =   600
         Left            =   2520
         TabIndex        =   18
         Top             =   135
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   1058
         _Version        =   196610
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
         Picture         =   "ModelOkazin.frx":21809
         Caption         =   "ÿ»«⁄… „” ‰œ "
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   7650
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   810
      Width           =   7485
      Begin VB.CommandButton Command1 
         Caption         =   " ÕœÌÀ"
         Height          =   285
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   585
         Width           =   735
      End
      Begin VB.CommandButton CMD_ALL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "„Õœœ ›ﬁÿ"
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   900
         Width           =   870
      End
      Begin VB.CommandButton CMD_ALL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "«·ﬂ·"
         Height          =   330
         Left            =   1170
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   900
         Width           =   555
      End
      Begin VB.CheckBox xisstop 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   "⁄—÷ „ Êﬁ›"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   180
         Width           =   1770
      End
      Begin VB.TextBox XDISC 
         Alignment       =   2  'Center
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
         Left            =   1800
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox XQTY2 
         Alignment       =   2  'Center
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
         Left            =   3780
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox XQTY1 
         Alignment       =   2  'Center
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
         Left            =   5445
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   990
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   5325
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4950
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "‰”»…  «·Œ’„ "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "«·⁄—÷ : "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   945
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "⁄œœ ﬁÿ⁄ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   945
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "»Ì«‰ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   705
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   1635
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   3240
      Top             =   -45
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   3420
      Top             =   90
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data3 
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
   Begin MSAdodcLib.Adodc data4 
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
   Begin MSAdodcLib.Adodc data5 
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   36
      Top             =   9150
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc DATA7 
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
   Begin MSAdodcLib.Adodc Adodc1 
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
End
Attribute VB_Name = "ModelOkazin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim Search31 As New Search3, search32 As New Search3, oSearchItem As New Search3
Dim CardTable As ADODB.Recordset
Dim cFilter As String
Dim cDefStore As String
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional nRow As Long = -1) As Boolean
Dim aInsert(8, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.text)

aInsert(2, 0) = "DESCA"
aInsert(2, 1) = addstring(xDescA.text)

aInsert(3, 0) = "QTY1"
aInsert(3, 1) = Val(XQTY1.text)

aInsert(4, 0) = "QTY2"
aInsert(4, 1) = Val(XQTY2.text)

aInsert(5, 0) = "DISC"
aInsert(5, 1) = Val(xDisc.text)

aInsert(6, 0) = "isstop"
aInsert(6, 1) = IIf(xisstop.Value, 1, 0)


aInsert(7, 0) = "Date1"
aInsert(7, 1) = addDate(xdate1.text)

aInsert(8, 0) = "Date2"
aInsert(8, 1) = addDate(xdate2.text)

If xDoc_No.Tag = DefineMode Then
    xDoc_No.text = RetZero(Newflag("FILE0_90H", "DOC_NO", con), 6)
    aInsert(0, 1) = addstring(xDoc_No.text)
    con.Execute CreateInsert(aInsert, "FILE0_90H")
Else
    con.Execute CreateUpdate(aInsert, "FILE0_90H", " where doc_no = " & addstring(xDoc_No.text))
End If
myreplaceGrd nRow
myreplace = True
Exit Function
myError:
prog1.Visible = False
MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
On Error GoTo myError
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    GrdDesc grid1.Row
    Grid1_AfterEdit grid1.Row, 0
    If Not bNew Then
'    Unload oSearchItem
'        CellPos 13, grid1.Row, 1
    Else
'        CellPos 13, grid1.Row, 8
    End If
Else
    xDoc_No.text = Search31.grid1.TextMatrix(Search31.grid1.Row, 0)
    myUndo
    Unload Search31
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
Unload Search
End Sub


Private Sub CD_ADDMODEL_Click()
    If Not xDoc_No.Enabled Then
        AddModelGrDisc.Show 1
        myUndo
    End If
End Sub

Private Sub CMD_AGE_Click()
    myLoadGrd 0, 0, 1
End Sub

Private Sub CMD_ALL_Click()
    con.Execute " UPDATE FILE0_90 SET DISC2 = " & Val(xDisc.text) & " WHERE DOC_NO = " & MyParn(xDoc_No.text)
    myLoadGrd 0, 0, 0
End Sub

Private Sub CMD_ALL2_Click()
    With grid1
        For nRow = 1 To .Rows - 1
            If .IsSelected(nRow) Then
                con.Execute " UPDATE FILE0_90 SET DISC2 = " & Val(xDisc.text) & " WHERE DOC_NO = " & MyParn(xDoc_No.text) & " AND MODELNO = " & MyParn(.TextMatrix(nRow, 0))
            End If
        Next nRow
    End With
    myLoadGrd 0, 0, 0
End Sub

Private Sub CMD_DELSELECT_Click()
If MsgBox(" Õ–› ”Ã·«  „Õœœ  ", vbYesNo + vbDefaultButton2) = vbYes Then
    With grid1
        For nRow = 1 To .Rows - 1
            If .IsSelected(nRow) Then
                con.Execute " DELETE FROM FILE0_90 WHERE ID = " & Val(.TextMatrix(nRow, .Cols - 1))
            End If
        Next nRow
        myLoadGrd 0, 0, 0
        .Select .Rows - 1, 0
    End With
End If
End Sub

Private Sub cmd_excel_Click()
    If SSTab1.Tab = 1 Then ToFileExel2 grid1, , , , , 1.1, , , , , , Me
    If SSTab1.Tab = 0 Then ToFileExel2 grid2, , , , , 1.1, , , , , , Me
End Sub

Private Sub Cmd_Print2_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String
    cHead1 = Me.Caption & " " & xDescA.text
    cHead2 = " „‰  «—ÌŒ " & Format(xdate1.text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(xdate2.text, "DD-MM-YYYY")
    cHead3 = " ⁄œœ ﬁÿ⁄ " & XQTY1.text & " «·⁄—÷ " & XQTY2.text & " Œ’„ " & xDisc.text
    
    printGrdNew.doprint grid1, 1, , cHead1, cHead2, , , , , 10, "Arial"
    printGrdNew.Show 1

End Sub

Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myError
        'con.BeginTrans
        con.Execute " Delete  From FILE0_90     where Doc_No = " & MyParn(xDoc_No.text)
        con.Execute " Delete  From FILE0_90H    where Doc_No = " & MyParn(xDoc_No.text)
        AddLod_Data cUserName, 2, " Õ–› " & Me.Caption, con, xDoc_No.text, xDate.text, , xDescA.text
        
        myUndo
        CmdNewInv_Click
        
        Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
    End If
    Exit Sub
myError:
    MsgBox Err.Description
    Err.Clear
    'con.RollbackTrans
End Sub

Private Sub cmdExcel_Click()
'If MsgBox("«÷«›… «·„ÊœÌ·«  ··„” ‰œ", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
Me.MousePointer = vbHourglass
cmdExcel.Enabled = False
getData
myUndo
cmdExcel.Enabled = True
Me.MousePointer = vbNormal
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(6, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),FILE0_90H.DESCA , isstop , date1 , date2 " & _
                  " FROM FILE0_90H "
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If

Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ-«·»Ì«‰"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]## or %%FILE0_90h.DESCA%%)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·»Ì«‰"
GrdArray(3, 1) = 4000

GrdArray(4, 0) = "„ Êﬁ›"
GrdArray(4, 1) = 800

GrdArray(5, 0) = "„‰  «—ÌŒ"
GrdArray(5, 1) = 1500

GrdArray(6, 0) = "≈·Ï"
GrdArray(6, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search31
Search31.Caption = "«” ⁄·«„"
Search31.Show 1
End Sub
Private Sub CmdNewInv_Click()
    myDefine
    On Error Resume Next
    xDoc_No.SetFocus
    Err.Clear
End Sub
Private Sub cmdSave_Click()
If Not myValid Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
AddLod_Data cUserName, 1, " Õ›Ÿ  " & Me.Caption, con, xDoc_No.text, xDate.text, , xDescA.text

myUndo
End Sub
Private Sub CmdUndo_Click()
    myUndo
End Sub

Private Sub Command1_Click()
    con.Execute " UPDATE  FILE6_20H SET DOC_DISC = " & addstring(xDoc_No.text) & " WHERE ( DOC_DISC is null or DOC_DISC = '' ) AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xdate2.text) & " AND NOTE_DISC = " & MyParn(xDescA.text), nRec
    MsgBox " ⁄œœ «·”Ã·«   " & nRec
End Sub

Private Sub Form_Load()
    bEdit = True
    openCon con
    Set grid1.DataSource = data1
    Set grid2.DataSource = DATA2
    Set GRID3.DataSource = DATA3
    myUndo
    Command1.Visible = (cBranch = "00") And bopt2
    cmdSave.Visible = (cBranch = "00")
    CmdDelInv.Visible = (cBranch = "00")
    CD_ADDMODEL.Visible = (cBranch = "00")
    CMD_DELSELECT.Visible = (cBranch = "00")
    CMD_ALL.Visible = (cBranch = "00")
    CMD_ALL2.Visible = (cBranch = "00")
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim bNew As Boolean
If grid1.Col = 0 Then
    GrdDesc grid1.Row
End If

'If Not validRow(Row) Then Exit Sub
If Row = grid1.Rows - 1 Then
'    myaddItem
    bNew = True
End If

With grid1
If myreplace(Row) Then
    If xDoc_No.Tag = DefineMode Then xDoc_No.Tag = LoadMode
Else
    bNew = False
End If
myLoadGrd 0, 0, 0
If bNew Then
    grid1.Row = grid1.Rows - 1
    grid1.ShowCell grid1.Rows - 1, 0
    grid1.Col = 0
End If
End With
End Sub
Private Sub grid1_EnterCell()
With grid1
If cBranch <> "00" Then
    .Editable = flexEDNone
    Exit Sub
End If
If (grid1.Col = 0 Or grid1.Col = 7) And cmdSave.Enabled Then
    .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = vbYellow
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End With
End Sub
Private Sub grid1_GotFocus()
With grid1
    If grid1.Row <= 1 Then
    .Select 1, 0, 1, 0
    .ShowCell 1, 0
    End If
End With
End Sub

Private Sub GRID3_DblClick()
    With GRID3
        myLoadGrd .Row, .Col, 0
    End With
End Sub

Private Sub xdate_GotFocus()
    myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
    myLostFocus xDate
End Sub
Private Sub xDate_Validate(Cancel As Boolean)
    myValidDate xDate
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then CmdInform_Click
End Sub
Private Function myValid(Optional bIgMsg As Boolean = False) As Boolean
If xDoc_No.text = "" Then
    If Not bIsgMsg Then Inform "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If xDescA.text = "" Then
    If Not bIsgMsg Then Inform "»Ì«‰ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    If Not bIsMsg Then Inform "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
'If grid1.Rows < 3 Then
'    If Not bIsMsg Then Inform "·«  ÊÃœ „ÊœÌ·«   „  ”ÃÌ·Â«"
'    Exit Function
'End If

myValid = True
End Function
Private Sub myload()
xDoc_No.text = CardTable!doc_no
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
xDescA.text = CardTable!DESCA & ""
xisstop.Value = IIf(CardTable!ISSTOP, 1, 0)
XQTY1.text = CardTable!QTY1 & ""
XQTY2.text = CardTable!QTY2 & ""
xDisc.text = CardTable!disc & ""

xdate1.text = Format(CardTable!Date1, "dd-mm-yyyy")
xdate2.text = Format(CardTable!date2, "dd-mm-yyyy")

xDoc_No.Enabled = False
myLoadGrd 0, 0, 0
myloadgrd2
myloadgrd3
Handlecontrols LoadMode
End Sub
Private Sub myDefine()
xDoc_No.Enabled = True
xDoc_No.text = RetZero(Newflag("FILE0_90H", "DOC_NO", con), 6)
xDate.text = Format(Date, "dd-mm-yyyy")
xdate1.text = ""
xdate2.text = ""

xisstop.Value = 0
XQTY1.text = ""
XQTY2.text = ""
xDisc.text = ""

xDescA.text = ""
grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
fixGrd
End Sub
Private Sub Handlecontrols(nMode)
Dim bClosed As Boolean
cmdNewInv.Enabled = nMode = LoadMode
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdExcel.Enabled = nMode = LoadMode
xDoc_No.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If xDoc_No.text = "" Then Exit Sub
xDoc_No.text = RetZero(xDoc_No.text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If cBranch <> "00" Then Exit Sub
If KeyCode = 112 And grid1.Col = 0 Then
    ModelLookupAll_bal Me, oSearchItem
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·„ÊœÌ· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myError
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "delete from FILE0_90 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And Col = 0 Then
     ModelLookupAll Me, oSearchItem
ElseIf KeyCode = 13 Then
'    CellPos KeyCode, Row, Col
End If
End Sub
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End Function
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub fixGrd()
Dim n1 As Double, n2 As Double
With grid1

.Cols = 9
'                   0            1      2           3           4           5           6
.FormatString = "„ÊœÌ· |" & "„Ê”„|" & "„’‰⁄|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "«·”⁄—|" & "⁄„— «·„ÊœÌ·|" & "‰”»… Œ’„|"
.WordWrap = True

.ColHidden(0) = True
.RowHeight(0) = 800
.ColWidth(0) = 2000
.ColWidth(1) = 800
.ColWidth(2) = 1800
.ColWidth(3) = 1500
.ColWidth(4) = 4100
.ColWidth(5) = 1000
.ColWidth(6) = 0
.ColWidth(7) = 1000

.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
End With
End Sub
Private Sub Fixgrd2()
Dim n1 As Double, n2 As Double
With grid2

.Cols = 13
'                   0            1      2           3           4           5           6           7           8           9           10              11          12
.FormatString = "„ÊœÌ· |" & "„Ê”„|" & "„’‰⁄|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "«··Ê‰|" & "„ „ﬁ«”|" & "„ﬁ«”|" & "«·”⁄—|" & "‰”»… Œ’„|" & "ﬂÊœ «·’‰›|" & "»«—ﬂÊœ „’‰⁄|" & "»«—ﬂÊœ GS1 "
.WordWrap = True

.RowHeight(0) = 800
.ColHidden(0) = True
.ColHidden(6) = True
.ColWidth(0) = 2000
.ColWidth(1) = 800
.ColWidth(2) = 1500
.ColWidth(3) = 1200
.ColWidth(4) = 3000
.ColWidth(5) = 1000
.ColWidth(6) = 0
.ColWidth(7) = 1000
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 1000
.ColWidth(11) = 1500
.ColWidth(12) = 1500

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
End With
End Sub
Private Sub myreplaceGrd(nRow)
Dim cStr1 As String, nCount As Double
Dim aInsert(2, 1)
With grid1
    For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        If .TextMatrix(i, 0) <> "" Then
        
            aInsert(0, 0) = "doc_no"
            aInsert(0, 1) = addstring(xDoc_No.text)
            
            aInsert(1, 0) = "MODELNO"
            aInsert(1, 1) = addstring(grid1.TextMatrix(i, 0))
            
            aInsert(2, 0) = "DISC2"
            aInsert(2, 1) = Val(grid1.TextMatrix(i, 7))
            
            
            If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
                cStr1 = cStr1 & CreateInsert(aInsert, "FILE0_90") & " ; "
'               con.Execute CreateInsert(aInsert, "FILE0_90")
            Else
                cStr1 = cStr1 & CreateUpdate(aInsert, "FILE0_90", " where ID = " & grid1.TextMatrix(i, .Cols - 1)) & " ; "
'               con.Execute CreateUpdate(aInsert, "FILE0_90", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
            End If
        End If
    Next
    If cStr1 <> "" Then con.Execute cStr1
End With
End Sub
Private Sub myLoadGrd(pRow, pCol, lAge As Boolean)
Dim cString As String

'If lAge Then
'    cField2 = "DATEDIFF(dd , (SELECT MIN(DATE) FROM FILE1_11 INNER JOIN FILE1_10 AS FILE1_10_2 ON FILE1_10_2.ITEM = FILE1_11.ITEM WHERE FILE1_10_2.MODELNO =  FILE0_90.MODELNO ) , GETDATE()) AS DATEDIFF"
'Else
'    cField2 = " ' '  "
'End If
'                       0               1               2               3                   4                   5                   6
cString = "SELECT FILE0_90.MODELNO , file1_10.mosm , fact.desca , file1_10.modelfact0 ,file1_10.desca , MIN(FILE1_10.PRICE) , ' ', FILE0_90.DISC2 , FILE0_90.ID from (FILE0_90 inner join file1_10 on file1_10.MODELNO = FILE0_90.MODELNO ) inner join fact on fact.code = file1_10.fact"
cString = cString & " WHERE DOC_NO = " & MyParn(xDoc_No.text)
If pRow > 0 Then
    If pCol = 1 Then
        cString = cString & " AND FILE1_10.[SECTION] = " & Val(GRID3.TextMatrix(pRow, 0))
    ElseIf pCol = 3 Then
        cString = cString & " AND FILE1_10.FACT = " & MyParn(GRID3.TextMatrix(pRow, 2))
    ElseIf pCol = 4 Then
        cString = cString & " AND FILE1_10.MOSM = " & MyParn(GRID3.TextMatrix(pRow, 4))
    Else
        cString = cString & " AND FILE1_10.[SECTION] = " & Val(GRID3.TextMatrix(pRow, 0))
        cString = cString & " AND FILE1_10.FACT = " & MyParn(GRID3.TextMatrix(pRow, 2))
        cString = cString & " AND FILE1_10.MOSM = " & MyParn(GRID3.TextMatrix(pRow, 4))
    End If
End If
cString = cString & " GROUP BY FILE0_90.MODELNO , file1_10.mosm , fact.desca , file1_10.modelfact0 ,file1_10.desca , FILE0_90.DISC2 ,FILE0_90.ID  ORDER BY FILE0_90.MODELNO "
Set data1.Recordset = myRecordSet(cString, con)
grid1.AddItem ""
fixGrd
End Sub
Private Sub myloadgrd2()
Dim cString As String
'                           0               1           2                   3               4                   5               6           7                   8                 9             10              11                  12
cString = "SELECT FILE0_90.MODELNO , file1_10.mosm , fact.desca , file1_10.modelfact0 ,file1_10.desca , FILE1_10.COLOR, FILE1_10.C_SCAL, FILE1_10.SCAL , FILE1_10.PRICE  , FILE0_90.DISC2 , FILE1_10.ITEM, FILE1_10.BARCODE13, FILE1_10.BARCODE_GS1 from (FILE0_90 inner join file1_10 on file1_10.MODELNO = FILE0_90.MODELNO ) inner join fact on fact.code = file1_10.fact"
cString = cString & turn(cString) & " DOC_NO = " & MyParn(xDoc_No.text)
cString = cString & " ORDER BY FILE0_90.MODELNO , FILE1_10.COLOR , FILE1_10.C_SCAL "
Set DATA2.Recordset = myRecordSet(cString, con)
Fixgrd2
End Sub

Private Function MYVALIDfix() As Boolean
If foundOther Then Exit Function
If xDoc_No.text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If

MYVALIDfix = True
End Function
Private Sub myUndo()
If IsNumeric(xDoc_No.text) Then
    openCardTable xDoc_No.text
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then myDefine Else myload
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xfilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter.text, 0
    End If
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not IsNumeric(.TextMatrix(Row, 0)) Then Exit Function
validRow = True
End With
End Function
Private Sub myAddItem()
grid1.Row = grid1.Row
grid1.Col = 8
grid1.Select grid1.Row, 8
grid1.AddItem ""
End Sub
Private Sub xdesca_GotFocus()
myGotFocus xDescA
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDescA
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Function openCardTable(Optional pDoc_no As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset

cString = "SELECT TOP 1 * from FILE0_90H "
If pSign = "=" Then
    If pDoc_no <> "" Then cWhere = "DOC_NO  " & pSign & addstring(pDoc_no)
Else
    If pDoc_no <> "" Then cWhere = "DOC_NO  " & pSign & addstring(pDoc_no)
End If
If sDoc_no <> "" Then cFilter = cFilter & turn(cFilter, " and ") & "DOC_NO = " & MyParn(sDoc_no)

' «·«Œ Ì«—« 
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by doc_no desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by doc_no ASC"
End If
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function
Private Sub CmdNext_Click()
openCardTable xDoc_No.text, ">"
If CardTable.EOF Then openCardTable xDoc_No.text
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xDoc_No.text, "<"
If CardTable.EOF Then openCardTable xDoc_No.text
myload
End Sub
Private Sub CmdFirst_Click()
openCardTable , ">"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
openCardTable , "<"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub
Private Sub GrdDesc(Row)
Dim nBalance As Double
Dim ItemTable As New ADODB.Recordset


aRet = aGetDesca("SELECT FILE1_10.MOSM , FACT.DESCA , FILE1_10.MODELFACT0 , FILE1_10.DESCA  FROM FILE1_10 INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT WHERE FILE1_10.MODELNO = " & MyParn(grid1.TextMatrix(Row, 0)), con)
If UBound(aRet) > 0 Then
With grid1
    grid1.TextMatrix(grid1.Row, 1) = aRet(1)
    grid1.TextMatrix(grid1.Row, 2) = aRet(2)
    grid1.TextMatrix(grid1.Row, 3) = aRet(3)
    grid1.TextMatrix(grid1.Row, 4) = aRet(4)
End With
End If
End Sub
Private Sub myloadgrd3()
    Dim cString As String
    '                           0               1           2                   3               4                   5               6           7                   8                 9             10              11                  12
    cString = "SELECT FILE1_10SC.CODE , file1_10SC.DESCA ,FILE1_10.FACT, fact.desca ,FILE1_10.MOSM,FILE0_90.DISC2, COUNT(DISTINCT MODEL)  from FILE0_90 inner join file1_10 on file1_10.MODELNO = FILE0_90.MODELNO inner join fact on fact.code = file1_10.fact LEFT JOIN FILE1_10SC ON FILE1_10SC.CODE = FILE1_10.[SECTION]"
    cString = cString & turn(cString) & " DOC_NO = " & MyParn(xDoc_No.text)
    cString = cString & " GROUP BY FILE1_10SC.CODE , file1_10SC.DESCA ,FILE1_10.FACT, fact.desca ,FILE1_10.MOSM,FILE0_90.DISC2 ORDER BY FILE1_10SC.CODE , file1_10SC.DESCA , fact.desca ,FILE0_90.DISC2"
    Set DATA3.Recordset = myRecordSet(cString, con)
    Fixgrd3
End Sub
Private Sub Fixgrd3()
Dim n1 As Double, n2 As Double
    With GRID3
    '                    0      1           2       3           4          5                6
    .FormatString = "ﬂÊœ|" & "«·ﬁ”„|" & "ﬂÊœ|" & "«·„’‰⁄|" & "«·„Ê”„|" & "«·Œ’„|" & "⁄œœ „ÊœÌ·« |"
    .WordWrap = True
    
    .RowHeight(0) = 800
    .ColHidden(0) = True
    .ColHidden(2) = True
    .Cols = 7
    .ColWidth(1) = 1700
    .ColWidth(3) = 1700
    .ColWidth(4) = 800
    .ColWidth(5) = 800
    .ColWidth(6) = 800
    
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ExplorerBar = flexExSortShow
    End With
End Sub
Private Function getData_org()
Dim cSv As New ChilkatCsv
Dim i As Long

Dim cFile As String

cFile = App.Path & "\csv\okaz.csv"

If Dir(cFile) = "" Then
    MsgBox "«·„·› €Ì— „ÊÃÊœ"
    Exit Function
End If

nAccess = cSv.LoadFile(cFile)

If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If


If cSv.NumRows < 2 Then Exit Function
Dim sb As New ChilkatStringBuilder

prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption

sb.Append "WITH tb AS ("
For i = 1 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & i & " „‰ " & cSv.NumRows - 1

    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    
    sb.Append IIf(i = 1, "", " UNION ALL ")
    sb.Append "SELECT " & cSv.GetCell(i, 0) & " AS ITEM,"
    sb.Append cSv.GetCell(i, 1) & " AS PRICE"
Next
Me.Caption = sCaption

sb.Append ")"
sb.Append vbCrLf

sb.Append "SELECT FILE1_10.MODEL,(100 * ROUND(1 - MIN(tb.PRICE/FILE1_10.PRICE),2))  AS Discount" & _
          " FROM FILE1_10 INNER JOIN tb ON FILE1_10.ITEM = tb.ITEM " & _
          " GROUP BY FILE1_10.MODEL"

Dim loctable As New ADODB.Recordset
Set loctable = mycmd(sb.GetAsString, con)

nRecordCount = loctable.RecordCount

Dim aInsert As Variant
con.BeginTrans
On Error GoTo myError
Dim nAffect As Long, nAffectAll As Long
i = 0
prog1.Value = 0
prog1.Visible = True
Do Until loctable.EOF
    i = i + 1
    Me.Caption = sCaption & " - " & "”Ã· " & i & " „‰ " & nRecordCount
    prog1.Value = Round(i / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
    aInsert = AddFlag(aInsert, "MODELNO", addstring(loctable!MODEL))
    aInsert = AddFlag(aInsert, "DISC2", loctable!discount)
    If IsEmpty(myField("SELECT DOC_NO FROM FILE0_90 WHERE DOC_NO = " & MyParn(xDoc_No.text) & " AND MODELNO = " & MyParn(loctable!MODEL), con)) Then
        con.Execute addInsert(aInsert, "FILE0_90"), nAffect
    Else
        con.Execute addUpdate(aInsert, "FILE0_90", "FILE0_90.DOC_NO = " & MyParn(xDoc_No.text) & " AND MODELNO = " & MyParn(loctable!MODEL)), nAffect
    End If
    nAffectAll = nAffectAll + nAffect
    loctable.MoveNext
Loop
con.CommitTrans
Me.Caption = sCaption
prog1.Visible = False
MsgBox " „  «÷«›… " & nAffectAll & " „ÊœÌ· »‰Ã«Õ"
getData = True
Finaly:
loctable.Close
Set loctable = Nothing
Exit Function
myError:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
GoTo Finaly
End Function
Private Function getData()
Dim cString As String
Dim i As Long

Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

Me.MousePointer = vbHourglass

Dim cFileName As String
cFileName = Common1.FileName
If cFileName = "" Then Exit Function

Dim cSv As New ChilkatCsv
nAccess = cSv.LoadFile(cFileName)

If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If


If cSv.NumRows < 2 Then Exit Function
Dim sb As New ChilkatStringBuilder

prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption

cString = "create table #Temp" & _
           "(" & _
           "ITEM int," & _
           "PRICE DECIMAL(18,2)" & _
           ")"
con.Execute cString

For Row = 1 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & Row & " „‰ " & cSv.NumRows - 1
    prog1.Value = Round(Row / (cSv.NumRows - 1), 2) * 100
    If IsNumeric(cSv.GetCell(Row, 0)) Then
        i = i + 1
        sb.Append "(" & _
                  cSv.GetCell(Row, 0) & "," & _
                  Val(cSv.GetCell(Row, 1)) & _
                  "),"
                  
    End If
    If i = 1000 Then
        i = 0
        sb.Shorten 1
        con.Execute "INSERT INTO #temp(ITEM,PRICE)" & _
                     "VALUES" & _
                     sb.GetAsString
        sb.Clear
    End If
Next

sb.Shorten 1
If sb.length > 0 Then
    con.Execute "INSERT INTO #temp(ITEM,PRICE)" & _
                 "VALUES" & _
                 sb.GetAsString
End If

cString = "SELECT FILE1_10.MODEL," & _
           "(100 * ROUND(1 - MIN(#temp.PRICE/FILE1_10.PRICE),2))  AS Discount" & _
          " FROM FILE1_10 INNER JOIN #temp ON FILE1_10.ITEM = #temp.ITEM " & _
          " GROUP BY FILE1_10.MODEL"

Dim loctable As New ADODB.Recordset
Set loctable = mycmd(cString, con)

nRecordCount = loctable.RecordCount

Dim aInsert As Variant

con.BeginTrans
On Error GoTo myError
Dim nAffect As Long, nAffectAll As Long
i = 0
prog1.Value = 0
prog1.Visible = True
Do Until loctable.EOF
    i = i + 1
    Me.Caption = sCaption & " - " & "”Ã· " & i & " „‰ " & nRecordCount
    prog1.Value = Round(i / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
    aInsert = AddFlag(aInsert, "MODELNO", addstring(loctable!MODEL))
    aInsert = AddFlag(aInsert, "DISC2", loctable!discount)
    If IsEmpty(myField("SELECT DOC_NO FROM FILE0_90 WHERE DOC_NO = " & MyParn(xDoc_No.text) & " AND MODELNO = " & MyParn(loctable!MODEL), con)) Then
        con.Execute addInsert(aInsert, "FILE0_90"), nAffect
    Else
        con.Execute addUpdate(aInsert, "FILE0_90", "FILE0_90.DOC_NO = " & MyParn(xDoc_No.text) & " AND MODELNO = " & MyParn(loctable!MODEL)), nAffect
    End If
    nAffectAll = nAffectAll + nAffect
    loctable.MoveNext
Loop
con.CommitTrans
Me.Caption = sCaption
prog1.Visible = False
MsgBox " „  «÷«›… " & nAffectAll & " „ÊœÌ· »‰Ã«Õ"
getData = True
Finaly:
clearTemp con
loctable.Close
Set loctable = Nothing
Exit Function
myError:
If TransCount(con) > 0 Then con.RollbackTrans
MsgBox Err.Description
Err.Clear
GoTo Finaly
End Function
