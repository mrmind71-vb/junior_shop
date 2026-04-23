VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form cash_salesfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”œ«œ ›« Ê—… „»Ì⁄« "
   ClientHeight    =   8145
   ClientLeft      =   105
   ClientTop       =   495
   ClientWidth     =   14895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   14895
   Begin VSFlex7Ctl.VSFlexGrid grdPhone 
      Height          =   2235
      Left            =   2295
      TabIndex        =   3
      Top             =   4005
      Visible         =   0   'False
      Width           =   6855
      _cx             =   12091
      _cy             =   3942
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
      BackColorBkg    =   16777215
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
      TabBehavior     =   0
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
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   3015
      Width           =   6945
      Begin VB.TextBox xE_mail 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Tag             =   "22"
         Top             =   990
         Visible         =   0   'False
         Width           =   5190
      End
      Begin VB.TextBox xName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "22"
         Top             =   585
         Width           =   5190
      End
      Begin VB.TextBox xPhone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   90
         MaxLength       =   11
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "22"
         Top             =   180
         Width           =   5190
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Email"
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
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1035
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«”„ «·⁄„Ì·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„  ·Ì›Ê‰ «·⁄„Ì·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1635
      Left            =   9180
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   1710
      Width           =   1770
      Begin VB.OptionButton XONLINE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«Ê‰ ·«Ì‰  Ê’Ì·"
         Enabled         =   0   'False
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1125
         Width           =   1545
      End
      Begin VB.OptionButton XONLINE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«Ê‰ ·«Ì‰ ‰ﬁœÏ"
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   630
         Width           =   1455
      End
      Begin VB.OptionButton XONLINE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "„»Ì⁄«  ‰ﬁœÏ"
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
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   180
         Value           =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   2805
      Left            =   10980
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1710
      Width           =   3840
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "€Ì— „”œœ"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2385
         Width           =   765
      End
      Begin VB.Label xtotal_install 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1980
         Width           =   1590
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬂ«‘"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   225
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰”»… «· ﬁ”Ìÿ"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   945
         Width           =   1020
      End
      Begin VB.Label xVisa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1260
         Width           =   1590
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·‰ﬁ«ÿ"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1710
         Width           =   435
      End
      Begin VB.Label xinstall 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   900
         Width           =   1590
      End
      Begin VB.Label xPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1620
         Width           =   1590
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì »⁄œ «· ﬁ”Ìÿ"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2025
         Width           =   1470
      End
      Begin VB.Label xCash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   180
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì «·›Ì“«"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1305
         Width           =   1020
      End
      Begin VB.Label xLate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   2340
         Width           =   1590
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·›Ì“«"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   585
         Width           =   420
      End
      Begin VB.Label xvisa_no_install 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   540
         Width           =   1590
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1035
      Width           =   8385
      Begin VB.TextBox xPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1590
      End
      Begin VB.Label xChange 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   180
         Width           =   1590
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·»«ﬁÌ"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„œ›Ê⁄ ‰ﬁœÌ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   9180
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3330
      Width           =   1770
      Begin Threed.SSCommand cmdExit 
         Height          =   465
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   675
         Width           =   1680
         _ExtentX        =   2963
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
         Picture         =   "cash_sales2.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   45
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
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
         Picture         =   "cash_sales2.frx":2323
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "cash_sales2.frx":4D18
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   8385
      Begin VB.Label xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·›« Ê—…"
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   1035
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   585
         Width           =   1590
      End
      Begin VB.Label Label1 
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
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   885
      End
      Begin VB.Label xDoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   1590
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì «·›« Ê—…"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   630
         Width           =   1200
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1035
      Top             =   45
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   360
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
      Top             =   360
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
      Caption         =   "data10"
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
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   0
      Top             =   360
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   2985
      Left            =   -7830
      TabIndex        =   35
      Top             =   -1710
      Visible         =   0   'False
      Width           =   8745
      _cx             =   15425
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
      BackColorSel    =   12648447
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      BackColorAlternate=   15921906
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3480
      Left            =   90
      TabIndex        =   1
      Top             =   4545
      Width           =   14730
      _cx             =   25982
      _cy             =   6138
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      BackColorAlternate=   15921906
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
Attribute VB_Name = "cash_salesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String, nTotal As Double, sDate As String, myform As salesfrm
Public bEditRecord As Boolean
Dim bAct As Boolean
Public nTag As Integer
Public nCash As Double
Dim oSearchVisa As New Search_abd
Dim oSearchBank As New Search_abd
Dim bIg As Boolean
Dim sLastControl As String
Dim nRound As Integer
Dim con As New ADODB.Connection
Const LoadMode = 0, DefineMode = 1
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not myValid Then Exit Sub
'If MsgBox("≈÷«›… «·”œ«œ ··›« Ê—… !! ‰⁄„ √„ ·« ", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub


If myreplaceGrd(, True) Then
    myform.xPay.Caption = xPay.text
    myform.xRest.Caption = xChange.Caption
    myform.xCash.Caption = xCash.Caption
    
    myform.xVisa.Caption = xvisa_no_install.Caption
    myform.XVISA2.Caption = xinstall.Caption
    myform.XPOINT.Caption = xPoints.Caption
    myform.xLate.Caption = xLate.Caption
    
    myform.xPhone.Caption = xPhone.text
    myform.XONLINE(0).Value = XONLINE(0).Value
    myform.XONLINE(1).Value = XONLINE(1).Value
    myform.XONLINE(2).Value = XONLINE(2).Value
        
    myform.lSave = True
    Unload Me
End If
End Sub
Private Sub cmdGo_Click()
myload
grid1.SetFocus
End Sub
Private Function myValidVisa() As Boolean
CalcTotals
If Val(xTotal.Caption) > 0 Then
    If Val(xPay.text) < 0 Then
        MsgBox "„»·€ „œ›Ê⁄ »«·”«·»"
        Exit Function
    End If
    
    If Val(xChange.Caption) >= 200 Then
        MsgBox "»«ﬁÌ «ﬂ»— „‰ 200 Ã‰ÌÂ"
        Exit Function
    End If
    
    If Val(xPay.text) >= Val(xTotal.Caption) And Val(xVisa.Caption) > 0 Then
         MsgBox "ﬁÌ„… «·„»·€ «·„œ›Ê⁄ «ﬂ»— „‰ ﬁÌ„… «·›« Ê—… „⁄ ”œ«œ ›Ì“«"
        Exit Function
    End If

    If Val(xvisa_no_install.Caption) > Val(xTotal.Caption) Then
         MsgBox "›Ì“« «ﬂ»— „‰ ﬁÌ„… «·›« Ê—…"
        Exit Function
    End If

    If Val(xvisa_no_install.Caption) = Val(xTotal.Caption) And Val(xPay.text) <> 0 Then
         MsgBox "›« Ê—… „œ›Ê⁄… »«·ﬂ«„· ›Ì“« »Â« ‰ﬁœÌ…"
        Exit Function
    End If
Else
    If Val(xvisa_no_install.Caption) < Val(xTotal.Caption) Then
         MsgBox "›Ì“« „”ÕÊ»… «ﬂ»— „‰ ﬁÌ„… «·›« Ê—…"
        Exit Function
    End If

    If Val(xLate.Caption) <> 0 And myform.xCode.text = cDefClient Then
         MsgBox "€Ì— „”„ÊÕ »¬Ã· ·⁄„Ì· ‰ﬁœÌ"
         Exit Function
    End If
End If

myValidVisa = True
End Function
Private Function myValid() As Boolean
If Val(xLate.Caption) <> 0 And myform.xCode.text = "0000" Then
     MsgBox "€Ì— „”„ÊÕ »¬Ã· ·⁄„Ì· ‰ﬁœÌ"
    Exit Function
End If


If Trim(xPhone.text) = "" Then
    If cBranch <> "" Then
        MsgBox " ·Ì›Ê‰ «·⁄„Ì· €Ì— „”Ã·"
        Exit Function
    End If
ElseIf Len(Trim(xPhone.text)) <> 11 Then
    MsgBox " ·Ì›Ê‰ «·⁄„Ì· " & Len(xPhone.text) & " —ﬁ„"
    Exit Function
End If

For i = 1 To grid1.Rows - 1
    If grid1.TextMatrix(i, 0) <> "" Then
        
        If grid1.TextMatrix(i, 8) = "" Then
            MsgBox "›Ì“« »œÊ‰ —ﬁ„ „” ‰œ"
            Exit Function
        End If
    
        If grid1.TextMatrix(i, 9) = "" Then
            MsgBox "›Ì“« »œÊ‰ —ﬁ„ „Ê«›ﬁ…"
            Exit Function
        End If
    End If
Next

If Not myValidVisa Then Exit Function
myValid = True
End Function

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    If grid1.Rows > 2 Then grid1.SetFocus
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()
openCon con

xCash.Enabled = bEditRecord
cmdSave.Enabled = bEditRecord

'CLIST = StrList2("SELECT * FROM VISA_CODES")

Set grid1.DataSource = data1
Set grdPhone.DataSource = data2
myload
End Sub
Public Sub myProc()
Dim nRest As Double
Dim nRateInstall As Double
Dim nRateVisa As Double

If grid1.Col = 1 Then
    grid1.TextMatrix(grid1.Row, 0) = oSearchBank.grid1.TextMatrix(oSearchBank.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 1) = oSearchBank.grid1.TextMatrix(oSearchBank.grid1.Row, 1)
   
    Unload oSearchBank
    nCount = mRound(myField("select count(*) from visa_codes where stoped = 0 and bank = " & MyParn(grid1.TextMatrix(grid1.Row, 0)), con))

    If nCount = 1 Then
        Dim loctable As New ADODB.Recordset
        Set loctable = mycmd("select * from visa_codes where stoped = 0 and bank = " & MyParn(grid1.TextMatrix(grid1.Row, 0)), con)
        If Not loctable.EOF Then
            If grid1.ValueMatrix(grid1.Row, 4) <> 0 Then
                nRest = mRound(grid1.ValueMatrix(grid1.Row, 4))
            Else
                nRest = mRound(xLate.Caption)
            End If
            
            nRateInstall = loctable!rate_install
            nRateVisa = loctable!Rate
            
            grid1.TextMatrix(grid1.Row, 2) = loctable!code
            grid1.TextMatrix(grid1.Row, 3) = loctable!desca
            
            If Val(xTotal.Caption) > 0 Then
                grid1.TextMatrix(grid1.Row, 4) = nRest
                grid1.TextMatrix(grid1.Row, 5) = nRateInstall
                grid1.TextMatrix(grid1.Row, 9 + 1) = nRateVisa
                grid1.TextMatrix(grid1.Row, 10 + 1) = IIf(loctable!with_points, 1, 0)
            Else
                nRest = Val(xCash.Caption) + mRound(grid1.ValueMatrix(grid1.Row, 4))
                grid1.TextMatrix(grid1.Row, 4) = nRest
                grid1.TextMatrix(grid1.Row, 5) = nRateInstall
                grid1.TextMatrix(grid1.Row, 9 + 1) = nRateVisa
                grid1.TextMatrix(grid1.Row, 10 + 1) = IIf(loctable!with_points, 1, 0)
            End If
            Grid1_AfterEdit grid1.Row, grid1.Col
            CellPos 13, grid1.Row, IIf(nRateInstall = 0, grid1.Cols - 1, 7)
        End If
    Else
        CellPos 13, grid1.Row, grid1.Col
    End If
Else
    grid1.TextMatrix(grid1.Row, 2) = oSearchVisa.grid1.TextMatrix(oSearchVisa.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 3) = oSearchVisa.grid1.TextMatrix(oSearchVisa.grid1.Row, 1)
    
    If grid1.ValueMatrix(grid1.Row, 4) <> 0 Then
        nRest = mRound(grid1.ValueMatrix(grid1.Row, 4))
    Else
        nRest = Val(xLate.Caption)
    End If
    
    nRateInstall = oSearchVisa.grid1.TextMatrix(oSearchVisa.grid1.Row, 4)
    nRateVisa = oSearchVisa.grid1.TextMatrix(oSearchVisa.grid1.Row, 5)
    
    If Val(xTotal.Caption) > 0 Then
        grid1.TextMatrix(grid1.Row, 4) = nRest
        grid1.TextMatrix(grid1.Row, 5) = nRateInstall
        grid1.TextMatrix(grid1.Row, 9 + 1) = nRateVisa
        grid1.TextMatrix(grid1.Row, 10 + 1) = IIf(oSearchVisa.grid1.ValueMatrix(oSearchVisa.grid1.Row, 6) = 0, 0, 1)
    Else
        nRest = Val(xCash.Caption) + mRound(grid1.ValueMatrix(grid1.Row, 4))
        grid1.TextMatrix(grid1.Row, 4) = nRest
        grid1.TextMatrix(grid1.Row, 5) = nRateInstall
        grid1.TextMatrix(grid1.Row, 9 + 1) = nRateVisa
        grid1.TextMatrix(grid1.Row, 10 + 1) = IIf(oSearchVisa.grid1.ValueMatrix(oSearchVisa.grid1.Row, 6) = 0, 0, 1)
    End If
    Unload oSearchVisa
    Grid1_AfterEdit grid1.Row, grid1.Col
    CellPos 13, grid1.Row, IIf(nRateInstall = 0, grid1.Cols - 1, 7)
End If
End Sub
Private Sub myload()
'Dim loctable As New ADODB.Recordset
'Set loctable = myCmd("[dbo].[sp_invoice_total]", con, adStoredProc, AddFlag(aPrm, "doc_no", sDoc_no))
'If loctable.EOF Then Exit Sub

xTotal.Caption = mRound(myform.xTotal.text)
xDoc_no.Tag = myform.xDoc_no.text
xDoc_no.Caption = myform.xDoc_no2.text
xDate.Caption = myFormat_p(myform.xDate.text)

bIg = True
xPhone.text = myform.xPhone.Caption
bIg = False

myform.XONLINE(0).Value = XONLINE(0).Value
myform.XONLINE(1).Value = XONLINE(1).Value
myform.XONLINE(2).Value = XONLINE(2).Value

xPay.Enabled = Val(xTotal.Caption) > 0

bIg = True
xPay.text = myform.xPay.Caption
bIg = False
If xPhone.text <> "" Then loadCust

myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1

CalcTotals
End Sub
Private Sub myLoadGrd()
Dim cString As String
cString = "SELECT VISA_CODES.BANK," & _
          "FILE5_10.DESCA," & _
          "FILE6_20_VISA.VISA_CODE," & _
          "VISA_CODES.DESCA," & _
          "FILE6_20_VISA.[VALUE]," & _
          "FILE6_20_VISA.RATE_INSTALL," & _
          "FILE6_20_VISA.VALUE_INSTALL," & _
          "FILE6_20_VISA.[VALUE] + FILE6_20_VISA.[VALUE_INSTALL]," & _
          "FILE6_20_VISA.DOC_VISA," & _
          "FILE6_20_VISA.APPROVE_NO," & _
          "FILE6_20_VISA.RATE_VISA," & _
          "VISA_CODES.WITH_POINTS," & _
          "FILE6_20_VISA.ID " & _
          " FROM FILE6_20_VISA INNER JOIN VISA_CODES ON FILE6_20_VISA.VISA_CODE = VISA_CODES.CODE" & _
          " INNER JOIN FILE5_10 ON VISA_CODES.BANK = FILE5_10.CODE" & _
          " WHERE FILE6_20_VISA.DOC_NO = " & MyParn(xDoc_no.Tag)

Set data1.Recordset = mycmd(cString, con)
grid1.AddItem ""
fixGrd
copyGrd grid1, grid2
End Sub
Private Sub fixGrd()
With grid1
'                   0          1             2                 3                 4                5                 6                  7                  8               9              10
.FormatString = "«·»‰ﬂ|" & "«·»‰ﬂ|" & "Ê”Ì·… «·”œ«œ|" & "Ê”Ì·… «·”œ«œ|" & "ﬁÌ„… «·›Ì“«|" & "‰”»… «· ﬁ”Ìÿ|" & "ﬁÌ„… «·‰”»…|" & "≈Ã„«·Ì «·›Ì“«|" & "„” ‰œ «·›Ì“«|" & "—ﬁ„ «·„Ê«›ﬁ…|" & "‰”»… «·»‰ﬂ|" & "‰ﬁ«ÿ|"


.WordWrap = True

.RowHeight(0) = 700

.ColWidth(0) = 1000
.ColWidth(1) = 2900
.ColWidth(2) = 1000
.ColWidth(3) = 2800
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 800
.ColWidth(7) = 1000
.ColWidth(8) = 2200
.ColWidth(9) = 2400

.ColFormat(5) = "##.##%"

.ColComboList(1) = "..."
.ColComboList(3) = "..."

.ColHidden(0) = True
.ColHidden(2) = True
'.ColHidden(9 + 1) = True

.ColHidden(.Cols - 3) = True
.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set cash_salesfrm = Nothing
closeCon con
End Sub
Private Sub grdPhone_LostFocus()
grdPhone.Visible = False
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If (Not myValidVisa) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    
    copyGrd grid2, grid1
    CalcTotals
    
    If Row < grid1.Rows - 1 Then
        grid1.Select Row, Col
    Else
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If

If Not validRow(Row) Then
    Exit Sub
End If

Dim bNew As Boolean
If Row = grid1.Rows - 1 Then
    myAddItem
    CalcTotals
End If
copyGrd grid1, grid2
'myreplaceGrd Row
'myLoadGrd
End Sub
Private Sub xfilter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xfilter.text, 2
End If
End Sub
Private Sub xfilterItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xfilterItem.text, 1
End If
End Sub
Private Function validRow(Row) As Boolean
If Trim(grid1.TextMatrix(Row, 0)) = "" Then Exit Function
If Trim(grid1.TextMatrix(Row, 2)) = "" Then Exit Function
If grid1.ValueMatrix(Row, 4) = 0 Then Exit Function
validRow = True
End Function
Private Sub CalcTotals(Optional bCalcRows As Boolean = True)
Dim nVisa_no_install  As Double, nInstall As Double
Dim nVisa  As Double, nPoints As Double

With grid1
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 6) = mRound(.ValueMatrix(i, 4) * .ValueMatrix(i, 5))
    grid1.TextMatrix(i, 7) = mRound(.ValueMatrix(i, 4) + .ValueMatrix(i, 6))
    If i < grid1.Rows - 1 Then
        If .ValueMatrix(i, .Cols - 2) = 0 Then
            nVisa_no_install = mRound(.ValueMatrix(i, 4)) + nVisa_no_install
            nInstall = mRound(.ValueMatrix(i, 6)) + nInstall
            nVisa = mRound(.ValueMatrix(i, 7)) + nVisa
        Else
            nPoints = mRound(.ValueMatrix(i, 7)) + nPoints
        End If
    End If
Next

xvisa_no_install.Caption = nVisa_no_install
xinstall.Caption = nInstall
xVisa.Caption = nVisa
xPoints.Caption = nPoints

If Val(xTotal.Caption) > 0 Then
    If mRound(Val(xTotal.Caption) - nVisa_no_install - nPoints) <= 0 Then
        xChange.Caption = ""
        xCash.Caption = ""
    ElseIf Val(xPay.text) > mRound(Val(xTotal.Caption) - nVisa_no_install - nPoints) Then
        xChange.Caption = mRound(mRound(xPay.text) - (mRound(xTotal.Caption) - nVisa_no_install - nPoints))
        xCash.Caption = mRound(mRound(xTotal.Caption) - nVisa_no_install - nPoints)
    Else
        xChange.Caption = ""
        xCash.Caption = xPay.text
    End If
Else
    xCash.Caption = mRound(Val(xTotal.Caption) - nVisa_no_install - nPoints)
End If

xtotal_install.Caption = mRound(xTotal.Caption) + nInstall
xLate.Caption = mRound(Val(xTotal.Caption) - Val(xCash.Caption) - mRound(xvisa_no_install.Caption) - Val(xPoints.Caption))
End With
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Col = 1 And Row = grid1.Rows - 1 And grid1.TextMatrix(Row, 3) = "" Then
    If Val(xTotal.Caption) < 0 And Val(xvisa_no_install.Caption) = Val(xTotal.Caption) And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then Exit Sub
    If Val(xLate.Caption) = 0 And Val(xTotal.Caption) > 0 And grid1.TextMatrix(grid1.Row, 0) = "" Then Exit Sub
    BankVisaLookup Me, oSearchBank
ElseIf Col = 3 And grid1.TextMatrix(Row, 0) <> "" Then
    If Val(xTotal.Caption) < 0 And Val(xvisa_no_install.Caption) = Val(xTotal.Caption) And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then Exit Sub
    If Val(xLate.Caption) = 0 And Val(xTotal.Caption) > 0 And grid1.TextMatrix(grid1.Row, 2) = "" Then Exit Sub
    visaLookup Me, oSearchVisa, grid1.TextMatrix(Row, 0), "VISA_CODES.STOPED = 0"
End If
End Sub
Private Sub grid1_EnterCell()
If Not bEditRecord Then
    grid1.Editable = flexEDNone
ElseIf grid1.Col = 1 Or grid1.Col = 3 Or grid1.Col = 4 Or grid1.Col = 8 Or grid1.Col = 9 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row > 0 And bEditRecord Then
    If grid1.Row <> grid1.Rows - 1 Then
        If MsgBox("Õ–› ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
            grid1.RemoveItem grid1.Row
            CalcTotals
        End If
    Else
        For Col = 0 To grid1.Cols - 1
            grid1.TextMatrix(grid1.Row, Col) = ""
        Next
        CalcTotals
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid1.Col = 1 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    If grid1.Col = 3 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If Col = 1 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    If Col = 3 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < 4 Then
     grid1.Col = NextEmpty(grid1, Row, Col + 1, 4)
ElseIf Col < 8 Then
     grid1.Col = 8
ElseIf Col = 8 Then
    grid1.Col = 9
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 4)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 4 Then
    If Val(grid1.EditText) = 0 Then
        Cancel = True
        Exit Sub
    End If
    
    If Val(xTotal.Caption) > 0 And Val(grid1.EditText) <= 0 Then
        Cancel = True
        Exit Sub
    ElseIf Val(xTotal.Caption) < 0 And Val(grid1.EditText) >= 0 Then
        Cancel = True
        Exit Sub
    End If
    
'    If Val(grid1.EditText) <> grid1.ValueMatrix(row, col) Then
'        grid1.TextMatrix(grid1.row, 6) = mRound(mRound(grid1.EditText) * grid1.ValueMatrix(row, 5))
'        grid1.TextMatrix(grid1.row, 7) = mRound(grid1.ValueMatrix(row, 5) + grid1.ValueMatrix(row, 6))
'    End If
End If
End Sub

Private Sub xcash_GotFocus()
myGotFocus xCash
End Sub
Private Sub xcash_LostFocus()
myLostFocus xCash
End Sub
Private Sub myAddItem()
grid1.AddItem ""
End Sub
Private Sub sendData()
'Dim Row As Long, Col As Long
'myform.xCash.text = xCash.text
'With myform.grdVisa
'    .Rows = 1
'    For Row = 1 To grid1.Rows - 2
'        .AddItem ""
'        For Col = 0 To grid1.Cols - 1
'            .TextMatrix(Row, Col) = grid1.TextMatrix(Row, Col)
'        Next
'    Next
'    myform.CalcTotals
'End With
End Sub

Private Sub Label15_Click()

End Sub

Private Sub Text1_Change()

End Sub
Private Sub xPay_Change()
If Not bIg Then
    CalcTotals
End If
End Sub
Private Function myreplaceGrd(Optional Row As Long = -1, Optional bDone As Boolean = False) As Boolean
Dim aInsert As Variant
Dim nAffect As Long
With grid1
con.BeginTrans

On Error GoTo myError
Dim sid As String
If bDone Then
    con.Execute "UPDATE FILE6_20H SET " & _
                "FILE6_20H.PAY = " & Val(xPay.text) & "," & _
                "FILE6_20H.REST = " & Val(xChange.Caption) & "," & _
                "FILE6_20H.CASH = " & Val(xCash.Caption) & "," & _
                "FILE6_20H.VISA = " & Val(xvisa_no_install.Caption) & "," & _
                "FILE6_20H.VISA2 = " & Val(xinstall.Caption) & "," & _
                "FILE6_20H.POINT = " & Val(xPoints.Caption) & "," & _
                "FILE6_20H.LATE = " & Val(xLate.Caption) & "," & _
                "FILE6_20H.PHONE = " & addstring(xPhone.text) & "," & _
                "FILE6_20H.ISNEW = 1 " & _
                " FROM FILE6_20H WHERE DOC_NO = " & MyParn(xDoc_no.Tag)

                aInsert = AddFlag(Empty, "desca", addstring(xName.text))
                aInsert = AddFlag(aInsert, "isnew", "1")
                con.Execute addUpdate(aInsert, "SUBCUST", "phone = " & MyParn(xPhone.text)), nAffect
                If nAffect = 0 Then
                    aInsert = AddFlag(aInsert, "phone", addstring(xPhone.text))
                    aInsert = AddFlag(aInsert, "e_mail", addDate(xPhone.text))
                    aInsert = AddFlag(aInsert, "branch", addstring(cBranch))
                    aInsert = AddFlag(aInsert, "f_Date", addDate(Date))
                    con.Execute addInsert(aInsert, "SUBCUST")
                End If
                
    con.Execute " update file6_20h set phone = " & addstring(xPhone.text) & " where doc_no = " & MyParn(xDoc_no.Tag)
End If
For i = 1 To grid1.Rows - 2
    If .TextMatrix(i, .Cols - 1) <> "" Then
        sid = sid & IIf(sid = "", "", ",") & .TextMatrix(i, .Cols - 1)
    End If
Next
If sid <> "" Then
    con.Execute "DELETE  FROM FILE6_20_VISA " & _
                " FROM FILE6_20_VISA  LEFT JOIN dbo.splitTable(" & MyParn(sid) & ",default) AS t  ON FILE6_20_VISA.ID = t.ITEM" & _
                " Where t.Item Is Null AND FILE6_20_VISA.DOC_NO = " & MyParn(xDoc_no.Tag)
Else
    con.Execute "DELETE  FROM FILE6_20_VISA WHERE FILE6_20_VISA.DOC_NO = " & MyParn(xDoc_no.Tag)
End If
    
For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
    aInsert = AddFlag(Empty, "[DOC_NO]", addstring(xDoc_no.Tag))
    aInsert = AddFlag(aInsert, "[VISA_CODE]", addvalue(grid1.TextMatrix(i, 2)))
    aInsert = AddFlag(aInsert, "[VALUE]", grid1.ValueMatrix(i, 4))
    aInsert = AddFlag(aInsert, "[RATE_INSTALL]", grid1.ValueMatrix(i, 5))
    aInsert = AddFlag(aInsert, "[VALUE_INSTALL]", grid1.ValueMatrix(i, 6))
    aInsert = AddFlag(aInsert, "[DOC_VISA]", addstring(grid1.TextMatrix(i, 8)))
    aInsert = AddFlag(aInsert, "[APPROVE_NO]", addstring(grid1.TextMatrix(i, 9)))
    aInsert = AddFlag(aInsert, "[RATE_VISA]", grid1.ValueMatrix(i, 10))
    If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
        con.Execute addInsert(aInsert, "FILE6_20_VISA")
    Else
        con.Execute addUpdate(aInsert, "FILE6_20_VISA", "ID = " & grid1.TextMatrix(i, .Cols - 1))
    End If
Next
End With
con.CommitTrans
myreplaceGrd = True
Exit Function
myError:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function
Private Sub xPay_GotFocus()
myGotFocus xPay
End Sub
Private Sub xPay_LostFocus()
myLostFocus xPay
End Sub
Private Sub FixGrdPhone()
grdPhone.RowHidden(0) = True
grdPhone.ColWidth(0) = 2000
grdPhone.ColWidth(1) = 2500
grdPhone.ColWidth(2) = 1000
grdPhone.ColWidth(3) = 1300
'grdPhone.ColHidden(2) = True
If grdPhone.Rows > 1 Then
    grdPhone.Cell(flexcpBackColor, 1, 2, grdPhone.Rows - 1, 2) = &H8000000F
    GrdHi grdPhone, 2, -1, grdPhone.Row
End If
For i = 0 To grdPhone.Cols - 1
    grdPhone.ColAlignment(i) = flexAlignRightCenter
Next
End Sub
Private Sub grdPhone_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    grdPhone_DblClick
ElseIf KeyCode = 27 Then
    grdPhone.Visible = False
End If
End Sub
Private Sub grdPhone_DblClick()
bIg = True
xPhone.text = grdPhone.TextMatrix(grdPhone.Row, 0)
xName.text = grdPhone.TextMatrix(grdPhone.Row, 1)
xE_mail.text = grdPhone.TextMatrix(grdPhone.Row, 2)
grdPhone.Visible = False
bIg = False
xName.SetFocus
End Sub
Private Sub gridPhone_LostFocus()
grdPhone.Visible = (ActiveControl.Name = xPhone.Name)
End Sub

Private Sub xPhone_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If grdPhone.Visible Then grdPhone.Visible = False
End If
End Sub
Private Sub XPHONE_LostFocus()
'sLastControl = xPhone.Name
myLostFocus xPhone

If Trim(xPhone.text) <> "" Then
    If IsValidMobile(xPhone.text) Then
        xPhone.text = ""
    End If
End If

grdPhone.Visible = (ActiveControl.Name = grdPhone.Name)
End Sub
Private Sub xPhone_Change()
If bIg Then Exit Sub
myLoadList , xPhone.text
End Sub
Private Sub myLoadList(Optional pName As String = "", Optional pPhone As String = "")
Dim cString As String
If Not grdPhone.Visible Then grdPhone.Visible = True
cString = "Select top 100 Phone,desca,e_mail,format(f_date,'yyyy/M/d') From SUBCUST"
If Trim(pName) <> "" Then
    cString = cString & " WHERE desca like " & MyParn(pName & "%")
ElseIf Trim(pPhone) <> "" Then
    cString = cString & " WHERE PHONE like " & MyParn(pPhone & "%")
Else
    cString = cString & " WHERE (NOT (desca IS NULL AND PHONE IS NULL )) "
End If
cString = cString & " order by " & IIf(Trim(pName) <> "", "[desca]", "Phone")
Set data2.Recordset = mycmd(cString, con)
FixGrdPhone
If grdPhone.Rows > 1 And grdPhone.Row < 1 Then
    grdPhone.Row = 1
ElseIf grdPhone.Rows = 1 Then
    grdPhone.Visible = False
End If
End Sub
Private Sub loadCust()
Dim loctable As New ADODB.Recordset
Set loctable = mycmd("select top 1 * from SUBCUST WHERE PHONE = " & MyParn(xPhone.text), con)
bIg = True
If Not loctable.EOF Then
    xName.text = loctable!desca & ""
    'xE_mail.text = loctable!E_MAIL & ""
End If
bIg = False
loctable.Close
Set loctable = Nothing
End Sub
Public Sub GrdHi(pGrd As Object, pCol As Long, pRow1 As Long, pRow2 As Long, Optional pColor As Long = &HC0FFFF)
    If pCol > 0 Then
        If pRow1 <> -1 Then pGrd.Cell(flexcpBackColor, pRow1, 0, pRow1, pCol - 1) = pGrd.BackColor
        If pRow2 <> -1 Then pGrd.Cell(flexcpBackColor, pRow2, 0, pRow2, pCol - 1) = pColor
    End If
    
    If pCol < pGrd.Cols - 1 Then
        If pRow1 <> -1 Then pGrd.Cell(flexcpBackColor, pRow1, pCol + 1, pRow1, pGrd.Cols - 1) = pGrd.BackColor
        If pRow2 <> -1 Then pGrd.Cell(flexcpBackColor, pRow2, pCol + 1, pRow2, pGrd.Cols - 1) = pColor
    End If
End Sub
Private Sub grdPhone_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow Then
    GrdHi grdPhone, 2, OldRow, NewRow
End If
End Sub

Private Sub xE_mail_GotFocus()
myGotFocus xE_mail
End Sub
Private Sub xE_mail_LostFocus()
myLostFocus xE_mail
End Sub
Private Sub xName_GotFocus()
myGotFocus xName
End Sub
Private Sub xname_LostFocus()
myLostFocus xName
End Sub
Private Sub xPhone_GotFocus()
myGotFocus xPhone
End Sub

