VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form stock_appFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20400
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
   ScaleHeight     =   11085
   ScaleWidth      =   20400
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5805
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   -45
      Width           =   4785
      Begin VB.TextBox txtModel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   270
         Width           =   1320
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2430
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ „ÊœÌ·"
         Height          =   285
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   315
         Width           =   825
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "»«—ﬂÊœ"
         Height          =   285
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   315
         Width           =   645
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   585
      Width           =   1770
      Begin Threed.SSCommand cmdCSVadd 
         Height          =   555
         Left            =   45
         TabIndex        =   67
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   979
         _Version        =   196610
         CaptionStyle    =   1
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
         Picture         =   "stock_app.frx":0000
         Caption         =   "«÷«›… „‰ „·› "
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   1
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":25D3
      End
      Begin Threed.SSCommand cmdCSVdel 
         Height          =   555
         Left            =   45
         TabIndex        =   68
         Top             =   765
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   979
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   192
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
         Picture         =   "stock_app.frx":4A56
         Caption         =   "Õ–› „‰ „·› "
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   1
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":723A
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   540
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   -450
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CheckBox xclosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "„€·ﬁ"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   315
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   1305
      Width           =   1590
      Begin VB.CheckBox xAll 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "All Zones"
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
         Height          =   360
         Left            =   90
         TabIndex        =   63
         Top             =   225
         Width           =   1455
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6675
      Left            =   5355
      TabIndex        =   4
      Top             =   1980
      Width           =   14730
      _cx             =   25982
      _cy             =   11774
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
      FixedCols       =   1
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
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   8685
      Width           =   5325
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "€Ì— „ÊÃÊœ »«·Ã—œ"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   180
         Width           =   1680
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«’‰«› „ÊÃÊœ… »«·Ã—œ"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   180
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·ﬂ·"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   180
         Value           =   -1  'True
         Width           =   780
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   10620
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   0
      Width           =   1545
      Begin Threed.SSCommand cmdDiffer 
         Height          =   510
         Left            =   45
         TabIndex        =   21
         Top             =   135
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   " ”ÊÌ… «·Ã—œ"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   12150
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   7935
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   6615
         TabIndex        =   12
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
         Picture         =   "stock_app.frx":9790
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "stock_app.frx":BB5B
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   5310
         TabIndex        =   13
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
         Picture         =   "stock_app.frx":DC04
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "stock_app.frx":FC0C
      End
      Begin Threed.SSCommand cmddelInv 
         Height          =   510
         Left            =   3960
         TabIndex        =   14
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
         Picture         =   "stock_app.frx":11BC3
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "stock_app.frx":1435F
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   15
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
         Picture         =   "stock_app.frx":167F3
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   510
         Left            =   2610
         TabIndex        =   56
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
         Picture         =   "stock_app.frx":18B16
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1350
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   900
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
         Picture         =   "stock_app.frx":1AF80
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":1D2F6
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   9630
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   630
      Width           =   10455
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6660
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   2805
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   330
         Left            =   6660
         TabIndex        =   2
         Top             =   540
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8100
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3780
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1545
      End
      Begin Threed.SSCommand cmdAddZone 
         Height          =   330
         Left            =   1800
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   900
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
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
         Picture         =   "stock_app.frx":1F479
         ButtonStyle     =   3
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":21790
      End
      Begin Threed.SSCommand cmdEditZone 
         Height          =   330
         Left            =   1260
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   " ⁄œÌ·"
         Top             =   900
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "stock_app.frx":239E2
         ButtonStyle     =   3
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":25CED
      End
      Begin Threed.SSCommand cmdDelZone 
         Height          =   330
         Left            =   720
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   900
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "stock_app.frx":27F46
         ButtonStyle     =   3
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":2A2C9
      End
      Begin Threed.SSCommand cmdDelApi 
         Height          =   330
         Left            =   1260
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   495
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "stock_app.frx":2C577
         ButtonStyle     =   3
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":2E8FA
      End
      Begin Threed.SSCommand cmdMerge 
         Height          =   330
         Left            =   90
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "œ„Ã"
         Top             =   900
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "stock_app.frx":30BA8
         ButtonStyle     =   3
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":3300D
      End
      Begin Threed.SSCommand cmdApi 
         Height          =   330
         Left            =   1800
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   495
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "stock_app.frx":35352
         ButtonStyle     =   3
         BevelWidth      =   0
      End
      Begin VB.Label xZone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         TabIndex        =   38
         Top             =   900
         Width           =   2940
      End
      Begin VB.Label xDoc_App_name 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         TabIndex        =   37
         Top             =   540
         Width           =   2940
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·“Ê‰"
         Height          =   270
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   945
         Width           =   435
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„” ‰œ «· ÿ»Ìﬁ"
         Height          =   270
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   585
         Width           =   1035
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "»Ì«‰"
         Height          =   240
         Left            =   9585
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   945
         Width           =   525
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Left            =   5415
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ „” ‰œ"
         Height          =   240
         Left            =   9585
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„Œ“‰"
         Height          =   270
         Left            =   9585
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   420
      End
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   -810
      Top             =   810
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
      Left            =   -1215
      Top             =   3510
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
   Begin MSAdodcLib.Adodc data7 
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   16605
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   8685
      Width           =   3480
      Begin Threed.SSCommand cmdUnpost 
         Height          =   420
         Left            =   45
         TabIndex        =   22
         Top             =   135
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   741
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
         Picture         =   "stock_app.frx":3766E
         Caption         =   " «·€«¡ «· —ÕÌ·"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   4
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "stock_app.frx":39A71
      End
      Begin Threed.SSCommand cmdPost 
         Height          =   420
         Left            =   1755
         TabIndex        =   23
         Top             =   135
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   741
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
         Picture         =   "stock_app.frx":3BD6D
         Caption         =   "  —ÕÌ· «·Ã—œ"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   4
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "stock_app.frx":3E18E
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   10710
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
         TabIndex        =   32
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
         TabIndex        =   33
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
         TabIndex        =   34
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   6675
      Left            =   45
      TabIndex        =   36
      Top             =   1980
      Width           =   5280
      _cx             =   9313
      _cy             =   11774
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
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
      AutoSearch      =   2
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
      ComboSearch     =   1
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc DATA12 
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   44
      Top             =   10560
      Visible         =   0   'False
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   13455
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   8640
      Width           =   3075
      Begin Threed.SSCommand cmdBranch 
         Height          =   420
         Left            =   45
         TabIndex        =   49
         Top             =   180
         Width           =   2985
         _ExtentX        =   5265
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
         Caption         =   "«Œ «— «·›—⁄"
         TagVariant      =   "«Œ «— «·„Œ“‰"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame Frame4 
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
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   8640
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   51
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
         Picture         =   "stock_app.frx":404A3
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":4264A
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   52
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
         Picture         =   "stock_app.frx":44691
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":4677C
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   53
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
         Picture         =   "stock_app.frx":48776
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":4A887
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   54
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
         Picture         =   "stock_app.frx":4C881
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "stock_app.frx":4EAA5
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   5715
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   585
      Width           =   2310
      Begin Threed.SSCommand cmdDelAdd 
         Height          =   555
         Left            =   45
         TabIndex        =   25
         Top             =   765
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   979
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   192
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
         Caption         =   "Õ–› «’‰«› »œÊ‰ —’Ìœ"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAddZero 
         Height          =   555
         Left            =   45
         TabIndex        =   26
         Top             =   180
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
         Caption         =   "√÷«›… «’‰«› »œÊ‰ —’Ìœ"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Index           =   4
      Left            =   8055
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   630
      Width           =   1545
      Begin Threed.SSCommand cmdSave 
         Height          =   555
         Left            =   45
         TabIndex        =   18
         Top             =   135
         Width           =   1455
         _ExtentX        =   2566
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
         Picture         =   "stock_app.frx":50B76
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "stock_app.frx":5349B
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   555
         Left            =   45
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
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
         Picture         =   "stock_app.frx":55CEF
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "stock_app.frx":57E4F
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   8640
      Width           =   1320
      Begin VB.CheckBox chkBarCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "»«·»«—ﬂÊœ"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   180
         Width           =   1050
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label XDOC_NO_FLAG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -900
      TabIndex        =   55
      Top             =   270
      Visible         =   0   'False
      Width           =   1995
   End
End
Attribute VB_Name = "Stock_appFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim bAct As Boolean
Dim bZone As Boolean
Dim oSearch As New Search_abd
Dim oSearchItem As New Search_abd
Dim oSearchBranch As New Search_abd
Dim oSearchApi As New Search_abd
Dim oSearchCSV As New Search_abd
Dim bStopCell As Boolean
Dim bEditRecord As Boolean
Dim CardTable As ADODB.Recordset
Dim oZoneAdd As stock_zone_addfrm
Dim oZoneEdit As stock_zonefrm
Dim oZoneMerge As stock_zone_mergefrm
Dim oStockZero As stock_zero
Dim oStockApi As New stock_doc_api
Dim cFilter As String
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Public Function myreplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[STORE]", addstring(xStore.BoundText))
aInsert = AddFlag(aInsert, "[DESCA]", addstring(xDesca.text))
aInsert = AddFlag(aInsert, "[IS_ZONE]", "1")
aInsert = AddFlag(aInsert, "[ISDATE]", "GETDATE()")
aInsert = AddFlag(aInsert, "[BRANCH]", addstring(cBranch))
con.BeginTrans
On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    xDoc_No.text = RetZero(NewflagBranch6("FILE0_10H", "DOC_NO", cBranch, con), 6)
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
    con.Execute addInsert(aInsert, "FILE0_10H")
Else
    cmd(addUpdate(aInsert, "FILE0_10H", "doc_no = N" & addstring(xDoc_No.text)), con).Execute
End If
If Row <> 0 Then
    myreplaceGrd Row
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
If prog1.Visible Then prog1.Visible = False
MsgBox Err.Description
If TransCount(con) > 0 Then
    con.RollbackTrans
End If
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 1), , 0)
    If nFound <> -1 Then
        Inform "«·’‰› „ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ " & nFound
        CellPos 13, nFound - 1, grid1.Cols - 1
        Unload oSearchItem
        Exit Sub
    End If


    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    Dim sMsg As String
    If Not GrdDesc(grid1.TextMatrix(grid1.Row, 1), grid1.Row, True, sMsg) Then
        Exit Sub
    End If
    
    Grid1_AfterEdit grid1.Row, grid1.col
    
    
    If bNew Then
        If chkBarCode.value = 1 Then
            CellPos 13, grid1.Rows - 2, grid1.Cols - 1
        Else
            CellPos 13, grid1.Row, grid1.col
            Unload oSearchItem
        End If
    End If
ElseIf ActiveControl.Name = cmdInform.Name Then
    openCardTable tbMode.tbFind, oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
ElseIf ActiveControl.Name = cmdBranch.Name Then
    If oSearchBranch.grid1.TextMatrix(oSearchBranch.grid1.Row, 0) = "" Then
        cmdBranch.Tag = ""
        cmdBranch.Caption = cmdBranch.TagVariant
    Else
        cmdBranch.Tag = oSearchBranch.grid1.TextMatrix(oSearchBranch.grid1.Row, 0)
        cmdBranch.Caption = oSearchBranch.grid1.TextMatrix(oSearchBranch.grid1.Row, 1)
    End If
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
    Unload oSearchBranch
ElseIf ActiveControl.Name = cmdDelApi.Name Then
    If MsgBox("Õ–› ﬂ· ”Ã·«  „” ‰œ «· ÿ»Ìﬁ „‰ „” ‰œ «·Ã—œ", vbOKCancel + vbDefaultButton2) Then
        Me.MousePointer = vbHourglass
        If delApp(oSearchApi.grid1.TextMatrix(oSearchApi.grid1.Row, 0)) Then
            MsgBox " „ Õ–› ”Ã·«  „” ‰œ «· ÿ»Ìﬁ »‰Ã«Õ"
            Unload oSearchApi
        End If
        Me.MousePointer = vbNormal
        myloadgrd2
        myLoadGrd
    End If
ElseIf ActiveControl.Name = cmdCSVdel.Name Then
    If MsgBox("Õ–› ﬂ· ”Ã·«  «·„·› „‰ „” ‰œ «·Ã—œ", vbOKCancel + vbDefaultButton2) Then
        Me.MousePointer = vbHourglass
        If delCSV(oSearchCSV.grid1.TextMatrix(oSearchCSV.grid1.Row, 0)) Then
            MsgBox " „ Õ–› ”Ã·«  «·„·› »‰Ã«Õ"
            Unload oSearchCSV
        End If
        Me.MousePointer = vbNormal
        myloadgrd2
        myLoadGrd
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Check3_Click()
    'addSetting "Check3", Check3.Value, App.Path & "\other.txt"
End Sub
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub CMD_FINDBARCODE_Click()
    FrmSubStock.Show
End Sub
Private Sub cmdAddZero_Click()
If xDoc_No.Tag = DefineMode Then
    MsgBox "«·—Ã«¡ Õ›Ÿ «·„·› «Ê·«"
    Exit Sub
End If
Set oStockZero = New stock_zero
Set oStockZero.myForm = Me
Set oStockZero.con = con
oStockZero.Show 1
End Sub

Private Sub cmdAddZone_Click()
Set oZoneAdd = Nothing
Set oZoneAdd = New stock_zone_addfrm
Set oZoneAdd.con = con
Set oZoneAdd.myForm = Me
oZoneAdd.pDoc_no = xDoc_No.text
oZoneAdd.Show 1
End Sub

Private Sub cmdApi_Click()
If Not myValid Then Exit Sub
cmdBranch.Tag = ""
cmdBranch.Caption = cmdBranch.TagVariant

If xDoc_No.Tag = DefineMode Then
    MsgBox "«·—Ã«¡ Õ›Ÿ «·„·› «Ê·«"
    Exit Sub
End If

Set oStockApi = New stock_doc_api
Set oStockApi.con = con
Set oStockApi.myForm = Me
oStockApi.Show 1
End Sub

Private Sub cmdBranch_Click()
StoreLookup Me, oSearchBranch, , , IIf(cmdBranch.Tag <> "", "ﬂ· «·„Œ«“‰", "")
End Sub

Private Sub cmdCSVadd_Click()
If getCSV Then
    myloadgrd2
    myLoadGrd
End If
End Sub

Private Sub cmdCSVdel_Click()
docCSVLookup
End Sub

Private Sub cmdDelAdd_Click()
If MsgBox("Õ–› «’‰«› €Ì— „ÊÃÊœ… »«·Ã—œ", vbOKCancel + vbDefaultButton) <> vbOK Then Exit Sub

Me.MousePointer = vbHourglass
cmdDelAdd.Enabled = False
On Error GoTo myerror
Dim nRecords As Long

con.Execute "Delete from file0_100 " & _
            " WHERE DOC_NO = " & MyParn(xDoc_No.text) & _
            " AND [AUTO] = 1", nRecords

If nRecords > 0 Then
    Inform "  „ Õ–› " & nRecords & " ”Ã·"
    myloadgrd2
    myLoadGrd
End If
Finally:
Me.MousePointer = vbNormal
cmdDelAdd.Enabled = True
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
GoTo Finally
End Sub

Private Sub cmdDelApi_Click()
docApiLookup
End Sub

Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        Me.MousePointer = vbHourglass
        CmdDelInv.Enabled = False
        con.BeginTrans
        con.Execute " Delete  From FILE0_100 where Doc_No = " & MyParn(xDoc_No.text)
        con.Execute " Delete  From FILE0_10 where Doc_No = " & MyParn(xDoc_No.text)
        con.Execute " Delete  From FILE0_10H where Doc_No = " & MyParn(xDoc_No.text)
        con.CommitTrans
        
        AddLod_Data cUserName & " (" & GetComputerName & ")", 2, " Õ–›  " & Me.Caption, con, xDoc_No.text, xDate.text, , xDesca.text
        
        Me.MousePointer = vbNormal
        Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
    
        If Not openCardTable(tbMode.tbPrevious, xdoc_no_flag.Caption) Then
            If Not openCardTable(tbMode.tbFirst) Then
                myDefine
            End If
        End If
    End If
Finally:
    Me.MousePointer = vbNormal
    CmdDelInv.Enabled = True
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    If TransCount(con) > 0 Then con.RollbackTrans
    'con.RollbackTrans
End Sub
Private Sub cmdDelZone_Click()
    If MsgBox("Õ–› √’‰«› «·“Ê‰ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
        
    Dim Row As Long
    Row = grid2.Row
    
    On Error GoTo myerror
    Me.MousePointer = vbHourglass
    con.Execute "Delete  From FILE0_100 " & _
                " where Doc_No = " & MyParn(xDoc_No.text) & _
                " and ZONE = " & MyParn(xZone.Caption)
    Me.MousePointer = vbNormal
    Inform " „ Õ–› √’‰«› «·“Ê‰ »‰Ã«Õ"
    
    myloadgrd2
    myLoadGrd
    Handlecontrols LoadMode
        
Finally:
    Me.MousePointer = vbNormal
    CmdDelInv.Enabled = True
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub cmdDiffer_Click()
Set ShowStockDiffer.con = con
ShowStockDiffer.sDoc_no = xDoc_No.text
ShowStockDiffer.sDate = xDate.text
ShowStockDiffer.sStore = xStore.text
ShowStockDiffer.Show 1
End Sub
Private Sub cmdEditZone_Click()
Set oZoneEdit = Nothing
Set oZoneEdit = New stock_zonefrm
Set oZoneEdit.con = con
Set oZoneEdit.myForm = Me
oZoneEdit.pDoc_no = xDoc_No.text
oZoneEdit.pZone = xZone.Caption
oZoneEdit.Show 1
End Sub

Private Sub cmdExcel_Click()
Me.MousePointer = 11

Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aRow = AddFlag(aRow, aSub)


ToFileExelNew grid1, , , aRow, , 0.9, , , , , , Me, Array("Ã—œ „Œ“‰ : " & xStore.text, "» «—ÌŒ : " & xDate.text, "Zone : " & xZone.Caption)

Me.MousePointer = 0
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFix_Click()

End Sub

Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO, FORMAT([DATE],'yyyy/M/d'),FILE0_40.DESCA,FILE0_10H.DESCA ,FILE0_10H.CLOSED " & _
                  " FROM (FILE0_10H INNER JOIN FILE0_40 ON FILE0_10H.Store = FILE0_40.CODE)"

Dim cWhere As String
If lIsBranchStore Then
    cWhere = " AND STORE = " & MyParn(cBranchStore)
End If

If cFilter <> "" Then
    cWhere = cWhere & " AND " & cFilter
End If

If cWhere <> "" Then
    Generalarray(1) = Generalarray(1) & " WHERE " & Mid(cWhere, 6)
End If

Generalarray(2) = "Order by Date DESC , DOC_NO DESC "
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ-«·»Ì«‰"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]## or %%FILE0_10h.DESCA%%)"

listarray(1, 0) = "«·›—⁄"
listarray(1, 1) = "%%FILE0_40.DESCA%%"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "„Œ“‰"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·»Ì«‰"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "„€·ﬁ"
GrdArray(4, 1) = 700

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.sCaption = "«” ⁄·«„"
oSearch.Show 1
End Sub
Private Sub docApiLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_APP_ID,DOC_APP_NAME,FORMAT(MIN(DOC_APP_DATE),'yyyy/M/d'),FORMAT(MAX(DOC_APP_DATE),'yyyy/M/d')" & _
                  " FROM FILE0_100 "

    Generalarray(1) = Generalarray(1) & " WHERE DOC_NO = " & MyParn(xDoc_No.text)

Generalarray(2) = "GROUP BY DOC_APP_ID,DOC_APP_NAME ORDER BY MIN(DOC_APP_DATE) DESC"
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DOC_APP_NAME%%)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 3000

GrdArray(2, 0) = "„‰  «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "Õ Ï  «—ÌŒ"
GrdArray(3, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchApi.sCaption = "«” ⁄·«„"
oSearchApi.Show 1
End Sub
Private Sub docCSVLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(0, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT FILE_NAME" & _
                  " FROM FILE0_100 "

Generalarray(1) = Generalarray(1) & " WHERE DOC_NO = " & MyParn(xDoc_No.text)

Generalarray(2) = "GROUP BY FILE_NAME"
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«”„ «·„·›"
listarray(0, 1) = "(%%FILE_NAME%%)"

GrdArray(0, 0) = "«”„ «·„·›"
GrdArray(0, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCSV.sCaption = "«” ⁄·«„"
oSearchCSV.Show 1
End Sub
Private Sub fixBalance()

End Sub
Private Sub cmdMerge_Click()
Set oZoneMerge = Nothing
Set oZoneMerge = New stock_zone_mergefrm
Set oZoneMerge.con = con
Set oZoneMerge.myForm = Me
oZoneMerge.pDoc_no = xDoc_No.text
oZoneMerge.pZone = xZone.Caption
oZoneMerge.Show 1
End Sub
Private Sub cmdPost_Click()
cmdPost.Enabled = False

If Not myValid Then Exit Sub
If Not myreplace(0) Then Exit Sub

If TransDoc("1") Then
    MsgBox " „  —ÕÌ· «·Ã—œ »‰Õ«Õ"
End If

If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub

Private Sub CmdNewInv_Click()
    myDefine
    On Error Resume Next
    xDoc_No.SetFocus
    Err.Clear
End Sub
Private Sub cmdPrint_Click()
doprint
End Sub
Private Sub doprint()
On Error GoTo myerror
Dim aHeader(2)
If Not myValid Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim locTable As New ADODB.Recordset

Dim cString As String
cString = "SELECT FILE0_10H.DATE," & _
          "FILE0_10H.DATE," & _
          "FILE0_40.DESCA AS STORE_DESCA," & _
          "FILE0_10H.DESCA AS DOC_DESCA," & _
          "FILE1_10.modelfact0," & _
          "FILE1_10.DESCA," & _
          "sum(FILE0_10.realBal) AS QUANT," & _
          "ROUND(SUM(FILE0_10.RealBal * FILE1_10.costitem), 2) AS TOTAL" & _
          " FROM FILE0_10 " & _
          " INNER JOIN FILE0_10H ON FILE0_10.DOC_NO = FILE0_10H.DOC_NO" & _
          " INNER JOIN FILE0_40 ON FILE0_10H.STORE = FILE0_40.CODE" & _
          " INNER JOIN FILE1_10 ON FILE1_10.item = FILE0_10.item " & _
          " WHERE FILE0_10H.DOC_NO = " & MyParn(xDoc_No.text) & _
          " GROUP BY " & _
          " FILE1_10.MODELFACT0," & _
          " FILE0_10H.DATE," & _
          " FILE0_10H.DESCA," & _
          " FILE0_40.DESCA," & _
          " FILE1_10.DESCA " & _
          " HAVING SUM(REALBAL) <> 0" & _
          " ORDER BY FILE1_10.modelfact0"


Set locTable = cmd(cString, con).Execute

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

With temptable
Do Until locTable.EOF
    .AddNew
    !STR21 = "ÿ»«⁄… „” ‰œ «·Ã—œ —ﬁ„ " & xDoc_No.text
    
    !STR4 = locTable!STORE_DESCA
    !STR5 = myFormat_p(locTable!Date)
    !str1 = locTable!MODELFACT0
    !str16 = locTable!DOC_DESCA
    !str2 = locTable!DESCA
    !VAL2 = locTable!Quant
    !VAL3 = locTable!TOTAL
    .Update
    locTable.MoveNext
Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

contemp.BeginTrans
contemp.CommitTrans
REPORT1.ReportFileName = App.Path & IIf(bOpt10, "\Reports\stock_cost.rpt", "\Reports\stock.rpt")
REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdSave_Click()
'foundOther
If Not myValid Then Exit Sub
If Not myreplace(0) Then Exit Sub
'AddLod_Data cusername, 1, " Õ÷Ÿ " & Me.Caption, con, xDoc_No.text, xDate.text, , xDesca.text

Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub CmdUndo_Click()
    myUndo
End Sub
Private Sub cmdunPost_Click()
cmdUnpost.Enabled = False
If TransDoc("0") Then
    MsgBox " „ «·€«¡  —ÕÌ· «·Ã—œ »‰Õ«Õ"
End If

If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub cmdAddItems_Click()
Dim locTable As New ADODB.Recordset, nRecordCount As Integer

cString = "SELECT FILE1_10.item,file1_10.mosm ,fact.desca as factDesca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color,Sum(FILE1_11.[IN] - FILE1_11.[out]) as Balance" & _
          " FROM (FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT"
cString = cString & turn(cString) & " FILE1_11.DATE <= " & DateSq(xDate.text)
cString = cString & turn(cString) & " FILE1_11.STORE = " & MyParn(xStore.BoundText)

cString = cString & " GROUP BY FILE1_10.item,file1_10.mosm ,fact.desca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color"
cString = cString & turn(cString, " HAVING ", " AND ") & " Sum(FILE1_11.[IN] - FILE1_11.[out]) <> 0"
locTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText


If Not (locTable.EOF And locTable.BOF) Then
    nRecordCount = locTable.RecordCount
End If
With grid1
    prog1.Visible = True
    prog1.value = 0
    Do Until locTable.EOF
        If grid1.FindRow(locTable!Item, , 0) = -1 Then
            i = i + 1
            prog1.value = Round(i / (nRecordCount), 2) * 100
            grid1.TextMatrix(.Rows - 1, 0) = locTable!Item
            grid1.TextMatrix(.Rows - 1, 1) = locTable!MOSM & ""
            grid1.TextMatrix(.Rows - 1, 2) = locTable!FACTDESCA
            grid1.TextMatrix(.Rows - 1, 3) = locTable!SUPP & ""
            grid1.TextMatrix(.Rows - 1, 4) = locTable!MODELFACT0 & ""
            grid1.TextMatrix(.Rows - 1, 5) = locTable!DESCA & ""
            grid1.TextMatrix(.Rows - 1, 6) = locTable!SCAL & ""
            grid1.TextMatrix(.Rows - 1, 7) = locTable!color & ""
            grid1.TextMatrix(.Rows - 1, 8) = 0
            grid1.AddItem ""
        End If
        locTable.MoveNext
    Loop
    prog1.Visible = False
End With
End Sub

Private Sub Command1_Click()
Dim locTable As New ADODB.Recordset, nRecordCount As Integer

If Not IsDate(xDate1.text) Then xDate1.text = xDate.text
If Not IsDate(xdate2.text) Then xdate2.text = xDate.text

cString = "SELECT FILE1_10.item,file1_10.mosm ,fact.desca as factDesca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color,Sum(FILE1_11.[IN] - FILE1_11.[out]) as Balance" & _
          " FROM (FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT where file1_10.item not in  (select file0_100.item from file0_100 inner join file0_10h on file0_10h.doc_no = file0_100.doc_no where file0_10h.date >= " & DateSq(xDate1.text) & " and date <= " & DateSq(xdate2.text) & " ) "
cString = cString & " and FILE1_11.DATE < " & DateSq(xDate.text)
cString = cString & " and FILE1_11.STORE = " & MyParn(xStore.BoundText)

cString = cString & " GROUP BY FILE1_10.item,file1_10.mosm ,fact.desca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color"
cString = cString & turn(cString, " HAVING ", " AND ") & " Sum(FILE1_11.[IN] - FILE1_11.[out]) <> 0"
locTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (locTable.EOF And locTable.BOF) Then
    nRecordCount = locTable.RecordCount
End If
With grid1
    prog1.Visible = True
    prog1.value = 0
    Do Until locTable.EOF
        i = i + 1
        prog1.value = Round(i / (nRecordCount), 2) * 100
        grid1.TextMatrix(.Rows - 1, 0) = locTable!Item
        grid1.TextMatrix(.Rows - 1, 1) = locTable!MOSM & ""
        grid1.TextMatrix(.Rows - 1, 2) = locTable!FACTDESCA
        grid1.TextMatrix(.Rows - 1, 3) = locTable!SUPP & ""
        grid1.TextMatrix(.Rows - 1, 4) = locTable!MODELFACT0 & ""
        grid1.TextMatrix(.Rows - 1, 5) = locTable!DESCA & ""
        grid1.TextMatrix(.Rows - 1, 6) = locTable!SCAL & ""
        grid1.TextMatrix(.Rows - 1, 7) = locTable!color & ""
        grid1.TextMatrix(.Rows - 1, 8) = 0
        grid1.AddItem ""
        locTable.MoveNext
    Loop
    prog1.Visible = False
End With
End Sub
Private Sub FIX_COST_Click()
    cStr1 = " UPDATE    FILE0_10 SET  cost = Coalesce((SELECT TOP 1 ( FILE7_20.cost_inv ) FROM FILE7_20 INNER JOIN FILE7_20H ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO WHERE FILE0_10.ITEM = FILE7_20.ITEM AND FILE7_20H.DATE <= FILE0_10H.DATE ORDER BY DATE DESC ) ,0) " & _
            " FROM      FILE0_10 INNER JOIN FILE0_10H ON FILE0_10.DOC_NO = FILE0_10H.DOC_NO"
    con.CommandTimeout = 600
    con.Execute cStr1, nRec
    Inform "  „ Ÿ»ÿ  ”ÊÌ… «·Ã—œ " & nRec

End Sub
Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xDoc_No.Tag = LoadMode Then
        grid1.SetFocus
    Else
        xStore.SetFocus
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
OpenCon con

Set data1.Recordset = cmd("SELECT * from file0_40 ORDER BY code ", con).Execute
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

If cBranch <> "00" Then
    xStore.BoundText = cBranchStore
    xStore.Enabled = False
    cmdBranch.Visible = False
End If

Set grid1.DataSource = DATA11
Set grid2.DataSource = data12
If Not openCardTable Then myDefine
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Err.Clear
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
On Error GoTo myerror
With grid1
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


If Not validRow(Row) Then
    CalcTotals
    Exit Sub
End If

Dim bNew As Boolean
If Row = grid1.Rows - 1 Then
    bNew = True
    myAddItem
End If

If myreplace(Row) Then
    myloadgrd2 True
    If xDoc_No.Tag = DefineMode Then
        myLoadGrd
        Handlecontrols LoadMode
        If bNew And chkBarCode.value = 1 Then
            bStopCell = True
            grid1.Select grid1.Rows - 1, 1
            grid1.ShowCell grid1.Rows - 1, 1
        End If
    ElseIf bNew Then
        If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
            myLoadGrd
        Else
            grid1.TextMatrix(Row, 0) = Row
        End If
        If bNew And chkBarCode.value = 1 Then
            grid1.Select grid1.Rows - 1, 1
        End If
        grid1.ShowCell grid1.Rows - 1, 1
    Else
        CalcTotals
    End If
Else
    myLoadGrd
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_EnterCell()
With grid1
If Not bEditRecord Then
    grid1.Editable = flexEDNone
ElseIf (grid1.col = 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Or grid1.col = 9 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End With
End Sub

Private Sub grid1_GotFocus()
grid1_EnterCell
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
        Cancel = True
    Else
        Dim sMsg As String
        grid1.EditText = retItemCode(grid1.EditText, con)
        If Not isNum(grid1.EditText) Then
            MsgBox "»«—ﬂÊœ €Ì— ’ÕÌÕ"
            Cancel = True
        ElseIf Not GrdDesc(grid1.EditText, Row, , sMsg) Then
            MsgBox IIf(sMsg = "", "»«—ﬂÊœ €Ì— ’ÕÌÕ", sMsg)
            Cancel = True
        End If
    End If
ElseIf col = 9 Then
    If Not IsNumeric(grid1.EditText) Then
        Cancel = True
    End If
End If
End Sub

Private Sub Grid2_dblClick()
If grid2.Row > 1 And xall.value = 0 Then
    SetZoneRow grid2.Row
End If
End Sub
Private Sub SetZoneRow(Row As Long)
If grid2.ValueMatrix(Row, 4) = 0 Then
    For i = 1 To grid2.Rows - 1
        If grid2.ValueMatrix(i, 4) = -1 Then
            grid2.TextMatrix(i, 4) = 0
        End If
    Next
    grid2.TextMatrix(Row, 4) = -1
End If
If xZone.Caption <> grid2.TextMatrix(Row, 1) Then
    xZone.Caption = grid2.TextMatrix(Row, 1)
    myLoadGrd
    CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End If
End Sub
Private Sub Option1_Click(index As Integer)
Me.MousePointer = vbHourglass
myloadgrd2
myLoadGrd
Me.MousePointer = vbNormal
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub txtBarcode_Change()
SearchGrd grid1, LCase(Trim(txtBarcode.text)), 1, True
End Sub

Private Sub txtModel_Change()
SearchGrd grid1, txtModel.text, 5
End Sub
Private Sub XALL_Click()
myLoadGrd
myloadgrd2
Handlecontrols LoadMode
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
Public Function myValid(Optional bIgMsg As Boolean = False) As Boolean
If xDoc_No.text = "" Then
    If Not bIsgMsg Then Inform "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    If Not bIsMsg Then Inform "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Not xStore.MatchedWithList Then
    If Not bIsMsg Then Inform "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If xZone.Caption = "" Then
    If Not bIsMsg Then Inform "·« ÌÊÃœ “Ê‰ „”Ã· "
    Exit Function
End If

myValid = True
End Function
Private Sub myload()
xDoc_No.text = CardTable!doc_no
xdoc_no_flag.Caption = CardTable!DOC_NO_FLAG & ""
xDate.text = myFormat_p(CardTable!Date)
xStore.BoundText = CardTable!STORE
xDesca.text = CardTable!DESCA & ""
xDoc_App_name.Caption = CardTable!doc_App_name & ""
xZone.Caption = ""
xClosed.value = IIf(CardTable!CLOSED, 1, 0)
xDoc_No.Enabled = False

myloadgrd2
myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1

CalcTotals

Handlecontrols LoadMode

On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
xDoc_No.Enabled = True
xDoc_No.text = RetZero(NewflagBranch6("FILE0_10H", "DOC_NO", cBranch, con), 6)
xdoc_no_flag.Caption = ""
xDoc_App_name.Caption = ""
xDate.text = myFormat_p(Date)
xZone.Caption = "Z1"
xDesca.text = ""
xClosed.value = 0
If cBranch = "00" Then
    xStore.BoundText = cmdBranch.Tag
End If
grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
Fixgrd
grid2.Rows = 1
Fixgrd2
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit And xClosed.value = 0 And xall.value = 0
grid2.Enabled = xall.value = 0
Dim nRecord As Long, nRecords As Long
retRecords xdoc_no_flag.Caption, nRecords, nRecord
If nMode = LoadMode Then
    panel1(0).Caption = ArbString("”Ã· " & nRecord & " „‰ " & nRecords)
Else
    panel1(0).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

'xStore.Enabled = cBranchStore = "000"
'xStore.BoundText = cBranchStore

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""

cmdUnpost.Enabled = bEdit And xClosed.value = 1
cmdDiffer.Enabled = bEdit And xClosed.value = 1

cmdNewInv.Enabled = nMode = LoadMode And bEdit And xall.value = 0

cmdAddZone.Enabled = bEditRecord And nMode = LoadMode

cmdAddZero.Enabled = nMode = LoadMode And bEditRecord
cmdDelAdd.Enabled = nMode = LoadMode And bEditRecord

cmdCSVadd.Enabled = nMode = LoadMode And bEditRecord
cmdCSVdel.Enabled = nMode = LoadMode And bEditRecord

cmdDelApi.Enabled = nMode = LoadMode And bEditRecord
cmdApi.Enabled = nMode = LoadMode And bEditRecord

xDoc_No.Enabled = (nMode = DefineMode)
CmdDelInv.Enabled = bEditRecord And nMode = LoadMode
cmdPost.Enabled = bEditRecord And nMode = LoadMode

'cmdFix.Enabled = bEditRecord And nMode = LoadMode

cmdSave.Enabled = bEditRecord
If cBranch = "00" Then
    xStore.Enabled = cmdBranch.Tag = ""
End If
xDoc_No.Tag = nMode

HandleZone
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If Trim(xDoc_No.text) = "" Then
     If xDoc_No.Tag = LoadMode Then
        myDefine
    End If
Else
    xDoc_No.text = RetZero(xDoc_No.text)
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
       If xDoc_No.Tag = LoadMode Then
            myDefine
       End If
    End If
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.col = 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" And bEditRecord Then
    ItemsLookup Me, oSearchItem
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bEditRecord Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "delete from file0_100 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        grid1.RemoveItem grid1.Row
        Fixgrd
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myRemove()
End Sub
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
Private Function GrdDesc(sitem As String, Row, Optional bSkip As Boolean = False, Optional pMsg As String = "") As Boolean
With grid1
If Trim(sitem) = "" Then Exit Function
.TextMatrix(Row, 2) = ""
.TextMatrix(Row, 3) = ""
.TextMatrix(Row, 5) = ""
.TextMatrix(Row, 7) = ""
.TextMatrix(Row, 7) = ""
.TextMatrix(Row, 8) = ""
.TextMatrix(Row, 9) = ""

Dim Com As New ADODB.command
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "ITEM", sitem)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
aPrm = AddFlag(aPrm, "STORE", TurnValue(xStore.BoundText))

Set Com = cmd("sp_item_balance", con, adStoredProc, aPrm)
Com.Execute

If Not Com.Parameters("@FOUND").value Then
    Exit Function
End If

If Com.Parameters("@ISSTOP").value Then
    pMsg = "’‰› „ Êﬁ›"
    Exit Function
End If

grid1.TextMatrix(grid1.Row, 2) = Com.Parameters("@MOSM").value & ""
grid1.TextMatrix(grid1.Row, 3) = Com.Parameters("@FACT").value & ""
grid1.TextMatrix(grid1.Row, 4) = Com.Parameters("@SUPP").value & ""
grid1.TextMatrix(grid1.Row, 5) = Com.Parameters("@modelfact0").value & ""
grid1.TextMatrix(grid1.Row, 6) = Com.Parameters("@DESCA").value & ""
grid1.TextMatrix(grid1.Row, 7) = Com.Parameters("@SCAL").value & ""
grid1.TextMatrix(grid1.Row, 8) = Com.Parameters("@COLOR").value & ""
grid1.TextMatrix(grid1.Row, 9) = IIf(grid1.ValueMatrix(grid1.Row, 9) = 0, 1, grid1.ValueMatrix(grid1.Row, 9))
GrdDesc = True
End With
End Function
Private Function CalcTotals()
'Dim nCount As Double
'Dim nTotalQuant As Double, nTotalCost As Double
'With grid1
'grid1.Redraw = flexRDNone
'For i = 1 To grid1.Rows - 2
''   nCount = nCount + 1
''    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(I, 8))
'Next
'Me.panel1(0).caption = IIf(nTotalQuant <> 0, "≈Ã„«·Ì «·ﬂ„Ì… : " & nTotalQuant, "")
'xCount.Caption = "⁄œœ «·«’‰«› " & Format(nCount, "#0")
'End With
End Function
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
Private Sub Fixgrd()
Dim n1 As Double, n2 As Double
With grid1
grid1.Redraw = flexRDNone
'                   0         1       2          3        4           5            6           7          8            9

.FormatString = "„|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "⁄œœ Ã—œ|" & "ZONE|"
.WordWrap = True

.RowHeight(0) = 500

.ColWidth(0) = 800
.ColWidth(1) = 1200
.ColWidth(2) = 1200
.ColWidth(3) = 2800
.ColWidth(4) = 1000
.ColWidth(5) = 1500
.ColWidth(6) = 2800 + IIf(xall.value = 1, 1500, 0)
.ColWidth(7) = 1000
.ColWidth(8) = 1000 + IIf(xall.value = 1, 1000, 0)
.ColWidth(9) = 900
.ColWidth(10) = 2500
'.ColWidth(11) = 1000
'.ColWidth(11) = 1000
'.ColHidden(9) = True
'.ColHidden(10) = True
.ColHidden(.Cols - 2) = xall.value = 0
.ColHidden(.Cols - 1) = True
Dim nDiffer As Long
nDiffer = .ColWidth(10) + 2500

.Left = 5355 - IIf(xall.value = 0, 0, nDiffer)
Dim i As Long
Dim nwidth As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
grid1.Width = 14730 + IIf(xall.value = 0, 0, nDiffer)

Dim nTotal As Long
For i = 1 To grid1.Rows - 1
    .TextMatrix(i, 0) = i
    'nTotal = nTotal + .ValueMatrix(i, 9)
Next
If grid2.Rows > 1 Then
    panel1(1).Caption = "⁄œœ «·«’‰«› :" & grid2.TextMatrix(1, 2)
    panel1(2).Caption = "ﬂ„Ì… «·Ã—œ :" & grid2.TextMatrix(1, 3)
Else
    panel1(1).Caption = ""
    panel1(2).Caption = ""
End If
End With
grid1.Redraw = flexRDBuffered
End Sub
Private Sub Fixgrd2()
With grid2
.FormatString = "„|" & "«·“Ê‰|" & "⁄œœ «’‰«›|" & "ﬂ„Ì…|" & "«Œ Ì«—"
.WordWrap = True

.RowHeight(0) = 500

.ColWidth(0) = 600
.ColWidth(1) = 2000
.ColWidth(2) = 800
.ColWidth(3) = 900
.ColWidth(4) = 600
.ColDataType(4) = flexDTBoolean
.Visible = xall.value = 0
.Redraw = flexRDNone
Dim i As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

.SubtotalPosition = flexSTAbove
.Subtotal flexSTSum, -1, 2, "##,##", &HE0E0E0, vbBlack, True, "«·≈Ã„«·Ï"
.Subtotal flexSTSum, -1, 3, "##,##", &HE0E0E0, vbBlack, True, "«·≈Ã„«·Ï"

If .Rows > 1 Then
    .Cell(flexcpFontSize, 1, 2, .Rows - 1, 3) = 11
End If

For i = 2 To .Rows - 1
    .TextMatrix(i, 0) = i - 1
Next
.Redraw = flexRDBuffered
End With
End Sub
Private Sub myreplaceGrd(Row As Long)
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "item", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "realbal", grid1.ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "zone", addstring(xZone.Caption))
        aInsert = AddFlag(aInsert, "row", i)
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            aInsert = AddFlag(aInsert, "doc_no", addstring(xDoc_No.text))
            con.Execute addInsert(aInsert, "FILE0_100")
            grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = get_Id(con)
        Else
            con.Execute addUpdate(aInsert, "FILE0_100", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If

    Next
End With
End Sub
Private Sub myLoadGrd()
Dim cString As String
cString = "SELECT FILE0_100.item ," & _
          "file1_10.mosm ," & _
          "fact.desca," & _
          "file1_10.supp," & _
          "file1_10.modelfact0," & _
          "file1_10.desca," & _
          "file1_10.scal," & _
          "file1_10.color," & _
          "FILE0_100.RealBal," & _
          "FILE0_100.ZONE," & _
          "FILE0_100.ID " & _
          " from FILE0_100 " & _
          " inner join file1_10 on file1_10.item = FILE0_100.item" & _
          " inner join fact on fact.code = file1_10.fact" & _
          " WHERE  DOC_NO = " & MyParn(xDoc_No.text)
If xall.value = 0 Then
    cString = cString & " AND ZONE = " & MyParn(xZone.Caption)
End If
If Option1(1).value Then
    cString = cString & " AND FILE0_100.AUTO = 0"
ElseIf Option1(2).value Then
    cString = cString & " AND FILE0_100.AUTO = 1"
End If
cString = cString & " ORDER BY ID"
Set DATA11.Recordset = cmd(cString, con).Execute

Fixgrd
myAddItem
End Sub
Private Sub myloadgrd2(Optional bNoFoundOnly As Boolean = False)
Dim Found As Long
If grid2.Rows > 2 And bNoFoundOnly Then
    nFound = grid2.FindRow(xZone.Caption, , 1)
    If nFound > 0 Then
        grid2.Select nFound, 1
        Exit Sub
    End If
End If

Dim cString As String
cString = "SELECT ZONE," & _
          " COUNT(*)," & _
          "SUM(REALBAL)" & _
          " from FILE0_100 " & _
          " WHERE  DOC_NO = " & MyParn(xDoc_No.text)

If Option1(1).value Then
    cString = cString & " AND FILE0_100.AUTO = 0"
ElseIf Option1(2).value Then
    cString = cString & " AND FILE0_100.AUTO = 1"
End If
cString = cString & _
          " GROUP BY ZONE" & _
          " ORDER BY ZONE"
Set data12.Recordset = cmd(cString, con).Execute
Fixgrd2

Dim i As Long
nFound = grid2.FindRow(xZone.Caption, , 1)
If nFound > 0 Then
    grid2.TextMatrix(nFound, 4) = -1
    grid2.Select nFound, 1
ElseIf grid2.Rows > 2 Then
    For i = 2 To grid2.Rows - 1
        If grid2.TextMatrix(i, 1) > xZone.Caption Then
            Exit For
        End If
    Next
    If i > grid2.Rows - 1 Then i = grid2.Rows - 1
    grid2.TextMatrix(i, 4) = -1
    grid2.Select i, 2
    xZone.Caption = grid2.TextMatrix(i, 1)
Else
    If IsEmpty(GetField("select id from file0_100 where doc_no = " & MyParn(xDoc_No.text), con)) Then
        xZone.Caption = "Z1"
    End If
End If
HandleZone
End Sub
Private Sub myAddItem()
grid1.AddItem ""
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

If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If


With grid1
For i = 1 To .Rows - 2
    If .TextMatrix(i, 0) = "" Then
        .Select i, 0, i, grid1.Cols - 1
        MsgBox "ﬂÊœ «·’‰› €Ì— „ÊÃÊœ"
        Exit Function
    Else
        If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & MyParn(.TextMatrix(i, 0)), con) = "" Then
            .Select i, 0, i, 2
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Exit Function
        End If
    End If
Next
End With
MYVALIDfix = True
End Function
Private Sub myUndo()
If xDoc_No.Tag = DefineMode Then
    If Not openCardTable Then
        myDefine
    End If
Else
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub
Private Sub xfilter_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        FilterGrd grid1, xfilter.text, 0
'    End If
End Sub
Private Sub SubReplaceItem(pItem, nReal, nComputer, nDiffer)
Dim aInsert(4, 1)
With grid1
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring(xDoc_No.text)
        
    aInsert(1, 0) = "item"
    aInsert(1, 1) = addstring(pItem)
        
    aInsert(2, 0) = "RealBal"
    aInsert(2, 1) = nReal
        
    aInsert(3, 0) = "ComputerBal"
    aInsert(3, 1) = nComputer
        
    aInsert(4, 0) = "Differ"
    aInsert(4, 1) = nDiffer
    con.Execute CreateInsert(aInsert, "FILE0_10")
End With
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not IsNumeric(.TextMatrix(Row, 1)) Then Exit Function
'If Not validItem(.TextMatrix(Row, 0), con) Then Exit Function
validRow = True
End With
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If grid1.TextMatrix(Row, 1) = "" Then
ElseIf bStopCell Then
    'bStopCell = False
ElseIf col < grid1.Cols - 3 Then
    grid1.col = IIf(col < 9, 9, col + 1)
    grid1.ShowCell Row, 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, IIf(NextEmpty(grid1, Row + 1, 1, 9) = 1, 1, 9)
    grid1.ShowCell grid1.Row, 1
End If
bStopCell = False
End Sub
Private Sub XTQ1_GotFocus()
myGotFocus XTQ1
End Sub
Private Sub XTQ1_LostFocus()
myLostFocus XTQ1
End Sub
Private Sub XTQ2_GotFocus()
myGotFocus XTQ2
End Sub
Private Sub XTQ2_LostFocus()
myLostFocus XTQ2
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore, False
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xfilter_GotFocus()
myGotFocus xfilter
End Sub
Private Sub xfilter_LostFocus()
myLostFocus xfilter
End Sub
Private Sub grid1_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
'grid1.EditMaxLength = IIf(Col = 8, 20, 0)
End Sub
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_no As String = "") As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_no) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

On Error GoTo myerror
cFilter = " AND FILE0_10H.IS_ZONE = 1"

If cmdBranch.Tag <> "" Then
    cFilter = cFilter & " AND FILE0_10H.STORE = " & MyParn(cmdBranch.Tag)
End If

cFilter = Mid(cFilter, 6)

If sDoc_no <> "" Then
    cFilter = "DOC_NO = " & MyParn(sDoc_no)
End If

cString = "SELECT TOP 1 *,[dbo].[fn_doc_app](DOC_NO)  AS DOC_APP_NAME FROM FILE0_10H"

If pMode = tbMode.tbFirst Then
    cOrder = "Order by DOC_NO_FLAG"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by DOC_NO_FLAG DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "Doc_no = " & MyParn(pDoc_no)
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "DOC_NO_FLAG < " & MyParn(pDoc_no)
    cOrder = "Order by DOC_NO_FLAG desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "DOC_NO_FLAG > " & MyParn(pDoc_no)
    cOrder = "Order by DOC_NO_FLAG"
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
If Not openCardTable(tbMode.tbNext, xdoc_no_flag.Caption) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, xdoc_no_flag.Caption) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Private Function retRecords(pCode, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, locTable As New ADODB.Recordset
If pCode <> "" Then
    cString = "SELECT Count(*) AS records,COUNT(CASE WHEN DOC_NO_FLAG <= " & MyParn(pCode) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM FILE0_10H"
If cFilter <> "" Then cString = cString & " WHERE " & cFilter

Set locTable = mycmd(cString, con)
If Not locTable.EOF Then
    nRecords = locTable!RECORDS
    nRecord = Val(locTable!Record & "")
End If
End Function
Public Sub myProcZoneAdd()
xZone.Caption = oZoneAdd.xZone.text
Unload oZoneAdd
myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
grid1.SetFocus
End Sub
Public Sub myProcZoneEdit()
'xZone.Caption = oZoneEdit.xZone.text

con.BeginTrans
On Error GoTo myerror
con.Execute "UPDATE FILE0_100 " & _
            " SET FILE0_100.ZONE = " & addstring(oZoneEdit.xZone.text) & _
            " WHERE DOC_NO = " & MyParn(xDoc_No.text) & _
            " AND FILE0_100.ZONE = " & MyParn(xZone.Caption)
con.CommitTrans
Inform " „  ⁄œÌ· «·“Ê‰"
xZone.Caption = oZoneEdit.xZone.text
Unload oZoneEdit
myloadgrd2
myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
grid1.SetFocus
Exit Sub
myerror:
MsgBox Err.Description
If TransCount(con) > 0 Then con.RollbackTrans
Err.Clear
End Sub
Public Sub myProcZoneMerge()
con.BeginTrans
con.Execute "UPDATE FILE0_100 " & _
            " SET FILE0_100.ZONE = " & addstring(oZoneMerge.xZone.BoundText) & _
            " WHERE DOC_NO = " & MyParn(xDoc_No.text) & _
            " AND FILE0_100.ZONE = " & MyParn(xZone.Caption)
con.CommitTrans
Inform " „ œ„Ã «·“Ê‰"
xZone.Caption = oZoneMerge.xZone.BoundText
Unload oZoneMerge
myloadgrd2
myLoadGrd

Handlecontrols LoadMode

CellPos 13, grid1.Rows - 2, grid1.Cols - 1
grid1.SetFocus
Exit Sub
myerror:
MsgBox Err.Description
If TransCount(con) > 0 Then con.RollbackTrans
Err.Clear
End Sub
Private Sub HandleZone()
cmdDelZone.Enabled = bEditRecord And grid2.Rows > 1
cmdEditZone.Enabled = bEditRecord And grid2.Rows > 1
cmdMerge.Enabled = bEditRecord And grid2.Rows > 2
End Sub
Private Function TransDoc(pClosed As String) As Boolean
Me.MousePointer = vbHourglass

If Not myValid Then Exit Function
Dim nCountInsert As Long
Dim nCountDelete As Long
Dim nCountInsertMove As Long
Dim nCountDeleteMove As Long


con.BeginTrans
TransStockAll xDoc_No.text, con, pClosed, nCountInsert, nCountDelete, nCountInsertMove, nCountDeleteMove
con.CommitTrans

Me.MousePointer = vbNormal

Exit Function
myerror:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
If TransCount(con) > 0 Then
    con.RollbackTrans
End If
End Function
Public Sub myProcZero()
Unload oStockZero
myloadgrd2
myLoadGrd
End Sub
Public Sub myProcAppAdd()
Unload oStockApi
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Function delApp(pDoc_app_id) As Boolean
On Error GoTo myerror
con.Execute "Delete From FILE0_100 " & _
            " where Doc_No = " & MyParn(xDoc_No.text) & _
            " and doc_app_id = " & pDoc_app_id
delApp = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function delCSV(pFile_Name As String) As Boolean
On Error GoTo myerror
con.Execute "Delete From FILE0_100 " & _
            " where Doc_No = " & MyParn(xDoc_No.text) & _
            " and FILE_NAME = " & MyParn(pFile_Name)
delCSV = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function getCSV() As Boolean
Dim cFileName As String
Dim sFileTitle As String
Dim sb As New ChilkatStringBuilder
Dim sbSql As New ChilkatStringBuilder

Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

cFileName = Common1.FileName
sFileTitle = Common1.FileTitle

If cFileName = "" Then Exit Function
    
dotPos = InStrRev(sFileTitle, ".")
If dotPos > 0 Then
    sFileTitle = Left(sFileTitle, dotPos - 1)
End If

Dim cSv As New ChilkatCsv
Dim i As Long, i2 As Long
SetKbLayout Lang_AR

cSv.HasColumnNames = 1
nAccess = sb.LoadFile(cFileName, "ansi")
If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If


nAccess = cSv.LoadFromString(sb.GetAsString)

If cSv.NumRows < 1 Then Exit Function

Dim Tb As New ChilkatStringBuilder
Dim cString As New ChilkatStringBuilder
Dim locTable As New ADODB.Recordset

prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption
con.BeginTrans
sb.Clear
For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows - 1
    prog1.value = Round(i / (cSv.NumRows), 2) * 100
    
    If Trim(cSv.GetCell(i, 0)) <> "" And Trim(cSv.GetCell(i, 2)) <> "" Then
        sitem = retItemCode(Trim(cSv.GetCell(i, 0)), con)
        If IsDgt(sitem) Then
            i2 = i2 + 1
            sb.Append "(" & _
                    addstring(xDoc_No.text) & "," & _
                    addstring(sitem) & "," & _
                    Val(cSv.GetCell(i, 1)) & "," & _
                    addstring(cSv.GetCell(i, 2)) & "," & _
                    addstring(sFileTitle) & "," & _
                    "1" & _
                   "),"
    
            If i2 Mod 1000 = 0 Then
                sb.Shorten 1
                con.Execute "INSERT INTO FILE0_100(" & _
                            "doc_no," & _
                            " item," & _
                            " RealBal," & _
                            " [ZONE]," & _
                            " [FILE_NAME]," & _
                            " [AUTO])" & _
                             "VALUES" & _
                             sb.GetAsString, nAdded
                sb.Clear
                nAffect = nAffect + nAdded
            End If
        End If
    End If
Next

If sb.length > 0 Then
    sb.Shorten 1
    con.Execute "INSERT INTO FILE0_100(" & _
                "doc_no," & _
                " item," & _
                " RealBal," & _
                " [ZONE]," & _
                " [FILE_NAME]," & _
                " [AUTO])" & _
                 "VALUES" & _
                 sb.GetAsString, nAdded
    sb.Clear
    nAffect = nAffect + nAdded
End If

con.CommitTrans
MsgBox " „  «÷«›… " & nAffect & "”Ã·"
Me.Caption = sCaption
prog1.Visible = False
getCSV = True
Finaly:
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GoTo Finaly
End Function

