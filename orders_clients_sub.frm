VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form orders_clients_sub 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·»Ì«  «·⁄„·«¡"
   ClientHeight    =   10065
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
   ScaleHeight     =   10065
   ScaleWidth      =   20400
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
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
      Height          =   780
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   -45
      Width           =   9330
      Begin Threed.SSCommand cmdExcelImport 
         Height          =   510
         Left            =   6795
         TabIndex        =   53
         Top             =   180
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
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
         Caption         =   "Excel «Õ÷«— ÿ·»Ì«  „‰"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrintScal 
         Height          =   510
         Left            =   3960
         TabIndex        =   54
         Top             =   180
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
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
         Caption         =   "ÿ»«⁄…  ›’Ì·Ï „ﬁ«”«  ÿ·»Ì… «·⁄„Ì·"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   510
         Left            =   1845
         TabIndex        =   55
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
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
         Caption         =   "ÿ»«⁄… ⁄—÷ ”⁄—"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdBalanceApp 
         Height          =   510
         Left            =   45
         TabIndex        =   65
         Top             =   180
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
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
         Caption         =   "—’Ìœ «’‰«› «·ÿ·»Ì« "
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame fmClose 
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
      Height          =   735
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   1620
      Width           =   2670
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«€·«ﬁ „” ‰œ"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   270
         Width           =   1230
      End
      Begin Threed.SSCommand cmdClosePeriod 
         Height          =   420
         Left            =   90
         TabIndex        =   40
         Top             =   225
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   741
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
         Caption         =   "«€·«ﬁ › —…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
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
      Height          =   1140
      Index           =   4
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   1215
      Width           =   1275
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   45
         TabIndex        =   5
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "orders_clients_sub.frx":0000
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "orders_clients_sub.frx":2925
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   465
         Left            =   45
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   630
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "orders_clients_sub.frx":5179
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "orders_clients_sub.frx":72D9
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   735
      Left            =   13320
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   -45
      Width           =   6900
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   5760
         TabIndex        =   12
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "orders_clients_sub.frx":95C6
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "orders_clients_sub.frx":B991
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   4635
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "orders_clients_sub.frx":DA3A
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "orders_clients_sub.frx":FA42
      End
      Begin Threed.SSCommand cmddelInv 
         Height          =   510
         Left            =   3510
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "orders_clients_sub.frx":119F9
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "orders_clients_sub.frx":14195
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         Picture         =   "orders_clients_sub.frx":16629
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   2340
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   180
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "orders_clients_sub.frx":1894C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "orders_clients_sub.frx":1ACC2
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   510
         Left            =   1125
         TabIndex        =   37
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "orders_clients_sub.frx":1CE45
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   7875
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   675
      Width           =   12345
      Begin VB.TextBox xOrder_Main 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9810
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox xdelivery_date 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1260
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   540
         Width           =   2265
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1260
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   2265
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9810
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1275
      End
      Begin VB.TextBox xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6345
         MaxLength       =   75
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1260
         Width           =   4740
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   330
         Left            =   1260
         TabIndex        =   2
         Top             =   900
         Width           =   2265
         _ExtentX        =   3995
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
      Begin MSDataListLib.DataCombo xCode 
         Height          =   315
         Left            =   6345
         TabIndex        =   41
         Top             =   900
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xsalesman 
         Height          =   330
         Left            =   90
         TabIndex        =   44
         Top             =   1260
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„” ‰œ —∆Ì”Ì"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   11175
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «· ”·Ì„"
         Height          =   270
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·«ÕŸ« "
         Height          =   270
         Left            =   11205
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1350
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„‰œÊ» «· ÿ»Ìﬁ"
         Height          =   270
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Label XDOC_APP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   180
         Width           =   2175
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "„” ‰œ «· ÿ»Ìﬁ"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„"
         Height          =   270
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   945
         Width           =   525
      End
      Begin VB.Label lblClient 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "«·⁄„Ì·"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   11205
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   900
         Width           =   450
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "—ﬁ„ „” ‰œ"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   11175
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   210
         Width           =   840
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   225
         Width           =   510
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   -90
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   90
      Top             =   1215
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid grid20 
      Height          =   405
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   465
      _cx             =   820
      _cy             =   714
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   3
      FixedCols       =   2
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
      Editable        =   2
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
   Begin MSComDlg.CommonDialog Common2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6180
      Left            =   45
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2385
      Width           =   20175
      _cx             =   35586
      _cy             =   10901
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
   Begin VB.Frame Frame5 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   14490
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   8550
      Width           =   5775
      Begin VB.Label lblTotalQuant 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì «·ﬂ„Ì…"
         Height          =   240
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label xtotal_Quant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label xcount_model 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "⁄œœ „ÊœÌ·« "
         Height          =   270
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   225
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   735
      Left            =   2295
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1620
      Width           =   1545
      Begin Threed.SSCommand cmdAddModel 
         Height          =   510
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   900
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
         Caption         =   "«÷«›… „ÊœÌ·"
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   9690
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   35
      Top             =   9585
      Visible         =   0   'False
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
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
      TabIndex        =   36
      Top             =   8505
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   8
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
         Picture         =   "orders_clients_sub.frx":1F2AB
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "orders_clients_sub.frx":21452
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   9
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
         Picture         =   "orders_clients_sub.frx":23499
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "orders_clients_sub.frx":25584
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   10
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
         Picture         =   "orders_clients_sub.frx":2757E
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "orders_clients_sub.frx":2968F
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   11
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
         Picture         =   "orders_clients_sub.frx":2B689
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "orders_clients_sub.frx":2D8AD
      End
   End
   Begin VB.Frame Frame6 
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
      Height          =   735
      Left            =   9675
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   -45
      Width           =   3615
      Begin Threed.SSCommand cmdClear 
         Height          =   420
         Left            =   90
         TabIndex        =   51
         Top             =   225
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
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
         Caption         =   "«·€«¡ «·‰”Œ"
         TagVariant      =   "«Œ «— «·⁄„Ì·"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPaste 
         Height          =   420
         Left            =   1845
         TabIndex        =   50
         Top             =   225
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
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
         Caption         =   "·’ﬁ «·ÿ·»Ì…"
         TagVariant      =   "«Œ «— «·⁄„Ì·"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCopy 
         Height          =   420
         Left            =   90
         TabIndex        =   56
         Top             =   225
         Width           =   3435
         _ExtentX        =   6059
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
         Caption         =   "‰”Œ ÿ·»»…"
         TagVariant      =   "«Œ «— «·⁄„Ì·"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   7020
      TabIndex        =   57
      Top             =   8550
      Width           =   4425
      Begin VB.OptionButton optChanged 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3555
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   690
      End
      Begin VB.OptionButton optChanged 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " „ «·«—”«·"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   225
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   225
         Width           =   1095
      End
      Begin VB.OptionButton optChanged 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " ⁄œÌ· »⁄œ «·«—”«·"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1755
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   225
         Width           =   1590
      End
   End
   Begin VB.Frame frClose 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   11475
      TabIndex        =   61
      Top             =   8550
      Width           =   2985
      Begin VB.OptionButton optclosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„› ÊÕ…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   990
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optclosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„€·ﬁ…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   135
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   270
         Width           =   735
      End
      Begin VB.OptionButton optclosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2115
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   270
         Value           =   -1  'True
         Width           =   690
      End
   End
End
Attribute VB_Name = "orders_clients_sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_No As String
Public nRound As Integer
Public bedit As Boolean
Dim bIg As Boolean
Dim bEditRecord As Boolean
Dim CardTable As ADODB.RecordSet, cFileHeader As String
Dim cFilter As String
Dim osearchDoc As New Search_db, oSearchClient As New Search_db
Dim con As New ADODB.Connection
Dim formMode
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[CODE]", addstring(xCode.BoundText))
aInsert = AddFlag(aInsert, "[MOSM]", addstring(xMosm.BoundText))
aInsert = AddFlag(aInsert, "[delivery_date]", addDate(xdelivery_date.text))
aInsert = AddFlag(aInsert, "[DESCA]", addstring(xdesca.text))
aInsert = AddFlag(aInsert, "[Deposit]", Val(xdeposit.text))
aInsert = AddFlag(aInsert, "[USERNAME]", addstring(sUserName))
aInsert = AddFlag(aInsert, "[CHANGED]", "1")
aInsert = AddFlag(aInsert, "[ISNEW]", "1")
Dim db As New clsDb
If xDoc_No.Tag = DefineMode Then
    xDoc_No.text = db.NewFlag("FILE6_50H", "doc_no", 6)
    If xDoc_No.text = "" Then GoTo CleanUp
    
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
    db.Sql = addInsert(aInsert, "FILE6_50H")
Else
    db.Sql = addUpdate(aInsert, "FILE6_50H", "doc_no = " & addstring(xDoc_No.text))
End If

If Not db.ExecuteTransaction Then GoTo CleanUp

myreplace = True
CleanUp:
Set db = Nothing
End Function
Private Sub myreplaceGrd(Row As Long, db As clsDb)
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addvalue(xDoc_No.text))
        aInsert = AddFlag(aInsert, "ITEM", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "QUANT", grid1.ValueMatrix(i, 5))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            db.Sql = addInsert(aInsert, "FILE6_50")
        Else
            db.Sql = addUpdate(aInsert, "FILE6_50", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = cmdInform.Name Then
    openCardTable tbMode.tbFind, osearchDoc.grid1.TextMatrix(osearchDoc.grid1.Row, 0)
    Unload osearchDoc
ElseIf ActiveControl.Name = xCode.Name Then
    xCode.BoundText = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
    Unload oSearchClient
End If
End Sub
Private Sub cmdAddGroup_Click()
If xDoc_No.text = "" Then cmdSave_Click
Dim db As New clsDb
db.Execute "UPDATE FILE6_50H SET TIME_POST = NULL WHERE DOC_NO = " & MyParn(xDoc_No.text)
Set db = Nothing

Set ModelAddOrder.myform = Me
ModelAddOrder.Show 1

If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If xDoc_No.Tag = LoadMode Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Sub cmdAddModel_Click()
'Set addmodelOrder.myform = Me
'Set addmodelOrder.pGrid = grid1
'addmodelOrder.pCol1 = 1
'addmodelOrder.pCol2 = 5
'addmodelOrder.Show 1
On Error GoTo myerror
If xDoc_No.text = "" Then
    If Not mySave Then Exit Sub
End If

Dim db As New clsDb
db.Execute "UPDATE FILE6_50H SET TIME_POST = NULL , time_save = getdate() WHERE DOC_NO = " & MyParn(xDoc_No.text)
Set db = Nothing

Set addmodel2.myform = Me
Set addmodel2.pGrid = Me.grid1
addmodel2.pCol1 = 1
addmodel2.pCol2 = 5
addmodel2.Show 1
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmdBalanceApp_Click()
grdBalanceAppfrm.Show
End Sub

Private Sub cmdClosePeriod_Click()
closefrm.sFile = "FILE6_50H"
closefrm.Show 1
myUndo
End Sub
Private Sub cmdCopy_Click()
If grid1.Rows = 1 Then Exit Sub
If xDoc_No.text <> "" Then
    cmdCopy.Tag = xDoc_No.text
    cmdCopy.Visible = False
    cmdPaste.Visible = True
    cmdClear.Visible = True
End If
End Sub
Private Sub CmdClear_Click()
cmdCopy.Tag = ""
cmdCopy.Visible = True
cmdPaste.Visible = False
cmdClear.Visible = False
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› ?", vbOKCancel + vbDefaultButton2 + vbCritical) <> vbOK Then Exit Sub
Dim db As New clsDb
db.Sql = "Delete From FILE6_50 where Doc_No = " & addstring(xDoc_No.text)
db.Sql = "Delete From FILE6_50H where Doc_No = " & addstring(xDoc_No.text)
If Not db.ExecuteTransaction Then GoTo CleanUp

If sDoc_No <> "" Then Exit Sub

If Not openCardTable(tbMode.tbPrevious, xDoc_No.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
CleanUp:
Set db = Nothing
End Sub
Private Sub cmdExcel_Click()
Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aRow = AddFlag(aRow, aSub)

grid1.RowHidden(grid1.Rows - 1) = True
ToFileExelNew grid1, , , aRow, , 0.9, , , , , , Me, Array(Me.Caption, "„” ‰œ —ﬁ„ : " & xDoc_No.text & " » «—ÌŒ : " & xDate.text, xdesca.text, retHeader(aHeader, 4, 5))
grid1.RowHidden(grid1.Rows - 1) = False
End Sub
Private Sub cmdExcelImport_Click()
    If grid1.Rows > 1 Then Exit Sub
    
    Dim xl1 As New Excel.Application, lAddOrder As Boolean
    Dim xlsheet1 As Excel.Worksheet
    Dim xlwbook1 As Excel.Workbook
    
    Dim xl2 As New Excel.Application
    Dim xlsheet2 As Excel.Worksheet
    Dim xlwbook2 As Excel.Workbook
    
    Dim cBacode1 As String, cBacode2 As String, nCount As Double
     
    Dim cFileName1 As String
    Dim cFileName2 As String
    Dim nCount1 As Double
    
    If xDoc_No.Tag = DefineMode Then
        If Not mySave Then Exit Sub
    End If
    
    If MsgBox("≈÷«›…  ›’Ì·Ï «’‰«› «·ÿ·»Ì«  „·› Excel", vbYesNo) = vbYes Then
        Common2.InitDir = ""
        Common2.FileName = ""
        Common2.Filter = "Excel (*.XLS*)|*.XLS*"
        Common2.ShowOpen
        cFileName2 = Common2.FileName

'''''''''''''''''''''

        Set xlwbook2 = xl2.Workbooks.Open(cFileName2)
        Set xlsheet2 = xlwbook2.Sheets.item(1)
        
        For nRow = 2 To 10000
            nCount = nCount + 1
            If Trim(xlsheet2.Cells(nRow, 1)) = "" Then Exit For
             
        Next nRow
        prog1.Visible = True
        prog1.Min = 0
        prog1.Value = 0
        prog1.Max = nCount
        cStr1 = ""
        
        For nRow = 2 To nCount
            If Trim(xlsheet2.Cells(nRow, 1)) <> "" Then
                Me.Caption = nRow
                prog1.Value = nRow
                 
                cBacode1 = xlsheet2.Cells(nRow, 1)
                cBacode2 = xlsheet2.Cells(nRow, 2)
                If cBacode1 <> "" Then
                    
                    Dim ItemTable As New ADODB.RecordSet
                    Set ItemTable = ItemFind(cBacode1, con)
                    If Not ItemTable.EOF Then
                        If ItemTable!ISDOZ Then
                            nQuant = Round(Val(xlsheet2.Cells(nRow, 3)) / 12, 2)
                        Else
                            nQuant = Val(xlsheet2.Cells(nRow, 3))
                        End If
                        cString = "Insert into FILE6_50 (DOC_NO,ITEM,QUANT)" & _
                               "Values(" & _
                               addstring(xDoc_No.text) & "," & _
                               addstring(ItemTable!item) & "," & _
                               nQuant & _
                               ")"
                        cStr1 = cStr1 & cString & " ; "
                        nCount1 = nCount1 + 1
                        If nCount1 = 500 Then
                            nCount1 = 0
                            If cStr1 <> "" Then con.Execute cStr1
                            Inform " ≈÷«›… 500 "
                            cStr1 = ""
                        End If
                    Else
                        MsgBox " ﬂÊœ €Ì— ’ÕÌÕ  " & cBacode1 & " ”ÿ—  " & nRow
                    End If
                End If
            End If
        Next nRow
        If cStr1 <> "" Then
            Dim db As New clsDb
            db.Execute cStr1
            Set db = Nothing
        End If
        
        MsgBox "   „  «·√÷«›… "
        
        myUndo
        
    End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdInform_Click()
ordersLookup Me, osearchDoc, cFilter
End Sub
Private Sub CmdNewInv_Click()
'On Error Resume Next
myDefine
xDoc_No.SetFocus
Err.Clear
End Sub
Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
Dim aRow As Variant, aSub As Variant
Set printGrdNew.myform = Me
grid1.RowHidden(grid1.Rows - 1) = True
   'cHead1 = " ›’Ì·Ï ≈Ã„«·Ï ÿ·»Ì… ⁄„Ì· " & xCodeDesca.Caption
   ' cHead2 = "„Ê”„  " & xMosm.text & " » «—ÌŒ :  " & xDate.text
printGrdNew.doPrint grid1, 1.1, 0, "«·⁄„Ì· : " & xCode.text & Space(3) & "ÿ·»Ì… —ﬁ„ : " & xDoc_No.text & Space(3) & "» «—ÌŒ : " & xDate.text, "„Ê”„ : " & xMosm.text, , , False, True, 10
grid1.RowHidden(grid1.Rows - 1) = False
printGrdNew.Show 1
End Sub
Private Sub cmdPaste_Click()
If grid1.Rows > 1 Then Exit Sub
If Not myValid Then Exit Sub
If xDoc_No.text = "" Or xDoc_No.text = cmdCopy.Tag Then Exit Sub
If MsgBox(cmdCopy.Tag & "   ”ÃÌ·  ﬂ—«— «·ÿ·»Ì…", vbYesNo + vbDefaultButton2) = vbYes Then
    
    If Not myreplace Then Exit Sub
    
    Dim strSql As String
    strSql = "INSERT FILE6_50(DOC_NO,ITEM,MODEL,PRICE)" & _
             "SELECT " & addstring(xDoc_No.text) & "," & _
             " ITEM," & _
             "MODEL," & _
             "PRICE" & _
             " FROM FILE6_50 WHERE DOC_NO = " & MyParn(cmdCopy.Tag)
    Dim db As New clsDb
    If db.Execute(strSql) > 0 Then
        MsgBox " „ ≈÷«›… ‰”Œ… ··ÿ·»Ì…"
        myUndo
        CmdClear_Click
    End If
    Set db = Nothing
    
                
    
'
'    CardTable.Requery
'    CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
'    If CardTable.EOF Then CardTable.MoveLast
'    myload
End If
End Sub

Private Sub cmdPrint_Click()
myPrint
End Sub

Private Sub cmdPrintModels_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    
    cHead1 = " ›’Ì·Ï ≈Ã„«·Ï ÿ·»Ì… ⁄„Ì· " & xCodeDesca.Caption
    cHead2 = "„Ê”„  " & xMosm.text & " » «—ÌŒ :  " & xDate.text
    
    PrintGrd.doPrint Me.grid1, 1, -2, cHead1, cHead2, , False, True, 10
    PrintGrd.Show 1
End Sub
Private Sub cmdPrintScal_Click()
Dim temptable As ADODB.RecordSet
Dim sourcetable As ADODB.RecordSet, nBalance As Single
Dim ModelsTable As New ADODB.RecordSet
Dim ScalTable As New ADODB.RecordSet
Dim ColorTable As New ADODB.RecordSet
Dim aHeader(3)
contemp.Execute "delete * from temp"
Set temptable = New ADODB.RecordSet
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

Dim strSql As String
strSql = "SELECT  FILE1_10.MODELNO," & _
        "FILE1_10.MODEL," & _
        " FILE1_10.desca," & _
        " FILE6_50H.mosm" & _
        " FROM  FILE6_50H " & _
        " INNER JOIN FILE6_50 ON FILE6_50H.DOC_NO = FILE6_50.DOC_NO" & _
        " INNER JOIN FILE3_10 ON FILE6_50H.CODE = FILE3_10.CODE" & _
        " INNER JOIN FILE1_10 ON FILE6_50.ITEM = FILE1_10.ITEM "
strSql = strSql & " WHERE FILE6_50H.DOC_NO = " & MyParn(xDoc_No.text)
strSql = strSql & " GROUP BY FILE1_10.MODELNO,FILE1_10.MODEL, FILE1_10.desca, FILE6_50H.mosm ORDER BY FILE1_10.MODELNO , FILE1_10.MODEL "

Dim db As New clsDb
Set ModelsTable = db.myRs(strSql)
Dim cModel As String

Dim TModelDem As RecordSet

Dim nTotal As Double
Dim nCol As Double
If ModelsTable.RecordCount = 0 Then
    MsgBox "·« ÌÊÃœ „ÊœÌ·« "
    Exit Sub
End If


db.Execute "IF OBJECT_ID('dbo.Dem_Rep1', 'U') IS NOT NULL" & _
           " DROP TABLE Dem_Rep1;"
           
           
strSql = "SELECT FILE1_10.UNIT," & _
        "FILE1_10.MODEL," & _
        "FILE1_10.C_SCAL," & _
        "FILE1_10.SCAL," & _
        "FILE1_10.COLOR," & _
        "SUM(FILE6_50.QUANT) AS T_Dem Into Dem_Rep1 " & _
        " FROM  FILE6_50H INNER JOIN FILE6_50 ON FILE6_50H.DOC_NO = FILE6_50.DOC_NO INNER JOIN FILE3_10 ON FILE6_50H.CODE = FILE3_10.CODE LEFT OUTER JOIN FILE1_10 ON FILE6_50.ITEM = FILE1_10.ITEM WHERE FILE6_50H.DOC_NO = " & MyParn(xDoc_No.text)
strSql = strSql & "GROUP BY FILE1_10.UNIT," & _
        "FILE1_10.MODEL," & _
        "FILE1_10.C_SCAL," & _
        "FILE1_10.SCAL," & _
        "FILE1_10.COLOR "

db.Execute strSql

Dim sCaption As String
sCaption = Me.Caption

With temptable
Dim comBalance As command

prog1.Visible = True
Do Until ModelsTable.EOF
    prog1.Value = (ModelsTable.AbsolutePosition / ModelsTable.RecordCount) * 100
    cModel = ModelsTable!MODEL & ""
    
    If ScalTable.State = adStateOpen Then ScalTable.Close
    strSql = "SELECT C_SCAL , SCAL FROM FILE1_10 WHERE MODEL =  " & MyParn(cModel) & " GROUP BY FILE1_10.C_SCAL , SCAL ORDER BY C_SCAL "
    
    Set ScalTable = db.myRs(strSql)
    
    strSql = "SELECT COLOR , C_COLOR FROM FILE1_10 WHERE MODEL =  " & MyParn(cModel) & " GROUP BY FILE1_10.COLOR , C_COLOR ORDER BY C_COLOR "
    If ColorTable.State = adStateOpen Then ColorTable.Close
    Set ColorTable = db.myRs(strSql)
        
    Do Until ColorTable.EOF
        Me.Caption = cModel
        .AddNew
        !STR15 = "  ›’Ì·Ï ÿ·»Ì… ··⁄„Ì· : " & xCode.text
                    
        !str17 = cModel
        !str14 = ModelsTable!MODELNO & cModel
        !str18 = ArbString(ModelsTable!DESCA)  '& " " & TurnValue(ModelsTable!UNIT, Null, "  "))
        !STR13 = TurnValue(xdesca.text, "", Null)
        !STR19 = " «—ÌŒ «· ”·Ì„ :  " & xdelivery_date.text
        
        !STR20 = " » «—ÌŒ : " & xDate.text
        !str22 = TurnValue(panel1(1).Caption)
        !str21 = xDoc_No.text
        !str23 = TurnValue(XDOC_APP.Caption, "", Null)
        nTotal = 0
        ScalTable.MoveFirst
        nCol = 0
        Do Until ScalTable.EOF
            nCol = nCol + 1
            .Fields("STR" & nCol) = ScalTable!SCAL
            .Fields("VAL" & nCol) = Val(db.rsField("SELECT T_DEM FROM Dem_Rep1 WHERE model =" & MyParn(cModel) & " and SCAL = " & MyParn(ScalTable!SCAL) & " AND COLOR = " & MyParn(ColorTable!color)) & "")
            
            aPrm = AddFlag(Empty, "MODEL", cModel)
            aPrm = AddFlag(aPrm, "SCAL", ScalTable!SCAL)
            
            Set comBalance = db.cmd("[dbo].[sp_scal_balance]", adStoredProc, aPrm)
            comBalance.Execute
            
            .Fields("VAL" & nCol + 15) = Val(comBalance.Parameters("@BALANCE") & "")
            nTotal = nTotal + Val((.Fields("VAL" & nCol)) & "")
            ScalTable.MoveNext
        Loop
        
        !STR16 = ColorTable!color
        !VAL13 = nTotal
        .Update
        ColorTable.MoveNext
    Loop
    ModelsTable.MoveNext
Loop

Me.Caption = sCaption
prog1.Visible = False

contemp.BeginTrans
contemp.CommitTrans

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    REPORT1.ReportFileName = App.Path & "\Reports\R_TDem.rpt"
    REPORT1.DataFiles(0) = tempFile
    REPORT1.Action = 1
End If
End With
Set temptable = Nothing
Set ModelsTable = Nothing
Set ScalTable = Nothing
Set ColorTable = Nothing
Set sourcetable = Nothing
Set db = Nothing
End Sub

Private Sub cmdSave_Click()
mySave
End Sub
Private Sub UpdateMobil()
Dim cError As String, pType As String, bSend As Boolean
cmdSave.Enabled = False
Me.MousePointer = vbHourglass
Dim clsApp As New clsAppOrders
If clsApp.SendAppOrder(xDoc_No.text, cError, pType) Then
    Inform " „ " & IIf(pType = "1", "«÷«›… «·ÿ·»Ì… · ÿ»Ìﬁ «·„Õ„Ê·", " ⁄œÌ· «·ÿ·»Ì… ⁄·Ì  ÿ»Ìﬁ «·„Õ„Ê·")
Else
    If cError <> "" Then MsgBox cError Else MsgBox "·„ Ì „ «÷«›… «·ÿ·»Ì… »‰Ã«Õ"
End If
Set clsApp = Nothing
Me.MousePointer = vbNormal
cmdSave.Enabled = True
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub Form_Load()
bedit = True

Dim db As New clsDb
Set xMosm.RowSource = db.myRs("SELECT * FROM MOSM ORDER BY DATE DESC")
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"

Set xsalesman.RowSource = db.myRs("SELECT * FROM auth_user order by username")
xsalesman.ListField = "username"
xsalesman.BoundColumn = "id"

Set xCode.RowSource = db.myRs("SELECT * FROM FILE3_10 order by desca")
xCode.ListField = "desca"
xCode.BoundColumn = "code"
Set db = Nothing

If Not openCardTable Then myDefine
End Sub
Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
Set orders_clients = Nothing
Err.Clear
End Sub

Private Sub grid1_DblClick()
If Not bEditRecord Then Exit Sub
If grid1.Row > 0 And grid1.Row < grid1.Rows - 1 Then
    Set addmodelOrder.myform = Me
    Set addmodelOrder.pGrid = grid1
    addmodelOrder.sModel = grid1.TextMatrix(grid1.Row, 0)
    addmodelOrder.pCol1 = 1
    addmodelOrder.pCol2 = 5
    addmodelOrder.Show 1
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And cmdSave.Enabled Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            Dim db As New clsDb
            If db.Execute("delete from FILE6_50 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) = -1 Then GoTo CleanUp
            db.Execute "UPDATE FILE6_50H SET CHANGED = 1 WHERE DOC_NO = " & MyParn(myform.xDoc_No.text)
            Set db = Nothing
        End If
        grid1.RemoveItem grid1.Row
        CalcTotals
    End If
End If
Exit Sub
CleanUp:
Set db = Nothing
End Sub
Private Function myValid(Optional bIgMsg As Boolean = True) As Boolean
If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Not xCode.MatchedWithList Then
    MsgBox "·„ Ì „  ”ÃÌ· «·⁄„Ì·"
    Exit Function
End If

If Not xMosm.MatchedWithList Then
    MsgBox "·„ Ì „  ”ÃÌ· «·„Ê”„"
    Exit Function
End If
myValid = True
End Function
Private Sub myload()
xDoc_No.text = CardTable!doc_no & ""
xDate.text = myFormat_p(CardTable!Date)
xCode.BoundText = CardTable!Code & ""
xDoc_No.text = CardTable!doc_no & ""
XDOC_APP.Caption = CardTable!DOC_APP & ""
xdelivery_date.text = myFormat_p(CardTable!delivery_date)
xdeposit.text = Myvalue(CardTable!deposit)
xsalesman.BoundText = CardTable!salesman_id & ""
xMosm.BoundText = CardTable!MOSM & ""
xcount_model.Caption = CardTable!count_model & ""

bIg = True
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
bIg = False

panel1(1).Caption = CardTable!UserName & ""
'panel1(1).Caption = CardTable!USERNAME2 & " " & myFormat_p(CardTable!TIME2, True)
Handlecontrols LoadMode
myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
xDoc_No.text = ""
xDate.text = myFormat_p(Date)
XDOC_APP.Caption = ""
xdelivery_date.text = ""
xMosm.BoundText = cPMosm
xCode.BoundText = ""
xcount_model.Caption = ""
xtotal_Quant.Caption = ""
xdesca.text = ""

bIg = True
xClosed.Value = 0
bIg = False

panel1(1).Caption = sUserName
grid1.Rows = 1

Handlecontrols DefineMode
Fixgrd
End Sub
Private Sub Handlecontrols(nMode)
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode

bEditRecord = bedit And xClosed.Value = 0
xClosed.Enabled = bopt1 And nMode = LoadMode

'cmdFilter.Visible = cmdFilter.Tag <> ""
cmdNewInv.Enabled = nMode = LoadMode And bedit
cmddelInv.Enabled = nMode = LoadMode And bEditRecord

cmdPrint.Enabled = nMode = LoadMode
cmdSave.Enabled = bEditRecord
cmddelInv.Enabled = nMode = LoadMode And bEditRecord
cmdAddModel.Enabled = bEditRecord And nMode = LoadMode
cmdAddGroup.Enabled = bEditRecord And nMode = LoadMode
cmdExcelImport.Enabled = bEditRecord And nMode = LoadMode
cmdPaste.Enabled = nMode = LoadMode

Dim nRecord As Long, nRecords As Long
retRecords xDoc_No.text, nRecords, nRecord

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
     panel1(0).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_No = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_No = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_No = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_No = ""

If cmdCopy.Visible Then
    cmdCopy.Enabled = nMode = LoadMode
End If
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub SSCommand3_Click()

End Sub

Private Sub optChanged_Click(Index As Integer)
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub

Private Sub optclosed_Click(Index As Integer)
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub

Private Sub xClosed_Click()
If bIg Then Exit Sub
Dim db As New clsDb
Dim strSql As String
strSql = "UPDATE FILE6_50H" & _
          " SET FILE6_50H.CLOSED = " & xClosed.Value & _
          " WHERE FILE6_50H.DOC_NO = " & MyParn(xDoc_No.text)

If db.Execute(strSql) > 0 Then
    Inform " „ " & IIf(xClosed.Value = 0, "«€·«ﬁ «·„” ‰œ", "› Õ «·„” ‰œ")
End If
Set db = Nothing

If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub

Private Sub xCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CustLookUp Me, oSearchClient
End Sub

Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No

If Not IsDgt(xDoc_No.text) Then
     If xDoc_No.Tag = LoadMode Then
        myDefine
    Else
        xDoc_No.text = ""
    End If
    Exit Sub
End If

xDoc_No.text = RetZero(xDoc_No.text)
If (Not (CardTable.EOF)) And xDoc_No.Tag = LoadMode Then
    If CardTable!doc_no = xDoc_No.text Then
        Exit Sub
    End If
End If

If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If xDoc_No.Tag = LoadMode Then
        If Not openCardTable Then myDefine
    Else
        xDoc_No.text = ""
    End If
End If
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "«·„ÊœÌ·|" & "«·’‰›|" & "«·’‰›|" & "«··Ê‰|" & "«·„ﬁ«”|" & "«·ﬂ„Ì…|" & "”⁄— „’‰⁄|" & "| " & "»«—ﬂÊœ|" & "”⁄— „” Â·ﬂ|"
.ColWidth(0) = 1500
.ColWidth(1) = 4000
.ColWidth(2) = 6000
.ColWidth(3) = 2000
.ColWidth(4) = 1300
.ColWidth(5) = 1300
.ColWidth(6) = 1300
.ColWidth(7) = 0
.ColHidden(7) = True
.ColWidth(8) = 1500
.ColWidth(9) = 1300
.ColWidth(10) = 1300


.ColHidden(8) = True
.ColHidden(9) = True
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(0) = flexDTString
.ColHidden(1) = True
For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ColHidden(.Cols - 1) = True

.SubtotalPosition = flexSTBelow
For i = 5 To 7
    .Subtotal flexSTSum, -1, i, "##0.00", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
Next

If .Rows > 1 Then
    For i = 5 To 7
        grid1.TextMatrix(grid1.Rows - 1, i) = mRound(grid1.ValueMatrix(grid1.Rows - 1, i))
    Next
End If
End With
End Sub
Private Sub myLoadGrd()
Dim strSql As String
strSql = "SELECT FILE1_10.MODEL," & _
          "FILE6_50.ITEM," & _
          "FILE1_10.DESCA," & _
          "FILE1_10.COLOR," & _
          "FILE1_10.SCAL," & _
          "FILE6_50.QUANT," & _
          "FILE6_50.PRICE," & _
          "FILE1_10.GRITEM," & _
          "FILE1_10.BARCODE13," & _
          "FILE1_10.PRICE2," & _
          "FILE6_50.ID " & _
          " FROM FILE6_50 " & _
          " INNER JOIN FILE1_10 ON FILE6_50.ITEM = FILE1_10.ITEM" & _
          " WHERE FILE6_50.DOC_NO = " & MyParn(xDoc_No.text) & _
          " ORDER BY FILE1_10.MODEL,FILE1_10.C_SCAL"

Dim db As New clsDb
Set grid1.DataSource = db.myRs(strSql)
Set db = Nothing
Fixgrd
CalcTotals
End Sub
Sub myaddModel()
With addmodel2.grid10
Dim db As New clsDb
db.Sql = "  DELETE  FROM FILE6_50 " & _
        " FROM FILE6_50 " & _
        " INNER JOIN FILE1_10 ON FILE6_50.ITEM = FILE1_10.ITEM" & _
        " WHERE DOC_NO = " & MyParn(xDoc_No.text) & _
        " AND FILE1_10.MODEL = " & MyParn(addmodel2.xModel.text)

Dim Row As Long
Dim col As Long

For Row = 3 To .Rows - 1
    For col = 2 To .Cols - 1
        If Val(.TextMatrix(Row, col)) <> 0 Then
            db.Sql = "Insert into FILE6_50(doc_no,Item,Model,Quant,Price)" & _
                       "Values(" & _
                           addstring(xDoc_No.text) & "," & _
                           addstring(addmodel2.grid20.TextMatrix(Row, col)) & "," & _
                           addstring(addmodel2.xModel.text) & "," & _
                           .ValueMatrix(Row, col) & "," & _
                           .ValueMatrix(1, col) & _
                       ")"
        End If
    Next
Next

db.Execute "UPDATE FILE6_50H SET CHANGED = 1 WHERE DOC_NO = " & MyParn(xDoc_No.text)

If Not db.ExecuteTransaction Then GoTo CleanUp

End With
myLoadGrd
CleanUp:
Set db = Nothing
End Sub
Sub myaddModel2()
With addmodel2.grid10
    con.BeginTrans
    cString = " DELETE  FROM " & cFile & " WHERE DOC_NO = " & MyParn(xDoc_No.text) & " AND MODEL = " & MyParn(addmodel2.xModel.text)
    con.Execute cString
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(.TextMatrix(nRow, nCol)) <> 0 Then
                cString = "Insert into " & cFile & "(doc_no,Item,Model,Quant,PRICE)" & _
                           "Values(" & _
                           addstring(xDoc_No.text) & "," & _
                           addstring(addmodel2.grid20.TextMatrix(nRow, nCol)) & "," & _
                           addstring(addmodel2.xModel.text) & "," & _
                           Val(.TextMatrix(nRow, nCol)) & "," & _
                           Val(.TextMatrix(1, nCol)) & _
                           ")"
                con.Execute cString
            End If
        Next
    Next
    con.CommitTrans
End With
myLoadGrd
CalcTotals
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Function mySave(Optional bIg As Boolean = False) As Boolean
If Not myValid(True) Then Exit Function

Me.MousePointer = vbHourglass

If Not myreplace Then Exit Function

'UpdateMobil

Me.MousePointer = vbNormal

Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"

If sDoc_No <> "" Then
    Unload Me
    Exit Function
End If

If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Function
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
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_no As String = "", Optional pDoc_no_Filter As String = "") As Boolean
Dim strSql As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_no) = "" And Trim(pDoc_no_Filter) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass


On Error GoTo myerror
Set CardTable = New ADODB.RecordSet
strSql = "SELECT  TOP 1 FILE6_50H.*," & _
         " FILE3_10.discitem," & _
         " (SELECT COUNT(DISTINCT FILE1_10.MODEL) FROM FILE6_50 INNER JOIN FILE1_10 ON FILE6_50.ITEM = FILE1_10.ITEM WHERE FILE6_50.DOC_NO = FILE6_50H.DOC_NO) AS COUNT_MODEL" & _
         " FROM FILE6_50H " & _
         " LEFT JOIN FILE3_10 ON File6_50H.Code = FILE3_10.CODE"
                   
If pMode = tbMode.tbFirst Then
    cOrder = "Order by FILE6_50H.DOC_NO"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by FILE6_50H.DOC_NO DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "FILE6_50H.Doc_no = " & MyParn(pDoc_no)
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "FILE6_50H.DOC_NO < " & MyParn(pDoc_no)
    cOrder = "Order by FILE6_50H.DOC_NO DESC"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "FILE6_50H.DOC_NO > " & MyParn(pDoc_no)
    cOrder = "Order by FILE6_50H.DOC_NO"
End If

cFilter = ""
If optclosed(1).Value Then
    cFilter = cFilter & Tr(cFilter) & "FILE6_50H.CLOSED = 0"
ElseIf optclosed(2).Value Then
    cFilter = cFilter & Tr(cFilter) & "FILE6_50H.CLOSED = 1"
End If

If optChanged(1).Value Then
    cFilter = cFilter & Tr(cFilter) & "FILE6_50H.CHANGED = 1"
ElseIf optChanged(2).Value Then
    cFilter = cFilter & Tr(cFilter) & "FILE6_50H.CHANGED = 0"
End If

If cWhere <> "" Then
    strSql = strSql & " WHERE " & cWhere
End If

If cFilter <> "" Then
    strSql = strSql & IIf(cWhere = "", " WHERE ", " AND ") & cFilter
End If

strSql = strSql & " " & cOrder
Dim db As New clsDb
Set CardTable = db.myRs(strSql)

If CardTable Is Nothing Then GoTo CleanUp

If (Not CardTable.EOF) Then
    myload
    openCardTable = True
End If
CleanUp:
Me.MousePointer = vbNormal
Set db = Nothing
Exit Function
myerror:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
Resume CleanUp:
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If grid1.Rows = 1 Then Exit Sub
If bStopCell Then
    bStopCell = False
ElseIf col < 1 Then
    grid1.col = 1
ElseIf col = 1 And grid1.TextMatrix(Row, col) = "" Then
ElseIf col < 5 Then
    grid1.col = 5
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 5)
    grid1.ShowCell grid1.Row, 0
End If
End Sub
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String
Dim loctable As New ADODB.RecordSet
If pDoc_no <> "" Then
    cString = "SELECT Count(*) AS records,COUNT(CASE WHEN DOC_NO <= " & MyParn(pDoc_no) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM FILE6_50H"
If cFilter <> "" Then cString = cString & " WHERE " & cFilter

Dim db As New clsDb
Set loctable = db.myRs(cString)
If Not loctable Is Nothing Then
    If Not loctable.EOF Then
        nRecords = loctable!records
        nRecord = Val(loctable!Record & "")
    End If
End If
Set db = Nothing
End Function
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
'    If col = 0 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
'    If col = 1 And (grid1.TextMatrix(Row, col) = "" Or grid1.TextMatrix(Row, grid1.Cols - 1) = "") Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'    If grid1.col = 0 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
'    If grid1.col = 1 And (grid1.TextMatrix(grid1.Row, grid1.col) = "" Or grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Then Exit Sub
    KeyAscii = 0
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
Private Sub CalcTotals()
Dim i As Long
Dim nTotal As Long
For i = 1 To grid1.Rows - 2
    nTotal = nTotal + grid1.ValueMatrix(i, 5)
Next
xtotal_Quant.Caption = Round(nTotal, 2)
End Sub
Private Sub xdesca_GotFocus()
myGotFocus xdesca
End Sub
Private Sub xdesca_LostFocus()
myLostFocus xdesca
End Sub

Private Sub xDone_Click()
If bIg Then Exit Sub
Dim nCount As Long
Dim db As New clsDb

If xDone.Value = 1 Then
    If Not db.rsValue("SELECT COUNT(*) FROM FILE1_61H WHERE ORDER_BRANCH = " & xDoc_No.text & " AND ISCLOSED = 0", nCount) Then GoTo CleanUp
    If nCount > 0 Then
        MsgBox "⁄œœ " & nCount & "  „” ‰œ  ÕÊÌ· €Ì— „€·ﬁ ·Â–… «·ÿ·»Ì…"
        Exit Sub
    End If
    
    If Not db.rsValue("SELECT COUNT(*) FROM FILE1_61H WHERE ORDER_BRANCH = " & xDoc_No.text & " AND ISACCEPT = 0", nCount) Then GoTo CleanUp
    If nCount > 0 Then
        MsgBox "⁄œœ " & nCount & "  „” ‰œ  ÕÊÌ· ·„ Ìﬁ»· »⁄œ ·Â–… «·ÿ·»Ì…"
        Exit Sub
    End If
End If

If db.Execute("UPDATE FILE6_50H" & _
            " SET FILE6_50H.DONE = " & xDone.Value & _
            " WHERE FILE6_50H.DOC_NO = " & MyParn(xDoc_No.text)) > -1 Then
    Inform " „ " & IIf(xClosed.Value = 0, " —«Ã⁄ ⁄‰ «·«‰Â«¡", "«‰Â«¡ «·„” ‰œ")
    myUndo
End If
CleanUp:
Set db = Nothing
End Sub
Private Sub xMosm_GotFocus()
myGotFocus xMosm
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xMosm
If Not xMosm.MatchedWithList Then xMosm.BoundText = ""
End Sub
Private Sub xStore_GotFocus()
myGotFocus xStore
End Sub
Private Sub xstore_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xdelivery_date_GotFocus()
myGotFocus xdelivery_date
End Sub
Private Sub xdelivery_date_LostFocus()
myLostFocus xdelivery_date
myValidDate xdelivery_date
End Sub
Private Sub xdeposit_GotFocus()
myGotFocus xdeposit
End Sub
Private Sub xdeposit_LostFocus()
myLostFocus xdeposit
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
If xCode.MatchedWithList Then Exit Sub
If Not IsDgt(Trim(xCode.text)) Then xCode.BoundText = ""
xCode.BoundText = RetZero(Trim(xCode.text))
If Not xCode.MatchedWithList Then xCode.BoundText = ""
End Sub
Private Sub xsalesman_GotFocus()
myGotFocus xsalesman
End Sub
Private Sub xsalesman_LostFocus()
myLostFocus xsalesman
If Not xsalesman.MatchedWithList Then xsalesman.BoundText = ""
End Sub
