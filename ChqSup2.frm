VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form chqsupfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "«Ê—«ﬁ œ›⁄"
   ClientHeight    =   6825
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   13935
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   13935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   0
      Width           =   2715
      Begin Threed.SSCommand cmdSum 
         Height          =   510
         Left            =   45
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   135
         Width           =   2625
         _ExtentX        =   4630
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
         Caption         =   "‘Ìﬂ«  „Ã„⁄…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   5805
      Width           =   10095
      Begin VB.OptionButton optclose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„—›Ê÷…"
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
         Index           =   1
         Left            =   2970
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   225
         Width           =   1215
      End
      Begin VB.OptionButton optclose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„Õ’·…"
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
         Index           =   2
         Left            =   5265
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   225
         Width           =   990
      End
      Begin VB.OptionButton optclose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "€Ì— „Õ’·…"
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
         Index           =   0
         Left            =   8325
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optclose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ‹‹‹·"
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
         Index           =   3
         Left            =   270
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   225
         Width           =   1215
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
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   5805
      Width           =   3615
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   32
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
         Picture         =   "ChqSup2.frx":0000
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ChqSup2.frx":21A7
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   33
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
         Picture         =   "ChqSup2.frx":41EE
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ChqSup2.frx":62D9
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   34
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
         Picture         =   "ChqSup2.frx":82D3
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ChqSup2.frx":A3E4
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   35
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
         Picture         =   "ChqSup2.frx":C3DE
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ChqSup2.frx":E602
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   3825
      Width           =   13695
      Begin VB.CheckBox xIns 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " √„Ì‰"
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   630
         Width           =   1260
      End
      Begin VB.TextBox XCHARGE 
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
         Left            =   11025
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   180
         Width           =   1230
      End
      Begin VB.TextBox XCHARGEDESCA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   8010
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   2985
      End
      Begin MSDataListLib.DataCombo XPART 
         Height          =   330
         Left            =   135
         TabIndex        =   20
         Top             =   180
         Width           =   4290
         _ExtentX        =   7567
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
      Begin MSDataListLib.DataCombo xacc 
         Height          =   330
         Left            =   8010
         TabIndex        =   19
         Top             =   540
         Width           =   4245
         _ExtentX        =   7488
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
         BackStyle       =   0  'Transparent
         Caption         =   "„’—Ê›"
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
         Left            =   12375
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   225
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·‘—Ìﬂ"
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
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   225
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Õ”«»"
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
         Left            =   12420
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   630
         Width           =   465
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   3165
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   675
      Width           =   13740
      Begin VB.TextBox xBnf 
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
         Left            =   180
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   900
         Width           =   4335
      End
      Begin VB.CheckBox xOld 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "‘Ìﬂ ”«»ﬁ"
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
         Left            =   8505
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   630
         Width           =   1260
      End
      Begin VB.TextBox xNAME4 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   180
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1260
         Width           =   4335
      End
      Begin VB.TextBox xBANK_REC 
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
         Left            =   180
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1620
         Width           =   4335
      End
      Begin VB.TextBox xdesca 
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
         Height          =   1065
         Left            =   180
         MaxLength       =   50
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1980
         Width           =   4335
      End
      Begin VB.TextBox xDATE_R 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   9855
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1980
         Width           =   2445
      End
      Begin VB.TextBox XCHK_ID 
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
         Left            =   9945
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   2355
      End
      Begin VB.TextBox XSER_NO 
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
         Left            =   9945
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "2"
         Top             =   180
         Width           =   2355
      End
      Begin VB.TextBox xValue 
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
         Left            =   9855
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   2340
         Width           =   2445
      End
      Begin VB.TextBox xDATE_1 
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
         Left            =   9855
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1620
         Width           =   2445
      End
      Begin VB.TextBox xNAME1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   6975
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   900
         Width           =   4065
      End
      Begin VB.TextBox xCode1 
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
         Left            =   11070
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   1230
      End
      Begin VB.TextBox xCode2 
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
         Left            =   11070
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1260
         Width           =   1230
      End
      Begin VB.TextBox xName2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   6975
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1260
         Width           =   4065
      End
      Begin MSDataListLib.DataCombo XID_BANK 
         Height          =   330
         Left            =   7020
         TabIndex        =   10
         Top             =   2700
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   330
         Left            =   180
         TabIndex        =   11
         Top             =   180
         Width           =   4335
         _ExtentX        =   7646
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
      Begin MSDataListLib.DataCombo XBRANCH 
         Height          =   330
         Left            =   180
         TabIndex        =   12
         Top             =   540
         Width           =   4335
         _ExtentX        =   7646
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„” ›Ìœ"
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
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê”„"
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
         Index           =   1
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   180
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
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
         Index           =   2
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   540
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·»‰ﬂ «·„”ÕÊ» ⁄·ÌÂ"
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
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   1665
         Width           =   1530
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„ŸÂ— „‰"
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
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   1305
         Width           =   675
      End
      Begin VB.Label label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·»Ì«‰"
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
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ  Õ—Ì—"
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
         Left            =   12420
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   2025
         Width           =   870
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„”·”· ‘Ìﬂ"
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
         Left            =   12420
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·‘Ìﬂ"
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
         Left            =   12420
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁÌ„…"
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
         Left            =   12420
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   2430
         Width           =   435
      End
      Begin VB.Label LabelCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·„Ê—œ"
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
         Left            =   12420
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   990
         Width           =   795
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
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
         Left            =   12375
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1665
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»‰ﬂ «·”Õ»"
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
         Index           =   0
         Left            =   12420
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   2790
         Width           =   825
      End
      Begin VB.Label LabelCode2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·⁄„Ì·"
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
         Left            =   12420
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1350
         Width           =   765
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   690
      Left            =   7470
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   0
      Width           =   6315
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   3195
         TabIndex        =   39
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   900
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
         Picture         =   "ChqSup2.frx":106D3
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ChqSup2.frx":130C8
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   900
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
         Picture         =   "ChqSup2.frx":15961
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1095
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   900
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
         Picture         =   "ChqSup2.frx":17C84
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ChqSup2.frx":1A420
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   2145
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   900
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
         Picture         =   "ChqSup2.frx":1C8B4
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ChqSup2.frx":1EAF5
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   510
         Left            =   4245
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   900
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
         Picture         =   "ChqSup2.frx":20DE2
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ChqSup2.frx":22DEA
      End
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   5265
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   900
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
         Picture         =   "ChqSup2.frx":24DA1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ChqSup2.frx":2716C
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -6120
      Top             =   1350
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
      Left            =   5850
      Top             =   6795
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   -6120
      Top             =   1665
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   4770
      Width           =   13695
      Begin VB.TextBox xMemo 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   135
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   585
         Width           =   4290
      End
      Begin VB.TextBox xDATE_3 
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
         Left            =   9000
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   630
         Width           =   1905
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "‘Ìﬂ €Ì— „Õ’·"
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
         Height          =   345
         Index           =   0
         Left            =   11385
         RightToLeft     =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   1740
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " ŸÂÌ— &  Õ’Ì· "
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
         Height          =   345
         Index           =   2
         Left            =   8775
         RightToLeft     =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   225
         Width           =   1740
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "—›÷ / —œ «·‘Ìﬂ"
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
         Height          =   345
         Index           =   1
         Left            =   6525
         RightToLeft     =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   180
         Width           =   1740
      End
      Begin MSDataListLib.DataCombo XBOX 
         Height          =   330
         Left            =   135
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   225
         Width           =   4290
         _ExtentX        =   7567
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
         BackStyle       =   0  'Transparent
         Caption         =   "«·Œ“‰…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   270
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   630
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·”œ«œ/  ŸÂÌ— / —›÷"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   270
         Left            =   11070
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   675
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -6120
      Top             =   1035
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
      Left            =   -6210
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
   Begin MSAdodcLib.Adodc DATA6 
      Height          =   330
      Left            =   -6210
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
      Left            =   -6210
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   67
      Top             =   6450
      Width           =   13935
      _ExtentX        =   24580
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
         TabIndex        =   68
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
         TabIndex        =   69
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
         TabIndex        =   70
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
End
Attribute VB_Name = "chqsupfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Dim cFilter As String
Dim con As New adodb.Connection
Dim CardTable As adodb.Recordset
Dim oSearch As New Search_abd
Dim bSumMode As Boolean
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
XSER_NO.Tag = nMode

Dim nRecord As Long, nRecords As Long
retRecords XSER_NO.text, nRecords, nRecord

bEditRecord = bedit

CmdSave.Enabled = bEditRecord
CmdAdd.Enabled = nMode = LoadMode And bedit

CmdDel.Enabled = bEditRecord And nMode = LoadMode

cmdNext.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdPrevious.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode
cmdLast.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdFirst.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    panel1(0).Caption = "”Ã· ÃœÌœ (" & (nRecords + 1) & ")"
End If
End Sub
Sub CLIENTLOOKUP(Optional nFlag As Integer = 1)

Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
If nFlag = 1 Then
    Generalarray(1) = "SELECT Code , Desca  From file4_10"
Else
    Generalarray(1) = "SELECT Code , Desca  From file3_10"
End If
Generalarray(2) = " Order by Code"
Generalarray(3) = 4000
Generalarray(5) = False

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 3000

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%desca%%)"

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„" & " " & IIf(nFlag = 1, "«·„Ê—œÌ‰", "«·⁄„·«¡")
Search3.Show 1
End Sub
Sub myDefine()
XSER_NO.text = RetZero(Val(Newflag("FILE5_21", "ser_no", con)), 6)
XCHARGE.text = ""
xBnf.text = ""
XCHARGEDESCA.text = ""
xBranch.BoundText = ""
xacc.BoundText = ""
XCHK_ID.text = ""
xOld.Value = 0
xIns.Value = 0
xCode1.text = ""
xCode2.text = ""
XID_BANK.BoundText = ""
xNAME1.text = ""
xName2.text = ""
xNAME4.text = ""
xBANK_REC.text = ""
xDATE_1.text = ""
xDATE_3.text = ""
xDATE_R.text = ""
XBOX.BoundText = ""
XPART.BoundText = ""
xValue.text = ""
xMEMO.text = ""
xMosm.BoundText = ""
xdesca.text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False
Handlecontrols DefineMode
End Sub
Sub myDefine2()
XSER_NO.text = RetZero(Val(Newflag("FILE5_21", "ser_no", con)), 6)
XCHK_ID.text = Val(XCHK_ID.text) + 1
xOld.Value = 0
xName2.text = ""
xNAME4.text = ""
xDATE_1.text = ""
xValue.text = ""
xMEMO.text = ""
xdesca.text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False
xacc.BoundText = ""
XPART.BoundText = ""
XCHARGE.text = ""

Handlecontrols DefineMode
End Sub
Sub myProc()
If TypeOf ActiveControl Is TextBox Then
    ActiveControl.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
Else
    openCardTable tbMode.tbFind, oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
End If
End Sub
Sub myload()
XSER_NO.text = CardTable!Ser_no
XCHARGE.text = CardTable!CHARGE & ""
'If Not IsNull(CardTable!CHARGE) Then
'    XCHARGEDESCA.text = GetDesca("SELECT DESCA FROM FILE8_51 WHERE CODE = " & MyParn(CardTable!CHARGE), con) & ""
'Else
'    XCHARGEDESCA.text = ""
'End If

XCHARGEDESCA.text = CardTable!CHARGE_DESCA & ""
XCHK_ID.text = CardTable!CHK_ID & ""
xCode1.text = CardTable!Code1 & ""
xBnf.text = CardTable!BNF & ""
xCode2.text = CardTable!CODE2 & ""
xNAME1.text = CardTable!desca1 & ""
xName2.text = CardTable!Desca2 & ""
xNAME4.text = CardTable!NAME4 & ""
xdesca.text = CardTable!DESCA & ""

xClosed(0).Value = IIf(CardTable!CLOSED = "0", True, False)
xClosed(1).Value = IIf(CardTable!CLOSED = "1", True, False)
xClosed(2).Value = IIf(CardTable!CLOSED = "2", True, False)

xOld.Value = IIf(CardTable!OLD, 1, 0)
xIns.Value = IIf(CardTable!ins, 1, 0)

xBANK_REC.text = CardTable!Bank_rec & ""
xValue.text = CardTable!Value
xDATE_1.text = myFormat_p(CardTable!date_1)
xDATE_3.text = myFormat_p(CardTable!date_3)
xDATE_R.text = myFormat_p(CardTable!date_R)
xMEMO.text = CardTable!Memo & ""
XID_BANK.BoundText = CardTable!ID_BANK & ""
XBOX.BoundText = CardTable!BOX & ""
XPART.BoundText = CardTable!PART & ""
xMosm.BoundText = CardTable!MOSM & ""
xBranch.BoundText = CardTable!branch & ""
xacc.BoundText = CardTable!ACC & ""

Handlecontrols LoadMode
End Sub
Function MYVALID() As Boolean
If XSER_NO.text = "" Then
    MsgBox "ÌÃ»  ”ÃÌ· „””·”· ··‘Ìﬂ"
    Exit Function
End If

If xMosm.BoundText = "" Then
    MsgBox "ÌÃ»  ”ÃÌ· «·„Ê”„ "
    Exit Function
End If

If Not IsDate(xDATE_R.text) Then
    MsgBox "ÌÃ»  ”ÃÌ·  «—ÌŒ «· Õ—Ì—"
    Exit Function
End If

If xCode1.text <> "" Then
    If GetDesca("select code from file4_10 where code = " & MyParn(xCode1.text), con) = "" Then Exit Function
End If

If xCode2.text <> "" Then
    If GetDesca("select code from file3_10 where code = " & MyParn(xCode2.text), con) = "" Then Exit Function
End If


If XSER_NO.Tag = DefineMode And Trim(xCode1.text) <> "" Then
    cString = GetDesca("select ser_no from FILE5_21 where code1 = " & MyParn(xCode1.text) & " and Chk_Id = " & MyParn(XCHK_ID.text), con) & ""
    If Trim(cString) <> "" Then
        MsgBox "‘Ìﬂ »‰›” «·—ﬁ„ ·‰›” «·≈”„"
        Exit Function
    End If
End If


If Trim(xCode1.text) = "" And Trim(xCode2.text) = "" And Trim(XCHARGE.text) = "" And (Not XPART.MatchedWithList) And (Not xacc.MatchedWithList) And (Not XBOX.MatchedWithList) Then
    MsgBox "ÌÃ»  ÕœÌœ «·„” ›Ìœ "
    Exit Function
End If

Dim sMsg As String
If Trim(xCode1.text) <> "" Then
    sMsg = "„Ê—œ"
End If

If Trim(xCode2.text) <> "" Then
    If sMsg <> "" Then
        MsgBox sMsg & " Ê " & "⁄„Ì· " & "·‰›” «·‘Ìﬂ"
        Exit Function
    Else
        sMsg = "⁄„Ì·"
    End If
End If


If Trim(XCHARGE.text) <> "" Then
    If sMsg <> "" Then
        MsgBox sMsg & " Ê " & "„’—Ê› " & "·‰›” «·‘Ìﬂ"
        Exit Function
    Else
        sMsg = "„’—Ê›"
    End If
End If

If XPART.MatchedWithList Then
    If sMsg <> "" Then
        MsgBox sMsg & " Ê " & "‘—Ìﬂ " & "·‰›” «·‘Ìﬂ"
        Exit Function
    Else
        sMsg = "‘—Ìﬂ"
    End If
End If

If xacc.MatchedWithList Then
    If sMsg <> "" Then
        MsgBox sMsg & " Ê " & "Õ”«» " & "·‰›” «·‘Ìﬂ"
        Exit Function
    Else
        sMsg = "Õ”«»"
    End If
End If

If XBOX.MatchedWithList Then
    If sMsg <> "" Then
        MsgBox sMsg & " Ê " & "Œ“‰… " & "·‰›” «·‘Ìﬂ"
        Exit Function
    End If
End If

MYVALID = True
End Function
Private Sub CmdAdd_Click()
myDefine
On Error Resume Next
XCHK_ID.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
On Error GoTo myError
con.BeginTrans
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.Execute "delete  From FILE5_21 where Ser_No = " & MyParn(XSER_NO.text)
End If
con.CommitTrans

CardTable.Requery
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.Find "SER_NO < " & MyParn(XSER_NO.text), , adSearchBackward, adBookmarkLast
    If CardTable.EOF Then CardTable.MoveFirst
    myload
Else
    If optclose(0).Value Then CmdAdd_Click Else myDefine
End If
Exit Sub
myError:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdInform_Click()
    ChqSupLookup Me, oSearch, cFilter
End Sub
Private Sub cmdSave_Click()

If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
If Not bIgMsg Then Inform " „ «·Õ›Ÿ »‰Ã«Õ"
If sCode <> "" Then
    Unload Me
    Exit Sub
End If
If XSER_NO.Tag = DefineMode Then
    If bSumMode Then cmdSum_Click Else CmdAdd_Click
Else
    If Not openCardTable(tbMode.tbFind, XSER_NO.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub cmdSum_Click()
myDefine2
On Error Resume Next
xDATE_1.SetFocus
Err.Clear
bSumMode = True
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
bedit = True

openCon con

Set data1.Recordset = cmd("FILE0_50", con, adTable).Execute

Set XBOX.RowSource = data1
XBOX.ListField = "Desca"
XBOX.BoundColumn = "Code"

Set DATA2.Recordset = cmd("FILE5_10", con, adTable).Execute

Set XID_BANK.RowSource = DATA2
XID_BANK.ListField = "Desca"
XID_BANK.BoundColumn = "code"

Set data4.Recordset = cmd("SELECT * FROM MOSM", con).Execute
Set xMosm.RowSource = data4
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"
xMosm = cPMosm

Set data5.Recordset = cmd("SELECT * FROM BRANCH ", con).Execute
Set xBranch.RowSource = data5
xBranch.ListField = "DESCA"
xBranch.BoundColumn = "CODE"

Set DATA6.Recordset = cmd("SELECT * FROM FILE8_71 ORDER BY DESCA", con).Execute
Set XPART.RowSource = DATA6
XPART.ListField = "DESCA"
XPART.BoundColumn = "CODE"

Set DATA7.Recordset = cmd("SELECT * FROM acc0_10 ORDER BY DESCA", con).Execute
Set xacc.RowSource = DATA7
xacc.ListField = "DESCA"
xacc.BoundColumn = "CODE"

If Not openCardTable Then myDefine

End Sub
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

If optclose(0).Value Then
    cFilter = "closed = '0'"
ElseIf optclose(1).Value Then
    cFilter = "closed = '1'"
ElseIf optclose(2).Value Then
    cFilter = "closed = '2'"
End If

cString = "SELECT TOP 1 FILE5_21.*," & _
          " FILE4_10.DESCA AS DESCA1," & _
          " FILE3_10.DESCA AS DESCA2," & _
          " FILE8_51.DESCA AS CHARGE_DESCA " & _
          " From FILE5_21 " & _
          " LEFT JOIN file4_10 ON FILE5_21.CODE1 = FILE4_10.CODE" & _
          " LEFT JOIN FILE3_10 ON FILE5_21.CODE2 = FILE3_10.CODE" & _
          " LEFT JOIN FILE8_51 ON FILE5_21.CHARGE = FILE8_51.CODE"

If pMode = tbMode.tbFirst Then
    cOrder = "Order by FILE5_21.SER_NO"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by FILE5_21.SER_NO DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "FILE5_21.SER_NO = " & MyParn(pDoc_no)
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "FILE5_21.SER_NO < " & MyParn(pDoc_no)
    cOrder = "Order by FILE5_21.SER_NO desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "FILE5_21.SER_NO > " & MyParn(pDoc_no)
    cOrder = "Order by FILE5_21.SER_NO"
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
myError:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function
Private Sub Form_Unload(Cancel As Integer)
Set CardTable = Nothing
closeCon con
End Sub

Private Sub optclose_Click(Index As Integer)
'cString = "Select FILE5_21.*,FILE4_10.DESCA AS DESCA1,FILE3_10.DESCA AS DESCA2" & _
'          " From (FILE5_21 left JOIN file4_10 ON FILE5_21.CODE1 = FILE4_10.CODE) LEFT JOIN FILE3_10 ON FILE5_21.CODE2 = FILE3_10.CODE"
'If Index <> 3 Then
'    cString = cString & turn(cString) & " closed = " & MyParn(Index)
'End If
'cString = cString & " Order by Ser_No"
'
'Set CardTable = New ADODB.Recordset
'CardTable.Open cString, con, adOpenKeyset, adLockOptimistic, adCmdText
'
'If Not (CardTable.EOF And CardTable.BOF) Then
'    CardTable.MoveLast
'    myload
'Else
'    If optclose(0).Value Then
'        CmdAdd_Click
'    Else
'        myDefine
'    End If
'End If
myUndo
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub XBOX_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then XBOX.BoundText = ""
End Sub

Private Sub xBOX2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xBox2.BoundText = ""
End Sub
Private Sub XCHK_ID_Validate(Cancel As Boolean)
If publicFlag = 2 And XSER_NO.Enabled Then
    If GetDesca("SELECT SER_NO FROM FILE5_21 WHERE SER_NO <> " & MyParn(XSER_NO.text) & " AND  CHK_ID = " & MyParn(XCHK_ID.text), con) <> "" Then
        MsgBox "—ﬁ„ «·‘Ìﬂ „ﬂ—— „‰ ﬁ»·"
    End If
End If
End Sub
Private Sub xClosed_Click(Index As Integer)
If Index = 0 Then
    xDATE_3.text = ""
End If
End Sub
Private Sub xCode1_Change()
If xCode1.text <> "" Then
    xCode2.text = ""
    xCode2.Enabled = False
    xName2.text = ""
Else
    xCode2.Enabled = True
End If
End Sub

Private Sub xCode1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP
End Sub
Private Sub xCode1_LostFocus()
myLostFocus xCode1
xNAME1.text = ""
If Trim(xCode1.text) = "" Then Exit Sub
xCode1.text = RetZero(xCode1.text, 3)
xNAME1.text = GetDesca("Select Desca from FILE4_10 where code = " & MyParn(xCode1.text), con)
End Sub
Private Sub xCode2_Change()
If xCode2.text <> "" Then
    xCode1.text = ""
    xCode1.Enabled = False
    xNAME1.text = ""
Else
    xCode1.Enabled = True
End If
End Sub

Private Sub xCode2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP 2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xCode2
If Trim(xCode2.text) = "" Then Exit Sub
xName2.text = ""
xCode2.text = RetZero(xCode2.text, 3)
xName2.text = GetDesca("Select Desca from FILE3_10 where code = " & MyParn(xCode2.text), con)
End Sub
Private Sub XID_BANK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then XID_BANK.BoundText = ""
End Sub

Private Sub XSER_NO_LostFocus()
myLostFocus XSER_NO
If Trim(XSER_NO.text) = "" Then
     If Xcode.Tag = LoadMode Then
        myDefine
    End If
Else
    XSER_NO.text = RetZero(XSER_NO.text)
'    If (Not (CardTable.EOF)) And xCode.Tag = LoadMode Then
'        If CardTable!CODE = xCode.text Then
'            Exit Sub
'        End If
'    End If
'
    If Not openCardTable(tbMode.tbFind, XSER_NO.text) Then
       If XSER_NO.Tag = LoadMode Then
            myDefine
       End If
    End If
End If
End Sub
Private Function myreplace() As Boolean
Dim aInsert(23, 1)
aInsert(0, 0) = "SER_NO"
aInsert(0, 1) = addstring(XSER_NO.text)

aInsert(1, 0) = "CHK_ID"
aInsert(1, 1) = addstring(XCHK_ID.text)

aInsert(2, 0) = "OLD"
aInsert(2, 1) = xOld.Value

aInsert(3, 0) = "code1"
aInsert(3, 1) = addstring(xCode1.text)

aInsert(4, 0) = "code2"
aInsert(4, 1) = addstring(xCode2.text)

aInsert(5, 0) = "[desca]"
aInsert(5, 1) = addstring(xdesca.text)

aInsert(6, 0) = "Bank_rec"
aInsert(6, 1) = addstring(xBANK_REC.text)

aInsert(7, 0) = "date_1"
aInsert(7, 1) = addDate(xDATE_1.text)

aInsert(8, 0) = "date_3"
aInsert(8, 1) = addDate(xDATE_3.text)

aInsert(9, 0) = "date_r"
aInsert(9, 1) = addDate(xDATE_R.text)

aInsert(10, 0) = "[VALUE]"
aInsert(10, 1) = Val(xValue.text)

aInsert(11, 0) = "NAME4"
aInsert(11, 1) = addstring(xNAME4.text)

aInsert(12, 0) = "BOX"
aInsert(12, 1) = addstring(XBOX.BoundText)

aInsert(13, 0) = "MEMO"
aInsert(13, 1) = addstring(xMEMO.text)

aInsert(14, 0) = "ID_BANK"
aInsert(14, 1) = addstring(XID_BANK.BoundText)

aInsert(18, 0) = "CHARGE"
aInsert(18, 1) = addstring(XCHARGE.text)


If Trim(xCode1.text) <> "" Then
    aInsert(15, 0) = "DESCA1"
    aInsert(15, 1) = addstring("„Ê—œ:" & xNAME1.text)
ElseIf Trim(xCode2.text) <> "" Then
    aInsert(15, 0) = "DESCA1"
    aInsert(15, 1) = addstring("⁄„Ì·:" & xNAME1.text)
Else
    aInsert(15, 0) = "DESCA1"
    aInsert(15, 1) = "NULL"
End If

aInsert(16, 0) = "Closed"
aInsert(16, 1) = addstring(retClose)

aInsert(17, 0) = "MOSM"
aInsert(17, 1) = addstring(xMosm.BoundText)

aInsert(19, 0) = "BRANCH"
aInsert(19, 1) = addstring(xBranch.BoundText)

aInsert(20, 0) = "PART"
aInsert(20, 1) = addstring(XPART.BoundText)

aInsert(21, 0) = "ACC"
aInsert(21, 1) = addstring(xacc.BoundText)

aInsert(22, 0) = "BNF"
aInsert(22, 1) = addstring(xBnf.text)

aInsert(2, 0) = "INS"
aInsert(2, 1) = xIns.Value

On Error GoTo myError
If XSER_NO.Tag = DefineMode Then
    XSER_NO.text = RetZero(Val(Newflag("FILE5_21", "ser_no", con)), 6)
    aInsert(0, 1) = addstring(XSER_NO.text)
    con.Execute CreateInsert(aInsert, "FILE5_21")
Else
    con.Execute CreateUpdate(aInsert, "FILE5_21", " where SER_NO = " & addstring(XSER_NO.text))
End If
myreplace = True
Exit Function
myError:
MsgBox Err.Description
Err.Clear
End Function
Private Function retClose() As String
Dim i As Integer
For i = 0 To 2
    If xClosed(i).Value Then
        retClose = i & ""
        Exit For
    End If
Next
End Function
Private Sub XCHARGE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ChargeLookup
End Sub
Private Sub ChargeLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE8_51 "
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
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
If Not openCardTable(tbMode.tbNext, XSER_NO.text) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, XSER_NO.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New adodb.Recordset
If pDoc_no <> "" Then
    cString = "SELECT Count(*) AS records,COUNT(CASE WHEN SER_NO <= " & MyParn(pDoc_no) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM FILE5_21 " & Tr(cFilter, " WHERE ") & cFilter
Set loctable = mycmd(cString, con)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Sub myUndo()
If XSER_NO.Tag = DefineMode Then
    If Not openCardTable Then
        CmdAdd_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, XSER_NO.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub


Private Sub XCHARGE_GotFocus()
myGotFocus XCHARGE
End Sub
Private Sub XCHARGE_LostFocus()
myLostFocus XCHARGE
If XCHARGE.text <> "" Then
    XCHARGEDESCA.text = GetDesca("SELECT DESCA FROM FILE8_51 WHERE CODE = " & MyParn(XCHARGE.text), con)
Else
    XCHARGEDESCA.text = ""
End If
End Sub
Private Sub XCHARGEDESCA_GotFocus()
myGotFocus XCHARGEDESCA
End Sub
Private Sub XCHARGEDESCA_LostFocus()
myLostFocus XCHARGEDESCA
End Sub
Private Sub XPART_GotFocus()
myGotFocus XPART
End Sub
Private Sub XPART_LostFocus()
myLostFocus XPART
If Not XPART.MatchedWithList Then XPART.BoundText = ""
End Sub
Private Sub xAcc_GotFocus()
myGotFocus xacc
End Sub
Private Sub xAcc_LostFocus()
myLostFocus xacc
If Not xacc.MatchedWithList Then xacc.BoundText = ""
End Sub
Private Sub xBnf_GotFocus()
myGotFocus xBnf
End Sub
Private Sub xBnf_LostFocus()
myLostFocus xBnf
End Sub
Private Sub xNAME4_GotFocus()
myGotFocus xNAME4
End Sub
Private Sub xNAME4_LostFocus()
myLostFocus xNAME4
End Sub
Private Sub xBANK_REC_GotFocus()
myGotFocus xBANK_REC
End Sub
Private Sub xBANK_REC_LostFocus()
myLostFocus xBANK_REC
End Sub
Private Sub xdesca_GotFocus()
myGotFocus xdesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xdesca
End Sub
Private Sub xDATE_R_GotFocus()
myGotFocus xDATE_R
End Sub
Private Sub xDATE_R_LostFocus()
myLostFocus xDATE_R
myValidDate xDATE_R
End Sub
Private Sub XCHK_ID_GotFocus()
myGotFocus XCHK_ID
End Sub
Private Sub XCHK_ID_LostFocus()
myLostFocus XCHK_ID
End Sub
Private Sub XSER_NO_GotFocus()
myGotFocus XSER_NO
End Sub
Private Sub xValue_GotFocus()
myGotFocus xValue
End Sub
Private Sub xValue_LostFocus()
myLostFocus xValue
End Sub
Private Sub xDATE_1_GotFocus()
myGotFocus xDATE_1
End Sub
Private Sub xDATE_1_LostFocus()
myLostFocus xDATE_1
myValidDate xDATE_1
End Sub
Private Sub xNAME1_GotFocus()
myGotFocus xNAME1
End Sub
Private Sub xNAME1_LostFocus()
myLostFocus xNAME1
End Sub
Private Sub xCode1_GotFocus()
myGotFocus xCode1
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xName2_GotFocus()
myGotFocus xName2
End Sub
Private Sub xName2_LostFocus()
myLostFocus xName2
End Sub
Private Sub XID_BANK_GotFocus()
myGotFocus XID_BANK
End Sub
Private Sub XID_BANK_LostFocus()
myLostFocus XID_BANK
If Not XID_BANK.MatchedWithList Then XID_BANK.BoundText = ""
End Sub
Private Sub xMosm_GotFocus()
myGotFocus xMosm
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xMosm
If Not xMosm.MatchedWithList Then xMosm.BoundText = ""
End Sub
Private Sub XBRANCH_GotFocus()
myGotFocus xBranch
End Sub
Private Sub XBRANCH_LostFocus()
myLostFocus xBranch
If Not xBranch.MatchedWithList Then xBranch.BoundText = ""
End Sub
Private Sub xbox_GotFocus()
myGotFocus XBOX
End Sub
Private Sub xbox_LostFocus()
myLostFocus XBOX
If Not XBOX.MatchedWithList Then XBOX.BoundText = ""
End Sub
Private Sub xMemo_GotFocus()
myGotFocus xMEMO
End Sub
Private Sub xMemo_LostFocus()
myLostFocus xMEMO
End Sub
Private Sub xDATE_3_GotFocus()
myGotFocus xDATE_3
End Sub
Private Sub xDATE_3_LostFocus()
myLostFocus xDATE_3
myValidDate xDATE_3
End Sub
