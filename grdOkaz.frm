VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdOkazfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ì «—»«Õ ‘Â—Ì… ··„Œ«“‰"
   ClientHeight    =   10290
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkBranch 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "⁄—÷ «·›—Ê⁄ „⁄ «·«ﬁ”«„"
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
      Height          =   240
      Left            =   5490
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   810
      Width           =   2310
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   1170
      Width           =   3840
      Begin VB.OptionButton xIs_Okaz 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "€Ì— «ﬂ«“ÌÊ‰"
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
         Height          =   285
         Index           =   2
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   180
         Width           =   1365
      End
      Begin VB.OptionButton xIs_Okaz 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«ﬂ«“ÌÊ‰"
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
         Height          =   285
         Index           =   1
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   180
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton xIs_Okaz 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
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
         Height          =   285
         Index           =   0
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   180
         Width           =   780
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   1170
      Width           =   3795
      Begin VB.OptionButton xonest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
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
         Height          =   285
         Index           =   2
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   180
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton xonest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«„«‰« "
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
         Height          =   285
         Index           =   1
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   180
         Width           =   825
      End
      Begin VB.OptionButton xonest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "€Ì— «„«‰« "
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
         Height          =   285
         Index           =   0
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   180
         Width           =   1185
      End
   End
   Begin VB.Frame fmBranch 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1710
      Width           =   3840
      Begin VB.OptionButton optStop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›—Ê⁄ „ Êﬁ›…"
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
         Height          =   270
         Index           =   2
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   225
         Width           =   1275
      End
      Begin VB.OptionButton optStop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›—⁄ Õ«·Ì…"
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
         Height          =   270
         Index           =   1
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton optStop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
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
         Height          =   270
         Index           =   3
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   225
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   1710
      Width           =   3795
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
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
         Height          =   270
         Index           =   0
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›—Ê⁄ ›ﬁÿ"
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
         Height          =   270
         Index           =   1
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " ÊﬂÌ·«  ›ﬁÿ"
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
         Height          =   270
         Index           =   2
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2295
      Width           =   7710
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   23
         TabStop         =   0   'False
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
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdOkaz.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1305
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
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
         Picture         =   "grdOkaz.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdOkaz.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   6390
         TabIndex        =   19
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "grdOkaz.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2565
         TabIndex        =   21
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "grdOkaz.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPdf 
         Height          =   555
         Left            =   3870
         TabIndex        =   20
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
         Picture         =   "grdOkaz.frx":BBFA
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSql 
         Height          =   555
         Left            =   5130
         TabIndex        =   42
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
         Picture         =   "grdOkaz.frx":E0A7
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   7875
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   45
      Width           =   12480
      Begin VB.TextBox XMOSM2 
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
         Left            =   720
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1260
         Width           =   1455
      End
      Begin VB.TextBox XCOUNT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   9675
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1980
         Width           =   1095
      End
      Begin VB.CheckBox xISDISC2_1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„»Ì⁄«  ⁄—÷ 2+1"
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
         Height          =   240
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1665
         Width           =   1815
      End
      Begin VB.TextBox xdate2 
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
         Height          =   330
         Left            =   7650
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox xDate1 
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
         Height          =   330
         Left            =   9225
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   7650
         TabIndex        =   4
         Top             =   1260
         Width           =   3120
         _ExtentX        =   5503
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
      Begin MSDataListLib.DataCombo xFact 
         Height          =   330
         Left            =   7650
         TabIndex        =   3
         Top             =   900
         Width           =   3120
         _ExtentX        =   5503
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
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   330
         Left            =   7650
         TabIndex        =   2
         Top             =   555
         Width           =   3120
         _ExtentX        =   5503
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
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   330
         Left            =   3465
         TabIndex        =   8
         Top             =   540
         Width           =   2805
         _ExtentX        =   4948
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
      Begin MSDataListLib.DataCombo xStore 
         Height          =   330
         Left            =   3465
         TabIndex        =   9
         Top             =   900
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
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
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   330
         Left            =   3465
         TabIndex        =   10
         Top             =   1260
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo xGroup_Sup 
         Height          =   330
         Left            =   3465
         TabIndex        =   7
         Top             =   180
         Width           =   2805
         _ExtentX        =   4948
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
      Begin MSDataListLib.DataCombo xage 
         Height          =   330
         Left            =   90
         TabIndex        =   11
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
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
      Begin MSDataListLib.DataCombo xsex 
         Height          =   315
         Left            =   90
         TabIndex        =   12
         Top             =   540
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo XDOC_DISC 
         Height          =   315
         Left            =   7650
         TabIndex        =   5
         Top             =   1620
         Width           =   3120
         _ExtentX        =   5503
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
      Begin MSDataListLib.DataCombo xMosm22 
         Height          =   315
         Left            =   90
         TabIndex        =   13
         Top             =   900
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo xCode 
         Height          =   315
         Left            =   3465
         TabIndex        =   67
         Top             =   1620
         Width           =   2805
         _ExtentX        =   4948
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
      Begin VB.Label Label6 
         Caption         =   "«·⁄„Ì·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   1665
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "‰Ê⁄ «·„Ê”„ "
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
         Index           =   9
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   1305
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«⁄·Ï ﬁÌ„… „»Ì⁄«  "
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
         Left            =   10845
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   2025
         Width           =   1350
      End
      Begin VB.Label Label2 
         Caption         =   "Œ’„ ⁄—Ê÷"
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
         Index           =   10
         Left            =   10890
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   1620
         Width           =   1230
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Index           =   7
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„Ê”„ œ«Œ·Ï"
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
         Height          =   240
         Index           =   8
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   945
         Width           =   1050
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "›∆… ⁄„—Ì…"
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
         Index           =   3
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "„Ã„Ê⁄… „Ê—œÌ‰"
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
         Index           =   0
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï"
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
         Height          =   270
         Left            =   10890
         TabIndex        =   34
         Top             =   180
         Width           =   300
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "„Ê”„ «·„ÊœÌ·"
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
         Index           =   6
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   585
         Width           =   1230
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
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
         Height          =   285
         Index           =   2
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   945
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "«·ﬁ”„"
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
         Index           =   1
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1305
         Width           =   690
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Ê—œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10875
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„Ã„Ê⁄… „ÊœÌ·« "
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
         Index           =   5
         Left            =   10875
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1305
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„’‰⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10875
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   930
         Width           =   690
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   5940
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   2700
      Top             =   405
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   2565
      Top             =   450
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   4005
      Top             =   135
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
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
      Caption         =   "data4"
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
      Left            =   900
      Top             =   270
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   405
      Top             =   315
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   330
      Left            =   1305
      Top             =   180
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc data8 
      Height          =   330
      Left            =   135
      Top             =   495
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc data9 
      Height          =   330
      Left            =   3510
      Top             =   90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   465
      Left            =   0
      TabIndex        =   35
      Top             =   9825
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   36
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
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
         Left            =   4095
         TabIndex        =   37
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
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
         Left            =   8100
         TabIndex        =   38
         Top             =   45
         Width           =   4000
         _ExtentX        =   7064
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
         Left            =   12150
         TabIndex        =   39
         Top             =   45
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
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
         Index           =   4
         Left            =   16155
         TabIndex        =   40
         Top             =   45
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   41
      Top             =   9630
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   2
      Left            =   14130
      TabIndex        =   47
      Top             =   2520
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   900
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
      Caption         =   "≈Ã„«·Ï „ÊœÌ·«  ⁄—Ê÷"
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   3
      Left            =   12015
      TabIndex        =   48
      Top             =   2520
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   900
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
      Caption         =   "≈Ã„«·Ï „Ê—œÌ‰"
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   5
      Left            =   7920
      TabIndex        =   49
      Top             =   2520
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   900
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
      Caption         =   "≈Ã„«·Ï „” ‰œ« "
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   4
      Left            =   9990
      TabIndex        =   50
      Top             =   2520
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   900
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
      Caption         =   "≈Ã„«·Ï ›—Ê⁄"
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6090
      Index           =   4
      Left            =   45
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3060
      Width           =   20220
      _cx             =   35666
      _cy             =   10742
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
      WordWrap        =   -1  'True
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
      Height          =   6000
      Index           =   5
      Left            =   45
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   3150
      Width           =   20220
      _cx             =   35666
      _cy             =   10583
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc DATA20 
      Height          =   330
      Index           =   1
      Left            =   4455
      Top             =   -135
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "data20"
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
   Begin MSAdodcLib.Adodc DATA20 
      Height          =   330
      Index           =   2
      Left            =   1665
      Top             =   45
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA20 
      Height          =   330
      Index           =   3
      Left            =   -225
      Top             =   585
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA20 
      Height          =   330
      Index           =   4
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA20 
      Height          =   330
      Index           =   5
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   1
      Left            =   18270
      TabIndex        =   55
      Top             =   2520
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   900
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
      Caption         =   "≈Ã„«·Ï „ÊœÌ·« "
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Width           =   1335
      _ExtentX        =   2355
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
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   6045
      Index           =   1
      Left            =   45
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3105
      Width           =   20220
      _cx             =   35666
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6090
      Index           =   2
      Left            =   45
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   3060
      Width           =   20220
      _cx             =   35666
      _cy             =   10742
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
      Cols            =   15
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6090
      Index           =   3
      Left            =   45
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   3060
      Width           =   20220
      _cx             =   35666
      _cy             =   10742
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
      WordWrap        =   -1  'True
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
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   6
      Left            =   16245
      TabIndex        =   73
      Top             =   2520
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   900
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
      Caption         =   "≈Ã„«·Ì «ﬁ”«„"
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6090
      Index           =   6
      Left            =   45
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   3060
      Width           =   20220
      _cx             =   35666
      _cy             =   10742
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
      WordWrap        =   -1  'True
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
   Begin MSAdodcLib.Adodc DATA20 
      Height          =   330
      Index           =   6
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
Attribute VB_Name = "grdOkazfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New adodb.Connection
Dim FieldTable As adodb.Recordset
Dim cString As String, aHeader()
Private Sub CMD_EXEL_Click()
End Sub
Private Sub CMD_PRINT_Click()
End Sub

Private Sub ckkBranch_Click()
If chkBranch.Value = 1 Then
    grid1(6).Tag = "rp.sp_okaz_section_store"
Else
    grid1(6).Tag = "rp.sp_okaz_section"
End If
End Sub

Private Sub cmdExcel_Click()
Me.MousePointer = 11
Dim i As Long
For i = 1 To cmdRep.UBound
    Dim aRow As Variant
    If cmdRep(i).Value = -1 Then
        If grid1(i).Rows > 1 Then
            aSub = AddFlag(Empty, "row", 1)
            aSub = AddFlag(aSub, "bold", True)
            aSub = AddFlag(aSub, "word_wrap", False)
            aSub = AddFlag(aSub, "back_color", 19)
            aRow = AddFlag(aRow, aSub)
        End If
        ToFileExelNew grid1(i), , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption & "-" & cmdRep(i).Caption, retHeader(aHeader, 0, 6), retHeader(aHeader, 4, 4), retHeader(aHeader, 8, 4))
    End If
Next
Me.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
If Not myValid Then Exit Sub
Me.MousePointer = vbHourglass
myLoadGrd
Me.MousePointer = vbNormal
End Sub
Private Function myValid() As Boolean
'If Not IsDate(xDate1.text) Then
'    MsgBox " «—ÌŒ «·„œ… «·«Ê· €Ì— „Õœœ"
'    Exit Function
'End If
myValid = True
End Function
Private Sub cmdPdf_Click()
myPrint "1", True
End Sub
Private Sub cmdPrint_Click()
myPrint
End Sub

Private Sub cmdRep_Click(Index As Integer, Value As Integer)
Handlecontrols
End Sub
Private Sub Handlecontrols()
Dim i As Long
For i = 1 To cmdRep.UBound
    cmdRep(i).BackColor = IIf(cmdRep(i).Value = -1, &HC0C0C0, &HFFFFFF)
    grid1(i).Visible = cmdRep(i).Value = -1
Next
End Sub
Private Sub cmdSql_Click()
myLoadGrd True
End Sub

Private Sub Form_Load()
openCon con

Set data1.Recordset = myCmd("SELECT * FROM FILE4_50 ", con)
Set xGroup_Sup.RowSource = data1
xGroup_Sup.ListField = "Desca"
xGroup_Sup.BoundColumn = "Code"

Set DATA2.Recordset = myCmd("Select Code,DescA From File4_10 order by Desca", con)
Set xSupp.RowSource = DATA2
xSupp.ListField = "Desca"
xSupp.BoundColumn = "Code"

Set DATA3.Recordset = myCmd("Select Code,DescA From File1_50 ORDER BY DESCA", con)
Set xGroup.RowSource = DATA3
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set DATA4.Recordset = myCmd("Select mosm ,descA From mosm ORDER BY date DESC ", con)
Set xMosm.RowSource = DATA4
xMosm.ListField = "Desca"
xMosm.BoundColumn = "MOSM"

Set DATA5.Recordset = myCmd("Select code ,desca From fact ORDER BY DESCA ", con)
Set xFact.RowSource = DATA5
xFact.ListField = "Desca"
xFact.BoundColumn = "Code"

Set DATA6.Recordset = myCmd("Select code ,desca From file3_10 ORDER BY DESCA ", con)
Set xCode.RowSource = DATA6
xCode.ListField = "Desca"
xCode.BoundColumn = "Code"

Set DATA7.Recordset = myCmd("Select Code,DescA From STORE_BR WHERE " & cUserStore, con)
Set xStore.RowSource = DATA7
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
If cBranch <> "00" Then
    xStore.BoundText = cBranchStore
    xStore.Enabled = False
End If

Set data8.Recordset = myCmd("Select Code,DescA From File1_10SC order by Desca", con)
Set XSECTION.RowSource = data8
XSECTION.ListField = "Desca"
XSECTION.BoundColumn = "Code"

Set data9.Recordset = myCmd("Select Code,DescA From File3_10  order by Desca", con)
Set xCode.RowSource = data9
xCode.ListField = "Desca"
xCode.BoundColumn = "Code"

Set data10.Recordset = myCmd("Select DOC_NO ,DescA From File0_90H order by DOC_NO DESC ", con)
Set XDOC_DISC.RowSource = data10
XDOC_DISC.ListField = "Desca"
XDOC_DISC.BoundColumn = "DOC_NO"

Set DATA11.Recordset = myCmd("Select Code,DescA From File0_82  order by Desca", con)
Set xage.RowSource = DATA11
xage.ListField = "Desca"
xage.BoundColumn = "Code"

Set data12.Recordset = myCmd("Select Code,DescA From File0_81  order by Desca", con)
Set xsex.RowSource = data12
xsex.ListField = "Desca"
xsex.BoundColumn = "Code"

Set DATA13.Recordset = myCmd("Select 'S' AS CODE UNION ALL SELECT 'W' AS CODE UNION ALL SELECT 'M' AS CODE", con)
Set xMosm22.RowSource = DATA13
xMosm22.ListField = "CODE"
xMosm22.BoundColumn = "CODE"

Dim i As Long
For i = 1 To grid1.UBound
    Set grid1(i).DataSource = DATA20(i)
    fixGrd (i)
Next

grid1(1).Tag = "rp.sp_okaz_model"
grid1(2).Tag = "rp.sp_okaz_doc"
grid1(3).Tag = "rp.sp_okaz_sup"
grid1(4).Tag = "rp.sp_okaz_store"
grid1(5).Tag = "rp.sp_okaz_doc_group"
grid1(6).Tag = "rp.sp_okaz_section"

fmBranch.Visible = lSupperVisor
optStop(1).Value = True

fixControls
cmdRep(1).Value = -1

End Sub
Private Sub myLoadGrd(Optional bString As Boolean = False)
ReDim aHeader(16)

Dim aPrm As Variant
Dim aPrm2 As Variant

If IsDate(xDate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
    aPrm2 = AddFlag(aPrm2, "DATE1", addDate(xDate1.text))
    aHeader(0) = BetweenString(xDate1.text, xdate2.text)
End If

If IsDate(xdate2.text) Then
    aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
    aPrm2 = AddFlag(aPrm2, "DATE2", addDate(xdate2.text))
    aHeader(0) = BetweenString(xDate1.text, xdate2.text)
End If

If XSECTION.MatchedWithList Then
    aPrm = AddFlag(aPrm, "SECTION", TurnValue(XSECTION.BoundText))
    aPrm2 = AddFlag(aPrm2, "SECTION", addvalue(XSECTION.BoundText))
    aHeader(1) = "«·ﬁ”„ : " & XSECTION.text
End If

If xSupp.MatchedWithList Then
    aPrm = AddFlag(aPrm, "SUP", TurnValue(xSupp.BoundText))
    aPrm2 = AddFlag(aPrm2, "SUP", addstring(xSupp.BoundText))
    aHeader(2) = "«·„Ê—œ : " & xSupp.text
End If

If xCode.MatchedWithList Then
    aPrm = AddFlag(aPrm, "CODE", TurnValue(xCode.BoundText))
    aPrm2 = AddFlag(aPrm2, "CODE", addstring(xCode.BoundText))
    aHeader(3) = "«·⁄„Ì· : " & xCode.text
End If

If xGroup.MatchedWithList Then
    aPrm = AddFlag(aPrm, "GROUP", TurnValue(xGroup.BoundText))
    aPrm2 = AddFlag(aPrm2, "GROUP", addstring(xGroup.BoundText))
    aHeader(4) = "„Ã„Ê⁄… «·’‰› : " & xGroup.text
End If

If xFact.MatchedWithList Then
    aPrm = AddFlag(aPrm, "FACT", TurnValue(xFact.BoundText))
    aPrm2 = AddFlag(aPrm2, "FACT", addstring(xFact.BoundText))
    aHeader(5) = "«·„’‰⁄ : " & xFact.text
End If

If xMosm.MatchedWithList Then
    aPrm = AddFlag(aPrm, "MOSM", TurnValue(xMosm.BoundText))
    aPrm2 = AddFlag(aPrm2, "MOSM", addstring(xMosm.BoundText))
    aHeader(6) = "«·„Ê”„ : " & xMosm.text
End If

If XMOSM2.text <> "" Then
    aPrm = AddFlag(aPrm, "MOSM2", TurnValue(XMOSM2.text))
    aPrm2 = AddFlag(aPrm2, "MOSM2", addstring(XMOSM2.text))
    aHeader(7) = "‰Ê⁄ «·„Ê”„ : " & XMOSM2.text
End If

If xMosm22.MatchedWithList Then
    aPrm = AddFlag(aPrm, "MOSM22", TurnValue(xMosm22.text))
    aPrm2 = AddFlag(aPrm2, "MOSM22", addstring(xMosm22.text))
    aHeader(8) = "„Ê”„ œ«Œ·Ì : " & XMOSM2.text
End If

If xStore.MatchedWithList Then
    aPrm = AddFlag(aPrm, "STORE", TurnValue(xStore.BoundText))
    aPrm2 = AddFlag(aPrm2, "STORE", addstring(xStore.BoundText))
    aHeader(9) = "„Œ“‰ : " & xStore.text
End If

If Option1(1).Value Or Option1(2).Value Then
    aPrm = AddFlag(aPrm, "ISBRANCH", IIf(Option1(1).Value, "1", "0"))
    aPrm2 = AddFlag(aPrm2, "ISBRANCH", IIf(Option1(1).Value, "1", "0"))
    aHeader(10) = IIf(Option1(1).Value, Option1(1).Caption, Option1(2).Caption)
End If

If optStop(1).Value Or optStop(2).Value Then
    aPrm = AddFlag(aPrm, "ISSTOP", IIf(optStop(1).Value, "0", "1"))
    aPrm2 = AddFlag(aPrm2, "ISSTOP", IIf(optStop(1).Value, "0", "1"))
    aHeader(11) = IIf(optStop(1).Value, optStop(1).Caption, optStop(2).Caption)
End If

If xsex.MatchedWithList Then
    aPrm = AddFlag(aPrm, "MODELSEX", TurnValue(xsex.BoundText))
    aPrm2 = AddFlag(aPrm2, "MODELSEX", addstring(xsex.BoundText))
    aHeader(12) = "«·‰Ê⁄ : " & xsex.text
End If

If xage.MatchedWithList Then
    aPrm = AddFlag(aPrm, "MODELAGE", TurnValue(xage.BoundText))
    aPrm2 = AddFlag(aPrm2, "MODELAGE", addstring(xage.BoundText))
    aHeader(13) = "«·›∆… «·⁄„—Ì… : " & xage.text
End If

If XDOC_DISC.MatchedWithList Then
    aPrm = AddFlag(aPrm, "DOC_DISC", TurnValue(XDOC_DISC.BoundText))
    aPrm2 = AddFlag(aPrm2, "DOC_DISC", addstring(XDOC_DISC.BoundText))
    aHeader(14) = "⁄—Ê÷ Œ’„ : " & XDOC_DISC.text
End If

If xISDISC2_1.Value = 1 Then
    aPrm = AddFlag(aPrm, "ISDISC2_1", xISDISC2_1.Value)
    aPrm2 = AddFlag(aPrm2, "ISDISC2_1", xISDISC2_1.Value)
    aHeader(15) = "⁄—÷ 2 1 "
End If

If xonest(1).Value Or optStop(2).Value Then
    aPrm = AddFlag(aPrm, "ISONEST", IIf(xonest(1).Value, "1", "0"))
    aPrm2 = AddFlag(aPrm2, "ISONEST", IIf(xonest(1).Value, "1", "0"))
    aHeader(16) = IIf(xonest(1).Value, xonest(1).Caption, xonest(2).Caption)
End If

If xIs_Okaz(1).Value Or xIs_Okaz(2).Value Then
    aPrm = AddFlag(aPrm, "Is_Okaz", IIf(xIs_Okaz(1).Value, "1", "0"))
    aPrm2 = AddFlag(aPrm2, "Is_Okaz", IIf(xIs_Okaz(1).Value, "1", "0"))
    aHeader(16) = IIf(xIs_Okaz(1).Value, xIs_Okaz(1).Caption, xIs_Okaz(2).Caption)
End If

If Val(XCOUNT.text) > 0 Then
    aPrm = AddFlag(aPrm, "top", Val(XCOUNT.text))
    aPrm2 = AddFlag(aPrm2, "top", Val(XCOUNT.text))
    aHeader(16) = "«ﬂ»— " & XCOUNT.text & " ”Ã·« "
End If

If Not bOpt5 Then
    aPrm2 = AddFlag(aPrm2, "usercode", TurnValue(nusercode))
End If
    
Dim cString As String
Dim i As Long
If Not bString Then
     For i = 1 To grid1.UBound
        cString = myPrcString(grid1(i).Tag, aPrm2)
        If cmdRep(i).Value = -1 Then
            Set DATA20(i).Recordset = myCmd(grid1(i).Tag, con, adStoredProc, aPrm, 1000)
            DATA20(i).Tag = cString
            fixGrd i
            panel1(0).Caption = "⁄œœ «·”Ã·«  : " & grid1(i).Rows - 1
        Else
            If DATA20(i).Tag <> cString Then grid1(i).Rows = 1
        End If
    Next
Else
    Clipboard.Clear
    For i = 1 To cmdRep.UBound
        If cmdRep(i).Value = -1 Then
            Clipboard.SetText myPrcString(grid1(i).Tag, aPrm)
        End If
    Next
End If
End Sub
Sub fixGrd(Index As Long)
With grid1(Index)
If Index = 1 Then
    .RowHeight(0) = 700
    
    .TextMatrix(0, 0) = "„"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„ÊœÌ·"
    .TextMatrix(0, 3) = "„Ê”„"
    .TextMatrix(0, 4) = "„ÊœÌ· "
    .TextMatrix(0, 5) = "«·’‰›"
    
    .TextMatrix(0, 6) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 7) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 8) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 9) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 10) = "‰”»… «·—»Õ"
    .TextMatrix(0, 11) = "Œ’„ «’‰«›"
    .TextMatrix(0, 12) = "Œ’„ »Ê‰"
    
    .TextMatrix(0, 13) = "ﬁÌ„… »”⁄— „” Â·ﬂ"
    .TextMatrix(0, 14) = "‰”»… „” Â·ﬂ"
    
    
    .ColHidden(0) = True
    .ColHidden(2) = True
    
    .ColHidden(8) = Not bOpt10
    .ColHidden(9) = Not bOpt10
    .ColHidden(10) = Not bOpt10

    .ColHidden(13) = Not bOpt10
    .ColHidden(14) = Not bOpt10

    .ColFormat(10) = "#.##%"
    .ColFormat(14) = "#.##%"
    
    .ColWidth(0) = 500
    .ColWidth(1) = 1000
    .ColWidth(2) = 1200
    .ColWidth(3) = 800
    .ColWidth(4) = 1500
    .ColWidth(5) = 2800
    
    .ColWidth(6) = 1100
    .ColWidth(7) = 1500
    .ColWidth(8) = 1400
    .ColWidth(9) = 1400
    .ColWidth(10) = 1200
    .ColWidth(11) = 1300
    .ColWidth(12) = 1300
    .ColWidth(13) = 1300
    .ColWidth(14) = 1000
        
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTAbove
    
    For Col = 6 To 9
        .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
        .ColDataType(Col) = flexDTDouble
    Next
    .ColDataType(10) = flexDTDouble
    For Col = 11 To 13
        .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
        .ColDataType(Col) = flexDTDouble
    Next
    
    If .Rows > 1 Then
        If .ValueMatrix(1, 8) <> 0 Then
            .TextMatrix(1, 10) = mRound(.ValueMatrix(1, 9) / .ValueMatrix(1, 8), 4)
            .TextMatrix(1, 14) = mRound((.ValueMatrix(1, 13) - .ValueMatrix(1, 8)) / .ValueMatrix(1, 8), 4)
        End If
    End If
ElseIf Index = 2 Then
    .RowHeight(0) = 700
    .TextMatrix(0, 0) = "⁄—÷ «·Œ’„"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„ÊœÌ·"
    .TextMatrix(0, 3) = "„Ê”„"
    .TextMatrix(0, 4) = "„ÊœÌ·"
    .TextMatrix(0, 5) = "«·’‰›"
    
    .TextMatrix(0, 6) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 7) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 8) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 9) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 10) = "‰”»… «·—»Õ"
    .TextMatrix(0, 11) = "Œ’„ «’‰«›"
    .TextMatrix(0, 12) = "Œ’„ »Ê‰"
    
    .TextMatrix(0, 13) = "ﬁÌ„… »”⁄— „” Â·ﬂ"
    .TextMatrix(0, 14) = "‰”»… „” Â·ﬂ"
    
    .ColFormat(10) = "#.##%"
    .ColFormat(14) = "#.##%"
    
    .ColWidth(0) = 2300
    .ColWidth(1) = 1000
    .ColWidth(2) = 1200
    .ColWidth(3) = 800
    .ColWidth(4) = 1500
    .ColWidth(5) = 2800
    
    .ColWidth(6) = 1100
    .ColWidth(7) = 1400
    .ColWidth(8) = 1400
    .ColWidth(9) = 1400
    .ColWidth(10) = 1200
    .ColWidth(11) = 1300
    .ColWidth(12) = 1300
    .ColWidth(13) = 1300
    .ColWidth(14) = 1000
       
    .ColHidden(2) = True
    .ColHidden(8) = Not bOpt10
    .ColHidden(9) = Not bOpt10
    .ColHidden(10) = Not bOpt10

    .ColHidden(13) = Not bOpt10
    .ColHidden(14) = Not bOpt10

    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTAbove
    For Col = 6 To 9
        .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    Next

    For Col = 11 To 13
        .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    Next

   .SubtotalPosition = flexSTBelow
    For Col = 6 To 9
        .Subtotal flexSTSum, 0, Col, "##,##", &HC0FFC0, vbBlack, True, "≈Ã„«·Ï " & "%s"
        .ColDataType(Col) = flexDTDouble
    Next
   .ColDataType(10) = flexDTDouble
    For Col = 11 To 13
        .Subtotal flexSTSum, 0, Col, "##,##", &HC0FFC0, vbBlack, True, "≈Ã„«·Ï " & "%s"
        .ColDataType(Col) = flexDTDouble
    Next
    
    .MergeCol(0) = True
    .MergeCells = flexMergeFree

    If .Rows > 1 Then
        For i = 1 To grid1(Index).Rows - 1
            If .ValueMatrix(i, 8) <> 0 And InStr(grid1(Index).TextMatrix(i, 0), "≈Ã„«·Ï") > 0 Then
                .TextMatrix(i, 10) = mRound(.ValueMatrix(i, 9) / .ValueMatrix(i, 8), 4)
                .TextMatrix(i, 14) = mRound((.ValueMatrix(i, 13) - .ValueMatrix(i, 8)) / .ValueMatrix(i, 8), 4)
            End If
        Next
    End If
ElseIf Index = 3 Or Index = 4 Or Index = 5 Or Index = 6 Then
    .RowHeight(0) = 700
    
    .TextMatrix(0, 0) = "«·ﬂÊœ"
    If Index = 6 Then
        .TextMatrix(0, 1) = "«·ﬁ”„"
    Else
        .TextMatrix(0, 1) = IIf(Index = 5, "⁄—÷ «·Œ’„", IIf(Index = 3, "«·„Ê—œ", "«·›—⁄"))
    End If
    
    .TextMatrix(0, 6 - 4) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 7 - 4) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 8 - 4) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 9 - 4) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 10 - 4) = "‰”»… «·—»Õ"
    .TextMatrix(0, 11 - 4) = "Œ’„ «’‰«›"
    .TextMatrix(0, 12 - 4) = "Œ’„ »Ê‰"
    
    .TextMatrix(0, 13 - 4) = "ﬁÌ„… »”⁄— „” Â·ﬂ"
    .TextMatrix(0, 14 - 4) = "‰”»… „” Â·ﬂ"
    
    
   
    .ColHidden(8 - 4) = Not bOpt10
    .ColHidden(9 - 4) = Not bOpt10
    .ColHidden(10 - 4) = Not bOpt10

    .ColHidden(13 - 4) = Not bOpt10
    .ColHidden(14 - 4) = Not bOpt10

    .ColFormat(10 - 4) = "#.##%"
    .ColFormat(14 - 4) = "#.##%"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 4000
    
    .ColWidth(6 - 4) = 1100
    .ColWidth(7 - 4) = 1500
    .ColWidth(8 - 4) = 1400
    .ColWidth(9 - 4) = 1400
    .ColWidth(10 - 4) = 1200
    .ColWidth(11 - 4) = 1300
    .ColWidth(12 - 4) = 1300
    .ColWidth(13 - 4) = 1300
    .ColWidth(14 - 4) = 1000
        
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTAbove
    
    For Col = 6 - 4 To 9 - 4
        .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
        .ColDataType(Col) = flexDTDouble
    Next
    .ColDataType(10 - 4) = flexDTDouble
    For Col = 11 - 4 To 13 - 4
        .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
        .ColDataType(Col) = flexDTDouble
    Next
    
    If .Rows > 1 Then
        If .ValueMatrix(1, 8 - 4) <> 0 Then
            .TextMatrix(1, 10 - 4) = mRound(.ValueMatrix(1, 9 - 4) / .ValueMatrix(1, 8 - 4), 4)
            .TextMatrix(1, 14 - 4) = mRound((.ValueMatrix(1, 13 - 4) - .ValueMatrix(1, 8 - 4)) / .ValueMatrix(1, 8 - 4), 4)
        End If
    End If
End If
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdBalancePrice = Nothing
End Sub

Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
Dim aRow As Variant, aSub As Variant

Dim i As Long
For i = 1 To cmdRep.UBound
    If cmdRep(i).Value = -1 Then
        aSub = AddFlag(Empty, "row", 1)
        aSub = AddFlag(aSub, "col", 0)
        aSub = AddFlag(aSub, "cols", 2)
        aSub = AddFlag(aSub, "text", "≈Ã„«·Ì")
        aRow = AddFlag(aRow, aSub)

        If grid1(i).Rows < 2 Then Exit Sub
        
        Set printGrdNew.myform = Me
        printGrdNew.DOPRINT grid1(i), 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption & " - " & cmdRep(i).Caption, ArbString(retHeader(aHeader, 0, 5)), , False, i = 5, 10, , aRow, Array(1)
    End If
Next

If Not bIgPreview Then
    printGrdNew.Show 1
Else
    Unload printGrdNew
End If
End Sub


Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub XMOSM2_GotFocus()
myGotFocus XMOSM2
End Sub
Private Sub XMOSM2_LostFocus()
myLostFocus XMOSM2
End Sub
Private Sub xModelFACT_GotFocus()
myGotFocus xModelFact
End Sub
Private Sub XModelFACT_LostFocus()
myLostFocus xModelFact
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub xFact_GotFocus()
myGotFocus xFact
End Sub
Private Sub xFACT_LostFocus()
myLostFocus xFact
If Not xFact.MatchedWithList Then xFact.BoundText = ""
End Sub
Private Sub xSupp_GotFocus()
myGotFocus xSupp
End Sub
Private Sub xSupp_LostFocus()
myLostFocus xSupp
If Not xSupp.MatchedWithList Then xSupp.BoundText = ""
End Sub
Private Sub xage_GotFocus()
myGotFocus xage
End Sub
Private Sub xage_LostFocus()
myLostFocus xage
If Not xage.MatchedWithList Then xage.BoundText = ""
End Sub
Private Sub xsex_GotFocus()
myGotFocus xsex
End Sub
Private Sub xsex_LostFocus()
myLostFocus xsex
If Not xsex.MatchedWithList Then xsex.BoundText = ""
End Sub
Private Sub xMosm_GotFocus()
myGotFocus xMosm
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xMosm
If Not xMosm.MatchedWithList Then xMosm.BoundText = ""
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xSection_GotFocus()
myGotFocus XSECTION
End Sub
Private Sub xSection_LostFocus()
myLostFocus XSECTION
If Not XSECTION.MatchedWithList Then XSECTION.BoundText = ""
End Sub
Private Sub xmosm20_GotFocus()
myGotFocus xmosm20
End Sub
Private Sub xmosm20_LostFocus()
myLostFocus xmosm20
If Not xmosm20.MatchedWithList Then xmosm20.BoundText = ""
End Sub
Private Sub fixControls()
If Dir(App.Path & "\sys_img\selected.jpg") = "" Then Exit Sub
Dim i As Long
For i = 1 To cmdRep.UBound
    cmdRep(i).PictureDn = LoadPicture(App.Path & "\sys_img\selected.jpg")
Next
End Sub
