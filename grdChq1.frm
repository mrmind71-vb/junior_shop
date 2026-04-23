VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdChqFrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ì „’«—Ì› ‘Â—Ì… ··›—Ê⁄"
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
   Begin VB.Frame fmType 
      BackColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   45
      Width           =   1140
      Begin VB.OptionButton chkType 
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
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   900
         Width           =   615
      End
      Begin VB.OptionButton chkType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " √„Ì‰« "
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
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   540
         Width           =   900
      End
      Begin VB.OptionButton chkType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "‘Ìﬂ« "
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
         Left            =   90
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   180
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   45
      Width           =   4560
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
         Height          =   270
         Index           =   1
         Left            =   135
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   180
         Width           =   945
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
         Height          =   315
         Index           =   2
         Left            =   1260
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   180
         Width           =   900
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
         Height          =   315
         Index           =   0
         Left            =   2250
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   180
         Width           =   1200
      End
      Begin VB.OptionButton optclose 
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
         Height          =   315
         Index           =   3
         Left            =   3690
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   180
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   630
      Width           =   4605
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "grdChq1.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1260
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "grdChq1.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdChq1.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   3375
         TabIndex        =   15
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "grdChq1.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2295
         TabIndex        =   23
         Top             =   135
         Width           =   1050
         _ExtentX        =   1852
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
         Picture         =   "grdChq1.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   5895
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   45
      Width           =   14460
      Begin VB.TextBox XDATE_ACTION1 
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
         Left            =   11610
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   900
         Width           =   1320
      End
      Begin VB.TextBox XDATE_ACTION2 
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
         Left            =   10215
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   900
         Width           =   1365
      End
      Begin VB.TextBox XDATE_DUE1 
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
         Left            =   11610
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1320
      End
      Begin VB.TextBox XDATE_DUE2 
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
         Left            =   10215
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   1365
      End
      Begin VB.TextBox xBnf 
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
         Left            =   6885
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   900
         Width           =   2400
      End
      Begin VB.TextBox xDate_R2 
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
         Left            =   10215
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox xDate_R1 
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
         Left            =   11610
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xBranch 
         Height          =   330
         Left            =   90
         TabIndex        =   13
         Top             =   540
         Width           =   2580
         _ExtentX        =   4551
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
      Begin MSDataListLib.DataCombo XID_BANK 
         Height          =   330
         Left            =   6885
         TabIndex        =   6
         Top             =   180
         Width           =   2400
         _ExtentX        =   4233
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
      Begin MSDataListLib.DataCombo xCharge 
         Height          =   330
         Left            =   90
         TabIndex        =   12
         Top             =   180
         Width           =   2580
         _ExtentX        =   4551
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
      Begin MSDataListLib.DataCombo xClient 
         Height          =   330
         Left            =   3690
         TabIndex        =   10
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
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
      Begin MSDataListLib.DataCombo xSup 
         Height          =   330
         Left            =   3690
         TabIndex        =   9
         Top             =   180
         Width           =   2445
         _ExtentX        =   4313
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
      Begin MSDataListLib.DataCombo xAcc 
         Height          =   330
         Left            =   3690
         TabIndex        =   11
         Top             =   900
         Width           =   2445
         _ExtentX        =   4313
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
      Begin MSDataListLib.DataCombo xBox 
         Height          =   330
         Left            =   6885
         TabIndex        =   7
         Top             =   540
         Width           =   2400
         _ExtentX        =   4233
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
      Begin MSDataListLib.DataCombo xPart 
         Height          =   330
         Left            =   90
         TabIndex        =   14
         Top             =   900
         Width           =   2580
         _ExtentX        =   4551
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
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «· Õ’Ì·"
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
         Left            =   13050
         TabIndex        =   35
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
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
         Left            =   13050
         TabIndex        =   34
         Top             =   585
         Width           =   1200
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„” ›Ìœ"
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
         Left            =   9405
         TabIndex        =   32
         Top             =   945
         Width           =   615
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Index           =   7
         Left            =   2790
         TabIndex        =   31
         Top             =   945
         Width           =   525
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   6
         Left            =   9405
         TabIndex        =   30
         Top             =   630
         Width           =   480
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Õ”«»"
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
         Index           =   5
         Left            =   6210
         TabIndex        =   29
         Top             =   945
         Width           =   585
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„’—Ê›"
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
         Index           =   4
         Left            =   2745
         TabIndex        =   28
         Top             =   180
         Width           =   750
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   3
         Left            =   6210
         TabIndex        =   27
         Top             =   585
         Width           =   450
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   2
         Left            =   6210
         TabIndex        =   26
         Top             =   225
         Width           =   480
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»‰ﬂ"
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
         Index           =   1
         Left            =   9405
         TabIndex        =   25
         Top             =   225
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «· Õ—Ì—"
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
         Left            =   13050
         TabIndex        =   19
         Top             =   180
         Width           =   990
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   2790
         TabIndex        =   18
         Top             =   540
         Width           =   405
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -720
      Top             =   -90
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   0
      Top             =   -135
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
      Height          =   7035
      Left            =   45
      TabIndex        =   16
      Top             =   1395
      Width           =   20280
      _cx             =   35772
      _cy             =   12409
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
      RightToLeft     =   0   'False
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1710
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   -495
      Top             =   -180
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
      Left            =   -135
      Top             =   -180
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   24
      Top             =   10095
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   -45
      Top             =   -270
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
   Begin MSAdodcLib.Adodc data6 
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
   Begin MSAdodcLib.Adodc DATA7 
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
   Begin MSAdodcLib.Adodc data8 
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
End
Attribute VB_Name = "grdChqFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearch As New Search_abd
Dim aHeader()

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

If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", grid1.Rows - 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 40)
    aRow = AddFlag(aRow, aSub)
End If
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), retHeader(aHeader, 4, 2))
Me.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
myload
End Sub

Private Sub cmdPdf_Click()
myPrint "1", True
End Sub
Private Sub cmdPrint_Click()
myPrint
End Sub
Private Sub Form_Load()
openCon con


Dim cWhere As String
If Not (lSupperVisor Or bOpt5 Or cBranch <> "00") Then
    cWhere = " Where BRANCH IN (SELECT FILE0_40.BRANCH FROM  USERSHOP INNER JOIN FILE0_40 WHERE USERSHOP.STORE = FILE0_40.CODE WHERE USERSHOP.CODE = " & MyParn(nusercode) & " ) " & "  ORDER BY CODE "
End If

Set data1.Recordset = mycmd("SELECT CODE,DESCA FROM BRANCH " & cWhere, con)
Set xBranch.RowSource = data1
xBranch.ListField = "Desca"
xBranch.BoundColumn = "Code"

Set DATA2.Recordset = cmd("SELECT CODE,DESCA FROM FILE5_10 ORDER BY DESCA", con).Execute
Set XID_BANK.RowSource = DATA2
XID_BANK.ListField = "Desca"
XID_BANK.BoundColumn = "code"

Set DATA3.Recordset = cmd("SELECT CODE,DESCA FROM FILE8_71 ORDER BY DESCA", con).Execute
Set XPART.RowSource = DATA3
XPART.ListField = "DESCA"
XPART.BoundColumn = "CODE"

Set DATA4.Recordset = cmd("SELECT CODE,DESCA FROM acc0_10 ORDER BY DESCA", con).Execute
Set xacc.RowSource = DATA4
xacc.ListField = "DESCA"
xacc.BoundColumn = "CODE"

Set DATA5.Recordset = cmd("SELECT CODE,DESCA FROM FILE3_10 WHERE CODE <> '0000' ORDER BY DESCA", con).Execute
Set xClient.RowSource = DATA5
xClient.ListField = "DESCA"
xClient.BoundColumn = "CODE"

Set DATA6.Recordset = cmd("SELECT CODE,DESCA FROM FILE4_10 ORDER BY DESCA", con).Execute
Set xSup.RowSource = DATA6
xSup.ListField = "DESCA"
xSup.BoundColumn = "CODE"

Set data7.Recordset = cmd("SELECT CODE,DESCA FROM FILE0_50 ORDER BY DESCA", con).Execute
Set XBOX.RowSource = data7
XBOX.ListField = "DESCA"
XBOX.BoundColumn = "CODE"

Set data8.Recordset = cmd("SELECT CODE,DESCA FROM FILE8_51 ORDER BY DESCA", con).Execute
Set XCHARGE.RowSource = data8
XCHARGE.ListField = "DESCA"
XCHARGE.BoundColumn = "CODE"

Set grid1.DataSource = DATA11

fixGrd
End Sub
Private Sub myload()
Dim cString As String, cWhere As String

ReDim aHeader(15)

If xBranch.MatchedWithList Then
    aPrm = AddFlag(aPrm, "BRANCH", xBranch.BoundText)
    aHeader(0) = "›—⁄ : " & xBranch.text
End If

If IsDate(xDate_R1.text) Then
    aPrm = AddFlag(aPrm, "DATE_R1", myFormat_sp(xDate_R1.text))
    aHeader(1) = " «—ÌŒ  Õ—Ì— : " & BetweenString(xDate_R1.text, xDate_R2.text)
End If

If IsDate(xDate_R2.text) Then
    aPrm = AddFlag(aPrm, "DATE_R2", myFormat_sp(xDate_R2.text))
    aHeader(1) = " «—ÌŒ  Õ—Ì— : " & BetweenString(xDate_R1.text, xDate_R2.text)
End If

If IsDate(XDATE_DUE1.text) Then
    aPrm = AddFlag(aPrm, "DATE_DUE1", myFormat_sp(XDATE_DUE1.text))
    aHeader(2) = " «—ÌŒ «” Õﬁ«ﬁ : " & BetweenString(XDATE_DUE1.text, XDATE_DUE2.text)
End If

If IsDate(XDATE_DUE2.text) Then
    aPrm = AddFlag(aPrm, "DATE_DUE2", myFormat_sp(XDATE_DUE2.text))
    aHeader(2) = " «—ÌŒ «” Õﬁ«ﬁ : " & BetweenString(XDATE_DUE1.text, XDATE_DUE2.text)
End If

If IsDate(XDATE_ACTION1.text) Then
    aPrm = AddFlag(aPrm, "DATE_ACTION1", myFormat_sp(XDATE_ACTION1.text))
    aHeader(3) = " «—ÌŒ ”œ«œ/—›÷ : " & BetweenString(XDATE_ACTION1.text, XDATE_ACTION2.text)
End If

If IsDate(XDATE_ACTION2.text) Then
    aPrm = AddFlag(aPrm, "DATE_ACTION2", myFormat_sp(XDATE_ACTION2.text))
    aHeader(3) = " «—ÌŒ ”œ«œ/—›÷ : " & BetweenString(XDATE_ACTION1.text, XDATE_ACTION2.text)
End If

If XID_BANK.MatchedWithList Then
    aPrm = AddFlag(aPrm, "BANK", XID_BANK.BoundText)
    aHeader(4) = "»‰ﬂ : " & XID_BANK.text
End If

If xClient.MatchedWithList Then
    aPrm = AddFlag(aPrm, "CLIENT", xClient.BoundText)
    aHeader(5) = "«·⁄„Ì· : " & xClient.text
End If

If xSup.MatchedWithList Then
    aPrm = AddFlag(aPrm, "SUP", xSup.BoundText)
    aHeader(6) = "«·„Ê—œ : " & xSup.text
End If

If XBOX.MatchedWithList Then
    aPrm = AddFlag(aPrm, "BOX", XBOX.BoundText)
    aHeader(7) = "«·Œ“‰… : " & XBOX.text
End If

If XPART.MatchedWithList Then
    aPrm = AddFlag(aPrm, "PART", XPART.BoundText)
    aHeader(8) = "«·‘—Ìﬂ : " & XPART.text
End If

If xacc.MatchedWithList Then
    aPrm = AddFlag(aPrm, "ACC", xacc.BoundText)
    aHeader(9) = "«·Õ”«» : " & xacc.text
End If

If XCHARGE.MatchedWithList Then
    aPrm = AddFlag(aPrm, "CHARGE", XCHARGE.BoundText)
    aHeader(10) = "«·„’—Ê› : " & XCHARGE.text
End If

If Trim(xBnf.text) <> "" Then
    aPrm = AddFlag(aPrm, "BNF", xBnf.text)
    aHeader(10) = "«·„” ›Ìœ : " & xBnf.text
End If

If chkType(0).Value Then
    aPrm = AddFlag(aPrm, "INS", 0)
    aHeader(11) = chkType(0).Caption
ElseIf chkType(1).Value Then
    aPrm = AddFlag(aPrm, "INS", 1)
    aHeader(11) = chkType(1).Caption
End If


If optclose(0).Value Then
    aPrm = AddFlag(aPrm, "CLOSED", "0")
    aHeader(11) = optclose(0).Caption
ElseIf optclose(1).Value Then
    aPrm = AddFlag(aPrm, "CLOSED", "1")
    aHeader(11) = optclose(0).Caption
ElseIf optclose(2).Value Then
    aPrm = AddFlag(aPrm, "CLOSED", "2")
    aHeader(11) = optclose(0).Caption
End If

If Not bOpt5 Then
    aPrm = AddFlag(aPrm, "USERCODE", nusercode)
End If

Set DATA11.Recordset = mycmd("[rp].[sp_chq_detail1]", con, adStoredProc, aPrm, 400)

fixGrd
End Sub
Sub fixGrd()
With grid1
    .RowHeight(0) = 600
    .WordWrap = True
    .TextMatrix(0, 0) = "Bank"
    .TextMatrix(0, 1) = "Supplier"
    .TextMatrix(0, 2) = "Cheque Amount"
    .TextMatrix(0, 3) = "Due Date"
    .TextMatrix(0, 4) = "Beneficiary"
    .TextMatrix(0, 5) = "Check No."
    .TextMatrix(0, 6) = "Doc. No."
    .TextMatrix(0, 7) = "Payment Status"
    .TextMatrix(0, 8) = "Action Done"
    
    .ColWidth(0) = 2500
    .ColWidth(1) = 2500
    .ColWidth(2) = 1400
    .ColWidth(3) = 1300
    .ColWidth(4) = 4500
    .ColWidth(5) = 2000
    .ColWidth(6) = 1700
    .ColWidth(7) = 1300
    .ColWidth(8) = 1300
                    
    .ColDataType(3) = flexDTDate
    .ColDataType(8) = flexDTDate
    .ColDataType(2) = flexDTDouble
    .ColFormat(2) = "#,###,##0"
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
        
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 2, "#0", &HC0FFC0, vbBlack, True
    
End With
End Sub
Private Sub Form_Resize()
grid1.Height = Me.Height - 2000
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdChargefrm1 = Nothing
End Sub
Private Function MYVALID() As Boolean
MYVALID = True
End Function

Private Sub grid1_DblClick()
'If grid1.Row > 1 And (grid1.Col = 2 Or grid1.Col = 3 Or grid1.Col = 4) Then
'    Dim aLocal As Variant
'    aLocal = AddFlag(aLocal, "DATE1", xDate1.text)
'    aLocal = AddFlag(aLocal, "DATE2", xDate2.text)
'    aLocal = AddFlag(aLocal, "MAN", grid1.TextMatrix(grid1.Row, 0))
'    aLocal = AddFlag(aLocal, "STORE", xBranch.BoundText)
'    If grid1.Col = 3 Then
'        aLocal = AddFlag(aLocal, "WHERE", "INV_TOTAL.DISCOUNT <> 0")
'        aLocal = AddFlag(aLocal, "CAPTION", " ›’Ì·Ì „»Ì⁄«  ·Â« Œ’„")
'    End If
'    showSales_inv.aData = aLocal
'    showSales_inv.Show 1
'ElseIf grid1.Col > 4 Then
'    aLocal = AddFlag(aLocal, "TYPE", 0)
'    aLocal = AddFlag(aLocal, "STORE", xBranch.BoundText)
'    aLocal = AddFlag(aLocal, "DATE1", xDate1.text)
'    aLocal = AddFlag(aLocal, "DATE2", xDate2.text)
'    aLocal = AddFlag(aLocal, "MAN", grid1.TextMatrix(grid1.Row, 0))
'    aLocal = AddFlag(aLocal, "GROUP", retFlag(aGroup(grid1.Col - 5), "CODE"))
'    showSales.aData = aLocal
'    showSales.Show 1
'End If
End Sub

Private Sub Option1_Click(Index As Integer)
'myload
End Sub
Private Sub CreateFieldTable(cWhere As String)
Dim cString As String
cString = "SELECT vw_item_exp.item ,vw_item_exp.item_desca as desca" & _
          " FROM vw_item_exp"
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " GROUP BY  vw_item_exp.item,vw_item_exp.item_desca"
cString = cString & " ORDER BY vw_item_exp.item"
Set FieldTable = Nothing
Set FieldTable = mycmd(cString, con)
End Sub
Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
If grid1.Rows < 3 Then Exit Sub
Dim aRow As Variant, aSub As Variant

For i = 1 To grid1.Rows - 2
    If grid1.TextMatrix(i, 2) = "" Then
        aSub = AddFlag(Empty, "row", i)
        aSub = AddFlag(aSub, "col", 1)
        aSub = AddFlag(aSub, "cols", 2)
        aSub = AddFlag(aSub, "text", grid1.TextMatrix(i, 1))
        aRow = AddFlag(aRow, aSub)
    End If
Next

aSub = AddFlag(Empty, "row", grid1.Rows - 1)
aSub = AddFlag(aSub, "col", 1)
aSub = AddFlag(aSub, "cols", 2)
aSub = AddFlag(aSub, "text", "≈Ã„«·Ì")
aRow = AddFlag(aRow, aSub)

'PrintGrdNew.sDeviceType = pDevice
'PrintGrdNew.Vp.Preview = Not bIgPreview

'PrintGrdNew.bIgPreview = bIgPreview
printGrdNew.DOPRINT Me.grid1, 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption, retHeader(aHeader, 0, 2), , True, True, 10, , aRow, Array(1)
If Not bIgPreview Then
    printGrdNew.Show 1
Else
    Unload printGrdNew
End If
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate2 xDate2
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xDate1
myValidDate2 xDate1
End Sub
Private Sub xProject_GotFocus()
myGotFocus xProject
End Sub
Private Sub xProject_LostFocus()
myLostFocus xProject
If Not xProject.MatchedWithList Then xProject.BoundText = ""
End Sub
Private Sub xitem_GotFocus()
myGotFocus xitem
End Sub
Private Sub xitem_LostFocus()
myLostFocus xitem
If Not xitem.MatchedWithList Then xitem.BoundText = ""
End Sub

Private Sub XID_BANK_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
     BankLookup Me, oSearch
End If
End Sub
Private Sub xBox_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
     BoxLookup Me, oSearch
End If
End Sub

Private Sub xSup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
     suplookup Me, oSearch
End If
End Sub
Private Sub xclient_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
     CLIENTLOOKUP Me, oSearch
End If
End Sub
Private Sub xACC_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
     AccLookup Me, oSearch
End If
End Sub
Private Sub xPART_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
     partLookup Me, oSearch
End If
End Sub
Private Sub xCHARGE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
     ChargeLookup Me, oSearch
End If
End Sub

Public Sub myProc()
ActiveControl.BoundText = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
Unload oSearch
End Sub
Private Sub XDATE_ACTION1_GotFocus()
myGotFocus XDATE_ACTION1
End Sub
Private Sub XDATE_ACTION1_LostFocus()
myLostFocus XDATE_ACTION1
myValidDate XDATE_ACTION1
End Sub
Private Sub XDATE_ACTION2_GotFocus()
myGotFocus XDATE_ACTION2
End Sub
Private Sub XDATE_ACTION2_LostFocus()
myLostFocus XDATE_ACTION2
myValidDate XDATE_ACTION2
End Sub
Private Sub XDATE_DUE1_GotFocus()
myGotFocus XDATE_DUE1
End Sub
Private Sub XDATE_DUE1_LostFocus()
myLostFocus XDATE_DUE1
myValidDate XDATE_DUE1
End Sub
Private Sub XDATE_DUE2_GotFocus()
myGotFocus XDATE_DUE2
End Sub
Private Sub XDATE_DUE2_LostFocus()
myLostFocus XDATE_DUE2
myValidDate XDATE_DUE2
End Sub
Private Sub xBnf_GotFocus()
myGotFocus xBnf
End Sub
Private Sub xBnf_LostFocus()
myLostFocus xBnf
End Sub
Private Sub xDate_R2_GotFocus()
myGotFocus xDate_R2
End Sub
Private Sub xDate_R2_LostFocus()
myLostFocus xDate_R2
myValidDate xDate_R2
End Sub
Private Sub xDate_R1_GotFocus()
myGotFocus xDate_R1
End Sub
Private Sub xDate_R1_LostFocus()
myLostFocus xDate_R1
myValidDate xDate_R1
End Sub
Private Sub XBRANCH_GotFocus()
myGotFocus xBranch
End Sub
Private Sub XBRANCH_LostFocus()
myLostFocus xBranch
If Not xBranch.MatchedWithList Then xBranch.BoundText = ""
End Sub
Private Sub XID_BANK_GotFocus()
myGotFocus XID_BANK
End Sub
Private Sub XID_BANK_LostFocus()
myLostFocus XID_BANK
If Not XID_BANK.MatchedWithList Then XID_BANK.BoundText = ""
End Sub
Private Sub XCHARGE_GotFocus()
myGotFocus XCHARGE
End Sub
Private Sub XCHARGE_LostFocus()
myLostFocus XCHARGE
If Not XCHARGE.MatchedWithList Then XCHARGE.BoundText = ""
End Sub
Private Sub xClient_GotFocus()
myGotFocus xClient
End Sub
Private Sub xClient_LostFocus()
myLostFocus xClient
If Not xClient.MatchedWithList Then xClient.BoundText = ""
End Sub
Private Sub xSup_GotFocus()
myGotFocus xSup
End Sub
Private Sub xSup_LostFocus()
myLostFocus xSup
If Not xSup.MatchedWithList Then xSup.BoundText = ""
End Sub
Private Sub xAcc_GotFocus()
myGotFocus xacc
End Sub
Private Sub xAcc_LostFocus()
myLostFocus xacc
If Not xacc.MatchedWithList Then xacc.BoundText = ""
End Sub
Private Sub xbox_GotFocus()
myGotFocus XBOX
End Sub
Private Sub xbox_LostFocus()
myLostFocus XBOX
If Not XBOX.MatchedWithList Then XBOX.BoundText = ""
End Sub
Private Sub XPART_GotFocus()
myGotFocus XPART
End Sub
Private Sub XPART_LostFocus()
myLostFocus XPART
If Not XPART.MatchedWithList Then XPART.BoundText = ""
End Sub
