VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form supfrm 
   Caption         =   "»Ì«‰«  «·„Ê—œÌ‰"
   ClientHeight    =   8640
   ClientLeft      =   420
   ClientTop       =   405
   ClientWidth     =   9495
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
   ScaleHeight     =   8640
   ScaleWidth      =   9495
   Begin VB.Frame Frame13 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   5085
      Width           =   9285
      Begin VB.CommandButton CMD_FIXSEC 
         Caption         =   " ⁄œÌ· «·„ÊœÌ·«  ··ﬁ”„"
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   330
         Left            =   4230
         TabIndex        =   14
         Top             =   180
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
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   225
         Width           =   405
      End
   End
   Begin VB.Frame Frame7 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   7335
      Width           =   9240
      Begin MSDataListLib.DataCombo xcust 
         Height          =   330
         Left            =   4185
         TabIndex        =   22
         Top             =   180
         Width           =   3210
         _ExtentX        =   5662
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
      Begin MSDataListLib.DataCombo xbranch 
         Height          =   330
         Left            =   90
         TabIndex        =   23
         Top             =   180
         Width           =   2355
         _ExtentX        =   4154
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
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   225
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· ÊﬂÌ·"
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
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   225
         Width           =   525
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   8010
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   27
         TabStop         =   0   'False
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
         Picture         =   "sup2.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "sup2.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   26
         TabStop         =   0   'False
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
         Picture         =   "sup2.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "sup2.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1575
         TabIndex        =   25
         TabStop         =   0   'False
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
         Picture         =   "sup2.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "sup2.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   24
         TabStop         =   0   'False
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
         Picture         =   "sup2.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "sup2.frx":EB26
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   2115
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   45
      Width           =   7215
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5970
         Picture         =   "sup2.frx":10C75
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4785
         MaskColor       =   &H00FFFFFF&
         Picture         =   "sup2.frx":13448
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1230
         MaskColor       =   &H00FFFFFF&
         Picture         =   "sup2.frx":159F4
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "sup2.frx":1828E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2415
         MaskColor       =   &H00FFFFFF&
         Picture         =   "sup2.frx":1A6FA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
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
         Height          =   510
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "sup2.frx":1CC73
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "—’Ìœ «›  «ÕÌ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   6300
      Width           =   9240
      Begin VB.CheckBox xisopen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "€Ì— „ Êﬁ›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   225
         Width           =   1275
      End
      Begin VB.TextBox xf_balance 
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
         Left            =   5850
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   1545
      End
      Begin VB.TextBox xf_Date 
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
         Left            =   5850
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Tag             =   "date"
         Top             =   585
         Width           =   1545
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ «Ê·"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «Ê·"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   630
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4380
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   720
      Width           =   9240
      Begin VB.CommandButton cmdAddSup 
         Caption         =   "+"
         Height          =   330
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1620
         Width           =   330
      End
      Begin VB.TextBox xRemark 
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
         Height          =   690
         Left            =   315
         MaxLength       =   200
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   3600
         Width           =   7080
      End
      Begin VB.TextBox xSubCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2700
         MaxLength       =   3
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   180
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CheckBox xisonest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„Ê—œ «„«‰«  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   1365
      End
      Begin VB.OptionButton xFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„’‰⁄"
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
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1620
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton xFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„ﬂ » Ã„·…"
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
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1185
      End
      Begin VB.TextBox xEMAIL 
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
         Left            =   315
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   3240
         Width           =   7080
      End
      Begin VB.TextBox xPhone1 
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
         Left            =   315
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2520
         Width           =   7080
      End
      Begin VB.TextBox xDescA 
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
         Left            =   2700
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   4695
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   6075
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "2"
         Top             =   180
         Width           =   1320
      End
      Begin VB.TextBox xFax 
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
         Left            =   315
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2880
         Width           =   7080
      End
      Begin VB.TextBox xManager 
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
         Left            =   2700
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   4695
      End
      Begin VB.TextBox xAddress 
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
         Height          =   510
         Left            =   315
         MaxLength       =   200
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1980
         Width           =   7080
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   "..."
         Height          =   330
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1260
         Width           =   330
      End
      Begin MSAdodcLib.Adodc data1 
         Height          =   330
         Left            =   360
         Top             =   585
         Visible         =   0   'False
         Width           =   1590
         _ExtentX        =   2805
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
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   315
         Left            =   4140
         TabIndex        =   4
         Top             =   1260
         Width           =   3255
         _ExtentX        =   5741
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
      Begin MSDataListLib.DataCombo xCode_main 
         Height          =   315
         Left            =   3780
         TabIndex        =   6
         Top             =   1620
         Width           =   3615
         _ExtentX        =   6376
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
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Height          =   270
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   3645
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê»«Ì·"
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   2925
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê—œ «·—∆Ì”Ì"
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1665
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "»—Ìœ «·Ìﬂ —Ê‰Ì"
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   3285
         Width           =   1050
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ"
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„"
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   615
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œÌ— «·„”∆Ê·"
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   945
         Width           =   1110
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄‰Ê«‰"
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ã„Ê⁄…"
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1305
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«· ·Ì›Ê‰"
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   2565
         Width           =   600
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   1710
      Top             =   495
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
      Left            =   0
      Top             =   225
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   5670
      Width           =   9240
      Begin VB.TextBox xDisc 
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
         Left            =   3420
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   180
         Width           =   1140
      End
      Begin VB.TextBox xRate 
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
         Left            =   6120
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   1275
      End
      Begin VB.TextBox xdisc2 
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
         Left            =   315
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·Œ’„"
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
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   225
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·„” Â·ﬂ"
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
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ «„«‰«  ⁄·Ï «·»Ì⁄"
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
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   225
         Width           =   1680
      End
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
   Begin VB.Label xRecord 
      Alignment       =   2  'Center
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
      Height          =   420
      Left            =   4815
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   8100
      Width           =   4515
   End
End
Attribute VB_Name = "supfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer
Dim con As New adodb.Connection
Dim cFilter As String
Dim formMode As Byte, cTableName As String, cGroupname As String
Dim oSearch As New Search_abd
Dim oSearchSup As New Search_abd
Dim CardTable As adodb.Recordset
Const LoadMode = 1, DefineMode = 2

Private Sub cmdAddSup_Click()
suplookup Me, oSearchSup, "IS_SECONDARY = 0"
End Sub

Private Sub cmdGroup_Click()
Dim sBound As String
sBound = xgroup.BoundText
Dim oFlagfrm As New flag_mainfrm
oFlagfrm.sFieldCode = "[CODE]"
oFlagfrm.sFieldDesca = "[DESCA]"
oFlagfrm.sFieldName1 = "«·ﬂÊœ"
oFlagfrm.sFieldName2 = "«·»Ì«‰"
oFlagfrm.sCaption = "„Ã„Ê⁄… «·„Ê—œÌ‰"
oFlagfrm.sTable = "FILE4_50"
oFlagfrm.nZero = 2
oFlagfrm.bedit = bedit
oFlagfrm.Show 1

data1.Recordset.Requery
xgroup.BoundText = sBound
If Not xgroup.MatchedWithList Then xgroup.BoundText = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()
openCon con

Set data1.Recordset = myRecordSet("SELECT * FROM file4_50", con)
Set xgroup.RowSource = data1
xgroup.ListField = "Desca"
xgroup.BoundColumn = "Code"

Set data2.Recordset = cmd("SELECT * FROM file1_10sc order by desca ", con).Execute
Set xSection.RowSource = data2
xSection.ListField = "DESCA"
xSection.BoundColumn = "code"

Set data3.Recordset = cmd("SELECT * FROM file3_10 order by desca ", con).Execute
Set xcust.RowSource = data3
xcust.ListField = "Desca"
xcust.BoundColumn = "Code"

Set data4.Recordset = cmd("SELECT * FROM branch_fr order by desca ", con).Execute
Set xbranch.RowSource = data4
xbranch.ListField = "Desca"
xbranch.BoundColumn = "Code"
    
Set DATA5.Recordset = cmd("SELECT * FROM branch_fr order by desca ", con).Execute
Set xCode_main.RowSource = DATA5
xCode_main.ListField = "Desca"
xCode_main.BoundColumn = "Code"
    
Set DATA6.Recordset = cmd("SELECT * FROM FILE4_10 WHERE IS_SECONDARY = 0  order by desca ", con).Execute
Set xCode_main.RowSource = DATA6
xCode_main.ListField = "Desca"
xCode_main.BoundColumn = "Code"
    
If Not openCardTable Then myDefine
End Sub
Private Sub CmdAdd_Click()
myDefine
xDescA.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myError
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.Execute "Delete  From FILE4_10  Where code = " & MyParn(xCode.text)
    If Not openCardTable(tbMode.tbPrevious, xCode.text) Then
        If Not openCardTable(tbMode.tbFirst) Then
            myDefine
        End If
    End If
End If
Exit Sub
myError:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not myValid Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ »Ì«‰«  «·„Ê—œ »‰Ã«Õ"
If Not openCardTable(tbMode.tbFind, xCode.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub CmdInform_Click()
CardLookup
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
If Not openCardTable(tbMode.tbNext, xCode.text) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, xCode.text) Then
       If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
CmdInform.Enabled = (nMode = LoadMode)

xCode.Tag = nMode
Dim nRecord As Long, nRecords As Long
retRecords xCode.text, nRecords, nRecord
If nMode = LoadMode Then
    xRecord.Caption = ArbString("”Ã· " & nRecord & " „‰ " & nRecords)
Else
    xRecord.Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""

'xCode.Enabled = Not (nMode = LoadMode)
'xCode.Enabled = False
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE4_10 WHERE IS_SECONDARY = 1"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «·„Ê—œÌ‰"
oSearch.Show 1
End Sub
Sub myDefine()
Dim sCodeNew As String
sCodeNew = GetField("SELECT MAX(CODE) FROM FILE4_10 WHERE IS_SECONDARY = 1", con) & ""
If sCodeNew <> "" Then
    sCodeNew = IncrementString(sCodeNew)
Else
    sCodeNew = "A01"
End If

xCode.text = sCodeNew
xCode_main.BoundText = ""
xDescA.text = ""
xisopen.Value = 1
xisonest.Value = 1
xf_Date.text = ""
xf_balance.text = ""

xManager.text = ""
xEMAIL.text = ""
xgroup.BoundText = ""
xcust.BoundText = ""
xbranch.BoundText = ""
xAddress.text = ""
xPhone1.text = ""
xFax.text = ""
xDisc.text = ""
xdisc2.text = ""
xRate.text = ""
xFact(0).Value = 1
xRemark.text = ""
xSubCode.text = ""
Handlecontrols DefineMode

End Sub
Sub myload()
xisopen.Value = IIf(CardTable!IsOpen, 1, 0)
xisonest.Value = IIf(CardTable!ISONEST, 1, 0)

xf_Date.text = myFormat_p(CardTable!F_DATE)
xf_balance.text = Myvalue(CardTable!F_Balance)

xCode.text = CardTable!Code & ""
xDescA.text = CardTable!DESCA
xManager.text = CardTable!Manager & ""
xAddress.text = CardTable!Address & ""
xEMAIL.text = CardTable!email & ""
xPhone1.text = CardTable!PHONE1 & ""
xFax.text = CardTable!Fax & ""
xDisc.text = Myvalue(CardTable!disc)
xdisc2.text = Myvalue(CardTable!disc2)
xRate.text = Myvalue(CardTable!Rate)
xFact(0).Value = IIf(CardTable!Fact, 1, 0)
xFact(1).Value = IIf(CardTable!SUPP, 1, 0)
xgroup.BoundText = CardTable!Group & ""
xcust.BoundText = CardTable!CUST & ""
xbranch.BoundText = CardTable!branch & ""
xRemark.text = CardTable!remark & ""
xSubCode.text = CardTable!SUBCODE & ""
xCode_main.BoundText = CardTable!CODE_MAIN & ""
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(xDescA.text))
aInsert = AddFlag(aInsert, "[MANAGER]", addstring(xManager.text))
aInsert = AddFlag(aInsert, "ADDRESS", addstring(xAddress.text))
aInsert = AddFlag(aInsert, "PHONE1", addstring(xPhone1.text))
aInsert = AddFlag(aInsert, "FAX", addstring(xFax.text))
aInsert = AddFlag(aInsert, "[GROUP]", addstring(xgroup.BoundText))
aInsert = AddFlag(aInsert, "CODE_MAIN", addstring(xCode_main.BoundText))
aInsert = AddFlag(aInsert, "EMAIL", addstring(xEMAIL.text))
aInsert = AddFlag(aInsert, "DISC", Val(xDisc.text))
aInsert = AddFlag(aInsert, "DISC2", Val(xdisc2.text))

aInsert = AddFlag(aInsert, "RATE", Val(xRate.text))
aInsert = AddFlag(aInsert, "FACT", IIf(xFact(0).Value, 1, 0))
aInsert = AddFlag(aInsert, "SUPP", IIf(xFact(1).Value, 1, 0))
aInsert = AddFlag(aInsert, "remark", addstring(xRemark.text))
aInsert = AddFlag(aInsert, "SUBCODE", IIf(xFact(0).Value, "1", "null"))
aInsert = AddFlag(aInsert, "ISOPEN", xisopen.Value)
aInsert = AddFlag(aInsert, "cust", addstring(xcust.BoundText))
aInsert = AddFlag(aInsert, "branch", addstring(xbranch.BoundText))
aInsert = AddFlag(aInsert, "ISONEST", xisonest.Value)
aInsert = AddFlag(aInsert, "F_DATE", addDate(xf_Date.text))
aInsert = AddFlag(aInsert, "F_BALANCE", Val(xf_balance.text))
aInsert = AddFlag(aInsert, "IS_SECONDARY", "1")
On Error GoTo myError
If xCode.Tag = DefineMode Then
    Dim sCodeNew As String
    sCodeNew = GetField("SELECT MAX(CODE) FROM FILE4_10 WHERE IS_SECONDARY = 1", con) & ""
    If sCodeNew <> "" Then
        sCodeNew = IncrementString(sCodeNew)
    Else
        sCodeNew = "A01"
    End If
    aInsert = AddFlag(aInsert, "CODE", addstring(sCodeNew))
    con.Execute addInsert(aInsert, "FILE4_10")
Else
    con.Execute addUpdate(aInsert, "FILE4_10", "code = " & addstring(xCode.text))
End If
myreplace = True
Exit Function
myError:
MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    openCardTable tbMode.tbFind, oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
ElseIf ActiveControl.Name = xCode_main.Name Then
    xCode_main.BoundText = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    Unload oSearchSup
ElseIf ActiveControl.Name = cmdAddSup.Name Then
    xCode_main.BoundText = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    myLoadSub oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    Unload oSearchSup
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set CardTable = Nothing
closeCon con
Set supfrm = Nothing
Err.Clear
End Sub

Private Sub xcode_LostFocus()
myLostFocus xCode
If Trim(xCode.text) = "" Then
     If xCode.Tag = LoadMode Then
        myDefine
    End If
Else
    xCode.text = RetZero(xCode.text)
    If Not openCardTable(tbMode.tbFind, xCode.text) Then
       If xCode.Tag = LoadMode Then
            myDefine
       End If
    End If
End If
End Sub
Function myValid() As Boolean
If xCode.text = "" Then
    MsgBox "«·ﬂÊœ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xgroup.BoundText = "" Then
    MsgBox " ”ÃÌ· «·„Ã„Ê⁄… "
    Exit Function
End If

If xDescA.text = "" Then
    MsgBox "«·≈”„ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If Not IsDate(xf_Date.text) Then
    MsgBox " «—ÌŒ «Ê· «·„œ… ÷—Ê—Ì"
    Exit Function
End If
myValid = True
End Function
Private Sub myUndo()
If xCode.Tag = DefineMode Then
    If Not openCardTable Then
        myDefine
    End If
Else
    If Not openCardTable(tbMode.tbFind, xCode.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pCode As String = "", Optional bDefine As Boolean = True) As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pCode) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

'On Error GoTo myError
cFilter = "IS_SECONDARY = 1"

If sCode <> "" Then
    cFilter = cFilter & " AND CODE = " & MyParn(sCode)
End If

cString = "SELECT TOP 1 * FROM FILE4_10"

If pMode = tbMode.tbFirst Then
    cOrder = "Order by CODE"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by CODE DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "CODE = " & MyParn(pCode)
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "CODE < " & MyParn(pCode)
    cOrder = "Order by CODE desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "CODE > " & MyParn(pCode)
    cOrder = "Order by CODE"
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
Private Function retRecords(pCode, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New adodb.Recordset
If pCode <> "" Then
    cString = "SELECT Count(*) AS records,COUNT(CASE WHEN CODE <= " & MyParn(pCode) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM FILE4_10"
If cFilter <> "" Then cString = cString & " WHERE " & cFilter

Set loctable = mycmd(cString, con)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function

Private Sub xPhone3_GotFocus()
myGotFocus xPhone3
End Sub
Private Sub xPhone3_LostFocus()
myLostFocus xPhone3
End Sub
Private Sub xCode_main_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    suplookup Me, oSearchSup, "IS_SECONDARY = 0"
End If
End Sub

Private Sub xSection_GotFocus()
myGotFocus xSection
End Sub
Private Sub xSection_LostFocus()
myLostFocus xSection
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
Private Sub xcust_GotFocus()
myGotFocus xcust
End Sub
Private Sub xcust_LostFocus()
myLostFocus xcust
If Not xcust.MatchedWithList Then xcust.BoundText = ""
End Sub
Private Sub XBRANCH_GotFocus()
myGotFocus xbranch
End Sub
Private Sub XBRANCH_LostFocus()
myLostFocus xbranch
If Not xbranch.MatchedWithList Then xbranch.BoundText = ""
End Sub
Private Sub xf_Balance_GotFocus()
myGotFocus xf_balance
End Sub
Private Sub xf_Balance_LostFocus()
myLostFocus xf_balance
End Sub
Private Sub xf_Date_GotFocus()
myGotFocus xf_Date
End Sub
Private Sub xf_Date_LostFocus()
myLostFocus xf_Date
myValidDate xf_Date
End Sub
Private Sub XREMARK_GotFocus()
myGotFocus xRemark
End Sub
Private Sub XREMARK_LostFocus()
myLostFocus xRemark
End Sub
Private Sub xSubCode_GotFocus()
myGotFocus xSubCode
End Sub
Private Sub xSubCode_LostFocus()
myLostFocus xSubCode
End Sub
Private Sub xEMAIL_GotFocus()
myGotFocus xEMAIL
End Sub
Private Sub xEMAIL_LostFocus()
myLostFocus xEMAIL
End Sub
Private Sub xPhone1_GotFocus()
myGotFocus xPhone1
End Sub
Private Sub xPhone1_LostFocus()
myLostFocus xPhone1
End Sub
Private Sub xdesca_GotFocus()
myGotFocus xDescA
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDescA
End Sub
Private Sub xcode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xFAx_GotFocus()
myGotFocus xFax
End Sub
Private Sub xFAx_LostFocus()
myLostFocus xFax
End Sub
Private Sub xManager_GotFocus()
myGotFocus xManager
End Sub
Private Sub xManager_LostFocus()
myLostFocus xManager
End Sub
Private Sub xAddress_GotFocus()
myGotFocus xAddress
End Sub
Private Sub xAddress_LostFocus()
myLostFocus xAddress
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xgroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xgroup
If Not xgroup.MatchedWithList Then xgroup.BoundText = ""
End Sub
Private Sub xCode_main_GotFocus()
myGotFocus xCode_main
End Sub
Private Sub xCode_main_LostFocus()
myLostFocus xCode_main
If Not xCode_main.MatchedWithList Then xCode_main.BoundText = ""
End Sub
Private Sub xDisc_GotFocus()
myGotFocus xDisc
End Sub
Private Sub xDisc_LostFocus()
myLostFocus xDisc
End Sub
Private Sub xRate_GotFocus()
myGotFocus xRate
End Sub
Private Sub xRate_LostFocus()
myLostFocus xRate
End Sub
Private Sub xdisc2_GotFocus()
myGotFocus xdisc2
End Sub
Private Sub xdisc2_LostFocus()
myLostFocus xdisc2
End Sub
Private Sub myLoadSub(pCode)
Dim loctable As New adodb.Recordset
Set loctable = cmd("SELECT * FROM FILE4_10 WHERE CODE = " & MyParn(pCode), con).Execute
If Not loctable.EOF Then
    xDescA.text = loctable!DESCA
    xManager.text = loctable!Manager & ""
    xAddress.text = loctable!Address & ""
    xEMAIL.text = loctable!email & ""
    xPhone1.text = loctable!PHONE1 & ""
    xFax.text = loctable!Fax & ""
    xDisc.text = Myvalue(loctable!disc)
    xdisc2.text = Myvalue(loctable!disc2)
    xRate.text = Myvalue(loctable!Rate)
    xFact(0).Value = IIf(loctable!Fact, 1, 0)
    xFact(1).Value = IIf(loctable!SUPP, 1, 0)
    xgroup.BoundText = loctable!Group & ""
    xcust.BoundText = loctable!CUST & ""
    xbranch.BoundText = loctable!branch & ""
    xRemark.text = loctable!remark & ""
    xSubCode.text = loctable!SUBCODE & ""
End If
loctable.Close
Set loctable = Nothing
End Sub
