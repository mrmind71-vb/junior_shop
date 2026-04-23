VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form BoxDayfrm3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÌÊ„Ì… «·Œ“‰…"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18330
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
   ScaleHeight     =   9270
   ScaleWidth      =   18330
   Tag             =   "1"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCheckBalance 
      Caption         =   "check Box Balance"
      Height          =   510
      Left            =   -270
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   405
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   8190
      Width           =   14955
      Begin VB.CheckBox chkBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   225
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Label Label13 
         Caption         =   "’«›Ì"
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
         Left            =   5805
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   225
         Width           =   750
      End
      Begin VB.Label xBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label Label11 
         Caption         =   "„”ÕÊ»« "
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
         Left            =   8415
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   225
         Width           =   750
      End
      Begin VB.Label xMinus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label xBalanceLast 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label lblBalance2 
         Caption         =   "—’Ìœ Œ «„Ì"
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
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   225
         Width           =   1110
      End
      Begin VB.Label xPlus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label Label7 
         Caption         =   "«Ìœ«⁄« "
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
         Left            =   11250
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   225
         Width           =   570
      End
      Begin VB.Label xBalanceBefore 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Left            =   12150
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label lblBalance1 
         Caption         =   "—’Ìœ ”«»ﬁ"
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
         Left            =   13815
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   225
         Width           =   975
      End
   End
   Begin VB.Frame Frame20 
      Height          =   1095
      Left            =   9585
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   945
      Width           =   2580
      Begin VB.CheckBox xDone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«‰ Â«¡"
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
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   225
         Width           =   780
      End
      Begin VB.CommandButton cmdOpenPeriod 
         Caption         =   "› Õ › —…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   135
         Width           =   1095
      End
      Begin VB.CheckBox xclosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«€·«ﬁ"
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
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   720
         Width           =   1050
      End
      Begin VB.CommandButton cmdClosePeriod 
         Caption         =   "«€·«ﬁ › —…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   585
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   7470
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1395
      Width           =   2085
      Begin VB.CheckBox chkClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„·›«  „› ÊÕ… ›ﬁÿ"
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   12825
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CmdInform 
         Height          =   510
         Left            =   4140
         Picture         =   "box_day3.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   510
         Left            =   2775
         MaskColor       =   &H00FFFFFF&
         Picture         =   "box_day3.frx":27D3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   510
         Left            =   1395
         MaskColor       =   &H00FFFFFF&
         Picture         =   "box_day3.frx":4D7F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "box_day3.frx":7619
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   8190
      Width           =   3255
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         TabIndex        =   10
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
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
         Picture         =   "box_day3.frx":9A37
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "box_day3.frx":BC07
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   825
         TabIndex        =   11
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
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
         Picture         =   "box_day3.frx":DD4F
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "box_day3.frx":FF17
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1605
         TabIndex        =   12
         Top             =   135
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   820
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
         Picture         =   "box_day3.frx":12066
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "box_day3.frx":14246
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2430
         TabIndex        =   13
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
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
         Picture         =   "box_day3.frx":163A1
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "box_day3.frx":1855D
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   13545
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   675
      Width           =   4650
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
         Height          =   330
         Left            =   1710
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox xDoc_No 
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
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   180
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo xBox 
         Height          =   330
         Left            =   90
         TabIndex        =   2
         Top             =   900
         Width           =   3075
         _ExtentX        =   5424
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   270
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   945
         Width           =   480
      End
      Begin VB.Label Label2 
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
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   585
         Width           =   510
      End
      Begin VB.Label Label1 
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   12195
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   945
      Width           =   1275
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
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "box_day3.frx":1A6AC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "box_day3.frx":1CA0F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   8370
      Top             =   360
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
      Left            =   1530
      Top             =   450
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   22
      Top             =   8925
      Width           =   18330
      _ExtentX        =   32332
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6090
      Left            =   45
      TabIndex        =   3
      Top             =   2070
      Width           =   18150
      _cx             =   32015
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
   Begin VSFlex7Ctl.VSFlexGrid GRID2 
      Height          =   6090
      Left            =   -16425
      TabIndex        =   42
      Top             =   1080
      Visible         =   0   'False
      Width           =   16530
      _cx             =   29157
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
      Cols            =   8
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
   Begin VB.Label xbranch 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
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
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   315
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label XFLAG 
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
      Height          =   330
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "BoxDayfrm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte
Dim rsStore As New adodb.Recordset
Dim rsAccount As New adodb.Recordset
Dim rsBox As New adodb.Recordset
Dim cm As New Command
Dim aPrm As Variant
Dim bChangeBal As Boolean
Dim CLIST As String, cList2 As String, cList3 As String
Dim CardTable As adodb.Recordset
Dim cFile As String, cFileHeader As String
Dim oSearch As New Search3, oSearchCode As New Search3, oSearchGrd As New Search_abd, oSearchDoc As New Search_abd
Dim oSearchAccount As New Search_abd
Dim cdef_Box As String, bigCheck As Boolean, cFilter As String
Dim bEditRecord As Boolean, bAct As Boolean
Dim formMode
Dim con As New adodb.Connection
Dim lCellButton As Boolean
Dim nUser  As Long
Const LoadMode = 0, DefineMode = 1
Private Enum enUser
Casher = 1
Admin = 2
User = 0
End Enum
Private Function myReplace(Optional Row As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[BOX]", addstring(xBox.BoundText))
aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, IIf(Me.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cusername))
aInsert = AddFlag(aInsert, IIf(Me.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")

con.BeginTrans
On Error GoTo myerror
If Me.Tag = DefineMode Then
'    If cBranch > "60" Then
'        xDoc_No.text = NewflagBranch(cFileHeader, "doc_no", "00", con)
'    Else
'        xDoc_No.text = NewflagBranch(cFileHeader, "doc_no", cBranch, con)
'    End If
    xDoc_No.text = NewflagBox
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(xbranch.Caption))
    con.Execute addInsert(aInsert, "ACCOUNT_H")
Else
    con.Execute addUpdate(aInsert, "ACCOUNT_H", "doc_no = " & addstring(xDoc_No.text))
End If
myreplaceGrd Row
con.CommitTrans
myReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function myRefresh()
xBox.Tag = xBox.BoundText
ChangeBalance
End Function
Private Sub myreplaceGrd(Row As Long)
Dim aInsert As Variant, sCode As String
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "ACCOUNT", addvalue(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "CODE", addstring(grid1.TextMatrix(i, 1 + 1)))
        'aInsert = AddFlag(aInsert, "DATE", addDate(grid1.TextMatrix(I, 3)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(i, 4 + 1)))
        aInsert = AddFlag(aInsert, "[VALUE]", Val(grid1.TextMatrix(i, 5 + 1)))
        aInsert = AddFlag(aInsert, "[VALUE2]", Val(grid1.TextMatrix(i, 6 + 1)))
        aInsert = AddFlag(aInsert, "[BRANCH]", addstring(grid1.TextMatrix(i, 7 + 1)))
        aInsert = AddFlag(aInsert, "[MOSM]", addstring(grid1.TextMatrix(i, 8 + 1)))
        
        sCode = Mid(Trim(grid1.TextMatrix(i, 1 + 1)), 2)
        
        aInsert = AddFlag(aInsert, "[CUSTOMER]", IIf(Val(grid1.TextMatrix(i, 0 + 1)) = 1, addstring(sCode), "NULL"))
        aInsert = AddFlag(aInsert, "[SUPLER]", IIf(Val(grid1.TextMatrix(i, 0 + 1)) = 2, addstring(sCode), "NULL"))
        aInsert = AddFlag(aInsert, "[BOX]", IIf(Val(grid1.TextMatrix(i, 0 + 1)) = 3, addstring(sCode), "NULL"))
        aInsert = AddFlag(aInsert, "[BANK]", IIf(Val(grid1.TextMatrix(i, 0 + 1)) = 4, addstring(sCode), "NULL"))
        aInsert = AddFlag(aInsert, "[PART]", IIf(Val(grid1.TextMatrix(i, 0 + 1)) = 5, addstring(sCode), "NULL"))
        aInsert = AddFlag(aInsert, "[CHARGE]", IIf(Val(grid1.TextMatrix(i, 0 + 1)) = 6, addstring(sCode), "NULL"))
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "ACCOUNT_D")
        Else
            con.Execute addUpdate(aInsert, "ACCOUNT_D", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    If grid1.col = 0 Then
        grid1.TextMatrix(grid1.Row, 0) = oSearchAccount.grid1.TextMatrix(oSearchAccount.grid1.Row, 1)
        grid1.TextMatrix(grid1.Row, 1) = oSearchAccount.grid1.TextMatrix(oSearchAccount.grid1.Row, 0)
        Unload oSearchAccount
        Grid1_AfterEdit grid1.Row, grid1.col
        CellPos 13, grid1.Row, grid1.col
    ElseIf grid1.col = 2 + 1 Then
        grid1.TextMatrix(grid1.Row, 1 + 1) = grid1.TextMatrix(grid1.Row, 0 + 1) & oSearchGrd.grid1.TextMatrix(oSearchGrd.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 2 + 1) = oSearchGrd.grid1.TextMatrix(oSearchGrd.grid1.Row, 1)
        Unload oSearchGrd
        Grid1_AfterEdit grid1.Row, grid1.col
        CellPos 13, grid1.Row, grid1.col
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    openCardTable tbMode.tbFind, oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
End If
End Sub

Private Sub chkBalance_Click()
chkBalance.Enabled = False
If chkBalance.Value = 1 Then Me.MousePointer = vbHourglass
ChangeBalance True
If chkBalance.Value = 1 Then Me.MousePointer = vbNormal
chkBalance.Enabled = True

xBalanceBefore.Visible = chkBalance.Value = 1
xBalanceLast.Visible = chkBalance.Value = 1
lblBalance1.Visible = chkBalance.Value = 1
lblBalance2.Visible = chkBalance.Value = 1

CalcTotals
End Sub
Private Sub chkClosed_Click()
myUndo
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From ACCOUNT_D where Doc_No = " & MyParn(xDoc_No.text)
    con.Execute "Delete  From ACCOUNT_H where Doc_No = " & MyParn(xDoc_No.text)
    con.CommitTrans
    
    If Not openCardTable(tbMode.tbPrevious, XFLAG.Caption) Then
        If Not openCardTable(tbMode.tbFirst) Then
            myDefine
        End If
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT ACCOUNT_H.Doc_No, CONVERT(VARCHAR(10),ACCOUNT_H.[Date],111),FILE0_50.DESCA " & _
                  " FROM ACCOUNT_H INNER JOIN FILE0_50 ON ACCOUNT_H.BOX = FILE0_50.CODE"
If cFilter <> "" Then Generalarray(1) = Generalarray(1) & " WHERE " & cFilter
Generalarray(2) = "ORDER BY ACCOUNT_H.DATE DESC ,ACCOUNT_H.DOC_NO DESC"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = " «—ÌŒ-«·Œ“‰…"
listarray(0, 1) = "##ACCOUNT_H.Date## " & _
                  " or %%FILE0_50.DESCA%%"

listarray(1, 0) = "≈·Ì Õ”«»"
listarray(1, 1) = " ACCOUNT_H.DOC_NO in (Select ACCOUNT_D.DOC_NO From ACCOUNT_D inner join ACCOUNT_CODES on ACCOUNT_D.code = ACCOUNT_CODES.code where %%ACCOUNT_CODES.desca%%)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·Œ“‰…"
GrdArray(2, 1) = 4000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "√” ⁄·«„ Õ—ﬂ… ÌÊ„ÌÂ «·Œ“‰"
oSearchDoc.Show 1
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
If Not openCardTable(tbMode.tbNext, XFLAG.Caption) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, XFLAG.Caption) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdNewInv_Click()
myDefine
'xAccount.SetFocus
End Sub
Private Sub cmdSave_Click()
If FoundDay Then Exit Sub
If Not MYVALID Then Exit Sub
If Not myReplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub cmdCheckBalance_Click()
Dim loctable As New adodb.Recordset
Set loctable = cmd("SELECT BOXMOVE.BOX,SUM(BOXMOVE.PLUS) AS PLUS,SUM(BOXMOVE.MINUS) AS MINUS FROM BOXMOVE where date <= " & DateSq(xDate.text) & "  GROUP BY BOX", con).Execute

Dim cm As New adodb.Command

For i = 0 To GRID2.Cols - 1
    GRID2.ColWidth(i) = 2000
Next

GRID2.Rows = 1
i = 0
Do Until loctable.EOF
    i = i + 1
    Me.Caption = i
    GRID2.AddItem ""
    GRID2.TextMatrix(GRID2.Rows - 1, 0) = loctable!box
    GRID2.TextMatrix(GRID2.Rows - 1, 1) = Round(loctable!plus, 2)
    
    
    GRID2.TextMatrix(GRID2.Rows - 1, 3) = Round(loctable!Minus, 2)
    
    Dim aPrm As Variant
    aPrm = AddFlag(aPrm, "BOX", loctable!box)
    aPrm = AddFlag(aPrm, "date", myFormat_sp(xDate.text))
    Set cm = cmd("sp_BOX_BAL", con, adStoredProc, aPrm)
    
    cm.Execute
    GRID2.TextMatrix(GRID2.Rows - 1, 2) = Round(cm.Parameters("@PLUS").Value, 2)
    GRID2.TextMatrix(GRID2.Rows - 1, 4) = Round(cm.Parameters("@MINUS").Value, 2)
    GRID2.TextMatrix(GRID2.Rows - 1, 5) = Round(GRID2.ValueMatrix(GRID2.Rows - 1, 1) - GRID2.ValueMatrix(GRID2.Rows - 1, 3), 2)
    GRID2.TextMatrix(GRID2.Rows - 1, 6) = Round(GRID2.ValueMatrix(GRID2.Rows - 1, 2) - GRID2.ValueMatrix(GRID2.Rows - 1, 4), 2)
    GRID2.TextMatrix(GRID2.Rows - 1, 7) = Round(GRID2.ValueMatrix(GRID2.Rows - 1, 5) - GRID2.ValueMatrix(GRID2.Rows - 1, 6), 2)
    
    
    loctable.MoveNext
'    MsgBox "DONE " & loctable!BOX
Loop
End Sub
Private Sub Form_Activate()
On Error Resume Next
If Not bAct Then
    bAct = True
    If Me.Tag = DefineMode Then
        If xBox.Enabled Then
            xBox.SetFocus
        Else
            grid1.SetFocus
        End If
    Else
        grid1.SetFocus
    End If
End If
Err.Clear
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
openCon con
bEdit = True

cFileHeader = "ACCOUNT_H"
If cBranch = "00" And bSupermode Then
    nUser = enUser.Admin
ElseIf cBranchBox <> "" And cBranch <> "" Then
    nUser = enUser.Casher
End If

Set rsBox = cmd("select * from file0_50", con).Execute

Set DATA1.Recordset = mycmd("SELECT * FROM FILE0_50", con)

Set xBox.RowSource = DATA1
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"
'If nUser = enUser.Casher Then
'    xBox.Enabled = False
'End If


CLIST = StrList2("Select code,Desca from account  ORDER BY CODE", con)
cList2 = StrList2("Select code,Desca from BRANCH WHERE STOP = 0  ORDER BY CODE", con)



'cList3 = StrList2("Select MOSM,Desca from MOSM WHERE CLOSED = 0  ORDER BY DATE DESC", con)

Set rsAccount = cmd("SELECT * from [ACCOUNT]", con, adText).Execute

Set grid1.DataSource = data11

If Not openCardTable Then myDefine
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'SaveText Me, , Array(xYear.Name)
CardTable.Close
Set CardTable = Nothing
closeCon con
Set BoxDayfrm = Nothing
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)

If FoundDay Then Exit Sub

If Not MYVALID(True) Then
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


If col = 0 And Left(grid1.TextMatrix(Row, 1 + 1), 1) <> grid1.TextMatrix(Row, 0 + 1) Then
    grid1.TextMatrix(Row, 1 + 1) = ""
    grid1.TextMatrix(Row, 2 + 1) = ""
End If

If Not validRow(Row) Then
    CalcTotals
    Exit Sub
End If

With grid1
If Row = grid1.Rows - 1 Then
    myAddItem
End If

If myReplace(Row) Then
    ChangeBalance
    If Me.Tag = DefineMode Then
        openCardTable tbMode.tbFind, xDoc_No.text, True
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myLoadGrd
        grid1.ShowCell grid1.Rows - 1, 0
    Else
        CalcTotals
    End If
Else
    myLoadGrd
End If
End With
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal col As Long)
Set oSearchGrd = Nothing
Set oSearchGrd = New Search_abd
If col = 0 Then
    Dim cWhere As String
    If sBranchCode = "00" Then
    ElseIf sBranchCode < "60" Then
        cWhere = "account.IS_branch = 1"
    Else
        cWhere = "account.IS_fran = 1"
    End If
    accountLookup Me, oSearchAccount, cWhere
ElseIf grid1.TextMatrix(grid1.Row, 0 + 1) = "1" Then
    CLIENTLOOKUP Me, oSearchGrd
ElseIf grid1.TextMatrix(grid1.Row, 0 + 1) = "2" Then
    suplookup Me, oSearchGrd
ElseIf grid1.TextMatrix(grid1.Row, 0 + 1) = "3" Then
    BoxLookup Me, oSearchGrd
ElseIf grid1.TextMatrix(grid1.Row, 0 + 1) = "4" Then
    BankLookup Me, oSearchGrd
ElseIf grid1.TextMatrix(grid1.Row, 0 + 1) = "5" Then
    partLookup Me, oSearchGrd
ElseIf grid1.TextMatrix(grid1.Row, 0 + 1) = "6" Then
    ChargeLookup Me, oSearchGrd
ElseIf grid1.TextMatrix(grid1.Row, 0 + 1) = "7" Then
    ChargeLookup Me, oSearchGrd
ElseIf grid1.TextMatrix(grid1.Row, 0 + 1) = "8" Then
    IncomeLookup Me, oSearchGrd
End If
End Sub
Private Sub grid1_EnterCell()
With grid1
If Not bEditRecord Then
    grid1.Editable = flexEDNone
ElseIf grid1.col = 0 Then
    grid1.Editable = IIf(grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "", flexEDNone, flexEDKbdMouse)
ElseIf grid1.TextMatrix(grid1.Row, 0 + 1) = "" Then
    grid1.Editable = flexEDNone
ElseIf grid1.col = 2 + 1 Then
    grid1.Editable = flexEDKbdMouse
ElseIf grid1.TextMatrix(grid1.Row, 1 + 1) = "" Then
    grid1.Editable = flexEDNone
ElseIf grid1.col = 7 + 1 And grid1.ValueMatrix(grid1.Row, 9 + 1) = 0 Then
    grid1.Editable = flexEDNone
ElseIf grid1.col = 8 + 1 And grid1.ValueMatrix(grid1.Row, 10 + 1) = 0 Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End With
End Sub
Private Sub grid1_GotFocus()
grid1_EnterCell
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not IsDate(xDate.text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— „”Ã·"
    Exit Function
End If
MYVALID = True
End Function
Private Sub myload(Optional bRefresh As Boolean = False)
bChangeBal = True
xDoc_No.text = CardTable!doc_no
xBox.BoundText = CardTable!box & ""
xDate.text = myFormat_p(CardTable!Date)

ChangeBalance

'If xDate.Tag <> xDate.text Or xDoc_No.Tag <> xDoc_No.text Or xBox.Tag <> xBox.BoundText Then
'    ChangeBalance
'End If

'xDoc_No.Tag = CardTable!doc_no
'xBox.Tag = CardTable!BOX & ""
'xDate.Tag = xDate.text


XFLAG.Caption = CardTable!Flag
xbranch.Caption = CardTable!branch & ""
bigCheck = True
xclosed.Value = IIf(CardTable!CLOSED, 1, 0)
xDone.Value = IIf(CardTable!DONE, 1, 0)
bigCheck = False
If Not IsNull(CardTable!UserName) Then
    StatusBar1.Panels(2).text = ArbString(CardTable!UserName & " [" & myFormat_p(CardTable!Time, True) & "]")
End If
If Not IsNull(CardTable!USERNAME2) Then
    StatusBar1.Panels(3).text = ArbString(CardTable!USERNAME2 & " [" & myFormat_p(CardTable!TIME2, True) & "]")
End If

Handlecontrols LoadMode
myLoadGrd

If Not bRefresh Then
    CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
End If
End Sub
Private Sub myLoadGrd()
Dim cString As String
With grid1
    cString = "SELECT " & _
              "ACCOUNT.DESCA," & _
              "ACCOUNT_D.ACCOUNT," & _
              "ACCOUNT_D.CODE," & _
              "ACCOUNT_CODES.DESCA," & _
              "NULL," & _
              " ACCOUNT_D.DESCA," & _
              " CASE WHEN VALUE = 0 THEN NULL ELSE VALUE END," & _
              " CASE WHEN VALUE2 = 0 THEN NULL ELSE VALUE2 END," & _
              " BRANCH," & _
              " MOSM," & _
              "ACCOUNT.WITH_BRANCH," & _
              "ACCOUNT.WITH_MOSM," & _
              " ACCOUNT_D.ID " & _
              " FROM ACCOUNT_D " & _
              " INNER JOIN ACCOUNT_CODES ON ACCOUNT_D.CODE =  ACCOUNT_CODES.CODE" & _
              " INNER JOIN ACCOUNT ON ACCOUNT_D.ACCOUNT = ACCOUNT.CODE" & _
              " WHERE DOC_NO = " & MyParn(xDoc_No.text)
    Set data11.Recordset = myRecordSet(cString, con)
    myAddItem
End With
CalcTotals
fixGrd
End Sub
Private Sub myDefine()
bChangeBal = True
xDoc_No.text = ""
bChangeBal = False

XFLAG.Caption = ""
xDate.text = myFormat_p(Date)
xBox.BoundText = cBranchBox
'xBox.Tag = xBox.BoundText

ChangeBalance

bLoadDate = False
xbranch.Caption = cBranch
bigCheck = True
xclosed.Value = 0
bCheck = False
fixGrd
grid1.Rows = 1
StatusBar1.Panels(2) = ""
StatusBar1.Panels(3) = ""
StatusBar1.Panels(4) = ""
myAddItem
Handlecontrols DefineMode
CalcTotals
On Error Resume Next
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
If xBox.MatchedWithList Then
    grid1.SetFocus
Else
    xBox.SetFocus
End If
Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit And xclosed.Value = 0
xBox.Enabled = bEditRecord And nMode = DefineMode
xBox.Enabled = xBox.Enabled And nUser = enUser.Admin And cBranchBox = ""

If nUser < enUser.Admin Then
    xDone.Enabled = nMode = LoadMode And cBranchBox <> "" And xBox.BoundText = cBranchBox And xDone.Value = 0
ElseIf nUser = enUser.Admin And xbranch.Caption = "00" Then
    xDone.Enabled = nMode = LoadMode And xclosed.Value = 0
Else
    xDone.Enabled = xDone.Value = 1 And xclosed.Value = 0
End If

xclosed.Enabled = nUser = enUser.Admin And nMode = LoadMode And xDone.Value = 1

cmdClosePeriod.Enabled = retFlag(aSec, "manager") And nMode = LoadMode
cmdOpenPeriod.Enabled = retFlag(aSec, "manager") And nMode = LoadMode

cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmdSave.Enabled = bEditRecord
CmdDelInv.Enabled = nMode = LoadMode And bEditRecord

Dim nRecord As Long, nRecords As Long
retRecords XFLAG.Caption, nRecords, nRecord
If nMode = LoadMode Then
    StatusBar1.Panels(1).text = ArbString("”Ã· " & nRecord & " „‰ " & nRecords)
Else
    StatusBar1.Panels(1).text = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""

'xDoc_No.Enabled = (nMode = DefineMode)
Me.Tag = nMode
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
ElseIf Not bEditRecord Then
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 Then
    If MsgBox("„‰ «·„” ‰œ ?", vbOKCancel + vbDefaultButton2) = vbOK Then
        con.BeginTrans
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from ACCOUNT_D where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        myRemove grid1.Row
        grid1_EnterCell
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If col = 0 And grid1.TextMatrix(Row, 1) = "" Then Exit Sub
    If col = 2 + 1 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    If col = 7 + 1 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    If col = 8 + 1 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid1.col = 0 And grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
    If grid1.col = 2 + 1 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    If grid1.col = 7 + 1 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    If grid1.col = 8 + 1 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 0 Then
'    If grid1.EditText = "" Then
'        Cancel = True
'    Else
'        'if rsField(rsaccount,grid1.edittext,"code","desca")
'    End If
        
ElseIf col = 5 + 1 Then
    If Val(grid1.EditText) <> 0 And grid1.ValueMatrix(Row, 6 + 1) <> 0 Then
        grid1.TextMatrix(Row, 6 + 1) = ""
    End If
ElseIf col = 6 + 1 Then
    If Val(grid1.EditText) <> 0 And grid1.ValueMatrix(Row, 5 + 1) <> 0 Then
        grid1.TextMatrix(Row, 5 + 1) = ""
    End If
End If
End Sub
Private Sub cmdOpenPeriod_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sCaption = Me.Caption
oClosefrm.nMode = 1
oClosefrm.Show 1
myUndo
End Sub
Private Sub cmdClosePeriod_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
'oClosefrm.bNoBranch = True
oClosefrm.sCaption = Me.Caption
oClosefrm.nMode = 0
oClosefrm.Show 1
myUndo
End Sub

Private Sub xBox_GotFocus()
Dim cString As String
cString = "SELECT * FROM FILE0_50"
If cBranch <> "" Then
    cString = cString & " WHERE BRANCH = " & cBranch
    If xBox.Tag <> "" Then cString = cString & " OR CODE = " & MyParn(xBox.Tag)
End If
cString = cString & " ORDER BY DESCA"
If DATA1.RecordSource <> cString Then
    Dim sBound As String
    sBound = xBox.BoundText
    Set DATA1.Recordset = mycmd(cString, con)

    xBox.BoundText = sBound
    If Not xBox.MatchedWithList Then xBox.BoundText = ""
End If
End Sub
Private Sub xBox_LostFocus()
Dim cString As String
cString = "SELECT * FROM FILE0_50"
If DATA1.Recordset.Source <> cString Then
    Dim sBound As String
    sBound = xBox.BoundText
    Set DATA1.Recordset = mycmd(cString, con)
    xBox.BoundText = sBound
    If Not xBox.MatchedWithList Then xBox.BoundText = ""
End If
End Sub
Private Sub xClosed_Click()
If bigCheck Then Exit Sub
On Error GoTo myerror
con.Execute " update " & cFileHeader & " set CLOSED = " & xclosed.Value & " WHERE doc_no = " & MyParn(xDoc_No.text)
myUndo
Inform " „ " & IIf(xclosed.Value = 1, "«€·«ﬁ", "› Õ") & " «·„” ‰œ »‰Ã«Õ"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function CalcTotals()
Dim nPlus As Double, nMinus As Double
With grid1
For i = 1 To grid1.Rows - 2
    nPlus = mRound(nPlus + grid1.ValueMatrix(i, 5 + 1))
    nMinus = mRound(nMinus + grid1.ValueMatrix(i, 6 + 1))
Next
xPlus.Caption = nPlus
xMinus.Caption = nMinus
xBalance.Caption = nPlus - nMinus
xBalanceLast.Caption = Val(xBalanceBefore.Caption) + (nPlus - nMinus)
StatusBar1.Panels(4).text = "≈Ã„«·Ì «Ìœ«⁄ : " & Myvalue(nTotal1)
End With
End Function
Private Sub fixGrd()
With grid1
.MergeCells = flexMergeFree
.MergeRow(0) = True
.FormatString = "‰Ê⁄ «·Õ—ﬂ…|" & "‰Ê⁄ «·Õ—ﬂ…|" & "«·ﬂÊœ|" & "«·«”„|" & "«· «—ÌŒ|" & "«·»Ì«‰|" & "«Ìœ«⁄« |" & "„”ÕÊ»« |" & "«·›—⁄|" & "«·„Ê”„|" & "»›—⁄|" & "»„Ê”„|"
.ColWidth(0) = 2000
.ColWidth(1 + 1) = 1200
.ColWidth(2 + 1) = 3500
.ColWidth(3 + 1) = 1400
.ColWidth(4 + 1) = 5500
.ColWidth(5 + 1) = 1400
.ColWidth(6 + 1) = 1400
.ColWidth(7 + 1) = 1600
.ColWidth(8 + 1) = 1800
.ColComboList(7 + 1) = cList2
.ColComboList(8 + 1) = cList3

.ColHidden(1) = True
.ColHidden(3 + 1) = True
.ColHidden(1 + 1) = True


.ColHidden(.Cols - 3) = True
.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
'.ColComboList(0) = CLIST
.ColComboList(0) = "..."
.ColComboList(2 + 1) = "..."
End With
End Sub
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_no As String = "", Optional bRefresh As Boolean = False) As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_no) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

'On Error GoTo myerror
cFilter = ""

'If (Not retFlag(aSec, "manager")) And (Not retFlag(aSec, "com")) Then
'    cFilter = " AND " & cFileHeader & ".USERCODE = " & nusercode
'ElseIf retFlag(aSec, "com") Then
'    cFilter = " AND " & cFileHeader & ".CLOSED  = 0"
'End If


If cBranch <> "" And nUser < enUser.Admin Then
    cFilter = cFilter & " AND " & "ACCOUNT_H.BRANCH = " & MyParn(cBranch)
End If

If nUser <= enUser.Casher Then
    cFilter = cFilter & " AND ACCOUNT_H.BOX = " & MyParn(cBranchBox)
End If

'If validYear(xYear.text) Then
'    cFilter = cFilter & " AND " & "YEAR(" & cFileHeader & ".[DATE]) = " & Val(xYear.text)
'End If

If chkClosed.Value = 1 Then
    cFilter = cFilter & " AND " & cFileHeader & ".CLOSED  = 0"
End If

cFilter = Mid(cFilter, 6)

If sDoc_no <> "" Then
    cFilter = "DOC_NO = " & MyParn(sDoc_no)
End If

cString = "SELECT TOP 1 * FROM " & cFileHeader

If pMode = tbMode.tbFirst Then
    cOrder = "Order by FLAG"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by FLAG DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "Doc_no = " & MyParn(pDoc_no)
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "FLAG < " & MyParn(pDoc_no)
    cOrder = "Order by FLAG desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "FLAG > " & MyParn(pDoc_no)
    cOrder = "Order by FLAG"
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
    myload bRefresh
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
If Me.Tag = DefineMode Then
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
Private Sub myAddItem()
With grid1
.AddItem ""
'If cdef_Box <> "" Then .TextMatrix(.Rows - 1, 3 + 1) = cdef_Box
If grid1.Rows > 2 Then
'    .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
'    .TextMatrix(.Rows - 1, 3) = .TextMatrix(.Rows - 2, 3)
End If
End With
End Sub
Private Function validRow(Row) As Boolean
With grid1
If Not ValidInt(.TextMatrix(Row, 0 + 1)) Then Exit Function
If Trim(.TextMatrix(Row, 1 + 1)) = "" Then Exit Function
If .ValueMatrix(Row, 5 + 1) = 0 And .ValueMatrix(Row, 6 + 1) = 0 Then Exit Function
If .ValueMatrix(Row, 5 + 1) <> 0 And .ValueMatrix(Row, 6 + 1) <> 0 Then Exit Function
'If Not MYVALID(True) Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
Dim nColEnd As Long

If grid1.ValueMatrix(Row, 10 + 1) <> 0 And grid1.TextMatrix(Row, 8 + 1) = "" Then
    nColEnd = grid1.Cols - 4
ElseIf grid1.ValueMatrix(Row, 9 + 1) <> 0 And grid1.TextMatrix(Row, 7 + 1) = "" Then
    nColEnd = grid1.Cols - 5
Else
    nColEnd = grid1.Cols - 6 - IIf(grid1.ValueMatrix(Row, 5 + 1) <> 0, 1, 0)
End If

If col < nColEnd Then
    If col = 5 + 1 And grid1.ValueMatrix(Row, 5 + 1) = 0 Then
        grid1.col = col + 1
    ElseIf col >= 5 + 1 Then
        grid1.col = nColEnd
    ElseIf col = 4 + 1 Then
        If (grid1.TextMatrix(Row, 0 + 1) = 2 Or grid1.TextMatrix(Row, 0 + 1) = 6 Or grid1.TextMatrix(Row, 0 + 1) = 7) And grid1.ValueMatrix(Row, 5 + 1) = 0 Then
            grid1.col = 6 + 1
        Else
            grid1.col = 5 + 1
        End If
    Else
        grid1.col = col + 1 + IIf(col = 0 Or col = 2 + 1, 1, 0) + IIf(col = 0, 1, 0)
    End If
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 4 + 1)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, col
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
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
CalcTotals
End Sub
Private Sub xYear_Validate(Cancel As Boolean)
If Me.Tag = DefineMode Then Exit Sub

If Not validYear(xYear.text) Then xYear.text = ""
If validYear(xYear.text) Or xYear.text = "" Then
    openCardTable
    myUndo
End If
End Sub
Private Function retRecords(pCode, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New adodb.Recordset
If pCode <> "" Then
    cString = "SELECT Count(*) AS records,COUNT(CASE WHEN FLAG <= " & MyParn(pCode) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM " & cFileHeader
If cFilter <> "" Then cString = cString & " WHERE " & cFilter

Set loctable = mycmd(cString, con)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Sub ChangeBalance(Optional bCheck As Boolean = False)
    If IsDate(xDate.text) And xBox.MatchedWithList And xDoc_No.text <> "" And chkBalance.Value = 1 Then
        If (Not IsEmpty(aPrm)) And bCheck = False Then
            If retFlag(aPrm, "BOX") = xBox.BoundText And _
               retFlag(aPrm, "DATE") = myFormat_sp(xDate.text) And _
               retFlag(aPrm, "DOC_NO") = xDoc_No.text Then
               Exit Sub
            End If
        End If
    
        aPrm = AddFlag(Empty, "BOX", xBox.BoundText)
        aPrm = AddFlag(aPrm, "date", myFormat_sp(xDate.text))
        aPrm = AddFlag(aPrm, "doc_no", xDoc_No.text)
        Set cm = cmd("sp_box_bal", con, adStoredProc, aPrm)
        cm.Execute
        xBalanceBefore.Caption = mRound(mRound(cm.Parameters("@PLUS").Value) - mRound(cm.Parameters("@MINUS").Value))
        CalcTotals
    Else
        xBalanceBefore.Caption = 0
        CalcTotals
    End If
End Sub
Private Function NewflagBox()
NewflagBox = Format(xDate.text, "yymmdd") & "-" & xBox.BoundText
End Function
Private Function FoundDay() As Boolean
If Me.Tag = LoadMode Then Exit Function
If Not IsDate(xDate.text) Then Exit Function
If Not xBox.MatchedWithList Then Exit Function
If myField("Select doc_no from ACCOUNT_H WHERE DOC_NO = " & MyParn(NewflagBox), con) = "" Then Exit Function

MsgBox "„” ‰œ »‰›” «·Œ“‰… ·‰›” «·ÌÊ„"
openCardTable tbMode.tbFind, NewflagBox
FoundDay = True
End Function
Private Sub xDone_Click()
If bigCheck Then Exit Sub
On Error GoTo myerror
con.Execute " update " & cFileHeader & " set DONE = " & xDone.Value & " WHERE doc_no = " & MyParn(xDoc_No.text)
myUndo
Inform " „ " & IIf(xclosed.Value = 1, "«·€«¡ «‰ Â«¡", "«·«‰ Â«¡ ") & " «·„” ‰œ »‰Ã«Õ"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
