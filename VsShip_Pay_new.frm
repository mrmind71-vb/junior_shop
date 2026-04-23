VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsShip_Pay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ «»⁄… »Ê‰«  «·»Ì⁄  ”·Ì„ «Ê‰ ·«Ì‰"
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   450
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   20400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   555
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   1125
      Width           =   3750
      Begin VB.OptionButton optShip 
         Alignment       =   1  'Right Justify
         Caption         =   "» «—ÌŒ ‘Õ‰"
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
         Index           =   1
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   135
         Width           =   1230
      End
      Begin VB.OptionButton optShip 
         Alignment       =   1  'Right Justify
         Caption         =   "»œÊ‰ ‘Õ‰"
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
         Index           =   2
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   135
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton optShip 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   0
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   135
         Width           =   690
      End
   End
   Begin VB.Frame Frame5 
      Height          =   555
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   45
      Width           =   3750
      Begin VB.OptionButton XPAY 
         Alignment       =   1  'Right Justify
         Caption         =   "€Ì— „Õ’·"
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
         Index           =   0
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   135
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton XPAY 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   2
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   645
      End
      Begin VB.OptionButton XPAY 
         Alignment       =   1  'Right Justify
         Caption         =   "„Õ’·"
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
         Index           =   1
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   135
         Width           =   870
      End
   End
   Begin VB.Frame Frame4 
      Height          =   555
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   585
      Width           =   3750
      Begin VB.OptionButton XONLINE 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   2
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   135
         Width           =   690
      End
      Begin VB.OptionButton XONLINE 
         Alignment       =   1  'Right Justify
         Caption         =   "‘Õ‰"
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
         Index           =   1
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   135
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton XONLINE 
         Alignment       =   1  'Right Justify
         Caption         =   "‰ﬁœÏ"
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
         Index           =   0
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   135
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   45
      Width           =   12075
      Begin VB.TextBox XPHONE 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   2220
      End
      Begin VB.TextBox XONLINE_DOC 
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
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   630
         Width           =   1635
      End
      Begin VB.TextBox XSHIP_NO 
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
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   3735
         Picture         =   "VsShip_Pay_new.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "⁄—÷"
         Top             =   1080
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   90
         Picture         =   "VsShip_Pay_new.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1080
         Width           =   1185
      End
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   2520
         Picture         =   "VsShip_Pay_new.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1080
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   7110
         TabIndex        =   0
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xship 
         Height          =   315
         Left            =   7110
         TabIndex        =   1
         Top             =   630
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   1305
         TabIndex        =   39
         Top             =   1080
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   820
         _Version        =   196610
         ForeColor       =   0
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
         Picture         =   "VsShip_Pay_new.frx":6D88
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «· ·Ì›Ê‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·ÿ·»Ì… "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   5625
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·»Ê»Ì’… "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "‘—ﬂ… «·‘Õ‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   10665
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "«·›—⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   10665
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   225
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   20
      Top             =   10710
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DATA1 
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
   Begin MSAdodcLib.Adodc DATA2 
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
   Begin VB.Frame Frame2 
      Height          =   1635
      Left            =   4005
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   45
      Width           =   4065
      Begin VB.TextBox xdate_ship1 
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
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1170
         Width           =   1275
      End
      Begin VB.TextBox xdate_ship2 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1170
         Width           =   1320
      End
      Begin VB.TextBox xDatePAY2 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   675
         Width           =   1320
      End
      Begin VB.TextBox xDatePAY1 
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
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   675
         Width           =   1275
      End
      Begin VB.TextBox xDate2 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   1320
      End
      Begin VB.TextBox xDate1 
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
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ ‘Õ‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ  Õ’Ì·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ „»Ì⁄«  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   270
         Width           =   1095
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   210
      Left            =   0
      TabIndex        =   37
      Top             =   10500
      Visible         =   0   'False
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7800
      Left            =   180
      TabIndex        =   38
      Top             =   1710
      Width           =   19995
      _cx             =   35269
      _cy             =   13758
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Cols            =   18
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
Attribute VB_Name = "VsShip_Pay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    cHead1 = Me.Caption & "  " & xStore.text
    grid1.ColHidden(15 + 1) = True
    
    grid1.ColHidden(0) = (xStore.BoundText <> "")
    
    Set printGrd3.grid1 = grid1
    Set printGrd3.myform = Me
    If xStore.BoundText = "" Then
        printGrd3.doprint 1, 0, cHead1, , , False, True, 10, "Arial", 0
    Else
        printGrd3.doprint 0.85, 0, cHead1, , , False, True, 10, "Arial"
    End If
    printGrd3.Show 1
    grid1.ColHidden(15 + 1) = False
    grid1.ColHidden(0) = False
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    
    Set data2.Recordset = myRecordSet("Select Code,DescA From STORE_BR ORDER BY CODE ", con)
    Set xStore.RowSource = data2
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    If cBranch <> "00" Then xStore.BoundText = cBranchStore
    xStore.Enabled = (cBranch = "00")
    
    Set DATA3.Recordset = myRecordSet("Select Code,DescA From ship ORDER BY CODE ", con)
    Set xship.RowSource = DATA3
    xship.ListField = "Desca"
    xship.BoundColumn = "Code"
    
    Set grid1.DataSource = data1
    grid1.Rows = 1
    fixGrd
End Sub
Private Sub myload()
    Dim cWhere As String
    Dim cString As New ChilkatStringBuilder
    
   '                        0       1       2       3       4       5       6       7                       8       9       10      11 + 1      12          13 + 1          14          15 + 1          16 + 1
'    cString = " SELECT  STORENAME, date, DOC_NO2, NAME, phone, CHARGE1, CHARGE2, TOTAL," & _
'    " (TOTAL+CHARGE1+CHARGE2 ), DESCA, SHIP_NO, BOXNAME , BANKNAME , DATE_PAY , DOC_NO , ' ', note_ret , ONLINE_DOC FROM Q_SHIP WHERE DOC_NO IS NOT NULL AND ONLINE > 0 "
    
    cString.Append "SELECT FILE0_40.DESCA," & _
              "CONVERT(VARCHAR(10),FILE6_20H.DATE,111)," & _
              "FILE6_20H.DOC_NO2," & _
              "FILE6_20H.NAME," & _
              "FILE6_20H.PHONE," & _
              "FILE6_20H.CHARGE1," & _
              "FILE6_20H.CHARGE2," & _
              "(FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT)," & _
              "(FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT) + FILE6_20H.CHARGE1 + FILE6_20H.CHARGE2," & _
              "SHIP.DESCA," & _
              "FILE6_20H.SHIP_NO," & _
              "FORMAT(FILE6_20H.DATE_SHIP ,'yyyy/M/d')," & _
              "FILE0_50.DESCA," & _
              "FILE5_10.DESCA," & _
              "CONVERT(VARCHAR(10),FILE6_20H.DATE_PAY,111)," & _
              "FILE6_20H.DOC_NO," & _
              "FILE6_20H.NOTE_RET," & _
              "FILE6_20H.ONLINE_DOC" & _
              " From FILE6_20H" & _
              " INNER JOIN FILE0_50 ON FILE6_20H.BOX = FILE0_50.CODE" & _
              " INNER JOIN FILE0_40 ON FILE6_20H.store = FILE0_40.CODE" & _
              " LEFT OUTER JOIN FILE5_10 ON FILE6_20H.BANK = FILE5_10.CODE" & _
              " LEFT OUTER JOIN SHIP  ON FILE6_20H.SHIP = SHIP.CODE"
              
    cString.Append " WHERE (FILE6_20H.[ONLINE] <> 0)"
    
    If xStore.MatchedWithList Then
        cString.Append " AND FILE6_20H.store = " & MyParn(xStore.BoundText)
    End If
    
    If xship.MatchedWithList Then
        cString.Append " AND FILE6_20H.ship = " & MyParn(xship.BoundText)
    End If
    
    If xship_no.text <> "" Then
        cString.Append " AND ship_NO = " & MyParn(xship_no.text)
    End If
    
    If xPhone.text <> "" Then
        cString.Append " AND phone = " & MyParn(xPhone.text)
    End If
    
    If xonline_doc.text <> "" Then
        cString.Append " AND ONLINE_DOC = " & MyParn(xonline_doc.text)
    End If
    
    If IsDate(xdate1.text) Then
        cString.Append " AND DATE >= " & DateSq(xdate1.text)
    End If
    
    If IsDate(xdate2.text) Then
        cString.Append " AND DATE <= " & DateSq(xdate2.text)
    End If
    
    If IsDate(xDatePAY1.text) Then
        cString.Append " AND DATE_PAY >= " & DateSq(xDatePAY1.text)
    End If
    
    If IsDate(xDatePAY2.text) Then
        cString.Append " AND DATE_PAY <= " & DateSq(xDatePAY2.text)
    End If
    
    If xPay(1).Value Then
        cString.Append " AND ((DATE_PAY IS NOT NULL) OR FILE6_20H.ONLINE = 1)"
    End If
    
    If xPay(0).Value Then
        cString.Append " AND (DATE_PAY IS NULL AND FILE6_20H.ONLINE =  2)"
    End If
    
    If XONLINE(0).Value Then
        cString.Append " AND FILE6_20H.[ONLINE] = 1 "
    End If
    
    If XONLINE(1).Value Then
        cString.Append " AND FILE6_20H.[ONLINE] = 2 "
    End If
    
    If IsDate(xdate1.text) Then
        cString.Append " AND DATE >= " & DateSq(xdate1.text)
    End If
    
    If IsDate(xdate2.text) Then
        cString.Append " AND DATE <= " & DateSq(xdate2.text)
    End If
    
    If IsDate(xdate_ship1.text) Then
        cString.Append " AND DATE_SHIP >= " & DateSq(xdate_ship1.text)
    End If
    
    If IsDate(xdate_ship2.text) Then
        cString.Append " AND DATE_SHIP <= " & DateSq(xdate_ship2.text)
    End If
    
    If optShip(1).Value Then
        cString.Append " AND DATE_SHIP IS NOT NULL"
    ElseIf optShip(2).Value Then
        cString.Append " AND DATE_SHIP IS NULL"
    End If
    
    cString.Append " order by FILE6_20H.STORE , FILE6_20H.DATE , FILE6_20H.doc_no2 "
    Set data1.Recordset = mycmd(cString.GetAsString, con)
    fixGrd
End Sub
Sub fixGrd()
    With grid1
    .ExplorerBar = flexExSortShow
    .WordWrap = True
    
    .RowHeight(0) = 1000
    .TextMatrix(0, 0) = "«·›—⁄"
    .TextMatrix(0, 1) = " «—ÌŒ"
    .TextMatrix(0, 2) = "—ﬁ„ "
    .TextMatrix(0, 3) = "«·⁄„Ì·"
    .TextMatrix(0, 4) = "«· ·Ì›Ê‰"
    .TextMatrix(0, 5) = "„’«—Ì› ‘Õ‰"
    .TextMatrix(0, 6) = "„’«—Ì› »Ê·Ì’…"
    .TextMatrix(0, 7) = " ﬁÌ„… «·›« Ê—…"
    .TextMatrix(0, 8) = " «·≈Ã„«·Ï"
    .TextMatrix(0, 9) = "‘—ﬂ… «·‘Õ‰"
    .TextMatrix(0, 10) = "—ﬁ„ »Ê·Ì’…"
    .TextMatrix(0, 11) = " «—ÌŒ «·‘Õ‰"
    .TextMatrix(0, 12) = "«·ﬂ«‘Ì—"
    
    .TextMatrix(0, 12 + 1) = "»‰ﬂ «·≈Ìœ«⁄"
    .TextMatrix(0, 13 + 1) = " «—ÌŒ  Õ’Ì·"
    
    .TextMatrix(0, 14 + 1) = "—ﬁ„ «·»Ê‰"
    
    .TextMatrix(0, 15 + 1) = "—›÷ «·»Ê‰"
    .TextMatrix(0, 16 + 1) = "—ﬁ„ «·ÿ·»Ì…"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 1300
    .ColWidth(2) = 1200
    .ColWidth(3) = 1800
    .ColWidth(4) = 1600
    .ColWidth(5) = 800
    .ColWidth(6) = 800
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 2000
    .ColWidth(11) = 1400
    .ColWidth(11 + 1) = 2000
    .ColWidth(12 + 1) = 1000
    .ColWidth(13 + 1) = 1300
    .ColWidth(14 + 1) = 1700
    .ColWidth(15 + 1) = 1500
    .ColWidth(16 + 1) = 2000
    
    .ColHidden(12 + 1) = cBranch <> "00"
    .ColHidden(13 + 1) = cBranch <> "00"
    
     
    .ColComboList(14 + 1) = "..."
    .ColDataType(0) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(1) = flexDTDate
    .ColDataType(13 + 1) = flexDTDate
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(8) = flexDTDouble
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 5, "#0", , , True, "«·«Ã„«·Ì "
    .Subtotal flexSTSum, -1, 6, "#0", , , True, "«·«Ã„«·Ì "
    .Subtotal flexSTSum, -1, 7, "#0", , , True, "«·«Ã„«·Ì "
    .Subtotal flexSTSum, -1, 8, "#0", , , True, "«·«Ã„«·Ì "
    .Subtotal flexSTCount, -1, 2, "#0", , , True, "«·«Ã„«·Ì "
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 11 Then
    con.Execute "UPDATE FILE6_20H SET " & _
                " FILE6_20H.DATE_SHIP = " & addDate(grid1.TextMatrix(Row, Col)) & _
                " WHERE DOC_NO = " & addstring(grid1.TextMatrix(Row, 14 + 1)), nAffect
End If
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With grid1
    If .Col = 14 + 1 Then
        Dim oSalesOnline As New sales_online_popfrm
        oSalesOnline.sDoc_no = .TextMatrix(.Row, 14 + 1)
        oSalesOnline.bUserAdmin = bopt2
        oSalesOnline.bedit = True
        oSalesOnline.Show 1
    End If
End With
myload
End Sub
Private Sub grid1_EnterCell()
With grid1
If .Col = 11 Or .Col = 14 + 1 Then
    .Editable = flexEDKbdMouse
Else
    .Editable = flexEDNone
End If
End With
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'If grid1.Col = 15 + 1 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 15 + 1 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < 11 Then
    grid1.Col = 11
'ElseIf Col < 14 + 1 Then
'    Grid1.Col = 14 + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 11)
Else
    grid1.Select Row, Col
End If
grid1.ShowCell grid1.Row, grid1.Col
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 11 Then
    If (Not IsDate(grid1.EditText)) And Trim(grid1.EditText) <> "" Then
        Cancel = True
    Else
        grid1.EditText = myFormat_p(grid1.EditText)
    End If
End If
End Sub
Private Sub xPhone_GotFocus()
myGotFocus xPhone
End Sub
Private Sub XPHONE_LostFocus()
myLostFocus xPhone
End Sub
Private Sub XONLINE_DOC_GotFocus()
myGotFocus xonline_doc
End Sub
Private Sub XONLINE_DOC_LostFocus()
myLostFocus xonline_doc
End Sub
Private Sub xship_no_GotFocus()
myGotFocus xship_no
End Sub
Private Sub xship_no_LostFocus()
myLostFocus xship_no
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xship_GotFocus()
myGotFocus xship
End Sub
Private Sub xship_LostFocus()
myLostFocus xship
If Not xship.MatchedWithList Then xship.BoundText = ""
End Sub
Private Sub xdate_ship1_GotFocus()
myGotFocus xdate_ship1
End Sub
Private Sub xdate_ship1_LostFocus()
myLostFocus xdate_ship1
myValidDate xdate_ship1
End Sub
Private Sub xdate_ship2_GotFocus()
myGotFocus xdate_ship2
End Sub
Private Sub xdate_ship2_LostFocus()
myLostFocus xdate_ship2
myValidDate xdate_ship2
End Sub
Private Sub xDatePAY2_GotFocus()
myGotFocus xDatePAY2
End Sub
Private Sub xDatePAY2_LostFocus()
myLostFocus xDatePAY2
myValidDate xDatePAY2
End Sub
Private Sub xDatePAY1_GotFocus()
myGotFocus xDatePAY1
End Sub
Private Sub xDatePAY1_LostFocus()
myLostFocus xDatePAY1
myValidDate xDatePAY1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
