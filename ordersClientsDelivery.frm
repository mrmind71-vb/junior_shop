VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form ordersClientsDelivery 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·»Ì«  ⁄„·«¡"
   ClientHeight    =   10095
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
   ScaleHeight     =   10095
   ScaleWidth      =   20400
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   14175
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   0
      Width           =   2085
      Begin Threed.SSCommand cmdAmzon 
         Height          =   510
         Left            =   45
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   180
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   900
         _Version        =   196610
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
         Caption         =   "Excel «” Ì—«œ „‰"
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
      Height          =   1185
      Index           =   0
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   855
      Width           =   3120
      Begin VB.CheckBox xisClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«€·«ﬁ «·„” ‰œ"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   1365
      End
      Begin Threed.SSCommand cmdClosePeriod 
         Height          =   510
         Index           =   0
         Left            =   1575
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   585
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«€·«ﬁ › —…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdClosePeriod 
         Height          =   510
         Index           =   1
         Left            =   90
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   585
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "› Õ › —…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   11205
      TabIndex        =   27
      Top             =   720
      Width           =   9060
      Begin VB.CheckBox xIsPosted 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«€·«ﬁ „·› «·«” ·«„"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   945
         Width           =   1815
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6660
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   540
         Width           =   1365
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   330
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   540
         Width           =   2175
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3240
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   900
         Width           =   4785
      End
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5355
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   28
         Top             =   2460
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox xDoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6660
         MaxLength       =   20
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·ÿ·»Ì…"
         Height          =   270
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   180
         Width           =   510
      End
      Begin VB.Label xPO_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         TabIndex        =   41
         Top             =   180
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "«·»Ì«‰"
         Height          =   270
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   990
         Width           =   420
      End
      Begin VB.Label xCode_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3240
         TabIndex        =   39
         Top             =   540
         Width           =   3390
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "«·⁄„Ì·"
         Height          =   270
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   585
         Width           =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "⁄„·«¡ ﬂ—Ê  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2430
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "—ﬁ„ „” ‰œ"
         Height          =   285
         Index           =   0
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   225
         Width           =   885
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
      Left            =   16245
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   45
      Width           =   4020
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   2700
         TabIndex        =   17
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
         Picture         =   "ordersClientsDelivery.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ordersClientsDelivery.frx":23CB
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   4635
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
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
         Picture         =   "ordersClientsDelivery.frx":4474
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ordersClientsDelivery.frx":647C
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1350
         TabIndex        =   19
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
         Picture         =   "ordersClientsDelivery.frx":8433
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ordersClientsDelivery.frx":ABCF
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   20
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
         Picture         =   "ordersClientsDelivery.frx":D063
         Alignment       =   8
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
      Left            =   9630
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   900
      Width           =   1545
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   45
         TabIndex        =   14
         Top             =   135
         Width           =   1455
         _ExtentX        =   2566
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
         Picture         =   "ordersClientsDelivery.frx":F386
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ordersClientsDelivery.frx":11CAB
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   465
         Left            =   45
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   630
         Width           =   1455
         _ExtentX        =   2566
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
         Picture         =   "ordersClientsDelivery.frx":144FF
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ordersClientsDelivery.frx":1665F
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   465
      Left            =   1755
      Top             =   405
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
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
   Begin Threed.SSCommand cmdFilter 
      Height          =   420
      Left            =   14535
      TabIndex        =   8
      Top             =   8685
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   741
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Undo Filter"
      ButtonStyle     =   3
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   9720
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   3
         Left            =   11520
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
   Begin Threed.SSCommand cmdClient 
      Height          =   420
      Left            =   16650
      TabIndex        =   26
      Top             =   8685
      Width           =   3660
      _ExtentX        =   6456
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
      Caption         =   "«Œ «— «·⁄„Ì·"
      TagVariant      =   "«Œ «— «·„Ê—œ"
      ButtonStyle     =   3
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   38
      Top             =   9615
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   465
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
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
      Height          =   6540
      Left            =   45
      TabIndex        =   7
      Top             =   2070
      Width           =   20220
      _cx             =   35666
      _cy             =   11536
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
      Cols            =   14
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
      RightToLeft     =   0   'False
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame3 
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
      TabIndex        =   21
      Top             =   8595
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   22
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
         Picture         =   "ordersClientsDelivery.frx":1894C
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ordersClientsDelivery.frx":1AAF3
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   23
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
         Picture         =   "ordersClientsDelivery.frx":1CB3A
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ordersClientsDelivery.frx":1EC25
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   24
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
         Picture         =   "ordersClientsDelivery.frx":20C1F
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ordersClientsDelivery.frx":22D30
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   25
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
         Picture         =   "ordersClientsDelivery.frx":24D2A
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ordersClientsDelivery.frx":26F4E
      End
   End
   Begin VB.Label XFLAG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3780
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   2625
   End
End
Attribute VB_Name = "ordersClientsDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nType As Byte, bedit As Boolean, sDoc_no As String
Public bIg As Boolean
Dim bEditRecord As Boolean, bAct As Boolean
Dim oSeachCode As New Search_abd, oSearchDoc As New Search_abd, oSearchProject As New Search_abd, osearchClient As New Search_abd
Dim cFilter As String, cSelect As String
Dim con As New adodb.Connection, CardTable As adodb.Recordset
Dim cFile As String, cFileHeader As String, aBoxs As Variant
Dim cList1 As String, cList2 As String, cList3 As String, sDef As String
Public sDoc_no_add As String
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[CODE]", addstring(xCode.text))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.text))
aInsert = AddFlag(aInsert, IIf(xdoc_no.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cusername))
aInsert = AddFlag(aInsert, IIf(xdoc_no.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")
con.BeginTrans
On Error GoTo myError
If xdoc_no.Tag = DefineMode Then
    xdoc_no.text = Newflag(cFileHeader, "DOC_NO", con)
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xdoc_no.text))
    con.Execute addInsert(aInsert, cFileHeader)
Else
    con.Execute addUpdate(aInsert, cFileHeader, "doc_no = " & addstring(xdoc_no.text))
End If
'myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc(Optional sControl As String = "")
If ActiveControl.Name = grid1.Name Then
    grid1.TextMatrix(grid1.Row, 2) = oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0)
    Grid1_AfterEdit grid1.Row, grid1.Col
    oSeachCode.Hide
    CellPos 13, grid1.Row, grid1.Col
ElseIf ActiveControl.Name = CmdInform.Name Then
    openCardTable tbMode.tbFind, oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
ElseIf ActiveControl.Name = cmdClient.Name Then
    If osearchClient.grid1.TextMatrix(osearchClient.grid1.Row, 0) = "" Then
        cmdClient.Tag = ""
        cmdClient.Caption = cmdClient.TagVariant
    Else
        cmdClient.Tag = osearchClient.grid1.TextMatrix(osearchClient.grid1.Row, 0)
        cmdClient.Caption = osearchClient.grid1.TextMatrix(osearchClient.grid1.Row, 1)
    End If
    If Not openCardTable(tbMode.tbFind, xdoc_no.text) Then
        If Not openCardTable Then myDefine
    End If
    Unload osearchClient
ElseIf ActiveControl.Name = xCode.Name Then
    xCode.text = osearchClient.grid1.TextMatrix(osearchClient.grid1.Row, 0)
    xCode_desca.Caption = osearchClient.grid1.TextMatrix(osearchClient.grid1.Row, 1)
    Unload osearchClient
ElseIf ActiveControl.Name = cmdProject.Name Then
    If oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 0) = "" Then
        cmdProject.Tag = ""
        cmdProject.Caption = cmdProject.TagVariant
    Else
        cmdProject.Tag = oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 0)
        cmdProject.Caption = oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 1)
    End If
    If Not openCardTable(tbMode.tbFind, xdoc_no.text) Then
        If Not openCardTable Then myDefine
    End If
    oSearchProject.Hide
End If
End Sub

Private Sub cmdAmzon_Click()
cmdAmzon.Enabled = False
Set OrdersClientsCsv.myform = Me
OrdersClientsCsv.Show 1
cmdAmzon.Enabled = True
If sDoc_no_add <> "" Then
    openCardTable tbMode.tbFind, sDoc_no_add
End If
End Sub

Private Sub cmdClosePeriod_Click(Index As Integer)
Dim oClosefrm As New closefrm
oClosefrm.sFile = "FILE6_52h"
oClosefrm.sCaption = DocTitle
oClosefrm.sFieldClose = "isclosed"
oClosefrm.nMode = Index
oClosefrm.Show 1
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› ?", vbOKCancel + vbDefaultButton2 + vbCritical) <> vbOK Then Exit Sub

con.BeginTrans
On Error GoTo myError
con.Execute "Delete  From " & cFile & " where Doc_No = " & addstring(xdoc_no.text)
con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & addstring(xdoc_no.text)
con.CommitTrans

If sDoc_no <> "" Then Exit Sub

If Not openCardTable(tbMode.tbPrevious, XFLAG.Caption) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
cString = "SELECT FILE6_52H.Doc_No,FILE6_52H.PO_NO, Convert(Varchar(10),FILE6_52H.DATE,111),FILE3_10.DESCA " & _
          " FROM FILE6_52H LEFT JOIN FILE3_10 ON FILE6_52H.CODE = FILE3_10.CODE"
If pFilter <> "" Then cString = cString & " WHERE " & pFilter
Generalarray(1) = cString
Generalarray(2) = "order by FILE6_52H.[DATE] DESC,FILE6_52H.DOC_NO DESC"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì·"
listarray(0, 1) = "%%FILE3_10.DESCA%%"

listarray(1, 0) = " «—ÌŒ «·„” ‰œ"
listarray(1, 1) = "(##[Date]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1600

GrdArray(2, 0) = "—ﬁ„ «·ÿ·»Ì…"
GrdArray(2, 1) = 1600

GrdArray(3, 0) = "«·⁄„Ì·"
GrdArray(3, 1) = 3000

Dim aFilter As Variant
aFilter = AddFlag(aFilter, "FILTER", True)
aFilter = AddFlag(aFilter, "FIELD", "DOC_NO")

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.sCaption = "«” ⁄·«„  ”·Ì„ ÿ·»Ì«  «·⁄„·«¡"
oSearchDoc.aFilter = aFilter
oSearchDoc.Show 1
End Sub
Private Sub CmdInform_Click()
CardLookup cFilter
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
End Sub
Private Sub cmdProject_Click()
'ProjectLookup Me, oSearchProject, , , cmdProject.Tag <> ""
End Sub
Private Sub cmdSave_Click()
If Not myValid Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If sDoc_no <> "" Then
    Unload Me
    Exit Sub
End If
If Not openCardTable(tbMode.tbFind, xdoc_no.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub cmdClient_Click()
CLIENTLOOKUP Me, osearchClient, , , IIf(cmdClient.Tag <> "", "ﬂ· «·⁄„·«¡", "")
End Sub

Private Sub CmdUndo_Click()
'openCardTable
myUndo
End Sub

Private Sub Command1_Click()
End Sub
Private Sub cmdYear_Click()
Set oSearchYear.myform = Me
oSearchYear.bEmpty = cmdYear.Tag <> ""
oSearchYear.Show 1
End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xdoc_no.Tag = LoadMode Then
        grid1.SetFocus
    Else
        xDate.SetFocus
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
bedit = True
cFile = "FILE6_52"
cFileHeader = "FILE6_52H"
openCon con

Set grid1.DataSource = data1


If Not openCardTable Then myDefine
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Set OrdersClientsfrm = Nothing
Err.Clear
End Sub
Private Function myValid(Optional bIgMsg As Boolean = True) As Boolean
If Trim(xdoc_no.text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Trim(xCode.text) = "" Then
    MsgBox "«·⁄„Ì· €Ì— „”Ã·"
    Exit Function
End If

If Trim(xCode_desca.Caption) = "" Then
    MsgBox "ﬂÊœ «·⁄„Ì· €Ì— ’«·Õ"
    Exit Function
End If

myValid = True
End Function
Private Sub myload()
xdoc_no.text = CardTable!doc_no
XFLAG.Caption = CardTable!Flag
xDate.text = myFormat_p(CardTable!Date)
xCode.text = CardTable!Code & ""
xCode_desca.Caption = CardTable!code_desca & ""
XPO_NO.Caption = CardTable!PO_NO & ""
xNotes.text = CardTable!NOTES & ""
bIg = True
xIsPosted.Value = IIf(CardTable!isPosted, 1, 0)
xisClosed.Value = IIf(CardTable!ISCLOSED, 1, 0)
bIg = False
panel1(0).Caption = CardTable!UserName & " " & myFormat_p(CardTable!Time, True)
panel1(1).Caption = CardTable!USERNAME2 & " " & myFormat_p(CardTable!TIME2, True)
Handlecontrols LoadMode
myLoadGrd
'CellPos 13, grid1.Rows - 2, grid1.Cols - 1
CellPos 13, 0, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
xdoc_no.text = ""
XFLAG.Caption = ""
xDate.text = myFormat_p(Date)
xCode.text = ""
xCode_desca.Caption = ""
fixGrd

grid1.Rows = 1


myAddItem

Handlecontrols DefineMode
'StatusBar1.Panels(2).text = "·« ÌÊÃœ ”œ«œ"
'StatusBar1.Panels(3).text = ""
'StatusBar1.Panels(4).text = ""

CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bedit

xisClosed.Enabled = nMode = LoadMode And (bEditRecord Or bopt2) And xIsPosted.Value = 0
xIsPosted.Enabled = nMode = LoadMode And (bopt2) And xisClosed.Value = 1

bEditRecord = bEditRecord And xisClosed.Value = 0 And xIsPosted.Value = 0 And nMode = LoadMode

cmdFilter.Visible = cmdFilter.Tag <> ""
'cmdNewInv.Enabled = nMode = LoadMode And bEdit
CmdDel.Enabled = nMode = LoadMode And bEditRecord

cmdSave.Enabled = bEditRecord
CmdDel.Enabled = nMode = LoadMode And bEditRecord

Dim nRecord As Long, nRecords As Long
retRecords XFLAG.Caption, nRecords, nRecord

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
     panel1(0).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""

xdoc_no.Enabled = (nMode = DefineMode)
xdoc_no.Tag = nMode
End Sub

Private Sub SSCommand5_Click()
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    CLIENTLOOKUP Me, osearchClient
End If
End Sub

Private Sub xDoc_No_LostFocus()
myLostFocus xdoc_no
If xdoc_no.text = "" Then
     If xdoc_no.Tag = LoadMode Then
        myDefine
    End If
Else
    If (Not (CardTable.EOF)) And xdoc_no.Tag = LoadMode Then
        If CardTable!doc_no = xdoc_no.text Then
            Exit Sub
        End If
    End If
    
    If openCardTable(tbMode.tbFind, xdoc_no.text) Then
        myload
    ElseIf xdoc_no.Tag = LoadMode Then
        myDefine
    Else
'        xDoc_No.Text = ""
    End If
End If
End Sub
Private Function CalcTotals(Optional Row As Long = 0)
Dim nTotal As Double
Dim i As Long
With grid1
For i = 1 To .Rows - 1
    '.TextMatrix(i, 12) = mRound(.ValueMatrix(i, 10) * .ValueMatrix(i, 11))
    nTotal = mRound(nTotal + .ValueMatrix(i, 12 + 1))
Next
panel1(3).Caption = "«·≈Ã„«·Ì : " & nTotal
End With
End Function
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_no As String = "", Optional bDefine As Boolean = True) As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_no) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

On Error GoTo myError
cFilter = ""
If cmdFilter.Tag <> "" Then cFilter = cFilter & Tr(cFilter) & "DOC_NO IN (" & cmdFilter.Tag & ")"
If cmdClient.Tag <> "" Then cFilter = cFilter & Tr(cFilter) & "FILE6_52H.CODE = " & MyParn(cmdClient.Tag)

'If cmdProject.Tag <> "" Then cFilter = cFilter & turn(cFilter, " And ") & "DOC_NO IN (SELECT " & cFile & " FROM " & cFile & " WHERE " & cFile & ".PROJECT = " & cmdProject.Tag & ")"
'If xYear.text <> "" Then cFilter = cFilter & Tr(cFilter) & "YEAR(DATE) = " & xYear.text
'If xMonth.text <> "" Then cFilter = cFilter & Tr(cFilter) & "MONTH(DATE) = " & xMonth.text

If sDoc_no <> "" Then cFilter = "DOC_NO = " & MyParn(sDoc_no)
cString = "SELECT TOP 1 FILE6_52H.*,FILE3_10.DESCA AS CODE_DESCA " & _
          " FROM FILE6_52H LEFT JOIN FILE3_10 ON FILE6_52H.CODE = FILE3_10.CODE"

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
Private Sub myUndo()
If xdoc_no.Tag = DefineMode Then
    If Not openCardTable Then
        CmdNewInv_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, xdoc_no.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub

Private Sub xcode_LostFocus()
myLostFocus xCode
If Trim(xCode.text) = "" Then
    xCode_desca.Caption = ""
Else
    xCode_desca.Caption = myField("select desca from file3_10 where code = " & MyParn(xCode.text), con) & ""
End If
End Sub
Private Sub chkDay_Click()
If Not openCardTable(tbMode.tbFind, xdoc_no.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub chkMonth_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, xdoc_no.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Sub chkOpen_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, xdoc_no.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Sub chkYear_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, xdoc_no.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New adodb.Recordset
If pDoc_no <> "" Then
    cString = "SELECT Count(*) AS records,COUNT(CASE WHEN FLAG <= " & MyParn(pDoc_no) & " THEN 1  END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM " & cFileHeader & Tr(cFilter, " WHERE ") & cFilter

Set loctable = mycmd(cString, con)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Sub cmdFilter_Click()
cmdFilter.Tag = ""
If Not openCardTable(tbMode.tbFind, xdoc_no.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Sub myproc2(pFilter As String)
oSearchDoc.Hide
cmdFilter.Tag = pFilter
If Not openCardTable(tbMode.tbFirst, xdoc_no.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub myreplaceGrd(Row As Long)
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "QUANT_CONFIRMED", grid1.ValueMatrix(i, 10 + 1))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, cFile)
        Else
            con.Execute addUpdate(aInsert, cFile, "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not myValid(True) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myLoadGrd
    If Row < grid1.Rows - 1 Then
        grid1.Select Row, Col
    Else
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If


If Not validRow(Row, Col) Then
    'CalcTotals Row
    Exit Sub
End If

With grid1
If Row = grid1.Rows - 1 Then
    myAddItem
ElseIf Row = grid1.Rows - 2 And (Col = 0 Or Col = 1) Then
    MyEditItem grid1, Row, Col
End If

If myreplace(Row) Then
    If xdoc_no.Tag = DefineMode Then
        Handlecontrols LoadMode
        myLoadGrd
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myLoadGrd
    Else
        CalcTotals Row
    End If
Else
    myLoadGrd
End If
End With
End Sub
Private Sub grid1_EnterCell()
If Not bEditRecord Then
    grid1.Editable = flexEDNone
ElseIf grid1.Col = 9 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub myLoadGrd()
Dim cString As String
Dim aFields(12)
aFields(0) = "FILE4_10.DESCA"
aFields(1) = "FILE1_10.BARCODE"
aFields(2) = "FILE6_52.ASIN"
aFields(3) = "FACT.DESCA"
aFields(4) = "FILE1_10.MODELFACT0"
aFields(5) = "FILE1_10.DESCA2"
aFields(6) = "file1_10.COLOR"
aFields(7) = "file1_10.SCAL"
aFields(8) = "FILE6_52.QUANT"
aFields(9) = "FILE6_52.QUANT_PO"
aFields(10) = "FILE6_52.PRICE"
aFields(11) = "FILE6_52.TOTAL"
aFields(12) = "FILE6_52.ID"

cString = "SELECT " & arString(aFields, ",") & _
         " FROM FILE6_52 " & _
         " inner join file1_10 on file1_10.item = FILE6_52.item " & _
         " INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT" & _
         " INNER JOIN FILE4_10 ON FILE4_10.CODE = FILE1_10.CODE " & _
         " WHERE  FILE6_52.DOC_NO = " & MyParn(xdoc_no.text) & _
         " ORDER BY FILE4_10.DESCA , FILE1_10.FACT , FILE1_10.MODEL , FILE1_10.COLOR , FILE1_10.C_SCAL "
Set data1.Recordset = mycmd(cString, con)
myAddItem

CalcTotals
fixGrd
End Sub
Private Sub grid1_GotFocus()
grid1_EnterCell
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
        On Error GoTo myError
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from " & cFile & " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        myRemove grid1.Row
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
        MsgBox "«·„‘—Ê⁄ €Ì— „”Ã·"
    End If
ElseIf Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
        MsgBox "«·Œ“‰… €Ì— „”Ã·…"
    End If
End If
End Sub
Private Sub fixGrd()
With grid1

.TextMatrix(0, 0) = "#"
.TextMatrix(0, 1) = "Supler"
.TextMatrix(0, 2) = "Barcode"
.TextMatrix(0, 3) = "ASIN"
.TextMatrix(0, 4) = "BRAND"
.TextMatrix(0, 5) = "SKU"
.TextMatrix(0, 6) = "product Description"
.TextMatrix(0, 7) = "Color"
.TextMatrix(0, 8) = "Size"
.TextMatrix(0, 9) = "Qty"
.TextMatrix(0, 10) = "Qty PO"
.TextMatrix(0, 11) = "Price"
.TextMatrix(0, 12) = "Total"

.WordWrap = True
.RowHeight(0) = 600

.ColWidth(0) = 600
.ColWidth(1) = 1600
.ColWidth(2) = 1500
.ColWidth(3) = 1600
.ColWidth(4) = 1600
.ColWidth(5) = 1600
.ColWidth(6) = 4000
.ColWidth(7) = 1600
.ColWidth(8) = 1600
.ColWidth(9) = 800
.ColWidth(10) = 800
.ColWidth(11) = 800
.ColWidth(12) = 1200

.ColHidden(10) = True
.ColHidden(.Cols - 1) = True

Dim i As Long
For i = 1 To .Cols - 1
    .ColAlignment(i) = flexAlignLeftCenter
Next

For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
Next
End With
End Sub
Private Sub myAddItem()
Exit Sub
With grid1
.AddItem ""
If grid1.Rows > 2 Then
    .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
    .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 2, 1)
End If
End With
End Sub
Private Function validRow(Row As Long, Optional Col As Long = -1) As Boolean
With grid1
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < 9 Then
    grid1.Col = 9
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 9, 9)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
CalcTotals
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
'    If Col = 0 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
'    If Col = 1 And (grid1.TextMatrix(Row, Col) = "" Or grid1.TextMatrix(Row, grid1.Cols - 1) = "") Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'    If grid1.Col = 0 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
'    If grid1.Col = 1 And (grid1.TextMatrix(grid1.Row, grid1.Col) = "" Or grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub xIsClosed_Click()
If bIg Then Exit Sub
If MsgBox(IIf(xisClosed.Value = 0, "› Õ «·„” ‰œ", "«€·«ﬁ «·„” ‰œ"), vbDefaultButton2 + vbOKCancel) <> vbOK Then Exit Sub
On Error GoTo myError
con.Execute "update " & cFileHeader & _
            " SET ISCLOSED = " & xisClosed.Value & _
            " WHERE DOC_NO = " & MyParn(xdoc_no.text)
Inform " „ " & IIf(xisClosed.Value = 0, "› Õ «·„” ‰œ", "«€·«ﬁ «·„” ‰œ") & " »‰Ã«Õ"
myUndo
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xisPosted_Click()
If bIg Then Exit Sub
If MsgBox(IIf(xIsPosted.Value = 0, "› Õ «·„” ‰œ", "«€·«ﬁ «·„” ‰œ"), vbDefaultButton2 + vbOKCancel) <> vbOK Then Exit Sub
On Error GoTo myError
con.Execute "update " & cFileHeader & _
            " SET isPosted = " & xIsPosted.Value & _
            " WHERE DOC_NO = " & MyParn(xdoc_no.text)
Inform " „ " & IIf(xIsPosted.Value = 0, "› Õ «·„” ‰œ", "«€·«ﬁ «·„” ‰œ") & " »‰Ã«Õ"
myUndo
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Function getAmazon() As String
Dim cFileName As String
Common1.InitDir = App.Path & "\AMAZON"
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

cFileName = Common1.FileName

If cFileName = "" Then Exit Function

Dim cSv As New ChilkatCsv
Dim i As Long

cSv.HasColumnNames = 1
nAccess = cSv.LoadFile(cFileName)
If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If


If cSv.NumRows < 1 Then Exit Function
Dim Doc_No_ As String, GS1_ As String, store_ As String
Dim date_ As String, notes_ As String, barcode_ As String, desca2_ As String
Dim asin_ As String, desca_ As String, quant_ As String, price_ As String


Doc_No_ = "Order/PO Number"
GS1_ = "External ID"
store_ = "Fulfillment Center"
date_ = "Order date"
notes_ = "Availability Status"

barcode_ = "External ID"
asin_ = "ASIN"
desca_ = "Title"
quant_ = "Quantity Ordered"
price_ = "Cost"
desca2_ = "Title"


Dim Tb As New ChilkatStringBuilder
Dim cString As New ChilkatStringBuilder
prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption

If cSv.GetCellByName(0, Doc_No_) = "" Then
   MsgBox "—ﬁ„ «·„” ‰œ €Ì— „”Ã·"
   Exit Function
End If
        
If myField("select * from FILE6_52H WHERE DOC_NO = " & MyParn(cSv.GetCellByName(0, Doc_No_)), con) & "" <> "" Then
   MsgBox "«·„” ‰œ „”Ã· „‰ ﬁ»·"
   Exit Function
End If

If Not IsDate(cSv.GetCellByName(0, date_)) Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If

cString.Append "INSERT INTO FILE6_52H"
cString.Append "("
cString.Append "DOC_NO,"
cString.Append "[DATE],"
cString.Append "[NOTES],"
cString.Append "STORE"
cString.Append ")"

cString.Append "VALUES"
cString.Append "("
cString.Append addstring(cSv.GetCellByName(0, Doc_No_)) & ","
cString.Append addDate(cSv.GetCellByName(0, date_)) & ","
cString.Append addstring(cSv.GetCellByName(0, notes_)) & ","
cString.Append addstring(cSv.GetCellByName(0, store_))
cString.Append ");"

On Error GoTo myError
con.BeginTrans
con.Execute cString.GetAsString

For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    cString.Clear
    
    If cSv.GetCellByName(i, barcode_) <> "" Then
        sitem = myField("SELECT * FROM FILE1_10 WHERE BARCODE_GS1 = " & MyParn(cSv.GetCellByName(i, barcode_)), con) & ""
        If sitem <> "" Then
            cString.Append "UPDATE FILE1_10 SET "
            If Trim(cSv.GetCellByName(i, desca2_)) <> "" Then
                cString.Append "DESCA2 = " & addstring(cSv.GetCellByName(i, desca_))
            End If
            If Trim(cSv.GetCellByName(i, asin_)) <> "" Then
                cString.Append ",ENA = " & addstring(cSv.GetCellByName(i, asin_))
            End If
            cString.Append "WHERE ITEM  = " & sitem & ";"
            
            cString.Append "INSERT INTO FILE6_52"
            cString.Append "("
            cString.Append "DOC_NO,"
            cString.Append "ITEM,"
            cString.Append "BARCODE,"
            cString.Append "[ASIN],"
            cString.Append "Quant,"
            cString.Append "Quant_confirmed,"
            cString.Append "PRICE,"
            cString.Append "ROW"
            cString.Append ")"
            
            cString.Append "VALUES"
            cString.Append "("
            cString.Append MyParn(cSv.GetCellByName(0, Doc_No_)) & ","
            cString.Append addstring(sitem) & ","
            cString.Append addstring(cSv.GetCellByName(i, barcode_)) & ","
            cString.Append addstring(cSv.GetCellByName(i, asin_)) & ","
            cString.Append Val(cSv.GetCellByName(i, quant_)) & ","
            cString.Append Val(cSv.GetCellByName(i, quant_)) & ","
            cString.Append Val(cSv.GetCellByName(i, price_)) & ","
            cString.Append i
            cString.Append ")"
        Else
            cString.Append "INSERT INTO FILE6_52E"
            cString.Append "("
            cString.Append "DOC_NO,"
            cString.Append "DESCA2,"
            cString.Append "BARCODE,"
            cString.Append "ASIN,"
            cString.Append "Quant,"
            cString.Append "PRICE,"
            cString.Append "ROW"
            cString.Append ")"
            
            cString.Append "VALUES"
            cString.Append "("
            cString.Append MyParn(cSv.GetCellByName(0, Doc_No_)) & ","
            cString.Append MyParn(cSv.GetCellByName(i, desca2_)) & ","
            cString.Append addstring(cSv.GetCellByName(i, barcode_)) & ","
            cString.Append addstring(cSv.GetCellByName(i, asin_)) & ","
            cString.Append Val(cSv.GetCellByName(i, quant_)) & ","
            cString.Append Val(cSv.GetCellByName(i, price_)) & ","
            cString.Append i
            cString.Append ")"
        End If
        con.Execute cString.GetAsString
    End If
Next
con.CommitTrans
prog1.Visible = False
Me.Caption = sCaption
getAmazon = cSv.GetCellByName(0, Doc_No_)
Exit Function
myError:
MsgBox Err.Description
Err.Clear
If TransCount(con) > 0 Then
    con.RollbackTrans
End If
End Function

Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xcode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xCode2
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xdoc_no
End Sub
