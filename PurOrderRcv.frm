VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form PurOrderRcvfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«” ·«„ ÿ·»Ì… ‘—«¡ „‰ „Ê—œ"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20835
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
   ScaleWidth      =   20835
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
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
      Height          =   1230
      Index           =   0
      Left            =   6975
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   900
      Width           =   3165
      Begin VB.CheckBox xisClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«€·«ﬁ «·„” ‰œ"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   270
         Width           =   1365
      End
      Begin Threed.SSCommand cmdClosePeriod 
         Height          =   510
         Index           =   0
         Left            =   1575
         TabIndex        =   35
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
         TabIndex        =   36
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
      Height          =   1230
      Index           =   4
      Left            =   10170
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   900
      Width           =   1365
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   45
         TabIndex        =   25
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
         Picture         =   "PurOrderRcv.frx":0000
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PurOrderRcv.frx":2925
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   45
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   675
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "PurOrderRcv.frx":5179
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PurOrderRcv.frx":72D9
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
      Left            =   13320
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   0
      Width           =   6990
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   5850
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
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
         Picture         =   "PurOrderRcv.frx":95C6
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PurOrderRcv.frx":B991
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   4725
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   135
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
         Picture         =   "PurOrderRcv.frx":DA3A
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PurOrderRcv.frx":FA42
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   3555
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   135
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "PurOrderRcv.frx":119F9
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PurOrderRcv.frx":14195
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   135
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
         Picture         =   "PurOrderRcv.frx":16629
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   2385
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
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
         Picture         =   "PurOrderRcv.frx":1894C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PurOrderRcv.frx":1ACC2
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   510
         Left            =   1170
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "PurOrderRcv.frx":1CE45
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PurOrderRcv.frx":1F259
      End
   End
   Begin VB.Frame Frame8 
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
      Height          =   780
      Left            =   1845
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1350
      Width           =   5100
      Begin VB.OptionButton TO_EXEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„Œ ·›… ⁄‰ «·«” ·«„"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   315
         Width           =   1770
      End
      Begin VB.OptionButton TO_EXEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«ﬂ»— „‰ «·ÿ·»Ì…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   315
         Width           =   1680
      End
      Begin VB.OptionButton TO_EXEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame nhlj 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   11565
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   675
      Width           =   8745
      Begin VB.CheckBox xMax_Rcv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Õœ ·· ”·Ì„"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   990
         Width           =   1500
      End
      Begin VB.TextBox XNOTES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3870
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   990
         Width           =   3705
      End
      Begin VB.TextBox xDoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6390
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Tag             =   "1"
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox XORDER_NO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3915
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   2310
      End
      Begin VB.Label xStore_Desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   585
         Width           =   3705
      End
      Begin VB.Label xcode_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   585
         Width           =   2310
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Ê—œ"
         Height          =   285
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·«ÕŸ«  "
         Height          =   270
         Left            =   7650
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   990
         Width           =   720
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ „” ‰œ"
         Height          =   285
         Left            =   7695
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·ÿ·»Ì…"
         Height          =   270
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   810
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
         TabIndex        =   7
         Top             =   -270
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Œ“‰"
         Height          =   270
         Left            =   7695
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   630
         Width           =   540
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -3825
      Top             =   6660
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Left            =   -1440
      Top             =   6885
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
      Left            =   -5355
      Top             =   8145
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Left            =   -3645
      Top             =   6975
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Left            =   -2340
      Top             =   6885
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
      Left            =   5670
      Top             =   -135
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
   Begin MSAdodcLib.Adodc data12 
      Height          =   330
      Left            =   -585
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   -585
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -585
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
      Height          =   465
      Left            =   0
      TabIndex        =   27
      Top             =   9630
      Width           =   20835
      _ExtentX        =   36751
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   28
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
         TabIndex        =   29
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
         TabIndex        =   30
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
         TabIndex        =   31
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
         TabIndex        =   22
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   23
      Top             =   9480
      Visible         =   0   'False
      Width           =   20835
      _ExtentX        =   36751
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin Threed.SSCommand cmdFilter 
      Height          =   465
      Left            =   18180
      TabIndex        =   32
      Top             =   8865
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   820
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6585
      Left            =   225
      TabIndex        =   1
      Top             =   2160
      Width           =   20160
      _cx             =   35560
      _cy             =   11615
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   13
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
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   8775
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   41
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
         Picture         =   "PurOrderRcv.frx":213DC
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PurOrderRcv.frx":23583
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   42
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
         Picture         =   "PurOrderRcv.frx":255CA
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PurOrderRcv.frx":276B5
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   43
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
         Picture         =   "PurOrderRcv.frx":296AF
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PurOrderRcv.frx":2B7C0
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   44
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
         Picture         =   "PurOrderRcv.frx":2D7BA
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PurOrderRcv.frx":2F9DE
      End
   End
End
Attribute VB_Name = "PurOrderRcvfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temptable As New ADODB.Recordset
Public sDoc_no As String
Dim cFilter As String
Dim bStopCell As Boolean
Dim bEditRecord As Boolean
Dim bIg As Boolean
Dim tCard As ADODB.Recordset
Dim nRound As Long
Dim CardTable As ADODB.Recordset
Dim nUser As Integer
Dim oSearchItem As New Search_abd, oSearchOrder   As New Search_abd
Dim oSearchDoc As New Search_abd
Dim bEdit As Boolean
Dim cFile As String, cCodeDesca As String
Dim con As New ADODB.Connection
Dim formMode, dDateLast As String
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1

Private Enum enUser
Admin = 2
Super = 1
User = 0
End Enum
Private Function myreplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.text))
'aInsert = AddFlag(aInsert, "[STORE]", addstring(xStore.BoundText))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(XNOTES.text))
aInsert = AddFlag(aInsert, "[ORDER_NO]", addstring(XORDER_NO.text))
'On Error GoTo myError
con.BeginTrans
If xDoc_no.Tag = DefineMode Then
    xDoc_no.text = IncrementString(RetZero(Newflag("FILE6_55H", "DOC_NO", con)))
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_no.text))
    aInsert = AddFlag(aInsert, "userName1", addstring(cUserName))
    aInsert = AddFlag(aInsert, "TIME1", addDate("GETDATE()"))
    con.Execute addInsert(aInsert, "file6_55H")
Else
    aInsert = AddFlag(aInsert, "userName2", addstring(cUserName))
    aInsert = AddFlag(aInsert, "TIME2", addDate("GETDATE()"))
    
    con.Execute addUpdate(aInsert, "file6_55H", "doc_no = " & addstring(xDoc_no.text))
End If

If Row > -2 Then myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myError:
prog1.Visible = False
MsgBox Err.Description
'con.RollbackTrans
Err.Clear
End Function
Sub myProc()
'''On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    'grid1.TextMatrix(grid1.Row, 2) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    Dim RowFound As Long
    Dim sMsg As String
    
    If GrdDesc(oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0), grid1.Row, RowFound, sMsg) Then
        If RowFound <> grid1.Row And Row > 0 Then
            If sMsg <> "" Then Inform sMsg
            Grid1_AfterEdit RowFound, grid1.Col
        Else
            Grid1_AfterEdit grid1.Row, grid1.Col
            CellPos 13, grid1.Row, grid1.Col
        End If
    Else
        If sMsg <> "" Then MsgBox sMsg
    End If
    oSearchItem.Hide
ElseIf ActiveControl.Name = cmdInform.Name Then
    openCardTable tbMode.tbFind, oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
ElseIf ActiveControl.Name = XORDER_NO.Name Then
    XORDER_NO.text = oSearchOrder.grid1.TextMatrix(oSearchOrder.grid1.Row, 0)
    XORDER_NO.Tag = oSearchOrder.grid1.TextMatrix(oSearchOrder.grid1.Row, 0)
    xcode_desca.Caption = oSearchOrder.grid1.TextMatrix(oSearchOrder.grid1.Row, 2)
    xStore_Desca.Caption = oSearchOrder.grid1.TextMatrix(oSearchOrder.grid1.Row, 3)
    xMax_Rcv.Value = IIf(oSearchOrder.grid1.ValueMatrix(oSearchOrder.grid1.Row, 4) = 0, 0, 1)
    Unload oSearchOrder
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CMD_PRINT_Click()
If myValid Then myreplace -2
doprint
End Sub

Private Sub cmdClosePeriod_Click(Index As Integer)
Dim oClosefrm As New closefrm
oClosefrm.sFile = "FILE6_55h"
oClosefrm.sCaption = "«” ·«„ ÿ·»Ì«  „‘ —Ì« "
oClosefrm.sFieldClose = "isclosed"
oClosefrm.nMode = Index
oClosefrm.Show 1
myUndo
End Sub

Private Sub CmdDel_Click()

If MsgBox("Õ–› ?", vbOKCancel + vbDefaultButton2 + vbCritical) <> vbOK Then Exit Sub

con.BeginTrans
On Error GoTo myError
con.Execute "Delete  FROM FILE6_55 where Doc_No = " & addstring(xDoc_no.text)
con.Execute "Delete  FROM file6_55H where Doc_No = " & addstring(xDoc_no.text)
con.CommitTrans

If sDoc_no <> "" Then Exit Sub

If Not openCardTable(tbMode.tbPrevious, xDoc_no.text) Then
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
Private Sub CmdFirst_Click()
If Not openCardTable(tbMode.tbFirst) Then
    myDefine
End If
End Sub

Private Sub CmdInform_Click()
CardLookup
End Sub

Private Sub CmdLast_Click()
If Not openCardTable Then
    myDefine
End If
End Sub
Private Sub CmdNext_Click()
If Not openCardTable(tbMode.tbNext, xDoc_no.text) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, xDoc_no.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdNewInv_Click()
myDefine
End Sub
Private Sub cmdSave_Click()
If Not myValid Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If sDoc_no <> "" Then
    Unload Me
    Exit Sub
End If

If Not openCardTable(tbMode.tbFind, xDoc_no.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xDoc_no.Tag = LoadMode Then
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
Dim ImgXTmp As ImgX
Set ImgXTmp = New ImgX

bEdit = True

openCon con
Set ImgXTmp = Nothing

nRound = 2

'FRM_CLOSED.Visible = bopt2
'Frame7.Visible = bopt2

'Set data1.Recordset = cmd("SELECT * FROM FILE0_40", con).Execute
'Set xStore.RowSource = data1
'xStore.ListField = "Desca"
'xStore.BoundColumn = "Code"

If cBranch = "00" And bSupermode Then
    nUser = enUser.Admin
ElseIf cBranch <> "" Then
    If bopt2 Then
        nUser = enUser.Super
    Else
        nUser = enUser.User
    End If
End If

Set grid1.DataSource = DATA11

If Not openCardTable Then myDefine
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Set Purch_recev = Nothing
Err.Clear
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


If Not validRow(Row) Then
    'CalcTotals Row
    Exit Sub
End If

With grid1
    If Row = grid1.Rows - 1 Then
        bNew = True
        myAddItem
    ElseIf Row = grid1.Rows - 2 And (Col = 0 Or Col = 1) Then
        'MyEditItem grid1, Row, col
    End If
    
    If myreplace(Row) Then
        If xDoc_no.Tag = DefineMode Then
            Handlecontrols LoadMode
            myLoadGrd
        ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
            bStopCell = True
            myLoadGrd
            grid1.Select grid1.Rows - 1, 1
            grid1.ShowCell grid1.Rows - 1, 1
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
ElseIf (grid1.Col = 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Or grid1.Col = 9 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub

Private Sub grid1_GotFocus()
grid1_EnterCell
End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'If grid1.Col = 0 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    'CellPos KeyAscii, Row, Col
    KeyAscii = 0
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
        Cancel = True
    Else
        Dim sMsg As String
        
'        grid1.EditText = retItemCode(grid1.EditText, con)
'        If Not isNum(grid1.EditText) Then
'            MsgBox "»«—ﬂÊœ €Ì— ’ÕÌÕ"
'            n = Beep(2000, 1000)
'            Cancel = True
        Dim RowFound As Long
        If Not GrdDesc(grid1.EditText, Row, RowFound, sMsg) Then
            MsgBox IIf(sMsg = "", "»«—ﬂÊœ €Ì— ’ÕÌÕ", sMsg)
            Cancel = True
            Exit Sub
        End If
        If RowFound <> Row And RowFound > 0 Then
            'grid1.Select RowFound, 0
            'grid1.ShowCell RowFound, 0
            Grid1_AfterEdit RowFound, Col
            If sMsg <> "" Then Inform sMsg
            Cancel = True
        Else
            grid1.EditText = grid1.TextMatrix(Row, 1)
        End If
    End If
ElseIf Col = 9 And Val(grid1.EditText) = 0 Then
    Cancel = True
ElseIf Col = 9 Then
    If xMax_Rcv.Value = 1 Then
        If Val(grid1.EditText) + grid1.ValueMatrix(Row, 10) - grid1.TextMatrix(Row, 8) > 0 Then
            Inform "«·’‰› Ê’· Õœ «· ”·Ì„"
            Cancel = True
        End If
    End If
End If
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And xClosed.Value = 0 Then SuppLookupAll Me, Search3
End Sub

Private Sub TO_EXEL_Click(Index As Integer)
fixGrd
End Sub

Private Sub xDate_Validate(Cancel As Boolean)
    If Not IsDate(xDate.text) Then Cancel = True
End Sub
Private Function myValid(Optional bIgMsg As Boolean = False) As Boolean
If Not IsDate(xDate.text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If


If XORDER_NO.text = "" Then
    If Not bIgMsg Then MsgBox " ”ÃÌ· —ﬁ„ ÿ·»Ì… «·„Ê—œ "
    Exit Function
End If


myValid = True
End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
xDoc_no.text = CardTable!DOC_NO
XORDER_NO.text = CardTable!ORDER_NO
XORDER_NO.Tag = CardTable!ORDER_NO
xMax_Rcv.Value = IIf(CardTable!max_rcv, 1, 0)
xDate.text = myFormat_p(CardTable!Date)
xStore_Desca.Caption = CardTable!STORE_DESCA
xcode_desca.Caption = CardTable!code_desca
'xStore.BoundText = CardTable!STORE & ""
XNOTES.text = CardTable!NOTES & ""
myLoadGrd
Handlecontrols LoadMode

bIg = True
xisClosed.Value = IIf(CardTable!ISCLOSED, 1, 0)
bIg = False

panel1(0).Caption = CardTable!UserName & " " & myFormat_p(CardTable!Time, True)
panel1(1).Caption = CardTable!USERNAME2 & " " & myFormat_p(CardTable!TIME2, True)
Handlecontrols LoadMode
myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myDefine()
xDoc_no.text = RetZero(IncRec(GetDesca("select MAX(doc_no) from file6_55h ", con)), 6)

If xDoc_no.text = "000000" Then xDoc_no.text = "000001"
XORDER_NO.text = ""
xDoc_no.Tag = DefineMode

xisClosed.Value = 0
xMax_Rcv.Value = 0
xDate.text = myFormat_p(Date)
xStore_Desca.Caption = ""
xcode_desca.Caption = ""
XNOTES.text = ""

grid1.Rows = 1
myAddItem
xDoc_no.Tag = 0

fixGrd
Handlecontrols DefineMode
cmdSave.Enabled = True
cmddel.Enabled = True
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_no
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload True
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 112 And grid1.Col = 1 And bEditRecord And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" And Trim(XORDER_NO.text) <> "" Then
    ItemsLook Me, oSearchItem, XORDER_NO.text
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        On Error GoTo myError
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from FILE6_55 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        myRemove grid1.Row
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Function GrdDesc(pItem As String, ByRef Row As Long, ByRef RowFound As Long, Optional ByRef sMsg As String) As Boolean
With grid1
If pItem = "" Or Len(pItem) < 3 Then
    sMsg = "ﬂÊœ €Ì— ’«·Õ"
    Exit Function
End If

If XORDER_NO.text = "" Then
    sMsg = "„” ‰œ «·ÿ·»Ì… €Ì— „”Ã·"
    Exit Function
End If

Dim nquant As Long
Dim cm As New ADODB.command
Set cm = cmd("dbo.sp_item_data", con, adStoredProc, AddFlag(Empty, "ITEM_FIND", pItem))
cm.Execute

If IsNull(cm.Parameters("@ITEM").Value) Then
    sMsg = "ﬂÊœ €Ì— ’«·Õ"
    Exit Function
End If

sitem = cm.Parameters("@ITEM").Value

If xMax_Rcv.Value = 1 Then
    Dim loctable As New ADODB.Recordset
    Dim cString As String
    cString = "SELECT ITEM,QUANT FROM FILE6_50 " & _
              " WHERE ITEM = " & MyParn(sitem) & _
              " AND DOC_NO = " & MyParn(XORDER_NO.text)
    
    Set loctable = cmd(cString, con).Execute
    If loctable.EOF Then
        sMsg = "’‰› €Ì— „”Ã· »«·ÿ·»Ì…"
        loctable.Close
        Set loctable = Nothing
        Exit Function
    Else
        nquant = loctable!Quant
    End If
End If

RowFound = grid1.FindRow(cm.Parameters("@ITEM").Value, , 1)
If RowFound > 0 And RowFound <> Row Then
    If xMax_Rcv.Value = 1 Then
        If grid1.ValueMatrix(RowFound, 12) <= 0 Then
            sMsg = "«·’‰› „ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ " & RowFound & vbCrLf & _
                   "Ê’· Õœ «· ”·Ì„"
            Exit Function
        End If
    End If

    grid1.TextMatrix(RowFound, 9) = grid1.ValueMatrix(RowFound, 9) + 1
    sMsg = "«·’‰› „ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ " & RowFound
Else
    grid1.TextMatrix(Row, 1) = cm.Parameters("@ITEM").Value
    grid1.TextMatrix(Row, 2) = cm.Parameters("@MOSM").Value
    grid1.TextMatrix(Row, 3) = cm.Parameters("@FACT_DESCA").Value
    grid1.TextMatrix(Row, 5) = cm.Parameters("@modelfact0").Value
    grid1.TextMatrix(Row, 5) = cm.Parameters("@DESCA").Value
    grid1.TextMatrix(Row, 6) = cm.Parameters("@SCAL").Value
    grid1.TextMatrix(Row, 7) = cm.Parameters("@COLOR").Value
    grid1.TextMatrix(Row, 8) = nQaunt
    grid1.TextMatrix(Row, 9) = 1
End If
GrdDesc = True
End With
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(5, 1)
Set Generalarray(0) = Me
'                                   0               1                   2                   3                                   4
Generalarray(1) = "SELECT FILE6_55H.DOC_NO," & _
                  "file6_55H.ORDER_NO," & _
                  " FILE4_10.DESCA," & _
                  " FILE0_40.DESCA," & _
                  " FORMAT(FILE6_55H.[DATE], 'yyyy/M/y')" & _
                  " FROM FILE6_55H " & _
                  " INNER JOIN FILE6_50H ON FILE6_55H.ORDER_NO = FILE6_50H.DOC_NO " & _
                  " INNER JOIN FILE4_10 ON FILE6_50H.CODE = FILE4_10.CODE" & _
                  " INNER JOIN FILE0_40 ON FILE6_50H.STORE = FILE0_40.CODE"

If cFilter <> "" Then Generalarray(1) = Generalarray(1) & " WHERE " & cFilter

Generalarray(2) = "ORDER BY FILE6_55H.DOC_NO DESC"
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„ -  «—ÌŒ"
listarray(0, 1) = "( %%FILE6_55H.ORDER_NO%% OR ##FILE6_55H.date## )"
   

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "—ﬁ„ ÿ·»Ì…"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·„Ê—œ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·„Œ“‰"
GrdArray(3, 1) = 1500

GrdArray(4, 0) = "—ﬁ„ «·›« Ê—… "
GrdArray(4, 1) = 1500

GrdArray(5, 0) = "«· «—ÌŒ"
GrdArray(5, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
'oSearchDoc.sId = "sales" & IIf(bopt3, "1", "0")
oSearchDoc.sCaption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_no.text = "" Then Cancel = True
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub fixGrd()
With grid1
.Redraw = flexRDNone
.RowHeight(0) = 800
.WordWrap = True

.ColWidth(0) = 700
.ColWidth(1) = 1200
.ColWidth(2) = 800
.ColWidth(3) = 2000 + IIf(nUser < enUser.Admin, 2000, 0)
.ColWidth(4) = 1500
.ColWidth(5) = 4000 + IIf(nUser < enUser.Admin, 3000, 0)
.ColWidth(6) = 1500
.ColWidth(7) = 1500
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 1000
.ColWidth(11) = 1000
.ColWidth(12) = 1000

.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "«·»«—ﬂÊœ"
.TextMatrix(0, 2) = "«·„Ê”„"
.TextMatrix(0, 3) = "«·„’‰⁄"
.TextMatrix(0, 4) = "«·„ÊœÌ·"
.TextMatrix(0, 5) = "«·’‰›"
.TextMatrix(0, 6) = "«··Ê‰"
.TextMatrix(0, 7) = "«·„ﬁ«”"
.TextMatrix(0, 8) = "ﬂ„Ì… «·ÿ·»Ì…"
.TextMatrix(0, 9) = "ﬂ„Ì… „” ·„…"
.TextMatrix(0, 10) = "„” ·„ ”«»ﬁ"
.TextMatrix(0, 11) = "≈Ã„«·Ì «·„” ·„"
.TextMatrix(0, 12) = "«·»«ﬁÌ"

.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter

.ColHidden(8) = nUser < enUser.Admin
.ColHidden(10) = nUser < enUser.Admin
.ColHidden(11) = nUser < enUser.Admin
.ColHidden(12) = nUser < enUser.Admin

.ColHidden(.Cols - 1) = True

Dim Row As Long
For Row = 1 To .Rows - 1
    .TextMatrix(Row, 0) = Row
    If TO_EXEL(0).Value Then
        .RowHidden(Row) = False
    ElseIf TO_EXEL(2).Value Then
        .RowHidden(Row) = .ValueMatrix(Row, 12) >= 0
    Else
        .RowHidden(Row) = .ValueMatrix(Row, 12) = 0
    End If
Next
.Redraw = flexRDBuffered
End With
End Sub
Private Function RemoveItem(nRow) As Boolean
grid1.RemoveItem Row
End Function
Private Function myreplaceGrd(nRow) As Boolean
Dim aInsert As Variant
With grid1
    For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_no.text))
        aInsert = AddFlag(aInsert, "ITEM", addstring(.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "QUANT", .ValueMatrix(i, 9))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "file6_55")
        Else
            con.Execute addUpdate(aInsert, "file6_55", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myLoadGrd()
With grid1
'    Dim cString As String
'    cString = "SELECT file6_55.ITEM," & _
'              " FILE1_10.MOSM," & _
'              " FACT.DESCA," & _
'              " FILE1_10.MODELFACT0," & _
'              " FILE1_10.DESCA," & _
'              " FILE1_10.COLOR," & _
'              " FILE1_10.SCAL," & _
'              " file6_55.Quant," & _
'              " FILE6_50.QUANT," & _
'              " file6_55.ID " & _
'              " FROM file6_55 " & _
'              " INNER JOIN FILE6_55H ON FILE6_55.DOC_NO = FILE6_55H.DOC_NO" & _
'              " INNER JOIN FILE1_10 ON FILE6_55.ITEM = FILE1_10.ITEM" & _
'              " INNER JOIN FACT ON file1_10.FACT = FACT.CODE" & _
'              " INNER JOIN FILE6_50 ON " & _
'              " FILE6_55H.ORDER_NO = FILE6_50.DOC_NO" & _
'              " AND FILE6_55.ITEM = FILE6_50.ITEM" & _
'              " WHERE FILE6_55.DOC_NO = " & MyParn(xDoc_no.text)

    Dim aPrm As Variant
    aPrm = AddFlag(aPrm, "DOC_NO", TurnValue(xDoc_no.text))
    aPrm = AddFlag(aPrm, "ORDER_NO", TurnValue(XORDER_NO.text))
    Set DATA11.Recordset = cmd("[dbo].[sp_sup_order_rcv]", con, adStoredProc, aPrm).Execute
    .AddItem ""
End With
fixGrd
End Sub
Private Sub myUndo()
If xDoc_no.Tag = DefineMode Then
    If Not openCardTable Then
        CmdNewInv_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, xDoc_no.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
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

On Error GoTo myError
cFilter = ""
If cmdFilter.Tag <> "" Then cFilter = cFilter & Tr(cFilter) & "DOC_NO IN (" & cmdFilter.Tag & ")"

If nUser < enUser.Admin Then
   cFilter = cFilter & Tr(cFilter) & "FILE0_40.BRANCH = " & MyParn(cBranch)
   'cFilter = cFilter & Tr(cFilter) & "FILE6_55H.ISCLOSED = 0"
End If

'If cmdClient.Tag <> "" Then cFilter = cFilter & Tr(cFilter) & "FILE6_52H.CODE = " & MyParn(cmdClient.Tag)

If sDoc_no <> "" Then cFilter = "DOC_NO = " & MyParn(sDoc_no)

cString = "SELECT TOP 1 FILE6_55H.*," & _
          "FILE6_50H.MAX_RCV," & _
          " FILE4_10.DESCA AS CODE_DESCA, " & _
          " FILE0_40.DESCA AS STORE_DESCA " & _
          " FROM FILE6_55H " & _
          " INNER JOIN FILE6_50H ON FILE6_55H.ORDER_NO = FILE6_50H.DOC_NO" & _
          " INNER JOIN FILE0_40 ON FILE6_50H.STORE = FILE0_40.CODE" & _
          " INNER JOIN FILE4_10 ON FILE6_50H.CODE = FILE4_10.CODE"

If pMode = tbMode.tbFirst Then
    cOrder = "Order by FILE6_55H.DOC_NO"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by FILE6_55H.DOC_NO DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "FILE6_55H.Doc_no = " & MyParn(pDoc_no)
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "FILE6_55H.DOC_NO < " & MyParn(pDoc_no)
    cOrder = "Order by FILE6_55H.DOC_NO desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "FILE6_55H.DOC_NO > " & MyParn(pDoc_no)
    cOrder = "Order by FILE6_55H.DOC_NO"
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
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not myValid(bIgMsg) Then Exit Function
If Not IsDgt(.TextMatrix(Row, 1)) Then Exit Function
If .ValueMatrix(Row, 9) = 0 Then Exit Function
'If Me.xMax_Rcv.Value = 1 Then
'    If .ValueMatrix(Row, 12) < 0 Then
'        Exit Function
'    End If
'End If

End With
validRow = True
End Function
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
'    If col = 0 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
ElseIf Col = 1 And grid1.TextMatrix(Row, Col) = "" Then
ElseIf Col < 9 Then
    grid1.Col = 9
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 10)
    grid1.ShowCell grid1.Row, 1
Else
    grid1.Select Row, Col
End If
End Sub

Private Sub xIsClosed_Click()
If bIg Then Exit Sub
If MsgBox(IIf(xisClosed.Value = 0, "› Õ «·„” ‰œ", "«€·«ﬁ «·„” ‰œ"), vbDefaultButton2 + vbOKCancel) <> vbOK Then Exit Sub
On Error GoTo myError
con.Execute "update FILE6_55H" & _
            " SET ISCLOSED = " & xisClosed.Value & _
            " WHERE DOC_NO = " & MyParn(xDoc_no.text)
Inform " „ " & IIf(xisClosed.Value = 0, "› Õ «·„” ‰œ", "«€·«ﬁ «·„” ‰œ") & " »‰Ã«Õ"
myUndo
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub xOrder_no_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then orderLookup
End Sub
Private Sub doprint()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
cStr1 = " SELECT FACT.DESCA AS fact_name, FILE1_10.MODELFACT0, FILE1_10.desca AS desca_item, FILE1_10.COLOR, FILE1_10.SCAL, FILE1_10.C_SCAL, FILE6_55.QUANT, FILE1_10.COST , FILE1_10.BARCODE_GS1 , FILE1_10.BARCODE , FILE1_10.BARCODE13 " & _
        " FROM   FILE1_10 INNER JOIN FILE6_55 ON FILE1_10.ITEM = FILE6_55.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE where file6_55.DOC_NO = " & MyParn(xDoc_no.text)
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText


With sourcetable
Do While Not .EOF
    If !Quant > 0 Then
        temptable.AddNew
        temptable!STR6 = TurnValue(xCode.text)
        temptable!str3 = TurnValue(xCode.text)
        temptable!Date1 = xDate.text
        temptable!STR4 = !fact_name
        temptable!str10 = !MODELFACT0
        temptable!str2 = !desca_item
        If Not IsNull(!BARCODE_GS1) Then
            temptable!str9 = !BARCODE_GS1
        ElseIf Not IsNull(!BARCODE13) Then
            temptable!str9 = !BARCODE13
        ElseIf Not IsNull(!BARCODE) Then
            temptable!str9 = !BARCODE
        End If
        temptable!str8 = xDoc_no.text
        
        temptable!str2 = !desca_item
        temptable!STR7 = !color
        temptable!STR5 = !SCAL
        temptable!val5 = !C_SCAL
        temptable!VAL1 = !Quant
        temptable!str11 = TurnValue(xDoc_no.text)
        temptable.Update
    End If
    .MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\Recevd_DOC.rpt"
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit

xisClosed.Enabled = nMode = LoadMode And (nUser = enUser.Admin Or (nUser = enUser.Super And xisClosed.Value = 0 And bEdit))
'xIsPosted.Enabled = nMode = LoadMode And (bopt2) And xisClosed.Value = 1
bEditRecord = bEditRecord And xisClosed.Value = 0
cmdFilter.Visible = cmdFilter.Tag <> ""
'cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmddel.Enabled = nMode = LoadMode And bEditRecord

cmdSave.Enabled = bEditRecord
cmddel.Enabled = nMode = LoadMode And bEditRecord

Dim nRecord As Long, nRecords As Long
retRecords xDoc_no.text, nRecords, nRecord

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
     panel1(0).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""

xDoc_no.Enabled = (nMode = DefineMode)
xDoc_no.Tag = nMode
End Sub
Private Function retRecords(pCode, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pCode <> "" Then
    cString = "SELECT Count(FILE6_55H.DOC_NO) AS records,COUNT(CASE WHEN FILE6_55H.DOC_NO <= " & MyParn(pCode) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(FILE6_55H.DOC_NO) AS records,0 as record"
End If

cString = cString & _
          " FROM FILE6_55H" & _
          " INNER JOIN FILE6_50H ON FILE6_55H.ORDER_NO = FILE6_50H.DOC_NO" & _
          " INNER JOIN FILE0_40 ON FILE6_50H.STORE = FILE0_40.CODE" & _
          " INNER JOIN FILE4_10 ON FILE6_50H.CODE = FILE4_10.CODE"
If cFilter <> "" Then cString = cString & " WHERE " & cFilter

Set loctable = mycmd(cString, con)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
grid1.Select Row, 1
Dim i As Long
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
On Error Resume Next
grid1.SetFocus
End Sub
Private Sub myAddItem()
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
End Sub
Private Function CalcTotals(Optional Row As Long = 0)
Dim nTotal As Double
Dim i As Long
With grid1
.TextMatrix(0, 8) = "ﬂ„Ì… «·ÿ·»Ì…"
.TextMatrix(0, 9) = "ﬂ„Ì… „” ·„…"
.TextMatrix(0, 10) = "„” ·„ ”«»ﬁ"
.TextMatrix(0, 11) = "≈Ã„«·Ì «·„” ·„"
.TextMatrix(0, 12) = "«·»«ﬁÌ"

For i = IIf(Row = 0, 1, Row) To IIf(Row = 0, .Rows - 1, Row)
    grid1.TextMatrix(i, 11) = grid1.ValueMatrix(i, 9) + grid1.ValueMatrix(i, 10)
    grid1.TextMatrix(i, 12) = grid1.ValueMatrix(i, 8) - grid1.ValueMatrix(i, 11)
Next
'panel1(3).Caption = "«·≈Ã„«·Ì : " & nTotal
End With
End Function
Private Sub xOrder_no_LostFocus()
myLostFocus xDoc_no
If XORDER_NO.Locked Then Exit Sub

If Trim(XORDER_NO.text) <> "" And XORDER_NO.Tag = XORDER_NO.text Then
    Exit Sub
End If

xcode_desca.Caption = ""
xStore_Desca.Caption = ""
XORDER_NO.Tag = ""
xMax_Rcv.Value = 0

If Trim(XORDER_NO.text) = "" Then Exit Sub

XORDER_NO.text = RetZero(XORDER_NO.text)

Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "SELECT FILE6_50H.MAX_RCV," & _
          "FILE4_10.DESCA AS CODE_DESCA," & _
          "FILE0_40.DESCA AS STORE_DESCA" & _
          " FROM FILE6_50H " & _
          " INNER JOIN FILE4_10 ON FILE6_50H.CODE = FILE4_10.CODE" & _
          " INNER JOIN FILE0_40 ON FILE6_50H.STORE = FILE0_40.CODE"
Set loctable = cmd(cString, con).Execute

If Not loctable.EOF Then
    xcode_desca.Caption = loctable!code_desca
    xStore_Desca.Caption = loctable!STORE_DESCA
    xMax_Rcv.Value = IIf(loctable!max_rcv, 1, 0)
    XORDER_NO.Tag = XORDER_NO.text
End If
End Sub
Private Sub orderLookup()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT FILE6_50H.DOC_NO," & _
                  "FORMAT(FILE6_50H.[DATE],'yyyy/M/d')," & _
                  " FILE4_10.Desca ," & _
                  " FILE0_40.desca ," & _
                  " FILE6_50H.MAX_RCV" & _
                  " FROM  FILE6_50H " & _
                  " INNER JOIN FILE4_10 ON FILE6_50H.CODE  = FILE4_10.CODE " & _
                  " INNER JOIN FILE0_40 on FILE6_50H.STORE = FILE0_40.CODE" & _
                  " WHERE ISEND_ORDER = 0 " & _
                  " AND FILE6_50H.ISCLOSED = 0"
If nUser < enUser.Admin Then
    Generalarray(1) = Generalarray(1) & " AND FILE0_40.BRANCH = " & MyParn(cBranch)
End If
Generalarray(2) = "Order by FILE6_50H.DATE DESC," & _
                  "FILE6_50H.DOC_NO DESC"
                  
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "—ﬁ„ «·„” ‰œ-«”„ «·„Ê—œ"
listarray(0, 1) = "(@@FILE6_50H.Doc_No@@6 OR %%FILE4_10.DESCA%%)"

listarray(1, 0) = "«· «—ÌŒ"
listarray(1, 1) = "(##DATE##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«”„ «·„Ê—œ"
GrdArray(2, 1) = 3000

GrdArray(3, 0) = "«·„Œ“‰"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "Õœ  ”·Ì„"
GrdArray(4, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchOrder.sCaption = "«” ⁄·«„"
oSearchOrder.Show 1
End Sub
Private Sub xNotes_GotFocus()
myGotFocus XNOTES
End Sub
Private Sub xNotes_LostFocus()
myLostFocus XNOTES
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_no
End Sub
Private Sub xOrder_no_GotFocus()
myGotFocus XORDER_NO
End Sub
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
