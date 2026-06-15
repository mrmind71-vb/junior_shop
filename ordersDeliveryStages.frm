VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form online_Stage_main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "ÿ·»Ì«  «Ê‰ ·«Ì‰"
   ClientHeight    =   11055
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
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   7650
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   90
      Width           =   4920
      Begin Threed.SSCommand cmdStage 
         Height          =   465
         Left            =   45
         TabIndex        =   32
         Top             =   135
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   820
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
         Caption         =   "«Œ «— „—Õ·… «· ÃÂÌ“"
         TagVariant      =   "«Œ «— «·„—Õ·…"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   7650
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   765
      Width           =   4920
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„€·ﬁ…"
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
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   225
         Width           =   735
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„› ÊÕ…"
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
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„·€«…"
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
         Index           =   4
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   225
         Width           =   690
      End
      Begin VB.OptionButton optType 
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
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   225
         Width           =   690
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„»«⁄…"
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
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   225
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1410
      Left            =   12600
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   7710
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
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Text            =   "2026-05-01"
         Top             =   225
         Width           =   1410
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
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1590
      End
      Begin VB.TextBox XDOC_NO 
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
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   1410
      End
      Begin VB.TextBox xphone 
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   945
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo xpay 
         Height          =   330
         Left            =   180
         TabIndex        =   3
         Top             =   585
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xMan 
         Height          =   330
         Left            =   180
         TabIndex        =   5
         Top             =   945
         Width           =   2670
         _ExtentX        =   4710
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ"
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ ÿ·»Ì…"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·”œ«œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   630
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «· ·Ì›Ê‰"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   945
         Width           =   945
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„‰œÊ»"
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   945
         Width           =   705
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   24
      Top             =   10440
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   7305
      Begin Threed.SSCommand cmdCSV 
         Height          =   510
         Left            =   4410
         TabIndex        =   14
         Top             =   135
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
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
         Picture         =   "ordersDeliveryStages.frx":0000
         Caption         =   "CSV ”Õ»"
         TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   510
         Left            =   2790
         TabIndex        =   15
         Top             =   135
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   900
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
         Picture         =   "ordersDeliveryStages.frx":25D3
         Caption         =   " ’œÌ— Excel"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1395
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
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
         Picture         =   "ordersDeliveryStages.frx":4A3D
         Caption         =   " ÿ»«⁄…"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ordersDeliveryStages.frx":6DB3
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   510
         Left            =   5940
         TabIndex        =   33
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
         Picture         =   "ordersDeliveryStages.frx":8F36
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   34
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
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "ordersDeliveryStages.frx":BDE9
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   -315
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
   Begin MSAdodcLib.Adodc data1 
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
   Begin MSAdodcLib.Adodc DATA6 
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
   Begin MSAdodcLib.Adodc DATA7 
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
   Begin MSAdodcLib.Adodc Adodc1 
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
      Caption         =   "data7"
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
   Begin MSAdodcLib.Adodc Adodc2 
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA4 
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7305
      Left            =   90
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1485
      Width           =   20220
      _cx             =   35666
      _cy             =   12885
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
      Cols            =   29
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   10635
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   741
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   21
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
         TabIndex        =   22
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
         TabIndex        =   23
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
End
Attribute VB_Name = "Online_Stage_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3, lIsBoxOnline As Boolean
Dim cString As String
Dim CLIST As String
Dim oSearchMan As New Search_abd
Dim oSearchStage As New Search_rs
Dim cStr1 As String, cStr2 As String
Dim oSearchItem As New Search3
Private Sub cmd_addexel_Click()
AddFromExel
Inform "  „ «÷«›… «·ÿ·»Ì«  "
myload
End Sub
Private Sub CMD_PRINT_Click()
doprint_day
End Sub
Private Sub CMD_SEND_Click()
'Dim cDocSalPost As String
If cBranch = "00" Then Exit Sub

If grid1.TextMatrix(grid1.Row, 16) <> "" Then Exit Sub

Dim oOnlineCheck As New OnlineCheck
Set oOnlineCheck.myForm = Me
oOnlineCheck.sDoc_no = grid1.TextMatrix(grid1.Row, 0)
oOnlineCheck.Show 1

'myreplace_Sales

'If grid1.TextMatrix(grid1.Row, 16) = "" Then
'    If MsgBox(" —ÕÌ· ·›« Ê—… „»Ì⁄«  ", vbYesNo + vbDefaultButton2) = vbYes Then
'        cDocSalPost = myreplace_Sales()
'        If cDocSalPost <> "" Then
'            MsgBox " „  —ÕÌ· ·»Ê‰ „»Ì⁄«  —ﬁ„ " & cDocSalPost
'        Else
'            MsgBox "·„ Ì „ ⁄„· „” ‰œ «·„»Ì⁄«   "
'        End If
'    End If
'    myload
'    grid2.Rows = 2
'End If
End Sub
Private Sub cmdCSV_Click()
Set onlineCSVfrm.myForm = Me
'Set onlineCSVfrm.con = con
onlineCSVfrm.Show 1
End Sub

Private Sub cmdExcel_Click()
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption)
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

Private Sub cmdStage_Click()
StagesOnlineLook Me, oSearchStage, , , IIf(cmdStage.Tag = "", "", "ﬂ· «·⁄„·«¡")
End Sub

Private Sub Form_Load()
    
'    opt(0).Visible = cBranch = "00"
'    opt(2).Visible = cBranch = "00"
'    opt(3).Visible = cBranch = "00"
'    opt(6).Visible = cBranch = "00"
    
'    Check1.Visible = (cBranch = "00")
'    cmdCSV.Enabled = (cBranch = "00")
    
    Dim con As New ADODB.Connection
    If openCn(con) Then
        Set xpay.RowSource = myRs("SELECT Payment_Method FROM file6_90h GROUP BY Payment_Method ", con)
        xpay.ListField = "Payment_Method"
        xpay.BoundColumn = "Payment_Method"
        
            
        Set xMan.RowSource = myRs("SELECT CODE,DESCA FROM FILE6_25 WHERE FILE6_25.BRANCH IN (SELECT FILE0_40.BRANCH FROM FILE0_40 WHERE FILE0_40.online =  1) AND FILE6_25.ISSTOP = 0 ORDER BY CODE", con)
        xMan.ListField = "Desca"
        xMan.BoundColumn = "Code"
        closeCon con
    End If
    Fixgrd
End Sub
Public Sub myload()
Dim i As Double
Dim cString  As String, cStr2 As String
Dim cWhere As String
Dim cSubTotal As String

With grid1
    cString = " SELECT DOC_NO," & _
              " FORMAT(DATE,'yyyy/M/d')," & _
              " NAME," & _
              " PHONE," & _
              "Shipping_City," & _
              "STREET, " & _
              "FILE6_90H.TOTAL_ITEM" & "," & _
              "DISCOUNT, " & _
              "FILE6_90H.TOTAL_ITEM - FILE6_90H.DISCOUNT ," & _
              "DISCOUNT_CODE," & _
              "SHIPPING, " & _
              "Payment_Method," & _
              "STORE," & _
              "STAGES_CODES.DESCA ," & _
              "FILE6_25.DESCA," & _
              "SEND_USER," & _
              "FORMAT(SEND_TIME,'yyyy/M/d')," & _
              "NOTES_ORDER ," & _
              "SUBSTRING(SALES_DOC,7,20) ," & _
              "FORMAT(SALES_DATE,'yyyy/M/d')," & _
              "SHIP_NO," & _
              "FORMAT(SHIP_DATE,'yyyy/M/d'),"
cString = cString & _
              " NOTES," & _
              " FORMAT(DelOrder_Date,'yyyy/M/d')," & _
              " FILE6_90H.MAN" & _
              " FROM FILE6_90H " & _
              " LEFT JOIN FILE6_25 ON FILE6_90H.MAN = FILE6_25.CODE" & _
              " INNER JOIN STAGES_CODES ON FILE6_90H.STAGE = STAGES_CODES.CODE" & _
              " LEFT JOIN MAN_ONLINE_CODES ON FILE6_90H.MAN_STAGE = MAN_ONLINE_CODES.CODE"
              
    If optType(1).Value Then
        cWhere = cWhere & Tr(cWhere) & "FILE6_90H.DOC_NO IN (SELECT ORDER_NO FROM vw_online_orders_open)"
    ElseIf optType(2).Value Then
        cWhere = cWhere & Tr(cWhere) & "(FILE6_90H.DOC_NO IN (SELECT ORDER_NO FROM vw_online_orders_closed) OR FILE6_90H.CANCELED = 1)"
    ElseIf optType(3).Value Then
        cWhere = cWhere & Tr(cWhere) & "FILE6_90H.DOC_NO IN (SELECT ORDER_NO FROM vw_online_orders_closed)"
    ElseIf optType(4).Value Then
        cWhere = cWhere & Tr(cWhere) & "FILE6_90H.CANCELED = 1"
    End If
    
    If cmdStage.Tag <> "" Then
        cWhere = cWhere & Tr(cWhere) & "FILE6_90H.STAGE = " & cmdStage.Tag
    End If
'    If opt(2).Value Then
'        cString = cString & _
'                 " INNER JOIN vw_online_orders_closed as v on FILE6_90H.DOC_NO = v.ORDER_NO"
'    ElseIf opt(3).Value Then
'        cWhere = cWhere & Tr(cWhere) & "CANCELED = 1"
'    ElseIf opt(4).Value Then
'        cString = cString & _
'                 " INNER JOIN vw_online_invoices_open as v on FILE6_90H.DOC_NO = v.ORDER_NO"
'    ElseIf opt(5).Value Then
'        cString = cString & _
'                 " INNER JOIN vw_online_orders_open as v on FILE6_90H.DOC_NO = v.ORDER_NO"
'    ElseIf opt(6).Value Then
'        cString = cString & _
'                 " INNER JOIN vw_online_invoices_closed as v on FILE6_90H.DOC_NO = v.ORDER_NO"
'    End If
    
    If xdoc_no.text <> "" Then
        cWhere = cWhere & Tr(cWhere) & " [DOC_NO] = " & MyParn(xdoc_no.text)
    End If
    
    If xphone.text <> "" Then
        cWhere = cWhere & Tr(cWhere) & " [phone] = " & MyParn(xphone.text)
    End If
    
    If xpay.BoundText <> "" Then
        cWhere = cWhere & Tr(cWhere) & " [Payment_Method] = " & MyParn(xpay.text)
    End If
    
    
    If xMan.MatchedWithList Then
        cWhere = cWhere & Tr(cWhere) & "MAN = " & MyParn(xMan.BoundText)
    End If
    
    If IsDate(xDate1.text) Then
        cWhere = cWhere & Tr(cWhere) & " [DATE] >= " & DateSq(xDate1.text)
    End If
    
    If IsDate(xdate2.text) Then
        cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xdate2.text)
    End If
            
        
    
    'If xSales(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " SALES_DOC IS NULL"
    'If xSales(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " SALES_DOC IS NOT NULL "
    
    'If RET_NOTSHIP.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " DelOrder_Date IS NOT NULL AND DelOrder_Date2 IS NULL "
    
    If cWhere <> "" Then cString = cString & " WHERE " & cWhere
    cString = cString & " ORDER BY DOC_NO "
    'On Error GoTo myerror
    Set grid1.DataSource = myRs(cString)
End With
Fixgrd
grid1.Cell(flexcpAlignment, 0, 0, grid1.Rows - 1, grid1.Cols - 1) = 7
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub Fixgrd()
With grid1
    .RowHeight(0) = 600
    .WordWrap = True
    .FrozenCols = 3
    
    .TextMatrix(0, 0) = "—ﬁ„"
    .TextMatrix(0, 1) = "«· «—ÌÕ"
    .TextMatrix(0, 2) = "«·≈”„"
    .TextMatrix(0, 3) = " ·Ì›Ê‰"
    .TextMatrix(0, 4) = "«·„œÌ‰…"
    .TextMatrix(0, 5) = "⁄‰Ê«‰"
    
    .TextMatrix(0, 6) = "≈Ã„«·Ï «’‰«›"
    .TextMatrix(0, 7) = "Œ’„"
    .TextMatrix(0, 8) = "«·≈Ã„«·Ï"
    .TextMatrix(0, 9) = "ﬂÊœ Œ’„"
    .TextMatrix(0, 10) = "«·‘Õ‰"
    
    .TextMatrix(0, 11) = "ÿ—Ìﬁ… «·œ›⁄"
    .TextMatrix(0, 12) = "«·›—⁄"
    .TextMatrix(0, 13) = "«·„—Õ·…"
    .TextMatrix(0, 14) = "«·„‰œÊ»"
    .TextMatrix(0, 15) = "«·„—”·"
    .TextMatrix(0, 16) = "SEND TIME "
    .TextMatrix(0, 17) = "„·«ÕŸ«  ÿ·»Ì…"
    .TextMatrix(0, 18) = "»Ê‰ „»Ì⁄« "
    .TextMatrix(0, 19) = " «—ÌŒ"
    
    .TextMatrix(0, 20) = "»Ê·Ì’… ‘Õ‰"
    .TextMatrix(0, 21) = " «—ÌŒ «·‘Õ‰"
        
    .ColHidden(9) = True
    .ColHidden(10) = True
    .ColHidden(12) = True
    .ColHidden(15) = True
    .ColHidden(21) = True
    .ColHidden(16) = True
    .ColHidden(24) = True
    
    .TextMatrix(0, 22) = "”»» «·≈·€«¡"
    .TextMatrix(0, 23) = "«·€«¡ «·ÿ·»"
    
    '.TextMatrix(0, 24) = " «—ÌŒ «” ·«„ «·„— Ã⁄"
    '.TextMatrix(0, 25) = "„·«ÕŸ«  «·—∆Ì”Ì"
    '.TextMatrix(0, 26) = "„·«ÕŸ«  «·›—Ê⁄"
    '.TextMatrix(0, 27) = "«·„‰œÊ»"

    
    .ColDataType(1) = flexDTDate
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(9) = flexDTDouble
    .ColDataType(10) = flexDTDouble
    
                    
    .ColWidth(0) = 800
    .ColWidth(1) = 1300
    .ColWidth(2) = 2000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 4000
    .ColWidth(6) = 1000
    .ColWidth(7) = 900
    .ColWidth(8) = 800
    .ColWidth(9) = 800
    .ColWidth(10) = 900
    .ColWidth(11) = 1400
    .ColWidth(12) = 1800
    .ColWidth(13) = 1200
    .ColWidth(14) = 1500
    .ColWidth(15) = 1800
    
    .ColWidth(18) = 1000
    .ColWidth(17) = 2000
    .ColWidth(19) = 1300
    .ColWidth(20) = 2000
    .ColWidth(22) = 1300
    .ColWidth(23) = 1300
    .ColWidth(24) = 1300
    
    For nRow = 1 To .Rows - 1
'       .TextMatrix(nRow, 6) = Val(.TextMatrix(nRow, 6)) + Val(.TextMatrix(nRow, 7))
    Next nRow
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    '.ColDataType(0) = flexDTDate
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTCount, -1, 0, "#", vbYellow, , True, ""
    
    .Subtotal flexSTSum, -1, 6, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 7, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 9, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 10, "#", vbYellow, , True, ""
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'closeCon con
    'If cBranch <> "00" Then closeCon con_MyShop
    SaveText Me, , Array(xDate1.Name, xdate2.Name)
End Sub
Sub AddFromExel()
    Dim xl As New Excel.Application, nREcOrder As Double
    Dim xlsheet As Excel.Worksheet, nItem As String, lAddOrder As Boolean
    Dim xlwbook As Excel.Workbook, cItem As String
    Dim cFileName As String, cBarCode As String, cDoc_No As String
    Dim nCount As Double, cModel As String, nquant As Double, nPrice As Double, nPriceC As Double
    Dim cMail As String, nSubTotal As Double, nTotal As Double, nDisc As Double, nShip As Double, cDidc_Code As String, dDate As Date, cCity As String, cAddress1 As String, cAddress2 As String, cPhone As String, cSku As String, cName As String, cStreet As String, cPayment   As String, cShipping_City As String
    Dim aInsert As Variant, aInsert2 As Variant

    If MsgBox("≈÷«›… »Ì«‰«  „‰ „·› Excel", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "Excel (*.XLS*)|*.XLS*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        If cFileName <> "" Then
            Set xlwbook = xl.Workbooks.Open(cFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
            With grid1
            i = 1
            Do While True
                i = i + 1
                Me.Caption = i
                If (xlsheet.Cells(i, 1)) <> "" Then
                    nCount = i
                Else
                    Exit Do
                End If
            Loop
            prog1.Visible = True
            prog1.Value = 0
            prog1.Max = nCount
            prog1.Min = 0
            .Rows = 1
            prog1.Min = 0
            prog1.Max = nCount
            For nRow = 2 To nCount
                prog1.Value = nRow
                If xlsheet.Cells(nRow, 1) <> "" Then
                    If cDoc_No <> Mid(xlsheet.Cells(nRow, 1), 2) Then
                        cDoc_No = Mid(xlsheet.Cells(nRow, 1), 2)
                        aInsert = AddFlag(Empty, "DOC_NO", addstring(cDoc_No))
                        For nCol = 2 To 90
                            Select Case xlsheet.Cells(1, nCol)
                            Case "Email"
                                    aInsert = AddFlag(aInsert, "E_MAIL", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Subtotal"
                                    aInsert = AddFlag(aInsert, "SUBTOTAL", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Shipping"
                                    aInsert = AddFlag(aInsert, "SHIPPING", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Total"
                                    aInsert = AddFlag(aInsert, "TOTAL", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Discount Code"
                                    aInsert = AddFlag(aInsert, "DISCOUNT_CODE", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Discount Amount"
                                    aInsert = AddFlag(aInsert, "DISCOUNT", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Created at"
                                    aInsert = AddFlag(aInsert, "DATE", addDate(DateValue(Mid(xlsheet.Cells(nRow, nCol), 1, 10))))
                                Case "Shipping Name"
                                    aInsert = AddFlag(aInsert, "NAME", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Shipping Street"
                                    cStreet = Trim(xlsheet.Cells(nRow, nCol))
                                    cStreet = Replace(cStreet, "'", " ")
                                    aInsert = AddFlag(aInsert, "STREET", addstring(cStreet))
                                Case "Shipping Address1"
                                    cAddress1 = Trim(xlsheet.Cells(nRow, nCol))
                                    cAddress1 = Replace(cAddress1, "'", " ")
                                    aInsert = AddFlag(aInsert, "Address1", addstring(cAddress1))
                                Case "Shipping Address2"
                                    cAddress2 = Trim(xlsheet.Cells(nRow, nCol))
                                    cAddress2 = Replace(cAddress1, "'", " ")
                                    aInsert = AddFlag(aInsert, "Address2", addstring(cAddress2))
                                Case "Shipping City"
                                    cCity = Trim(xlsheet.Cells(nRow, nCol))
                                    cCity = Replace(cCity, "'", " ")
                                    aInsert = AddFlag(aInsert, "CITY", addstring(cCity))
                                Case "Shipping Phone"
                                    cPhpne = Trim(xlsheet.Cells(nRow, nCol))
                                    cPhpne = Replace(cPhpne, " ", "")
                                    aInsert = AddFlag(aInsert, "PHONE", addstring(cPhpne))
                                Case "Payment Method"
                                    aInsert = AddFlag(aInsert, "Payment_Method", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Shipping Province Name"
                                    aInsert = AddFlag(aInsert, "Shipping_City", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Notes"
                                    aInsert = AddFlag(aInsert, "NOTES_order", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Payment ID"
                                    aInsert = AddFlag(aInsert, "PAYMENT_ID", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                            End Select
                        Next nCol
                        lAddOrder = False
                        aRet = aGetDesca("SELECT DOC_NO , STORE FROM FILE6_90H WHERE DOC_NO = " & MyParn(cDoc_No), con)
                        If UBound(aRet) = 0 Then
                            lAddOrder = True
                            con.Execute addInsert(aInsert, "FILE6_90H")
                        Else
                            If TurnValue(aRet(2), Null, "") = "" Then
                                lAddOrder = True
                                con.Execute addUpdate(aInsert, "FILE6_90H", " DOC_NO = " & MyParn(cDoc_No))
                            End If
                        End If
                        If lAddOrder Then
                            nREcOrder = nREcOrder + 1
                            con.Execute " DELETE FROM FILE6_90 WHERE DOC_NO = " & MyParn(cDoc_No)
                        End If
                    End If
                    If lAddOrder Then
                                                aInsert2 = AddFlag(Empty, "DOC_NO", addstring(cDoc_No))
                        cItem_NAME = ""
                        For nCol = 2 To 90
                            Select Case xlsheet.Cells(1, nCol)
                                Case "Lineitem price"
                                    aInsert2 = AddFlag(aInsert2, "PRICE", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Lineitem quantity"
                                    aInsert2 = AddFlag(aInsert2, "QUANT", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Lineitem name"
                                    cItem_NAME = Trim(xlsheet.Cells(nRow, nCol))
                                    aInsert2 = AddFlag(aInsert2, "ITEM_NAME", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Lineitem pricename"
                                    aInsert2 = AddFlag(aInsert2, "PRICE", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Lineitem sku"
                                    cSku = xlsheet.Cells(nRow, nCol)
                                    aInsert2 = AddFlag(aInsert2, "sku", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                            End Select
                        Next nCol
                        
                        cItem = ""
                        Dim loctable As ADODB.Recordset
                        If Len(cSku) > 8 Then
                            Set loctable = ItemFind_BARCODE(cSku, con)
                            If loctable.RecordCount > 0 Then cItem = loctable!Item
                        Else
                            cItem = cSku
                        End If
                        If cItem = "" Then MsgBox "ERROR SKU ORDER NO # " & cDoc_No & " - " & cItem_NAME
                        aInsert2 = AddFlag(aInsert2, "ITEM", addstring(cItem))
                        con.Execute addInsert(aInsert2, "FILE6_90")
                    End If
                End If
            Next nRow
            MsgBox " „ «÷«›… «·ÿ·»Ì«  ⁄œœ " & nREcOrder
            End With
        End If
    End If
End Sub
Private Sub grid1_DblClick()
If grid1.Row < 1 Or grid1.Row = grid1.Rows - 1 Then Exit Sub
If cBranch <> "00" Then
    Dim aValues As Variant
    If IsNull(rsValue("select order_no from vw_online_orders_open where order_no = " & MyParn(grid1.TextMatrix(grid1.Row, 0)))) Then
        Exit Sub
    End If
    
    Dim cString As String
    cString = "SELECT ORDER_NO " & _
              " FROM vw_online_invoices_open" & _
              " WHERE ORDER_NO = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
    If Not IsNull(rsValue(cString)) Then
        orders_online_invoices.bEdit = True
        orders_online_invoices.sStore = grid1.TextMatrix(grid1.Row, 12)
        orders_online_invoices.sOrder_No = grid1.TextMatrix(grid1.Row, 0)
        orders_online_invoices.Show 1
        Exit Sub
    End If
Else
    orders_online_items.sDoc_no = grid1.TextMatrix(grid1.Row, 0)
    'orders_online_items.sStore = grid1.TextMatrix(grid1.Row, 12)
    orders_online_items.Show 1
End If


'If cBranch <> "00" Then
'    If IsNull(rsValue("select order_no from vw_online_orders_open where order_no = " & MyParn(grid1.TextMatrix(grid1.Row, 0)))) Then
'        Exit Sub
'    End If
'    Dim cString As String
'    cString = "SELECT ORDER_NO " & _
'              " FROM vw_online_invoices_open" & _
'              " WHERE ORDER_NO = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
'    If Not IsNull(rsValue(cString)) Then
'        orders_online_invoices.bEdit = True
'        orders_online_invoices.sStore = grid1.TextMatrix(grid1.Row, 12)
'        orders_online_invoices.sOrder_No = grid1.TextMatrix(grid1.Row, 0)
'        orders_online_invoices.Show 1
'        Exit Sub
'    End If
'End If
End Sub
Private Sub grid1_EnterCell()
'With grid1
'    If (.col = 2 Or .col = 3 Or .col = 4 Or .col = 5 Or .col = 12 Or .col = 13 Or .col = 19 + 1 Or .col = 22 + 1) And cBranch = "00" Then
'        .Editable = flexEDKbdMouse
'    ElseIf .col = 19 And cBranch = "00" Then
'        .Editable = flexEDKbdMouse
'    ElseIf (.col = 23 + 1 Or .col = 24 + 1) And cBranch <> "00" Then
'        .Editable = flexEDKbdMouse
'    Else
'        .Editable = flexEDNone
'    End If
'End With
End Sub
Private Function LoadConString_OnLine()
Dim cServerName As String, cUserId As String, cPassword As String, sCatalog_Online As String
If Not vpn Then
    cServerName = "junior-sql.database.windows.net"
    cUserId = decrypt("062F8C407C77", "dr")
    cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
Else
    cServerName = servername_vpn
    cUserId = userid_vpn
    cPassword = password_vpn
End If
sCatalog_Online = "JUNIOR_SHOP"
LoadConString_OnLine = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
                        & "catalog=" & sCatalog_Online & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Private Sub grid1_GotFocus()
grid1_EnterCell
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal col As Long)
With grid2
    If grid1.TextMatrix(grid1.Row, 16) = "" Then
        If .col = 0 Then
            Dim loctable As ADODB.Recordset
            Set loctable = ItemFind(Val(.TextMatrix(Row, 0)), con)
            If Not (loctable.EOF And loctable.BOF) Then
                grid2.TextMatrix(Row, 1) = loctable!BARCODE13 & ""
                grid2.TextMatrix(Row, 3) = loctable!DESCA & ""
                grid2.TextMatrix(Row, 7) = loctable!price & ""
                            
                If Val(.TextMatrix(Row, 9)) <> 0 Then
                    con.Execute " UPDATE FILE6_90 SET ITEM = " & addvalue(.TextMatrix(Row, 0)) & " , ITEM_NAME = " & addstring(.TextMatrix(Row, 3)) & " , PRICE = " & Val(.TextMatrix(Row, 7)) & " , SKU = " & addstring(.TextMatrix(Row, 1)) & " WHERE ID = " & Val(.TextMatrix(Row, 9))
                Else
                    grid2.TextMatrix(Row, 6) = 1
                    cStr1 = "INSERT INTO FILE6_90 ( doc_no , sku , item , item_name , price , quant )" & _
                                "VALUES( " & _
                                addstring(grid1.TextMatrix(grid1.Row, 0)) & "," & _
                                addstring(.TextMatrix(.Row, 1)) & "," & _
                                addvalue(.TextMatrix(.Row, 0)) & "," & _
                                addstring(.TextMatrix(.Row, 3)) & "," & _
                                Val(.TextMatrix(.Row, 7)) & "," & _
                                Val(.TextMatrix(.Row, 6)) & _
                                ")"
                    con.Execute cStr1
                End If
                MYLOAD2 (grid1.TextMatrix(grid1.Row, 0))
            Else
                MsgBox " »«—ﬂÊœ €Ì— ’ÕÌÕ "
                MYLOAD2 (grid1.TextMatrix(grid1.Row, 0))
            End If
        End If
        If .col = 6 Then
            con.Execute " UPDATE FILE6_90 SET QUANT = " & Val(.TextMatrix(Row, 6)) & " WHERE ID = " & Val(.TextMatrix(Row, 9))
        End If
        If .col = 7 Then
            con.Execute " UPDATE FILE6_90 SET PRICE = " & Val(.TextMatrix(Row, 7)) & " WHERE ID = " & Val(.TextMatrix(Row, 9))
        End If
    End If
End With
End Sub

Private Sub grid2_EnterCell()
With grid2
'    If (.Col = 0 Or .Col = 7 Or .Col = 6) And cBranch = "00" Then
'        .Editable = flexEDKbdMouse
'    Else
'        .Editable = flexEDNone
'    End If
End With
End Sub
Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
With grid2
    If .col = 0 And cBranch = "00" Then
        If KeyCode = 112 Then ItemsLookupAll Me, oSearchItem
    End If
    If KeyCode = 45 And cBranch = "00" And col = 0 Then
        grid2.AddItem grid2.Row
        grid2.TextMatrix(grid2.Row, 0) = ""
    End If
    If KeyCode = 46 And cBranch = "00" And grid1.TextMatrix(grid1.Row, 16) = "" Then
        If MsgBox("  Õ–›  «·’‰› „‰ «·ÿ·»Ì… ", vbYesNo + vbDefaultButton2) = vbYes Then
            con.Execute " DELETE FROM FILE6_90 WHERE ID = " & Val(grid2.TextMatrix(grid2.Row, 9))
            .RemoveItem grid2.Row
        End If
    End If
End With
End Sub
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = cmdStage.Name Then
    If oSearchStage.grid1.TextMatrix(oSearchStage.grid1.Row, 0) = "" Then
        cmdStage.Tag = ""
        cmdStage.Caption = cmdStage.TagVariant
    Else
        cmdStage.Tag = oSearchStage.grid1.TextMatrix(oSearchStage.grid1.Row, 0)
        cmdStage.Caption = oSearchStage.grid1.TextMatrix(oSearchStage.grid1.Row, 1)
    End If
    Unload oSearchStage
    myload
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub doprint_day(Optional pMan As String = "")
Dim aHeader(1)
'f Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim SourchTable As New ADODB.Recordset, i As Double
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

Dim cString As New ChilkatStringBuilder
cString.Append "SELECT  FILE6_90H.DOC_NO," & _
               "FILE1_10.MODELFACT0 ," & _
               "FILE6_90H.NAME," & _
               "FILE6_90H.PHONE ," & _
               "FILE1_10.MOSM," & _
               "FILE6_90H.DATE," & _
               "FILE6_90H.STREET," & _
               "FILE6_90H.CITY," & _
               "FILE6_90H.Shipping_City," & _
               "FILE6_90H.Payment_Method," & _
               "FILE6_90.ITEM," & _
               "FACT.DESCA AS FACT_NAME," & _
               "FILE1_10.[GROUP]," & _
               "FILE1_10.COLOR," & _
               "FILE1_10.SCAL," & _
               "FILE6_90.PRICE," & _
               "FILE6_90.QUANT," & _
               "FILE1_10.DESCA," & _
               "FILE6_90H.NOTES," & _
               "FILE0_40.DESCA AS STORENAME  " & _
               "FROM  FILE6_90H " & _
               "INNER JOIN FILE6_90 ON FILE6_90H.DOC_NO = FILE6_90.DOC_NO " & _
               "INNER JOIN FILE1_10 ON FILE6_90.ITEM = FILE1_10.ITEM " & _
               "INNER JOIN FACT ON FILE1_10.code = FACT.CODE " & _
               "LEFT JOIN FILE0_40 ON FILE0_40.CODE = FILE6_90H.STORE "
cString.Append "WHERE  SALES_DOC IS NULL"

If xdoc_no.text <> "" Then
    cString.Append " AND [DOC_NO] = " & MyParn(xdoc_no.text)
End If

If xpay.MatchedWithList Then
    cString.Append " AND [Payment_Method] = " & MyParn(xpay.text)
End If

If xStore.MatchedWithList Then
    cString.Append " AND [STORE] = " & MyParn(xStore.BoundText)
End If
If IsDate(xDate1.text) Then
    cString.Append " AND [DATE] >= " & DateSq(xDate1.text)
End If
If IsDate(xdate2.text) Then
    cString.Append " AND [DATE] <= " & DateSq(xdate2.text)
End If
If cBranch = "00" Then
    If xIssend(0).Value <> 0 Then
        cString.Append " AND STORE IS NULL  and DelOrder_Date is null "
    End If
    If xIssend(1).Value <> 0 Then
        cString.Append " AND STORE IS NOT  NULL AND DelOrder_Date IS NULL "
    End If
    If xIssend(3).Value <> 0 Then
        cString.Append " AND DelOrder_Date IS NOT  NULL "
    End If
Else
    cString.Append " AND STORE = " & MyParn(cBranchStore)
End If

If xSales(0).Value <> 0 Then
    cString.Append " AND SALES_DOC IS NULL "
End If
If xSales(1).Value <> 0 Then
    cString.Append " AND SALES_DOC IS NOT NULL "
End If

If xMan.MatchedWithList Then
    cString.Append " AND FILE6_90H.MAN = " & MyParn(xMan.BoundText)
End If

SourchTable.Open cString.GetAsString(), con, adOpenStatic, adLockReadOnly, adCmdText

nCount = SourchTable.RecordCount
prog1.Min = 0
prog1.Value = 0
If nCount > 0 Then
    SourchTable.MoveFirst
Else
    MsgBox " ·« ÌÊÃœ »Ì«‰«  "
    Exit Sub
End If
prog1.Max = nCount
With SourchTable
    Do Until .EOF
        i = i + 1
        prog1.Value = i
        temptable.AddNew
        temptable!STR4 = !DOC_NO
        temptable!Date1 = !Date
        temptable!STR7 = !Name
        temptable!STR6 = !phone
        
        temptable!str14 = !Shipping_City
        temptable!str15 = !Street
        temptable!str16 = !Payment_Method
        If xMan.MatchedWithList Then
            temptable!STR20 = "«·„‰œÊ» : " & xMan.text
        End If
        
        temptable!str1 = !Item
        temptable!str11 = !MOSM
        temptable!str12 = !fact_name
        temptable!str13 = !MODELFACT0
        temptable!str3 = !DESCA
        temptable!str9 = !NOTES
        temptable!str2 = !STORENAME
        
        temptable!str8 = !color
        temptable!STR5 = !SCAL
        
        temptable!VAL1 = !Quant
        temptable!VAL3 = !price
        'If cBranch <> "00" Then temptable!VAL4 = LastBalance(!Item, cBranchStore, con_MyShop)
        temptable!STR19 = " ÿ·»«  «Ê‰ ·«Ì‰  €Ì— „‰›–… " & xStore.text
        temptable.Update
    .MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\rep_order_online.rpt"
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub

Private Sub xPhone_GotFocus()
myGotFocus xphone
End Sub
Private Sub XPHONE_LostFocus()
myLostFocus xphone
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xdoc_no
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xdoc_no
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xPay_GotFocus()
myGotFocus xpay
End Sub
Private Sub xPay_LostFocus()
myLostFocus xpay
If Not xpay.MatchedWithList Then xpay.BoundText = ""
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid1.col = 12 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    If grid1.col = 13 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    If grid1.col = 19 + 1 Then Exit Sub
    If grid1.col = 21 + 1 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    If grid1.col = 24 + 1 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If col = 12 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    If col = 13 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    If col = 19 + 1 Then Exit Sub
    If col = 21 + 1 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    If col = 24 + 1 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If col < 2 Then
    grid1.col = 2
ElseIf col > 2 And col < 5 Then
    grid1.col = grid1.col + 1
ElseIf col < 12 Then
    grid1.col = 12
    grid1.ShowCell Row, grid1.col
ElseIf col < 13 Then
    grid1.col = 13
    grid1.ShowCell Row, grid1.col
'ElseIf Col < 19 And cBranch = "00" Then
'    grid1.Col = 19
'    grid1.ShowCell Row, grid1.Col
ElseIf col < 22 + 1 And cBranch = "00" Then
    grid1.col = 22 + 1
    grid1.ShowCell Row, grid1.col
ElseIf col < 23 + 1 And cBranch <> "00" Then
    grid1.col = 23 + 1
    grid1.Row = Row
    grid1.ShowCell Row, grid1.col
ElseIf col < 24 + 1 Then
    grid1.col = 24 + 1
    grid1.Row = Row
    grid1.ShowCell Row, grid1.col
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 2)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, col
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 4 Then
    If Val(grid1.EditText) = 0 Then
        Cancel = True
        Exit Sub
    End If
    
    If Val(xTotal.Caption) > 0 And Val(grid1.EditText) <= 0 Then
        Cancel = True
        Exit Sub
    ElseIf Val(xTotal.Caption) < 0 And Val(grid1.EditText) >= 0 Then
        Cancel = True
        Exit Sub
    End If
ElseIf col = 19 Then
    If (Not IsDate(grid1.EditText)) And Trim(grid1.EditText) <> "" Then
        Cancel = True
    Else
        grid1.EditText = myFormat_p(grid1.EditText)
    End If
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
End If
End Sub
Private Function getData()
If MsgBox("≈÷«›… «’‰«› „‰ „·› CSV", vbYesNo) <> vbYes Then
    Exit Function
End If
    
Dim cFileName As String
Common1.InitDir = ""
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


If cSv.NumRows < 2 Then Exit Function
Dim Tb As New ChilkatStringBuilder
Dim cString As New ChilkatStringBuilder


prog1.Visible = True
Dim sCaption As String
Dim sitem As String
sCaption = Me.Caption
For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & i & " „‰ " & cSv.NumRows - 1
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    If Trim(cSv.GetCellByName(i, "Payment ID")) <> "" Then
        If Trim(Replace(cSv.GetCellByName(i, "Name"), "#", "")) <> sName Then
            If myField("select sales_doc from file6_90H WHERE Payment_ID = " & MyParn(cSv.GetCellByName(i, "Payment ID")), con) & "" = "" Then
                sName = Trim(Replace(cSv.GetCellByName(i, "Name"), "#", ""))
                Tb.Append "SELECT " & addstring(sName) & " AS DOC_NO,"
                Tb.Append addstring(cSv.GetCellByName(i, "Email")) & " AS EMAIL,"
                Tb.Append addstring(cSv.GetCellByName(i, "Subtotal")) & " AS Subtotal,"
                Tb.Append addstring(cSv.GetCellByName(i, "Shipping")) & " AS Shipping,"
                Tb.Append addstring(cSv.GetCellByName(i, "Total")) & " AS Total,"
                Tb.Append addstring(cSv.GetCellByName(i, "Discount Code")) & " AS Discount_Code,"
                Tb.Append addstring(cSv.GetCellByName(i, "Discount Amount")) & " AS Discount,"
                Tb.Append addDate(Mid(cSv.GetCellByName(i, "Created at"), 1, 10)) & " AS [date],"
                Tb.Append addstring(cSv.GetCellByName(i, "Shipping Name")) & " AS NAME,"
                Tb.Append addstring(cSv.GetCellByName(i, "Shipping Street")) & " AS STREET,"
                Tb.Append addstring(cSv.GetCellByName(i, "Shipping Address1")) & " AS ADDRESS1,"
                Tb.Append addstring(cSv.GetCellByName(i, "Shipping Address2")) & " AS ADDRESS2,"
                Tb.Append addstring(cSv.GetCellByName(i, "Shipping City")) & " AS CITY,"
                Tb.Append addstring(cSv.GetCellByName(i, "Shipping Phone")) & " AS PHONE,"
                Tb.Append addstring(cSv.GetCellByName(i, "Payment Method")) & " AS Payment_Method,"
                Tb.Append addstring(cSv.GetCellByName(i, "Shipping Province Name")) & " AS Shipping_City,"
                Tb.Append addstring(cSv.GetCellByName(i, "Notes")) & " AS Notes_Order,"
                Tb.Append addstring(cSv.GetCellByName(i, "Payment ID")) & " AS Payment_ID"
                Tb.Append " UNION ALL "
            End If
        End If
    End If
Next
Tb.Shorten 11

If Tb.length = 0 Then Exit Function
Dim loctable As New ADODB.Recordset
Set loctable = mycmd(Tb.GetAsString, con)

nRecordCount = loctable.RecordCount

Dim aInsert As Variant
con.BeginTrans
On Error GoTo myerror
Dim nAffect As Long, nAffectAll As Long
i = 0
prog1.Value = 0
prog1.Visible = True
Do Until loctable.EOF
    i = i + 1
    Me.Caption = sCaption & " - " & "”Ã· " & i & " „‰ " & nRecordCount
    prog1.Value = Round(i / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "ITEM", addvalue(loctable!Item))
    aInsert = AddFlag(aInsert, "PRICE", loctable!price)
    aInsert = AddFlag(aInsert, "QUANT", loctable!Quant)
    aInsert = AddFlag(aInsert, "MODEL", addstring(loctable!MODEL))
    If IsEmpty(myField("SELECT DOC_NO FROM " & cFile & " WHERE DOC_NO = " & MyParn(xdoc_no.text) & " AND ITEM = " & addvalue(loctable!Item), con)) Then
        aInsert = AddFlag(aInsert, "DOC_NO", addstring(xdoc_no.text))
        con.Execute addInsert(aInsert, cFile), nAffect
        nAffect = 1
    Else
        con.Execute addUpdate(aInsert, cFile, "DOC_NO = " & MyParn(xdoc_no.text) & " AND ITEM = " & addvalue(loctable!Item)), nAffect
    End If
    nAffectAll = nAffectAll + nAffect
    loctable.MoveNext
Loop
con.CommitTrans
Me.Caption = sCaption
prog1.Visible = False
MsgBox " „  «÷«›… " & nAffectAll & " ’‰› »‰Ã«Õ"
getData = True
Finaly:
loctable.Close
Set loctable = Nothing
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
GoTo Finaly
End Function


