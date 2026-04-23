VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form OrderOnline_New 
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   32
      Top             =   10485
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   675
      Width           =   5100
      Begin Threed.SSCommand cmdExcel 
         Height          =   510
         Left            =   90
         TabIndex        =   34
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
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
         Picture         =   "OrderOnline_New3.frx":0000
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label XBRANCH 
         Alignment       =   2  'Center
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   225
         Width           =   3390
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1860
      Left            =   5310
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   45
      Width           =   13425
      Begin Threed.SSFrame SSFrame1 
         Height          =   465
         Left            =   1665
         TabIndex        =   36
         Top             =   1350
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   820
         _Version        =   196610
         Begin VB.CheckBox RET_NOTSHIP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "ÿ·»Ì«  „—›Ê÷… Ê ·„ Ì „ «” ·«„Â« „‰ «·‘Õ‰"
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
            Left            =   1170
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   45
            Width           =   3705
         End
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
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   990
         Width           =   2130
      End
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   135
         Picture         =   "OrderOnline_New3.frx":252B
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1305
         Width           =   1455
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
         Left            =   10845
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   630
         Width           =   1365
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   135
         Picture         =   "OrderOnline_New3.frx":4955
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "⁄—÷"
         Top             =   810
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Height          =   600
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   720
         Width           =   5010
         Begin VB.OptionButton XSALES 
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
            Height          =   375
            Index           =   2
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   180
            Width           =   780
         End
         Begin VB.OptionButton XSALES 
            Alignment       =   1  'Right Justify
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
            Height          =   375
            Index           =   1
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   180
            Width           =   780
         End
         Begin VB.OptionButton XSALES 
            Alignment       =   1  'Right Justify
            Caption         =   "€Ì— „»«⁄…"
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
            Index           =   0
            Left            =   3690
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   180
            Value           =   -1  'True
            Width           =   1140
         End
      End
      Begin VB.Frame Frame2 
         Height          =   600
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   135
         Width           =   5010
         Begin VB.OptionButton XISSEND 
            Alignment       =   1  'Right Justify
            Caption         =   "ÿ·»Ì«  „·€«…"
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
            Index           =   3
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   180
            Width           =   1500
         End
         Begin VB.OptionButton XISSEND 
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
            Height          =   285
            Index           =   2
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton XISSEND 
            Alignment       =   1  'Right Justify
            Caption         =   "„—Õ·… "
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
            Left            =   2430
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   180
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton XISSEND 
            Alignment       =   1  'Right Justify
            Caption         =   "·„  —Õ· ··›—⁄"
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
            Index           =   0
            Left            =   3375
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   180
            Width           =   1500
         End
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
         Left            =   7830
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   270
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
         Left            =   10845
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xpay 
         Height          =   330
         Left            =   6750
         TabIndex        =   8
         Top             =   630
         Width           =   2625
         _ExtentX        =   4630
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
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   330
         Left            =   6750
         TabIndex        =   5
         Top             =   990
         Width           =   2625
         _ExtentX        =   4630
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
         Left            =   6750
         TabIndex        =   6
         Top             =   1350
         Width           =   5460
         _ExtentX        =   9631
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
         Left            =   12330
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1395
         Width           =   660
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
         Left            =   12330
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1035
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "«·›—⁄ "
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
         Left            =   9450
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1035
         Width           =   525
      End
      Begin VB.Label Label7 
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
         Left            =   9450
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   720
         Width           =   660
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
         Left            =   12330
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   675
         Width           =   690
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
         Left            =   9450
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   315
         Width           =   750
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
         Left            =   12330
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   660
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   10725
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin VSFlex7Ctl.VSFlexGrid GRID2 
      Bindings        =   "OrderOnline_New3.frx":6E47
      Height          =   3285
      Left            =   45
      TabIndex        =   19
      Top             =   7515
      Width           =   18690
      _cx             =   32967
      _cy             =   5794
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16776960
      ForeColorSel    =   64
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
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
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   5055
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "⁄—÷ —’Ìœ ﬂ· «·›—Ê⁄"
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
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   2175
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "OrderOnline_New3.frx":6E5B
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_addexel 
         Height          =   480
         Left            =   1260
         TabIndex        =   25
         Top             =   135
         Visible         =   0   'False
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   847
         _Version        =   196610
         Font3D          =   3
         CaptionStyle    =   1
         ForeColor       =   64
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«÷«›… „‰ «ﬂ”Ì·"
         ButtonStyle     =   2
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "OrderOnline_New3.frx":92C7
      Height          =   4860
      Left            =   45
      TabIndex        =   9
      Top             =   1935
      Width           =   18690
      _cx             =   32967
      _cy             =   8572
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16776960
      ForeColorSel    =   64
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   27
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
   Begin VB.Frame Frame6 
      Height          =   690
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   6795
      Width           =   18690
      Begin Threed.SSCommand CMD_SEND 
         Height          =   465
         Left            =   90
         TabIndex        =   43
         Top             =   135
         Width           =   2310
         _ExtentX        =   4075
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
         Picture         =   "OrderOnline_New3.frx":92DB
         Caption         =   " —ÕÌ· ·›« Ê—… „»Ì⁄« "
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "OrderOnline_New3.frx":B9FE
      End
      Begin VB.Label XUSER 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   420
         Left            =   14760
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   180
         Width           =   3840
      End
      Begin VB.Label XTIME 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   420
         Left            =   11835
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   180
         Width           =   2895
      End
      Begin VB.Label XNOTES 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   420
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   180
         Width           =   9375
      End
   End
   Begin Threed.SSCommand cmdCSV 
      Height          =   510
      Left            =   45
      TabIndex        =   44
      Top             =   1395
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   900
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
      Caption         =   "CSV ”Õ» „‰ „·›"
      TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
      ButtonStyle     =   3
   End
End
Attribute VB_Name = "OrderOnline_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3, lIsBoxOnline As Boolean
Dim cString As String
Dim CLIST As String
Dim oSearchMan As New Search_abd
Dim cStr1 As String, cStr2 As String
Dim con As New adodb.Connection
Dim con_MyShop As New adodb.Connection, oSearchItem As New Search3
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
Set oOnlineCheck.myform = Me
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

Private Sub cmdBank_Click()

End Sub

Private Sub cmdCSV_Click()
Set onlineCSVfrm.myform = Me
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
Private Sub Form_Load()
'    On Error GoTo myerror
    
    'openCon con, LoadConString_OnLine
    
    openCon con
    'If cBranch <> "00" Then openCon con_MyShop
    
    Set grid1.DataSource = DATA1
    Set grid2.DataSource = data2
    
    Frame2.Visible = (cBranch = "00")
    Check1.Visible = (cBranch = "00")
    'cmd_addexel.Visible = (cBranch = "00")
    cmdCSV.Enabled = (cBranch = "00")
    
    Set DATA3.Recordset = myRecordSet("SELECT Payment_Method FROM file6_90h GROUP BY Payment_Method ", con)
    Set xPay.RowSource = DATA3
    xPay.ListField = "Payment_Method"
    xPay.BoundColumn = "Payment_Method"
    
    Set data4.Recordset = myRecordSet("SELECT CODE , DESCA FROM FILE0_40 WHERE online =  1 ORDER BY CODE ", con)
    Set xStore.RowSource = data4
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
        
    Set data5.Recordset = myCmd("SELECT CODE,DESCA FROM FILE6_25 WHERE FILE6_25.BRANCH IN (SELECT FILE0_40.BRANCH FROM FILE0_40 WHERE FILE0_40.online =  1) AND FILE6_25.ISSTOP = 0 ORDER BY CODE ", con)
    Set xMan.RowSource = data5
    xMan.ListField = "Desca"
    xMan.BoundColumn = "Code"
        
    If GetBoolean("SELECT ONLINE FROM FILE0_40 WHERE CODE = " & MyParn(cBranchStore), con) Then
        lIsBoxOnline = True
        If cBranchBox <> "" Then
            lIsBoxOnline = GetBoolean("SELECT ISBOXONLINE FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con)
        End If
    End If
    
    'CMD_SEND.Visible = (cBranch <> "00") And lIsBoxOnline And cManBox <> "" And Not lSupperVisor
    CMD_SEND.Enabled = (cBranch <> "00") And lIsBoxOnline And cManBox <> "" And Not lSupperVisor
    'If cBranch <> "00" Then grid2.Width = grid2.Width - 3000
    fixGrd
    XBRANCH.Caption = GetDesca("SELECT DESCA FROM BRANCH WHERE CODE = " & MyParn(cBranch), con)
    CLIST = StrListA("SELECT CODE , DESCA FROM FILE0_40 WHERE ONLINE = 1 ORDER BY CODE ", con)
    LoadText Me
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Public Sub myload()
On Error GoTo myerror
Dim i As Double
Dim cString  As String, cStr2 As String
Dim cWhere As String
Dim cF1 As String, cF2 As String, cF3 As String, cF4 As String, cF5 As String
Dim cSubTotal As String


With grid1
    cSubTotal = "(SELECT SUM(TOTAL) FROM FILE6_90 WHERE FILE6_90.DOC_NO = FILE6_90H.DOC_NO )    "
    cTotal = "(SELECT SUM(TOTAL) FROM FILE6_90 WHERE FILE6_90.DOC_NO = FILE6_90H.DOC_NO )  - DISCOUNT   "
    
    cString = " SELECT DOC_NO," & _
              " DATE," & _
              " NAME," & _
              " PHONE," & _
              "Shipping_City," & _
              "STREET, " & _
              cSubTotal & "," & _
              "DISCOUNT, " & _
              "DISCOUNT_CODE," & _
              "SHIPPING, " & _
              cTotal & "," & _
              "Payment_Method," & _
              "STORE," & _
              "SEND_USER," & _
              "SEND_TIME," & _
              "NOTES_ORDER ," & _
              "SUBSTRING(SALES_DOC,7,20) ," & _
              "SALES_DATE," & _
              "SHIP_NO," & _
              "SHIP_DATE," & _
              " NOTES,DelOrder_Date,DelOrder_Date2,Note_main,Note_branch," & _
              " FILE6_25.DESCA,FILE6_90H.MAN" & _
              " FROM FILE6_90H LEFT JOIN FILE6_25 ON FILE6_90H.MAN = FILE6_25.CODE"
    If xDoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " [DOC_NO] = " & MyParn(xDoc_no.text)
    If xphone.text <> "" Then cWhere = cWhere & Tr(cWhere) & " [phone] = " & MyParn(xphone.text)
    If xPay.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " [Payment_Method] = " & MyParn(xPay.text)
    If xStore.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " [STORE] = " & MyParn(xStore.BoundText)
    If xMan.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & "MAN = " & MyParn(xMan.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " [DATE] >= " & DateSq(xDate1.text)
    If IsDate(xDate2.text) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.text)
    
    If cBranch = "00" Then
        If xIssend(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " STORE IS NULL  and DelOrder_Date is null "
        If xIssend(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " STORE IS NOT  NULL AND DelOrder_Date IS NULL "
        If xIssend(3).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " DelOrder_Date IS NOT  NULL "
    Else
        cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(cBranchStore)
    End If
    
    If xsales(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " SALES_DOC IS NULL "
    If xsales(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " SALES_DOC IS NOT NULL "
    
    If RET_NOTSHIP.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " DelOrder_Date IS NOT NULL AND DelOrder_Date2 IS NULL "
    
    If cWhere <> "" Then cString = cString & " WHERE " & cWhere
    cString = cString & " ORDER BY DOC_NO "
    
    
    Set DATA1.Recordset = myCmd(cString, con)

End With
fixGrd
grid1.Cell(flexcpAlignment, 0, 0, grid1.Rows - 1, grid1.Cols - 1) = 7
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrd()
With grid1
    .RowHeight(0) = 1000
    .WordWrap = True
    '.Cols = 22
    .FrozenCols = 3
    
    .TextMatrix(0, 0) = "—ﬁ„"
    .TextMatrix(0, 1) = "«· «—ÌÕ"
    .TextMatrix(0, 2) = "«·≈”„"
    .TextMatrix(0, 3) = " ·Ì›Ê‰"
    .TextMatrix(0, 4) = "«·„œÌ‰…"
    .TextMatrix(0, 5) = "⁄‰Ê«‰"
    
    .TextMatrix(0, 6) = "≈Ã„«·Ï «’‰«›"
    .TextMatrix(0, 7) = "Œ’„"
    .TextMatrix(0, 8) = "ﬂÊœ Œ’„"
    .TextMatrix(0, 9) = "«·‘Õ‰"
    
    .TextMatrix(0, 10) = "«·≈Ã„«·Ï"
    .TextMatrix(0, 11) = "ÿ—Ìﬁ… «·œ›⁄"
    
    .TextMatrix(0, 12) = "«·›—⁄"
    .TextMatrix(0, 13) = "«ŸÂ«— ··›—⁄"
    .TextMatrix(0, 14) = "SEND TIME "
    .TextMatrix(0, 15) = "„·«ÕŸ«  ÿ·»Ì…"
    .ColComboList(12) = CLIST
    .ColComboList(13) = "..."
    
    .ColComboList(19 + 1) = "..."
    .ColComboList(21 + 1) = "..."
    .ColComboList(24 + 1) = "..."
    
    .TextMatrix(0, 16) = "»Ê‰ „»Ì⁄« "
    .TextMatrix(0, 17) = " «—ÌŒ"
    
    .TextMatrix(0, 18) = "»Ê·Ì’… ‘Õ‰"
    .TextMatrix(0, 19) = " «—ÌŒ «·‘Õ‰"
    
    .ColHidden(19) = True
    .TextMatrix(0, 19 + 1) = "”»» «·≈·€«¡"
    .TextMatrix(0, 20 + 1) = "«·€«¡ «·ÿ·»"
    
    .TextMatrix(0, 21 + 1) = " «—ÌŒ «” ·«„ «·„— Ã⁄"
    .TextMatrix(0, 22 + 1) = "„·«ÕŸ«  «·—∆Ì”Ì"
    .TextMatrix(0, 23 + 1) = "„·«ÕŸ«  «·›—Ê⁄"
    .TextMatrix(0, 24 + 1) = "«·„‰œÊ»"
    .TextMatrix(0, 25 + 1) = "«·„‰œÊ»"

    
    .ColDataType(1) = flexDTDate
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(9) = flexDTDouble
    .ColDataType(10) = flexDTDouble
    
    .ColHidden(14) = True
    .ColHidden(25 + 1) = True
    
    .ColHidden(19 + 1) = cBranch <> "00"
    .ColHidden(20 + 1) = cBranch <> "00"
                
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
    .ColWidth(16) = 1000
    .ColWidth(15) = 2000
    .ColWidth(17) = 1300
    .ColWidth(18) = 2000
    .ColWidth(19 + 1) = 1300
    .ColWidth(20 + 1) = 1300
    .ColWidth(21 + 1) = 1300
    .ColWidth(22 + 1) = 2500
    .ColWidth(23 + 1) = 2500
    .ColWidth(24 + 1) = 3000
    .ColDataType(21 + 1) = flexDTDate
    
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
    closeCon con
    'If cBranch <> "00" Then closeCon con_MyShop
    SaveText Me, , Array(xDate1.Name, xDate2.Name)
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
                        Dim loctable As adodb.Recordset
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
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
    If Col = 23 + 1 Then
        con.Execute " UPDATE FILE6_90H SET [NOTE_BRANCH] = " & addstring(.TextMatrix(Row, Col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
    End If

    'If cBranch <> "00" Then Exit Sub
    Select Case Col
        Case 2
            con.Execute " UPDATE FILE6_90H SET [NAME] = " & addstring(.TextMatrix(Row, Col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        Case 3
            con.Execute " UPDATE FILE6_90H SET [PHONE] = " & addstring(.TextMatrix(Row, Col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        Case 4
            con.Execute " UPDATE FILE6_90H SET [CITY] = " & addstring(.TextMatrix(Row, Col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        Case 5
            con.Execute " UPDATE FILE6_90H SET [ADDRESS1] = " & addstring(.TextMatrix(Row, Col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        Case 19
            con.Execute " UPDATE FILE6_90H SET [SHIP_DATE] = " & addDate(.TextMatrix(Row, Col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        Case 22 + 1
            con.Execute " UPDATE FILE6_90H SET [NOTE_MAIN] = " & addstring(.TextMatrix(Row, Col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
    End Select
End With
End Sub
Private Sub grid1_DblClick()
    MYLOAD2 grid1.TextMatrix(grid1.Row, 0)
End Sub
Private Sub MYLOAD2(pDoc_No)
On Error GoTo myerror
Dim StoreTable As New adodb.Recordset
Dim cBalNotShip As String
If cBranch = "00" Then
    If Check1.Value = 0 Then
        StoreTable.Open " SELECT * FROM FILE0_40 WHERE ONLINE = 1 ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
    Else
        StoreTable.Open " SELECT * FROM FILE0_40 WHERE ISONEST2 = 0 AND ISSTOP = 0 AND ISONEST = 0 AND BRANCH IS NOT NULL  AND CODE <> '000' ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
    End If
Else
    StoreTable.Open " SELECT * FROM FILE0_40 WHERE CODE = " & MyParn(cBranchStore), con, adOpenStatic, adLockReadOnly, adCmdText
End If
Dim cString  As String, cStrStore As String
With StoreTable
    cStrStore = " , (SELECT SUM(BALITEM) FROM [BAL_ITEM_ORDERONLINE] WHERE [BAL_ITEM_ORDERONLINE].ITEM = FILE6_90.ITEM AND [BAL_ITEM_ORDERONLINE].STORE  = " & MyParn(StoreTable!CODE) & " ) AS 'Total All Order' "
    Do While Not .EOF
        cStrStore = cStrStore & " , (SELECT SUM([IN]-[OUT]) FROM FILE1_11 WHERE FILE1_11.ITEM = FILE6_90.ITEM AND FILE1_11.STORE  = " & MyParn(StoreTable!CODE) & " ) AS '" & StoreTable!DESCA & "'"
        .MoveNext
    Loop
End With

grid2.Rows = 2
grid2.FixedRows = 2

cBalNotShip = " ,(SELECT SUM(QUANT ) FROM Q_BAL_NOTSHIP WHERE Q_BAL_NOTSHIP.ITEM = FILE6_90.ITEM)"
With grid1
'                               0           1               2           3                                                                                   4               5               6            7                  8                   9           10          11
    cString = " SELECT    FILE6_90.ITEM, FILE6_90.SKU, FACT.DESCA , CASE WHEN FILE1_10.desca IS NULL THEN FILE6_90.ITEM_NAME ELSE FILE1_10.desca END , FILE1_10.COLOR, FILE1_10.SCAL, FILE6_90.QUANT, FILE6_90.PRICE, file6_90.ITEM_NAME, FILE6_90.ID " & cStrStore & cBalNotShip & " FROM            FILE6_90 LEFT  JOIN FILE1_10 ON FILE6_90.ITEM = FILE1_10.ITEM left  JOIN  FACT ON FILE1_10.code = FACT.CODE WHERE DOC_NO =  " & MyParn(pDoc_No)
    Set data2.Recordset = myRecordSet(cString, con)
End With
Fixgrd2
grid2.Cell(flexcpAlignment, 0, 0, grid2.Rows - 1, grid2.Cols - 1) = 7
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub Fixgrd2()
Dim LShowColAll As Boolean, LShowCol As Boolean
With grid2
        
    .FixedRows = 2
    .RowHeight(0) = 700
    .RowHidden(1) = True
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "SKU"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "«·’‰›"
    .TextMatrix(0, 4) = "«··Ê‰"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«·ﬂ„Ì…"
    .TextMatrix(0, 7) = "«·”⁄—"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 4000
    .ColWidth(4) = 1500
    .ColWidth(5) = 800
    .ColWidth(6) = 600
    .ColWidth(7) = 700
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(.Cols - 1) = 1000
    .TextMatrix(0, .Cols - 1) = "—’Ìœ ‘Õ‰ €Ì— „” ·„"
    LShowColAll = False
    For nCol = 11 To .Cols - 1
        If Check1.Value <> 0 Then
            .TextMatrix(1, nCol) = IIf(GetBoolean("select online from file0_40 where desca = " & MyParn(.TextMatrix(0, nCol)), con), 1, 0)
        Else
            .TextMatrix(1, nCol) = 1
        End If
        If cBranch = "00" Then
            .ColWidth(nCol) = 800
        Else
            .ColWidth(nCol) = 1200
        End If
        LShowCol = True
        For nRow = 2 To .Rows - 1
            If .ValueMatrix(nRow, nCol) = 0 Then
                LShowCol = False
            End If
        Next nRow
'        .ColHidden(nCol) = Not LShowCol Or .TextMatrix(1, nCol) = "0"
        If LShowCol And .TextMatrix(1, nCol) = "1" Then LShowColAll = True
    Next nCol
    If Not LShowColAll Then
        For nCol = 11 To .Cols - 1
            .ColHidden(nCol) = False
        Next nCol
    End If
    For nCol = 11 To .Cols - 1
        If Not .ColHidden(nCol) Then
            LShowCol = True
            For nRow = 2 To .Rows - 1
                If .ValueMatrix(nRow, nCol) = 0 Then
                    LShowCol = False
                End If
            Next nRow
            If LShowCol Then
                .Cell(flexcpBackColor, 0, nCol, .Rows - 1, nCol) = vbGreen
            Else
                .Cell(flexcpBackColor, 0, nCol, .Rows - 1, nCol) = &H8080FF
            End If
        End If
    Next nCol
    
'    If cBranch <> "00" Then
'        .Cols = .Cols + 1
'        .ColWidth(.Cols - 1) = 1200
'        .TextMatrix(0, .Cols - 1) = "—’Ìœ «·›—⁄"
'        For nRow = 2 To .Rows - 1
'            .TextMatrix(nRow, .Cols - 1) = LastBalance(grid2.TextMatrix(nRow, 0), cBranchStore, con_MyShop)
'        Next nRow
'    End If
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .ColDataType(0) = flexDTDate
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTCount, -1, 0, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 6, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 8, "#", vbYellow, , True, ""
    .Select .Rows - 1, 0
    .ShowCell .Rows - 1, 0
    End With
End Sub
Private Sub grid1_EnterCell()
With grid1
    CMD_SEND.Visible = (cBranch <> "00") And lIsBoxOnline And cManBox <> "" And (Not lSupperVisor) And .TextMatrix(.Row, 16) = ""
    If (.Col = 2 Or .Col = 3 Or .Col = 4 Or .Col = 5 Or .Col = 12 Or .Col = 13 Or .Col = 19 + 1 Or .Col = 22 + 1) And cBranch = "00" Then
        .Editable = flexEDKbdMouse
    ElseIf .Col = 19 And cBranch = "00" Then
        .Editable = flexEDKbdMouse
    ElseIf (.Col = 23 + 1 Or .Col = 24 + 1) And cBranch <> "00" Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
    xNotes.Caption = .TextMatrix(.Row, 15)
    xUser.Caption = .TextMatrix(.Row, 13)
    xtime.Caption = .TextMatrix(.Row, 14)
End With
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim cNoteS As String
With grid1
    If Col = 13 And .TextMatrix(Row, 19 + 1) = "" Then
        If .TextMatrix(Row, 12) <> "" Then
            If MsgBox(" «ŸÂ«— ÿ·»Ì…  «Ê‰ ·«Ì‰ ··›—⁄ ", vbYesNo + vbDefaultButton2) = vbYes Then
                cNoteS = InputBox("", " „·«ÕŸ«  Œ«’… »«·ÿ·» ··›—⁄ ")
                
                .TextMatrix(Row, 13) = cusername
                .TextMatrix(Row, 14) = Date
                .TextMatrix(Row, 15) = cusername
                .TextMatrix(Row, 15) = cNoteS
                
                If Trim(.TextMatrix(Row, 12)) = "" Then
                    con.Execute "UPDATE FILE6_90H SET STORE = NULL , SEND_USER = NULL , SEND_TIME = NULL , NOTES = NULL  WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
                Else
                    con.Execute "UPDATE FILE6_90H SET STORE = " & addstring(.TextMatrix(Row, 12)) & " , SEND_USER = " & addstring(cusername) & " , SEND_TIME = GETDATE()  , NOTES= " & addstring(cNoteS) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
                End If
                xNotes.Caption = cNoteS
                xUser.Caption = cusername
                xtime.Caption = Date & " " & Time
            
            End If
        Else
            If MsgBox(" «·€«¡  —ÕÌ· «·ÿ·»Ì…  ··›—⁄ ", vbYesNo + vbDefaultButton2) = vbYes Then
                con.Execute " UPDATE FILE6_90H SET STORE = NULL , SEND_USER = NULL , SEND_TIME = NULL , NOTES = NULL  WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
                xNotes.Caption = ""
                xUser.Caption = ""
                xtime.Caption = ""
            
                .TextMatrix(Row, 13) = ""
                .TextMatrix(Row, 14) = ""
                .TextMatrix(Row, 15) = ""
                .TextMatrix(Row, 15) = ""
            
            End If
        End If
    ElseIf Col = 19 + 1 And .TextMatrix(Row, 12) <> "" Then
        If MsgBox(" «·€«¡ «·ÿ·»Ì… „‰ «·⁄„Ì· ", vbYesNo + vbDefaultButton2) = vbYes Then
            cNoteS = InputBox("", " ”»» «·≈·€«¡ ")
            .TextMatrix(Row, 19 + 1) = cNoteS
            .TextMatrix(Row, 20 + 1) = Date
            con.Execute " UPDATE FILE6_90H SET store = null, DelOrder_Date = " & DateSq(Date) & "  , NOTES = " & addstring(cNoteS) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
            .TextMatrix(Row, 12) = ""
        Else
            cNoteS = ""
            .TextMatrix(Row, 19 + 1) = ""
            .TextMatrix(Row, 20 + 1) = ""
            con.Execute " UPDATE FILE6_90H SET DelOrder_Date = null , NOTES = null WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        End If
    ElseIf Col = 21 + 1 And IsDate(.TextMatrix(Row, 20 + 1)) Then
        If MsgBox("  «—ÌŒ «” ·«„ «·„— Ã⁄ „‰ ‘—ﬂ… «·‘Õ‰ ", vbYesNo + vbDefaultButton2) = vbYes Then
            DDate2 = InputBox("", " «· «—ÌÕ ")
            .TextMatrix(Row, 21 + 1) = DDate2
            con.Execute " UPDATE FILE6_90H SET DelOrder_Date2 = " & DateSq(sDate2) & "  WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        Else
            con.Execute " UPDATE FILE6_90H SET DelOrder_Date2 = null  WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        End If
    ElseIf Col = 24 + 1 And grid1.TextMatrix(Row, 12) <> "" Then
        SalesManLookup Me, oSearchMan, "FILE6_25.BRANCH IN (SELECT BRANCH FROM FILE0_40 WHERE STORE = " & MyParn(grid1.TextMatrix(Row, 12)) & ") AND FILE6_25.ISSTOP = 0", , IIf(grid1.TextMatrix(grid1.Row, 24 + 1) = "", "", "»œÊ‰ »«∆⁄Ì‰")
    End If
End With
End Sub
Private Function LoadConString_OnLine()
Dim cServerName As String, cUserId As String, cPassword As String, sCatalog_Online As String
    cServerName = "junior-sql.database.windows.net"
    cUserId = decrypt("062F8C407C77", "dr")
    cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
    sCatalog_Online = "JUNIOR_SHOP"
    LoadConString_OnLine = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
        & "catalog=" & sCatalog_Online & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Private Sub grid1_GotFocus()
grid1_EnterCell
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid2
    If grid1.TextMatrix(grid1.Row, 16) = "" Then
        If .Col = 0 Then
            Dim loctable As adodb.Recordset
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
        If .Col = 6 Then
            con.Execute " UPDATE FILE6_90 SET QUANT = " & Val(.TextMatrix(Row, 6)) & " WHERE ID = " & Val(.TextMatrix(Row, 9))
        End If
        If .Col = 7 Then
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
    If .Col = 0 And cBranch = "00" Then
        If KeyCode = 112 Then ItemsLookupAll Me, oSearchItem
    End If
    If KeyCode = 45 And cBranch = "00" And Col = 0 Then
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
If ActiveControl.Name = grid2.Name Then
    grid2.TextMatrix(grid2.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    grid2_AfterEdit grid2.Row, 0
    Unload oSearchItem
ElseIf ActiveControl.Name = grid1.Name And grid1.Col = 24 + 1 Then
    If oSearchMan.grid1.TextMatrix(oSearchMan.grid1.Row, 0) = "" Then
        grid1.TextMatrix(grid1.Row, 24 + 1) = ""
        grid1.TextMatrix(grid1.Row, 25 + 1) = ""
        Unload oSearchMan
    Else
        grid1.TextMatrix(grid1.Row, 24 + 1) = oSearchMan.grid1.TextMatrix(oSearchMan.grid1.Row, 1)
        grid1.TextMatrix(grid1.Row, 25 + 1) = oSearchMan.grid1.TextMatrix(oSearchMan.grid1.Row, 0)
        con.Execute " UPDATE FILE6_90H SET [Man] = " & addstring(grid1.TextMatrix(grid1.Row, 25 + 1)) & " WHERE DOC_NO = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
        oSearchMan.Hide
        If grid1.Row < grid1.Rows - 1 Then
            CellPos 13, grid1.Row + 1, 23 + 1
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub doprint_day(Optional pMan As String = "")
Dim aHeader(1)
'f Not MYVALID Then Exit Sub
Dim temptable As New adodb.Recordset
Dim SourchTable As New adodb.Recordset, i As Double
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

If xDoc_no.text <> "" Then
    cString.Append " AND [DOC_NO] = " & MyParn(xDoc_no.text)
End If

If xPay.MatchedWithList Then
    cString.Append " AND [Payment_Method] = " & MyParn(xPay.text)
End If

If xStore.MatchedWithList Then
    cString.Append " AND [STORE] = " & MyParn(xStore.BoundText)
End If
If IsDate(xDate1.text) Then
    cString.Append " AND [DATE] >= " & DateSq(xDate1.text)
End If
If IsDate(xDate2.text) Then
    cString.Append " AND [DATE] <= " & DateSq(xDate2.text)
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

If xsales(0).Value <> 0 Then
    cString.Append " AND SALES_DOC IS NULL "
End If
If xsales(1).Value <> 0 Then
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
        temptable!str4 = !doc_no
        temptable!Date1 = !Date
        temptable!STR7 = !Name
        temptable!str6 = !phone
        
        temptable!str14 = !Shipping_City
        temptable!str15 = !Street
        temptable!str16 = !Payment_Method
        If xMan.MatchedWithList Then
            temptable!STR20 = "«·„‰œÊ» : " & xMan.text
        End If
        
        temptable!str1 = !Item
        temptable!str11 = !MOSM
        temptable!str12 = !fact_name
        temptable!str13 = !modelfact0
        temptable!str3 = !DESCA
        temptable!str9 = !NOTES
        temptable!str2 = !STORENAME
        
        temptable!str8 = !color
        temptable!str5 = !SCAL
        
        temptable!VAL1 = !Quant
        temptable!val3 = !price
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
myGotFocus xDoc_no
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_no
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xPay_GotFocus()
myGotFocus xPay
End Sub
Private Sub xPay_LostFocus()
myLostFocus xPay
If Not xPay.MatchedWithList Then xPay.BoundText = ""
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
    If grid1.Col = 12 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    If grid1.Col = 13 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    If grid1.Col = 19 + 1 Then Exit Sub
    If grid1.Col = 21 + 1 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    If grid1.Col = 24 + 1 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If Col = 12 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    If Col = 13 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    If Col = 19 + 1 Then Exit Sub
    If Col = 21 + 1 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    If Col = 24 + 1 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < 2 Then
    grid1.Col = 2
ElseIf Col > 2 And Col < 5 Then
    grid1.Col = grid1.Col + 1
ElseIf Col < 12 Then
    grid1.Col = 12
    grid1.ShowCell Row, grid1.Col
ElseIf Col < 13 Then
    grid1.Col = 13
    grid1.ShowCell Row, grid1.Col
'ElseIf Col < 19 And cBranch = "00" Then
'    grid1.Col = 19
'    grid1.ShowCell Row, grid1.Col
ElseIf Col < 22 + 1 And cBranch = "00" Then
    grid1.Col = 22 + 1
    grid1.ShowCell Row, grid1.Col
ElseIf Col < 23 + 1 And cBranch <> "00" Then
    grid1.Col = 23 + 1
    grid1.Row = Row
    grid1.ShowCell Row, grid1.Col
ElseIf Col < 24 + 1 Then
    grid1.Col = 24 + 1
    grid1.Row = Row
    grid1.ShowCell Row, grid1.Col
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 2)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 4 Then
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
ElseIf Col = 19 Then
    If (Not IsDate(grid1.EditText)) And Trim(grid1.EditText) <> "" Then
        Cancel = True
    Else
        grid1.EditText = myFormat_p(grid1.EditText)
    End If
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
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

If Tb.Length = 0 Then Exit Function
Dim loctable As New adodb.Recordset
Set loctable = myCmd(Tb.GetAsString, con)

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
    If IsEmpty(myField("SELECT DOC_NO FROM " & cFile & " WHERE DOC_NO = " & MyParn(xDoc_no.text) & " AND ITEM = " & addvalue(loctable!Item), con)) Then
        aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_no.text))
        con.Execute addInsert(aInsert, cFile), nAffect
        nAffect = 1
    Else
        con.Execute addUpdate(aInsert, cFile, "DOC_NO = " & MyParn(xDoc_no.text) & " AND ITEM = " & addvalue(loctable!Item)), nAffect
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


