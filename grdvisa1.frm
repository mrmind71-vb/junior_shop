VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdvisafrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "„—«Ã⁄… «·›Ì“«"
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   45
      Width           =   5685
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdvisa1.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1170
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdvisa1.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdvisa1.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   4545
         TabIndex        =   10
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdvisa1.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2295
         TabIndex        =   12
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdvisa1.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPdf 
         Height          =   555
         Left            =   3420
         TabIndex        =   11
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdvisa1.frx":BBFA
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   8955
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   0
      Width           =   11355
      Begin VB.TextBox xApprove_no 
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
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   945
         Width           =   3930
      End
      Begin VB.TextBox xphone 
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
         Left            =   7335
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   900
         Width           =   3075
      End
      Begin VB.TextBox xDate2 
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
         Left            =   7335
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1455
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
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1590
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   330
         Left            =   7335
         TabIndex        =   3
         Top             =   540
         Width           =   3075
         _ExtentX        =   5424
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
      Begin MSDataListLib.DataCombo xbank 
         Height          =   330
         Left            =   2115
         TabIndex        =   2
         Top             =   225
         Width           =   3930
         _ExtentX        =   6932
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
      Begin MSDataListLib.DataCombo xvisa_code 
         Height          =   330
         Left            =   2115
         TabIndex        =   4
         Top             =   585
         Width           =   3930
         _ExtentX        =   6932
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         DragMode        =   1  'Automatic
         Height          =   1140
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   135
         Width           =   1905
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»œ· ‰ﬁ«ÿ"
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
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   810
            Width           =   1185
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "›Ì“«"
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
            Height          =   210
            Index           =   1
            Left            =   675
            RightToLeft     =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   495
            Width           =   1050
         End
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
            Height          =   210
            Index           =   0
            Left            =   495
            RightToLeft     =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   225
            Value           =   -1  'True
            Width           =   1230
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„Ê«›ﬁ…"
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
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   945
         Width           =   915
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·›Ì“«"
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
         Left            =   6120
         TabIndex        =   24
         Top             =   585
         Width           =   750
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
         Left            =   6165
         TabIndex        =   23
         Top             =   225
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ „Ê»Ì·"
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
         Left            =   10485
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   10530
         TabIndex        =   19
         Top             =   225
         Width           =   510
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
         Left            =   10530
         TabIndex        =   18
         Top             =   585
         Width           =   405
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   810
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Index           =   1
      Left            =   2745
      Top             =   360
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
      Left            =   2475
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   2025
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   1755
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
   Begin Threed.SSRibbon cmdRep 
      Height          =   465
      Index           =   2
      Left            =   3240
      TabIndex        =   16
      Top             =   855
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   820
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
      Caption         =   "›Ì“ €Ì— „ ÿ«»ﬁ…"
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin Threed.SSRibbon cmdRep 
      Height          =   465
      Index           =   1
      Left            =   6120
      TabIndex        =   15
      Top             =   855
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   820
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
      Caption         =   " ›’Ì·Ì ›Ì“ ÌÊ„Ì…"
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
      Value           =   -1  'True
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Index           =   2
      Left            =   180
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   27
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
         TabIndex        =   32
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
      TabIndex        =   33
      Top             =   9630
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8385
      Index           =   1
      Left            =   45
      TabIndex        =   25
      Top             =   1395
      Width           =   20220
      _cx             =   35666
      _cy             =   14790
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
      Cols            =   14
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8250
      Index           =   2
      Left            =   45
      TabIndex        =   26
      Top             =   1395
      Width           =   20220
      _cx             =   35666
      _cy             =   14552
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "grdvisafrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADOdb.Connection
Dim FieldTable As ADOdb.Recordset
Dim cString As String, aHeader()
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
If Not MYVALID Then Exit Sub
Me.MousePointer = vbHourglass
myLoadGrd
Me.MousePointer = vbNormal
End Sub

Private Sub cmdPdf_Click()
myPrint "1", True
End Sub
Private Sub cmdPrint_Click()
myPrint
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
        
        Set printGrdNew.myForm = Me
        printGrdNew.doprint grid1(i), 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption & " - " & cmdRep(i).Caption, ArbString(retHeader(aHeader, 0, 5)), , False, i = 5, 10, , aRow, Array(1)
    End If
Next

If Not bIgPreview Then
    printGrdNew.Show 1
Else
    Unload printGrdNew
End If
End Sub
Private Sub Form_Load()
openCon con

Set DATA1.Recordset = mycmd("SELECT * FROM FILE0_40", con)
Set XSTORE.RowSource = DATA1
XSTORE.ListField = "Desca"
XSTORE.BoundColumn = "Code"
If cBranch <> "00" Then
    XSTORE.BoundText = cBranchStore
    XSTORE.Enabled = False
End If

Set DATA2.Recordset = myRecordSet("SELECT FILE5_10.CODE,FILE5_10.DESCA FROM FILE5_10 INNER JOIN VISA_CODES ON FILE5_10.CODE = VISA_CODES.BANK WHERE IS_VISA = 1 GROUP BY FILE5_10.CODE,FILE5_10.DESCA ", con)
Set xBank.RowSource = DATA2
xBank.ListField = "Desca"
xBank.BoundColumn = "Code"

Set DATA3.Recordset = myRecordSet("SELECT CODE,DESCA FROM VISA_CODES WHERE VISA_CODES.STOPED = 0 ORDER BY DESCA", con)
Set xvisa_code.RowSource = DATA3
xvisa_code.ListField = "Desca"
xvisa_code.BoundColumn = "Code"

Dim i As Long
For i = 1 To grid1.UBound
    Set grid1(i).DataSource = DATA11(i)
    fixGrd (i)
Next

xDate1.text = myFormat_p(Date)
xDate2.text = myFormat_p(Date)
grid1(1).Tag = "rp.sp_visa_totals"
grid1(2).Tag = "rp.sp_visa_totals_is_Differ"
End Sub
Private Sub myLoadGrd(Optional bString As Boolean = False)
ReDim aHeader(6)
Dim aPrm As Variant

If XSTORE.MatchedWithList Then
    aPrm = AddFlag(Empty, "STORE", addstring(XSTORE.BoundText))
    aHeader(0) = "«·„Œ“‰ : " & XSTORE.text
End If

If xBank.MatchedWithList Then
    aPrm = AddFlag(aPrm, "bank", addstring(xBank.BoundText))
    aHeader(1) = "«·»‰ﬂ : " & xBank.text
End If

If xvisa_code.MatchedWithList Then
    aPrm = AddFlag(aPrm, "visa_code", addvalue(xvisa_code.BoundText))
    aHeader(2) = "‰Ê⁄ «·›Ì“« : " & xvisa_code.text
End If

If Trim(xPhone.text) <> "" Then
    aPrm = AddFlag(aPrm, "phone", addstring(xPhone.text))
    aHeader(3) = "—ﬁ„ «· ·Ì›Ê‰ : " & xPhone.text
End If

If Trim(xApprove_no.text) <> "" Then
    aPrm = AddFlag(aPrm, "APPROVE_NO", addstring(xApprove_no.text))
    aHeader(4) = "—ﬁ„ «·„Ê«›ﬁ… : " & xApprove_no.text
End If

If IsDate(xDate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", addDate(xDate1.text))
    aHeader(5) = BetweenString(xDate1.text, xDate2.text)
End If

If IsDate(xDate2.text) Then
    aPrm = AddFlag(aPrm, "DATE2", addDate(xDate2.text))
    aHeader(5) = BetweenString(xDate1.text, xDate2.text)
End If

If Not Option1(0).Value Then
    aPrm = AddFlag(aPrm, "WITH_POINTS", IIf(Option1(1).Value, "0", "1"))
    aHeader(6) = IIf(Option1(1).Value, Option1(1).Caption, Option1(2).Caption)
End If

Dim cString As String
Dim i As Long
If Not bString Then
    For i = 1 To cmdRep.UBound
        cString = myPrcString(grid1(i).Tag, aPrm)
        If cmdRep(i).Value = -1 Then
            Set DATA11(i).Recordset = mycmd(cString, con, , , 600)
            fixGrd i
            panel1(i - 1).Caption = "⁄œœ «·”Ã·«  : " & grid1(i).Rows - 1
        Else
            If DATA11(i).RecordSource <> cString Then grid1(i).Rows = 1
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
Sub fixGrd(Index)
With grid1(Index)
If Index = 1 Then
    .RowHeight(0) = 800
    .WordWrap = True
    .TextMatrix(0, 0) = "«· «—ÌŒ"
    .TextMatrix(0, 1) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 2) = "«·›—⁄"
    .TextMatrix(0, 3) = "—ﬁ„ «· ·Ì›Ê‰"
    .TextMatrix(0, 4) = "≈Ã„«·Ì «·›« Ê—…"
    .TextMatrix(0, 5) = "«·»‰ﬂ"
    .TextMatrix(0, 6) = "‰Ê⁄ «·›Ì“«"
    .TextMatrix(0, 7) = "ﬁÌ„… «·›Ì“«"
    .TextMatrix(0, 8) = "≈Ã„«·Ì ‰”»… «·⁄„Ì·"
    .TextMatrix(0, 9) = "≈Ã„«·Ì «·›Ì“«"
    .TextMatrix(0, 10) = "‰”»… «·»‰ﬂ"
    .TextMatrix(0, 11) = "—ﬁ„ «·⁄„·Ì…"
    .TextMatrix(0, 12) = "—ﬁ„ «·„Ê«›ﬁ…"
    .TextMatrix(0, 13) = "≈”„ «·ÃÂ«“"
    .ColWidth(0) = 1300
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1100
    .ColWidth(5) = 1900
    .ColWidth(6) = 2000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColWidth(11) = 1700
    .ColWidth(12) = 1700
    .ColWidth(12) = 1700
    .MergeCells = flexMergeRestrictAll
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
         
    For i = 7 To .Cols - 2
        .ColDataType(i) = flexDTDouble
    Next
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 4, "#0.00", &HC0FFC0, vbBlack, True, "  "
    For i = 7 To .Cols - 3
        .Subtotal flexSTSum, -1, i, "#0.00", &HC0FFC0, vbBlack, True, "  "
    Next
    
    If .Rows > 1 Then
        .TextMatrix(1, 1) = "«·≈Ã„«·Ï"
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
ElseIf Index = 2 Then
    .RowHeight(0) = 800
    .WordWrap = True
    .TextMatrix(0, 0) = "«· «—ÌŒ"
    .TextMatrix(0, 1) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 2) = "«·›—⁄"
    .TextMatrix(0, 3) = "≈”„ «·ÃÂ«“"
    .TextMatrix(0, 3 + 1) = "—ﬁ„ «· ·Ì›Ê‰"
    .TextMatrix(0, 4 + 1) = "≈Ã„«·Ì «·›« Ê—…"
    .TextMatrix(0, 5 + 1) = "≈Ã„«·Ì ›Ì“ «·›« Ê—…"
    .TextMatrix(0, 6 + 1) = "≈Ã„«·Ì «·›Ì“"
    .ColWidth(0) = 1300
    .ColWidth(1) = 1600
    .ColWidth(2) = 1600
    .ColWidth(3) = 3000
    .ColWidth(3 + 1) = 1600
    .ColWidth(4 + 1) = 1300
    .ColWidth(5 + 1) = 1300
    .ColWidth(6 + 1) = 1300
    .MergeCells = flexMergeRestrictAll
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3 + 1) = True
    .MergeCol(4 + 1) = True
         
    For i = 4 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 4 + 1, "#0.00", &HC0FFC0, vbBlack, True, "  "
    For i = 4 + 1 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0.00", &HC0FFC0, vbBlack, True, "  "
    Next
    
    If .Rows > 1 Then
        .TextMatrix(1, 1) = "«·≈Ã„«·Ï"
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
End If
End With
End Sub

Private Sub Form_Resize()
If Me.Height < 5000 Then Exit Sub
For i = 1 To grid1.UBound
    grid1(i).Height = Me.Height - 2000
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdChargefrm1 = Nothing
End Sub
Private Function MYVALID() As Boolean
MYVALID = True
End Function


Private Sub grid1_DblClick(Index As Integer)
'salesfrm.isOld = grid1(Index).Col = 1
salesfrm.sDoc_no = grid1(Index).TextMatrix(grid1(Index).Row, 1)
salesfrm.Show
End Sub

'Private Sub grid1_DblClick()
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
'End Sub
Private Sub Option1_Click(Index As Integer)
'myloadgrd
End Sub
Private Sub xbank_Change()
If xBank.MatchedWithList Or Trim(xBank.text) = "" Then
    Dim sSql As String
    sSql = "SELECT CODE,DESCA FROM VISA_CODES WHERE VISA_CODES.STOPED = 0"
    sSql = sSql & " AND VISA_CODES.BANK = " & MyParn(xBank.BoundText) & _
            " ORDER BY DESCA"
    If sSql <> DATA3.Recordset.Source Then
        Dim sBound As String
        sBound = xvisa_code.BoundText
        Set DATA3.Recordset = mycmd(sSql, con)
        xvisa_code.BoundText = sBound
        If Not xvisa_code.MatchedWithList Then xvisa_code.BoundText = ""
    End If
End If
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
Private Sub Text1_GotFocus()
myGotFocus Text1
End Sub
Private Sub Text1_LostFocus()
myLostFocus Text1
End Sub
Private Sub xPhone_GotFocus()
myGotFocus xPhone
End Sub
Private Sub XPHONE_LostFocus()
myLostFocus xPhone
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus XSTORE
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus XSTORE
If Not XSTORE.MatchedWithList Then XSTORE.BoundText = ""
End Sub
Private Sub xbank_GotFocus()
myGotFocus xBank
End Sub
Private Sub xbank_LostFocus()
myLostFocus xBank
If Not xBank.MatchedWithList Then xBank.BoundText = ""
End Sub
Private Sub xvisa_code_GotFocus()
myGotFocus xvisa_code
End Sub
Private Sub xvisa_code_LostFocus()
myLostFocus xvisa_code
If Not xvisa_code.MatchedWithList Then xvisa_code.BoundText = ""
End Sub
Private Sub xApprove_no_GotFocus()
myGotFocus xApprove_no
End Sub
Private Sub xApprove_no_LostFocus()
myLostFocus xApprove_no
End Sub

