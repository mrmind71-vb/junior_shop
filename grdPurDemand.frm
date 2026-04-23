VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdPurDemand 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ì ÿ·»Ì«  ‘—«¡"
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   900
      Width           =   2310
      Begin VB.CheckBox chkValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ﬁÌ„… «·„‘ —Ì« "
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   540
         Width           =   1995
      End
      Begin VB.CheckBox chkSeason 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„Ê«”„"
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
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   180
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   6345
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   540
      Width           =   3840
      Begin VB.OptionButton optOrd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "€Ì—„‰›–…"
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
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   225
         Width           =   1050
      End
      Begin VB.OptionButton optOrd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„‰›–…"
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   225
         Width           =   915
      End
      Begin VB.OptionButton optOrd 
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   780
      End
   End
   Begin VB.Frame fmBranch 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   540
      Width           =   3885
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
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   225
         Width           =   1365
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   31
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
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   225
         Width           =   645
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1170
      Width           =   7710
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   11
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
         Picture         =   "grdPurDemand.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1305
         TabIndex        =   10
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
         Picture         =   "grdPurDemand.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdPurDemand.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   6420
         TabIndex        =   7
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
         Picture         =   "grdPurDemand.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2565
         TabIndex        =   9
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
         Picture         =   "grdPurDemand.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPdf 
         Height          =   555
         Left            =   3870
         TabIndex        =   8
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
         Picture         =   "grdPurDemand.frx":BBFA
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSql 
         Height          =   555
         Left            =   5145
         TabIndex        =   28
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
         Picture         =   "grdPurDemand.frx":E0A7
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   10215
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   45
      Width           =   10095
      Begin VB.TextBox xOrder_no 
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
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1815
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
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   330
         Left            =   4860
         TabIndex        =   2
         Top             =   540
         Width           =   3660
         _ExtentX        =   6456
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
         Left            =   135
         TabIndex        =   4
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
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
         Left            =   135
         TabIndex        =   5
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
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
         Left            =   135
         TabIndex        =   6
         Top             =   900
         Width           =   3435
         _ExtentX        =   6059
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —…"
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
         Left            =   8685
         TabIndex        =   20
         Top             =   225
         Width           =   420
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
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "«·ﬁ”„ "
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
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   990
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
         Left            =   8625
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„” ‰œ «·ÿ·»Ì…"
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
         Left            =   8625
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   945
         Width           =   1275
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   630
      Top             =   -45
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
      Left            =   6750
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
      Left            =   225
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   270
      Top             =   765
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
      TabIndex        =   21
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
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   25
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
         TabIndex        =   26
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
      Height          =   240
      Left            =   225
      TabIndex        =   27
      Top             =   9585
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   423
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   2
      Left            =   13545
      TabIndex        =   33
      Top             =   1395
      Width           =   3345
      _ExtentX        =   5900
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
      Caption         =   "≈Ã„«·Ì «·„” ‰œ«  - √ﬁ”«„"
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Index           =   1
      Left            =   3375
      Top             =   135
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
      Index           =   2
      Left            =   3960
      Top             =   -45
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
      Index           =   3
      Left            =   8415
      Top             =   135
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
      Left            =   16920
      TabIndex        =   34
      Top             =   1395
      Width           =   3390
      _ExtentX        =   5980
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
      Caption         =   "≈Ã„«·Ì «·„” ‰œ« "
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   3
      Left            =   10170
      TabIndex        =   42
      Top             =   1395
      Width           =   3345
      _ExtentX        =   5900
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
      Caption         =   " ›’Ì·Ì «’‰«›"
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6900
      Index           =   2
      Left            =   45
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1935
      Width           =   20220
      _cx             =   35666
      _cy             =   12171
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
      Cols            =   21
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
      Height          =   6900
      Index           =   1
      Left            =   45
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1935
      Width           =   20220
      _cx             =   35666
      _cy             =   12171
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
      Cols            =   21
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
      Height          =   6900
      Index           =   3
      Left            =   45
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1935
      Width           =   20220
      _cx             =   35666
      _cy             =   12171
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
      Cols            =   24
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
End
Attribute VB_Name = "grdPurDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim FieldTable As ADODB.Recordset
Dim sOrder_No As String
Dim oSearch As New Search_abd
Dim cString As String, aHeader()
Private Sub chkSeason_Click()
grid1(1).Tag = IIf(chkSeason.Value = 0, "rp.sp_pur_demand", "rp.sp_pur_demand_season")
grid1(2).Tag = IIf(chkSeason.Value = 0, "rp.sp_pur_demand_section", "rp.sp_pur_demand_section_season")
End Sub

Private Sub chkValue_Click()
Dim i As Long
For i = 1 To cmdRep.UBound
    fixGrd i
Next
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
If Not IsDate(xdate1.text) Then
    MsgBox " «—ÌŒ «·„œ… «·«Ê· €Ì— „Õœœ"
    Exit Function
End If
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

Set data1.Recordset = mycmd("Select Code,DescA From File1_10SC ORDER BY code ", con)
Set xSection.RowSource = data1
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

Set data2.Recordset = mycmd("SELECT CODE,DESCA FROM STORE_BR", con)
Set xStore.RowSource = data2
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"


Set DATA3.Recordset = mycmd("Select Code,DescA From File4_10 order by Desca", con)
Set xSupp.RowSource = DATA3
xSupp.ListField = "Desca"
xSupp.BoundColumn = "Code"

Set data4.Recordset = mycmd("Select mosm ,descA From mosm ORDER BY date DESC ", con)
Set xMosm.RowSource = data4
xMosm.ListField = "Desca"
xMosm.BoundColumn = "MOSM"

Dim i As Long
For i = 1 To grid1.UBound
    Set grid1(i).DataSource = DATA11(i)
    fixGrd (i)
Next

grid1(1).Tag = "rp.sp_pur_demand"
'grid1(2).Tag = "[rp].[sp_pur_demand_season]"
grid1(2).Tag = "rp.sp_pur_demand_section"
grid1(3).Tag = "rp.sp_pur_demand_items"

fmBranch.Visible = lSupperVisor
optStop(1).Value = True

fixControls
cmdRep(1).Value = -1
End Sub
Private Sub myLoadGrd(Optional bString As Boolean = False, Optional bNoRemove As Boolean = False)
ReDim aHeader(13)

Dim aPrm As Variant
If IsDate(xdate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", addDate(xdate1.text))
    aHeader(0) = BetweenString(xdate1.text, xDate2.text)
End If

If IsDate(xDate2.text) Then
    aPrm = AddFlag(aPrm, "DATE2", addDate(xDate2.text))
    aHeader(0) = BetweenString(xdate1.text, xDate2.text)
End If

If xSection.MatchedWithList Then
    aPrm = AddFlag(aPrm, "SECTION", addvalue(xSection.BoundText))
    aHeader(1) = "«·ﬁ”„ : " & xSection.text
End If

If xSupp.MatchedWithList Then
    aPrm = AddFlag(aPrm, "SUP", addstring(xSupp.BoundText))
    aHeader(2) = "«·„Ê—œ : " & xSupp.text
End If

If Trim(xOrder_no.text) <> "" Then
    aPrm = AddFlag(aPrm, "ORDER_NO", addstring(xOrder_no.text))
    aHeader(2) = "—ﬁ„ «·ÿ·»Ì… : " & xOrder_no.text
End If

If sOrder_No <> "" Then
    aPrm = AddFlag(aPrm, "ORDER_NO", addstring(sOrder_No))
    aHeader(2) = "—ﬁ„ «·ÿ·»Ì… : " & sOrder_No
    sOrder_No = ""
End If

If xMosm.MatchedWithList Then
    aPrm = AddFlag(aPrm, "MOSM", addstring(xMosm.BoundText))
    aHeader(3) = "«·„Ê”„ : " & xMosm.text
End If

If xStore.MatchedWithList Then
    aPrm = AddFlag(aPrm, "STORE", addstring(xStore.BoundText))
    aHeader(4) = "„Œ“‰ : " & xStore.text
End If

If optOrd(1).Value Or optOrd(2).Value Then
    aPrm = AddFlag(aPrm, "ISEND_ORDER", IIf(optOrd(1).Value, "0", "1"))
    aHeader(5) = IIf(optOrd(1).Value, optOrd(1).Caption, optOrd(2).Caption)
End If

If optStop(1).Value Or optStop(2).Value Then
    aPrm = AddFlag(aPrm, "ISSTOP", IIf(optStop(1).Value, "0", "1"))
    aHeader(6) = IIf(optStop(1).Value, optStop(1).Caption, optStop(2).Caption)
End If


If Not bOpt5 Then
    aPrm = AddFlag(aPrm, "usercode", TurnValue(nusercode))
End If

Dim cString As String
Dim i As Long

If Not bString Then
    For i = 1 To cmdRep.UBound
        cString = myPrcString(grid1(i).Tag, aPrm)
        If cmdRep(i).Value = -1 Then
            Set DATA11(i).Recordset = mycmd(cString, con, , , 600)
            fixGrd i
            panel1(0).Caption = "⁄œœ «·”Ã·«  : " & grid1(i).Rows - 1
        ElseIf Not bNoRemove Then
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
Sub fixGrd(Index As Long)
With grid1(Index)
If Index = 1 Then
    .RowHeight(0) = 700
    .TextMatrix(0, 0) = "«·›—⁄"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    .TextMatrix(0, 2) = " «—ÌŒ ÿ·»Ì…"
    .TextMatrix(0, 3) = " „” ‰œ"
    .TextMatrix(0, 4) = "«·„Ê”„"
    
    .ColHidden(4) = (chkSeason.Value = 0)
    
    .TextMatrix(0, 5) = "⁄œœ „ÊœÌ·« "
    .TextMatrix(0, 6) = "ﬂ„Ì… ÿ·»Ì…"
    .TextMatrix(0, 7) = "ﬁÌ„… «·ÿ·»Ì…"
    
    .TextMatrix(0, 8) = "ﬂ ≈” ·«„"
    .TextMatrix(0, 9) = "»«ﬁÏ ≈” ·«„"
    
    .TextMatrix(0, 10) = "„‘ —Ì« "
    .TextMatrix(0, 11) = "›Ì„… „‘ —Ì« "
    
    .TextMatrix(0, 12) = " ÕÊÌ·« "
    
    .TextMatrix(0, 13) = "ﬂ„Ì… „‰›–…"
    .TextMatrix(0, 14) = "“Ì«œ… ⁄‰ «·ÿ·»Ì…"
    .TextMatrix(0, 15) = "«·»«ﬁÏ"
    
    .TextMatrix(0, 16) = "«Œ—  «—ÌŒ"
    .TextMatrix(0, 17) = "„‰›–"
    
    .TextMatrix(0, 18) = "‰”»…"
    .TextMatrix(0, 19) = "„·«ÕŸ« "
    
    .ColHidden(7) = chkValue.Value = 0
    .ColHidden(11) = chkValue.Value = 0
    
    .ColWidth(0) = 1300
    .ColWidth(1) = 4000
    .ColWidth(2) = 1300
    .ColWidth(3) = 1000
    .ColWidth(4) = 800
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1400
    .ColWidth(9) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColWidth(11) = 1400
    
    .ColWidth(12) = 1000
    .ColWidth(13) = 1000
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    .ColWidth(16) = 1500
    .ColWidth(17) = 1000
    .ColWidth(18) = 1000
    .ColWidth(19) = 3000
        
    
    .ColDataType(15) = flexDTDate
    .ColFormat(18) = "#.##%"
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    
    .MergeCells = flexMergeRestrictAll
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = chkSeason.Value = 1
    
    .SubtotalPosition = flexSTAbove
    
    
    .Subtotal flexSTSum, -1, 5, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 6, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 7, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 8, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 9, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 10, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 11, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 12, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 13, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 14, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 15, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    
    If .Rows > 1 Then
        .TextMatrix(1, 18) = Round(.ValueMatrix(1, 13) / .ValueMatrix(1, 6), 2)
    End If
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .SubtotalPosition = flexSTAbove
ElseIf Index = 2 Then
    .RowHeight(0) = 700
    .TextMatrix(0, 0) = "«·›—⁄"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    .TextMatrix(0, 2) = " «—ÌŒ ÿ·»Ì…"
    .TextMatrix(0, 3) = " „” ‰œ"
    .TextMatrix(0, 4) = "«·ﬁ”„"
    .TextMatrix(0, 5) = "«·„Ê”„"
    
    .TextMatrix(0, 6) = "⁄œœ „ÊœÌ·« "
    .TextMatrix(0, 7) = "ﬂ„Ì… ÿ·»Ì…"
    .TextMatrix(0, 8) = "ﬁÌ„… ÿ·»Ì…"
    .TextMatrix(0, 9) = "ﬂ ≈” ·«„"
    .TextMatrix(0, 10) = "»«ﬁÏ ≈” ·«„"
    .TextMatrix(0, 11) = "„‘ —Ì« "
    .TextMatrix(0, 12) = "ﬁÌ„… „‘ —Ì« "
    .TextMatrix(0, 13) = " ÕÊÌ·« "
    
    .TextMatrix(0, 14) = "ﬂ„Ì… „‰›–…"
    .TextMatrix(0, 15) = "“Ì«œ… ⁄‰ «·ÿ·»Ì…"
    .TextMatrix(0, 16) = "«·»«ﬁÏ"
        
    .TextMatrix(0, 17) = "«Œ—  «—ÌŒ"
    .TextMatrix(0, 18) = "„‰›–"
    .TextMatrix(0, 19) = "‰”»…"
    .TextMatrix(0, 20) = "„·«ÕŸ« "
    
    .ColHidden(8) = chkValue.Value = 0
    .ColHidden(12) = chkValue.Value = 0
    
    .ColWidth(0) = 1300
    .ColWidth(1) = 4000
    .ColWidth(2) = 1300
    .ColWidth(3) = 1000
    .ColWidth(4) = 2000
    .ColWidth(5) = 1000
        
    .ColWidth(6) = 1000
    .ColWidth(8) = 1400
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColWidth(11) = 1000
    .ColWidth(12) = 1400
    .ColWidth(13) = 1400
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    .ColWidth(16) = 1000
    .ColWidth(17) = 1500
    .ColWidth(18) = 1000
    .ColWidth(19) = 1000
    .ColWidth(20) = 3000
        
    
    .ColDataType(17) = flexDTDate
    .ColFormat(19) = "#.##%"
    
     .ColHidden(5) = chkSeason.Value = 0
  
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    
    .MergeCells = flexMergeRestrictAll
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = chkSeason.Value = 1
    
    .SubtotalPosition = flexSTAbove
    
    .Subtotal flexSTSum, -1, 7, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 8, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 9, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 10, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 11, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 12, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 13, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 14, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 15, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 16, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    
    If .Rows > 1 Then
        .TextMatrix(1, 19) = Round(Val(.TextMatrix(1, 14)) / Val(.TextMatrix(1, 7)), 2)
    End If
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .SubtotalPosition = flexSTAbove
ElseIf Index = 3 Then
    .RowHeight(0) = 700
    .TextMatrix(0, 0) = "«·›—⁄"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    .TextMatrix(0, 2) = " «—ÌŒ ÿ·»Ì…"
    .TextMatrix(0, 3) = " „” ‰œ"
    
    .TextMatrix(0, 4) = "«·„ÊœÌ·"
    .TextMatrix(0, 5) = "«·«”„"
    .TextMatrix(0, 6) = "«·„ﬁ«”"
    .TextMatrix(0, 7) = "«·»«—ﬂÊœ"
    
    .TextMatrix(0, 8) = "⁄œœ „ÊœÌ·« "
    .TextMatrix(0, 9) = "ﬂ„Ì… ÿ·»Ì…"
    .TextMatrix(0, 10) = "ﬁÌ„… ÿ·»Ì…"
    .TextMatrix(0, 11) = "ﬂ ≈” ·«„"
    .TextMatrix(0, 12) = "»«ﬁÏ ≈” ·«„"
        
    .TextMatrix(0, 13) = "„‘ —Ì« "
    .TextMatrix(0, 14) = "ﬁÌ„… „‘ —Ì« "
    .TextMatrix(0, 15) = " ÕÊÌ·« "

    
    .TextMatrix(0, 16) = "ﬂ„Ì… „‰›–…"
    .TextMatrix(0, 17) = "“Ì«œ… ⁄‰ «·ÿ·»Ì…"
    .TextMatrix(0, 18) = "«·»«ﬁÏ"
    .TextMatrix(0, 19) = "—’Ìœ „Œ“‰"
    
    .TextMatrix(0, 19 + 1) = "«Œ—  «—ÌŒ"
    .TextMatrix(0, 20 + 1) = "„‰›–"
    .TextMatrix(0, 21 + 1) = "‰”»…"
    .TextMatrix(0, 22 + 1) = "„·«ÕŸ« "
    
    .ColHidden(10) = chkValue.Value = 0
    .ColHidden(14) = chkValue.Value = 0
    
    .ColWidth(0) = 1300
    .ColWidth(1) = 2000
    .ColWidth(2) = 1300
    .ColWidth(3) = 1000
    .ColWidth(4) = 1500
    .ColWidth(5) = 2500
    .ColWidth(6) = 1000
    .ColWidth(7) = 1100
    
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    
    .ColWidth(10) = 1400
    .ColWidth(11) = 1000
    .ColWidth(12) = 1000
    .ColWidth(13) = 1000
    
    .ColWidth(14) = 1400
    .ColWidth(15) = 1000
    .ColWidth(16) = 1000
    .ColWidth(17) = 1000
    .ColWidth(18) = 1000
    .ColWidth(19) = 1000
    
    .ColWidth(19 + 1) = 1500
    .ColWidth(20 + 1) = 1000
    .ColWidth(21 + 1) = 1000
    .ColWidth(22 + 1) = 3000
        
    
    .ColDataType(18) = flexDTDate
    .ColFormat(21 + 1) = "#.##%"
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    
    .MergeCells = flexMergeRestrictAll
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .SubtotalPosition = flexSTAbove
    
    .Subtotal flexSTSum, -1, 9, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 10, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 11, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 12, "#,###,##0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 13, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 14, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 15, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 16, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 17, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 18, "#0", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    
    If .Rows > 1 Then
        .TextMatrix(1, 21 + 1) = Round(Val(.TextMatrix(1, 16)) / Val(.TextMatrix(1, 9)), 2)
    End If
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .SubtotalPosition = flexSTAbove
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

Private Sub grid1_DblClick(Index As Integer)
If Index = 1 Then
    If grid1(1).Row > 1 Then
        cmdRep(3).Value = True
        Handlecontrols
        'DATA11(3).Recordset.Filter = "ORDER_NO = " & grid1(Index).TextMatrix(grid1(Index).Row, 3)
        sOrder_No = grid1(Index).TextMatrix(grid1(Index).Row, 3)
        
        Me.MousePointer = vbHourglass
        myLoadGrd False, True
        Me.MousePointer = vbNormal
    End If
End If
End Sub

Private Sub xOrder_no_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    orderLookup
End If
End Sub
Private Sub fixControls()
If Dir(App.Path & "\sys_img\selected.jpg") = "" Then Exit Sub
Dim i As Long
For i = 1 To cmdRep.UBound
    cmdRep(i).PictureDn = LoadPicture(App.Path & "\sys_img\selected.jpg")
Next
End Sub
Private Sub orderLookup()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(6, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,[DATE] AS DATE1," & _
                   " FILE4_10.Desca ," & _
                   " STORE_BR.desca " & _
                   "FILE6_50H.REMARK" & _
                  " FROM  file6_50h " & _
                  " INNER JOIN file4_10 ON file6_50h.CODE  = FILE4_10.CODE" & _
                  " inner join STORE_BR on STORE_BR.code = file6_50h.store "
If optOrd(1).Value Then
    Generalarray(1) = Generalarray(1) & " where ISEND_ORDER = 0"
ElseIf optOrd(2).Value Then
    Generalarray(1) = Generalarray(1) & " where ISEND_ORDER = 1"
End If

Generalarray(2) = "Order by DATE DESC , doc_no desc "
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·—ﬁ„-≈”„ " & cCodeDesca & "-«· «—ÌŒ"
listarray(0, 1) = "(inv_no Like '%cFilter%' or Doc_No Like '%cFilter%' or  FILE4_10.DESCA LIKE '%cFilter%' OR " & _
                  "##date##)"

listarray(1, 0) = "«·›—⁄"
listarray(1, 1) = "STORE_BR.desca Like '%cFilter%' "

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "≈”„ " & cCodeDesca
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "—ﬁ„ ›« Ê—…"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "„Ê”„"
GrdArray(5, 1) = 0

GrdArray(6, 0) = "«·„Œ“‰"
GrdArray(6, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.sCaption = "«” ⁄·«„"
oSearch.Show 1
End Sub
Public Sub myProc()
xOrder_no.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
oSearch.Hide
End Sub
Private Sub xOrder_no_GotFocus()
myGotFocus xOrder_no
End Sub
Private Sub xOrder_no_LostFocus()
myLostFocus xOrder_no
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xSupp_GotFocus()
myGotFocus xSupp
End Sub
Private Sub xSupp_LostFocus()
myLostFocus xSupp
If Not xSupp.MatchedWithList Then xSupp.BoundText = ""
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
myGotFocus xSection
End Sub
Private Sub xSection_LostFocus()
myLostFocus xSection
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
