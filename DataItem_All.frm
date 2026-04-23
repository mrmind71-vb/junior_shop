VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form DataItem_All 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ›’Ì·Ï »Ì«‰«  «·„ÊœÌ·« "
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
   Begin VB.Frame Frame5 
      Height          =   555
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   315
      Width           =   2895
      Begin VB.CheckBox chkBarCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ »«—ﬂÊœ œ«Œ·Ì"
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
         Height          =   285
         Left            =   675
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   180
         Width           =   2040
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1230
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   810
      Width           =   2895
      Begin VB.CheckBox XNODEM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ «’‰«› »œÊ‰ ÿ·»Ì…"
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
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   765
         Width           =   2445
      End
      Begin VB.CheckBox XISSTOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ «’‰«› „ Êﬁ›…"
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
         Left            =   495
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   450
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CheckBox XNOMOVE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«’‰«› ·Ì” ·Â« «Ï Õ—ﬂ…"
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   180
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Height          =   600
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   -45
      Width           =   3525
      Begin VB.CommandButton CMD_DELITEM 
         BackColor       =   &H008080FF&
         Caption         =   "Õ–› «·«’‰«›"
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
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   135
         Width           =   3390
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   720
      Width           =   3390
      Begin VB.CommandButton cmd_excel 
         Height          =   555
         Left            =   1170
         Picture         =   "DataItem_All.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   1620
         Picture         =   "DataItem_All.frx":2404
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "⁄—÷"
         Top             =   720
         Width           =   1680
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   90
         Picture         =   "DataItem_All.frx":48F6
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   720
         Width           =   1500
      End
      Begin Threed.SSCommand cmdScv 
         Height          =   555
         Left            =   90
         TabIndex        =   43
         Top             =   135
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   979
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
         Picture         =   "DataItem_All.frx":6D62
         Caption         =   "CSV"
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "DataItem_All.frx":9335
      End
      Begin Threed.SSCommand cmdSql 
         Height          =   555
         Left            =   2250
         TabIndex        =   44
         Top             =   135
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   979
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
         Picture         =   "DataItem_All.frx":B62B
         ButtonStyle     =   1
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   6705
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   11805
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "»«—ﬂÊœ Œ«—ÃÏ „ﬂ——"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   135
         Width           =   1950
      End
      Begin VB.TextBox XBARCODE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   135
         MaxLength       =   15
         TabIndex        =   8
         Top             =   540
         Width           =   2445
      End
      Begin VB.TextBox xColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7065
         MaxLength       =   15
         TabIndex        =   3
         Top             =   900
         Width           =   1545
      End
      Begin VB.TextBox XITEM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   135
         MaxLength       =   15
         TabIndex        =   12
         Top             =   1305
         Width           =   1815
      End
      Begin VB.TextBox xScal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   8955
         MaxLength       =   15
         TabIndex        =   2
         Top             =   900
         Width           =   1545
      End
      Begin VB.TextBox xModelFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4095
         MaxLength       =   15
         TabIndex        =   7
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox xDesca 
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   922
         Width           =   5505
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   7065
         TabIndex        =   1
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   7065
         TabIndex        =   0
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   3330
         TabIndex        =   6
         Top             =   180
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XBRNCH 
         Height          =   315
         Left            =   7065
         TabIndex        =   4
         Top             =   1305
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   315
         Left            =   3465
         TabIndex        =   10
         Top             =   1305
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xage 
         Height          =   315
         Left            =   7065
         TabIndex        =   5
         Top             =   1665
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xsex 
         Height          =   315
         Left            =   3465
         TabIndex        =   11
         Top             =   1665
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "›∆… ⁄„—Ì… :"
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
         Left            =   10710
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1665
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "«·‰Ê⁄ :"
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
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1665
         Width           =   555
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "»«—ﬂÊœ Œ«—ÃÏ-GS1"
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
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   585
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁ”„ : "
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
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1305
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄ : "
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
         Left            =   10635
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1305
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "»«—ﬂÊœ "
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
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1350
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„ﬁ«” - ·Ê‰ :"
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
         Left            =   10635
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   945
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ÊœÌ· :"
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
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ ’‰› :"
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
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… :"
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
         Index           =   1
         Left            =   10635
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·„’‰⁄ :"
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
         Left            =   10635
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   255
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ê”„ :"
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
         Index           =   3
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   780
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   13
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
            Key             =   ""
            Object.Tag             =   ""
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   3105
      Top             =   1665
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
      Left            =   2430
      Top             =   1305
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
      Left            =   4080
      Top             =   1440
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
      Left            =   3495
      Top             =   1440
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
      Left            =   3195
      Top             =   1440
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
   Begin MSAdodcLib.Adodc data6 
      Height          =   330
      Left            =   3795
      Top             =   1440
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
      Left            =   3195
      Top             =   1395
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   25
      Top             =   10560
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   265
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   330
      Left            =   3195
      Top             =   1395
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
   Begin MSAdodcLib.Adodc DATA8 
      Height          =   330
      Left            =   3195
      Top             =   1395
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
      Height          =   7125
      Left            =   180
      TabIndex        =   33
      Top             =   2115
      Width           =   18375
      _cx             =   32411
      _cy             =   12568
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
      Cols            =   23
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc DATA9 
      Height          =   330
      Left            =   3195
      Top             =   1395
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
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   3195
      Top             =   1395
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
   Begin VB.Label XMODEL 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1575
      Width           =   2085
   End
End
Attribute VB_Name = "DataItem_All"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Variant, sM_Fact As String
Dim oSearch As New Search3
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Private Sub CMD_DELITEM_Click()
On Error GoTo myerror
With grid1
    If MsgBox(" Õ–› «·„ÊœÌ·«  ›Ï «· ﬁ—Ì— ", vbYesNo + vbDefaultButton2) = vbYes Then
        prog1.Visible = True
        prog1.Value = 0
        prog1.Min = 0
        prog1.Max = .Rows - 1
        For nRow = 1 To .Rows - 1
            xModel.Caption = Val(.TextMatrix(nRow, 0))
            If GetDesca("SELECT ITEM FROM FILE1_11_ALL WHERE  ITEM = " & Val(.TextMatrix(nRow, 0)), con) = "" Then
                con.Execute " DELETE FROM FILE6_50 WHERE ITEM = " & Val(.TextMatrix(nRow, 0))
                con.Execute " DELETE FROM FILE6_51 WHERE ITEM = " & Val(.TextMatrix(nRow, 0))
                con.Execute " DELETE FROM FILE6_52 WHERE ITEM = " & Val(.TextMatrix(nRow, 0))
                con.Execute " DELETE FROM FILE1_10 WHERE ITEM = " & Val(.TextMatrix(nRow, 0))
            End If
            prog1.Value = nRow
        Next nRow
        Inform "  „ Õ–› «·„ÊœÌ·«  "
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub

Private Sub cmdScv_Click()
myLoadCsv
End Sub

Private Sub cmdSql_Click()
myload , True
End Sub

Private Sub Form_Load()
    openCon con
    
    CMD_DELITEM.Visible = bSupermode
    StoreTable.Open "SELECT * FROM FILE0_40 WHERE isstop = 0 ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
    
    
    Set DATA3.Recordset = myRecordSet("Select Code,DescA From File1_50 ORDER BY DESCA", con)
    Set xgroup.RowSource = DATA3
    xgroup.ListField = "Desca"
    xgroup.BoundColumn = "Code"
    
    Set data4.Recordset = myRecordSet("Select mosm ,descA From mosm ORDER BY date DESC ", con)
    Set xMosm.RowSource = data4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    Set data5.Recordset = myRecordSet("Select code ,desca From fact ORDER BY code ", con)
    Set xFact.RowSource = data5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    xFact.BoundText = sM_Fact
    
    
    Set DATA7.Recordset = myRecordSet("Select Code,DescA From STORE_BR ORDER BY CODE ", con)
    Set XBRNCH.RowSource = DATA7
    XBRNCH.ListField = "Desca"
    XBRNCH.BoundColumn = "Code"
    
    Set data8.Recordset = myRecordSet("Select Code,DescA From FILE1_10SC ", con)
    Set xSection.RowSource = data8
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    
    Set data9.Recordset = mycmd("Select Code,DescA From FILE0_81 ", con)
    Set xsex.RowSource = data9
    xsex.ListField = "Desca"
    xsex.BoundColumn = "Code"
    
    Set data10.Recordset = mycmd("Select Code,DescA From FILE0_82 ", con)
    Set xage.RowSource = data10
    xage.ListField = "Desca"
    xage.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA6
    
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .RowHeight(0) = 800
    .FrozenCols = 7
    End With
    

    fixGrd
    
    
    If xFact.BoundText <> "" Then cmdGo_Click
End Sub
Private Sub myload(Optional bTemp As Boolean = False, Optional bSql As Boolean = False)
    Dim cWhere As String, cFiled1 As String, cString  As String
    Dim aFields(22)
    If bTemp Then
       aFields(0) = "CASE WHEN FILE1_10.ITEM IS NULL THEN #TEMP.ITEM2 ELSE  CAST(FILE1_10.ITEM AS VARCHAR(30)) END"
    Else
       aFields(0) = "FILE1_10.ITEM"
    End If
    aFields(1) = "FACT.DESCA AS FACT_DESCA"
    aFields(2) = "FILE1_10.MOSM"
    aFields(3) = "FILE1_10.MODELFACT0"
    aFields(4) = "FILE1_10.desca"
    aFields(5) = "FILE1_10.SCAL"
    aFields(6) = "FILE1_10.COLOR"
    If bTemp Then
       aFields(7) = "#TEMP.QUANT"
    Else
       aFields(7) = "NULL AS QUANT"
    End If
    
    aFields(7 + 1) = "FILE1_10.PRICE"
    
    
    aFields(8 + 1) = "FILE1_10.GrModel"
    aFields(9 + 1) = "FILE1_10.PRICE_2"
    
    aFields(10 + 1) = "FILE1_10.BARCODE"
    aFields(11 + 1) = "FILE1_10.BARCODE2"
    aFields(12 + 1) = "FILE1_10.BARCODE13"
    aFields(13 + 1) = "FILE1_10.BARCODE_GS1"
    
    aFields(14 + 1) = "FILE1_10.DESCE"
    aFields(15 + 1) = "FILE1_10.COLOR_E"
    aFields(16 + 1) = "CASE WHEN PRICE_P = 0 THEN PRICE ELSE  FILE1_10.PRICE_P END  AS PRICE_P"
    aFields(17 + 1) = "FILE1_50.DESCA"
    aFields(18 + 1) = "FILE1_10.MODEL"
    
    aFields(19 + 1) = "FILE1_10SC.DESCA"
    aFields(20 + 1) = "FILE0_81.DESCA AS [SEX]"
    aFields(21 + 1) = "FILE0_82.DESCA AS [AGE]"
      
    '                        0               1                           2               3                   4                       5                   6               7           8                       9               10      11          12          13          14      15                      16                                                                  17  18
    'cString = " SELECT  FILE1_10.ITEM , FACT.DESCA AS FACTDESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca AS DESCA,  FILE1_10.SCAL ,  FILE1_10.COLOR, FILE1_10.PRICE , FILE1_10.PRICE2 , FILE1_10.PRICE_2  , BARCODE , BARCODE2 , BARCODE13 , BARCODE_GS1 , DESCE , COLOR_E ,CASE WHEN PRICE_P = 0 THEN PRICE ELSE  FILE1_10.PRICE_P END  AS PRICE_P ,file1_50.desca, MODEL,FILE1_10SC.DESCA "
    
    If bTemp Then
       cString = "SELECT " & arString(aFields, "," & vbCrLf) & _
                   " FROM #TEMP LEFT JOIN FILE1_10 ON #TEMP.ITEM = FILE1_10.ITEM " & _
                   " LEFT JOIN FACT ON FILE1_10.FACT = FACT.CODE " & _
                   " LEFT JOIN FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]" & _
                   " LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE" & _
                   " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                   " LEFT JOIN FILE0_82 ON FILE1_10.MODELAGE = FILE0_82.CODE"
    Else
       cString = "SELECT " & arString(aFields, "," & vbCrLf) & vbCrLf & _
                   "FROM  FILE1_10 " & sBt & _
                   "INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE " & sBt & _
                   "LEFT JOIN FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]" & sBt & _
                   "LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE" & sBt & _
                   "LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & sBt & _
                   "LEFT JOIN FILE0_82 ON FILE1_10.MODELAGE = FILE0_82.CODE"
    End If
    
    xModelFact.text = DelZero(xModelFact.text)
    
    If xMosm.MatchedWithList Then
       cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText) & vbCrLf
    End If
    
    If Trim(xModelFact.text) <> "" Then
       cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text) & vbCrLf
    End If
    
    If xgroup.MatchedWithList Then
       cWhere = cWhere & Tr(cWhere) & "  FILE1_10.[GROUP] = " & MyParn(xgroup.BoundText) & vbCrLf
    End If
       
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & "  FILE1_10.FACT = " & MyParn(xFact.BoundText) & vbCrLf
    
    If xSection.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & "  FILE1_10.SECTION = " & Val(xSection.BoundText) & vbCrLf
    
    If XBRNCH.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ITEM IN (SELECT ITEM FROM FILE1_11_ALL WHERE STORE = " & MyParn(XBRNCH.BoundText) & ")" & vbCrLf
    
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & "  FILE1_10.FACT = " & MyParn(xFact.BoundText) & vbCrLf
    
    If xdesca.text <> "" Then cWhere = cWhere & Tr(cWhere) & " " & MyParnAnd(xdesca.text, "file1_10.desca") & vbCrLf
    
    If xItem.text <> "" Then cWhere = cWhere & Tr(cWhere) & "  FILE1_10.item = " & Val(xItem.text) & vbCrLf
    
    If xbarcode.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( FILE1_10.BARCODE = " & MyParn(xbarcode.text) & " OR  FILE1_10.BARCODE_GS1 = " & MyParn(xbarcode.text) & " ) " & vbCrLf
    
    If Trim(xScal.text) <> "" Then cWhere = cWhere & Tr(cWhere) & "  FILE1_10.SCAL = " & MyParn(xScal.text) & vbCrLf
    
    If Trim(xColor.text) <> "" Then cWhere = cWhere & Tr(cWhere) & "  " & MyParnAnd(xColor, "FILE1_10.COLOR") & vbCrLf
    
    'If Check1.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " BARCODE IN (SELECT BARCODE FROM QD_BARCODE )" & vbCrLf
    If Check1.Value Then
        cWhere = cWhere & Tr(cWhere) & " FILE1_10.BARCODE IN (SELECT BARCODE FROM FILE1_10 AS F GROUP BY F.BARCODE HAVING COUNT(*) > 1  )"
    End If

    
    If XNOMOVE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " ITEM not IN (SELECT ITEM FROM FILE1_11) AND ITEM not IN (SELECT ITEM FROM FR1_11 ) " & vbCrLf
    
    If XNODEM.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " ISNODEM = 0 " & vbCrLf
    If xisstop.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " ISSTOP = 0 " & vbCrLf
    
    If xage.MatchedWithList Then
        cWhere = cWhere & Tr(cWhere) & " MODELAGE =  " & xage.BoundText & vbCrLf
    End If
    If xsex.MatchedWithList Then
        cWhere = cWhere & Tr(cWhere) & " MODELSEX =  " & xsex.BoundText & vbCrLf
    End If
    
    If chkBarCode.Value = 1 Then
        cWhere = cWhere & Tr(cWhere) & " BARCODE_GS1 IS NULL" & vbCrLf
    End If
    If cWhere <> "" Then cString = cString & vbCrLf & "WHERE " & cWhere
    
    cString = cString & " ORDER BY FACT.DESCA , FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.COLOR ,  FILE1_10.SCAL , PRICE_P"
    
    If Not bSql Then
       Set DATA6.Recordset = cmd(cString, con).Execute
    Else
       Clipboard.Clear
       Clipboard.SetText cString
       Inform "done ..."
       Exit Sub
    End If
    
    fixGrd bTemp
    CellPos 13, 0, grid1.Cols - 1
    StatusBar1.Panels(1).text = "⁄œœ «·«’‰«› : " & grid1.Rows - 1
End Sub
Sub fixGrd(Optional bTemp As Boolean = False)
With grid1
    .TextMatrix(0, 0) = "»«—ﬂÊœ"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "„ÊœÌ·"
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = " „ﬁ«”"
    .TextMatrix(0, 6) = " «··Ê‰"
    .TextMatrix(0, 7) = "«·ﬂ„Ì…"
    .TextMatrix(0, 7 + 1) = " ”⁄— «·»Ì⁄"
    .TextMatrix(0, 8 + 1) = "„Ã„Ê⁄… «·„ÊœÌ·"
    .TextMatrix(0, 9 + 1) = " ”⁄— √Êﬂ«“ÊÌ‰2"
        
    .TextMatrix(0, 10 + 1) = " »«—ﬂÊœ Œ«—ÃÏ"
    .TextMatrix(0, 11 + 1) = " »«—ﬂÊœ „’‰⁄"
    .TextMatrix(0, 12 + 1) = " »«—ﬂÊœ 13"
    .TextMatrix(0, 13 + 1) = "BARCODE GS1"

    .TextMatrix(0, 14 + 1) = "Description"
    .TextMatrix(0, 15 + 1) = "COLOR"
    .TextMatrix(0, 16 + 1) = "”⁄— ÿ»«⁄… 2"
    .TextMatrix(0, 17 + 1) = "„Ã„Ê⁄…"
    .TextMatrix(0, 18 + 1) = "„Ã„Ê⁄…"
    .TextMatrix(0, 19 + 1) = "«·ﬁ”„"
    .TextMatrix(0, 20 + 1) = "«·‰Ê⁄"
    .TextMatrix(0, 21 + 1) = "«·›∆… «·⁄„—Ì…"
            
    .ColHidden(7) = Not bTemp
    .ColHidden(9 + 1) = True
    .ColHidden(18 + 1) = True
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 2000
    .ColWidth(2) = 600
    .ColWidth(3) = 1500
    .ColWidth(4) = 4000
    .ColWidth(5) = 1000
    .ColWidth(6) = 2500
    .ColWidth(7 + 1) = 800
    .ColWidth(8 + 1) = 2000
    .ColWidth(9 + 1) = 900
    .FrozenCols = 7 + 1
    .ColWidth(10 + 1) = 2200
    .ColWidth(11 + 1) = 1000
    .ColWidth(12 + 1) = 1700
    .ColWidth(13 + 1) = 1700
    .ColWidth(14 + 1) = 3000
    .ColWidth(15 + 1) = 1100
    .ColWidth(16 + 1) = 1000
    .ColWidth(17 + 1) = 2000
    .ColWidth(19 + 1) = 2000
    .ColWidth(20 + 1) = 2000
    .ColWidth(21 + 1) = 2000
        
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTString
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
   ' If .Rows > 1 Then .Row = 1
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Sub myProc()
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
Dim cColor As String
Dim nRec As Long
On Error GoTo myerror
With grid1
    If col = 10 + 1 Then
        con.Execute "UPDATE FILE1_10 SET " & _
                   " BARCODE = " & addstring(.TextMatrix(Row, 10 + 1)) & _
                   " WHERE ITEM = " & Val(.TextMatrix(Row, 0))
    
        For sbr = 3 To nCountBranch
            cBr = RetZero(sbr, 2)
            con.Execute " INSERT INTO NEW_ITEM (ITEM ,  BRANCH) " & _
                        " VALUES ( " & Val(.TextMatrix(Row, 0)) & " , " & addstring(cBr) & ")"
        Next sbr
        
        For sbr = 61 To nCountBranch_fr
            cBr = RetZero(sbr, 2)
            con.Execute "  INSERT INTO NEW_ITEM_FR (ITEM ,  BRANCH)" & _
                        "  VALUES ( " & Val(.TextMatrix(Row, 0)) & " , " & addstring(cBr) & ")"
        Next sbr
    ElseIf col = 14 + 1 Then
        cString = " UPDATE FILE1_10 SET " & _
                    " DESCE = " & addstring(.TextMatrix(Row, 14 + 1)) & _
                    "  WHERE MODEL = " & MyParn(.TextMatrix(Row, 18 + 1))
        con.Execute cString, nRec
        myload
        'Inform "  „  ⁄œÌ· " & nRec
    ElseIf col = 15 + 1 Then
        cColor = .TextMatrix(Row, 15 + 1)
        con.Execute " UPDATE FILE1_10 SET " & _
                    "  COLOR_E = " & addstring(.TextMatrix(Row, 15 + 1)) & _
                    " WHERE MODEL = " & MyParn(.TextMatrix(Row, 18 + 1)) & _
                    " AND COLOR = " & MyParn(.TextMatrix(Row, 6)), nRec
        'Inform "  „  ⁄œÌ· " & nRec
    ElseIf col = 4 Then
        con.Execute " UPDATE FILE1_10 SET " & _
                    " DESCA = " & addstring(.TextMatrix(Row, 4)) & _
                    " WHERE MODEL = " & MyParn(.TextMatrix(Row, 18 + 1)), nRec
        'Inform "  „  ⁄œÌ· " & nRec
    ElseIf col = 6 Then
        con.Execute " UPDATE FILE1_10 SET " & _
                    " COLOR = " & addstring(.TextMatrix(Row, 6)) & _
                    " WHERE COLOR_E = " & MyParn(.TextMatrix(Row, 15 + 1)) & _
                    " AND MODEL = " & MyParn(.TextMatrix(Row, 18 + 1)), nRec
        'Inform "  „  ⁄œÌ· " & nRec
    ElseIf col = 13 + 1 Then
        con.Execute " UPDATE FILE1_10 SET " & _
                    " BARCODE_GS1 = " & addstring(.TextMatrix(Row, 13 + 1)) & _
                    " WHERE ITEM = " & .ValueMatrix(Row, 0)
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_DblClick()
    items.sModelItem = grid1.TextMatrix(grid1.Row, 0)
    items.Show 1
End Sub
Private Sub grid1_EnterCell()
With grid1
    If .col = 10 + 1 Or (.col = 13 + 1 And bOpt10) Or .col = 14 + 1 Or .col = 15 + 1 Or .col = 4 Or .col = 6 Then
        .Editable = flexEDKbdMouse
    ElseIf grid1.TextMatrix(0, 1) = "" And grid1.col = 0 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub

Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 13 Then
    If grid1.EditText = grid1.TextMatrix(Row, col) Then
        Cancel = True
    ElseIf Len(grid1.EditText) < 13 And Trim(grid1.EditText) <> "" Then
        Cancel = True
    End If
End If
End Sub

Private Sub xFACT_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Set Generalarray(0) = Me
Generalarray(1) = "Select code ,DescA From FACT"
Generalarray(2) = "Order by DESCA"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1

End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 And ActiveControl.Name <> grid1.Name Then
 '   cmdGo_Click
'End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'If grid1.Col = 3 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 3 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
     CellPos KeyCode, grid1.Row, grid1.col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If col < grid1.Cols - 7 Then
    If col < 4 Then
        grid1.col = 4
    ElseIf col < 6 Then
        grid1.col = 6
    ElseIf col < 10 + 1 Then
        grid1.col = 10 + 1
    ElseIf col < 13 + 1 Then
        grid1.col = 13 + 1
    Else
        grid1.col = col + 1
    End If
    grid1.ShowCell grid1.Row, grid1.col
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 4)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, col
End If
End Sub
Private Function myLoadCsv()
Dim cString As String
Dim i As Long

Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

Me.MousePointer = vbHourglass

'On Error GoTo myError

Dim cFileName As String
cFileName = Common1.FileName
If cFileName = "" Then
    GoTo Finally
End If

Dim cSv As New ChilkatCsv
nAccess = cSv.LoadFile(cFileName)

If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    GoTo Finally
End If


If cSv.NumRows < 2 Then Exit Function
Dim sb As New ChilkatStringBuilder

prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption

clearTemp con

cString = "create table #Temp" & _
           "(" & _
           "ITEM int," & _
           "ITEM2 NVARCHAR(50)," & _
           "QUANT INT" & _
           ")"
con.Execute cString
For Row = 1 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & Row & " „‰ " & cSv.NumRows - 1
    prog1.Value = Round(Row / (cSv.NumRows - 1), 2) * 100
    
    i = i + 1
    sb.Append "(" & _
              "COALESCE([dbo].[fn_ret_item](" & addstring(cSv.GetCell(Row, 0)) & "),'0')" & "," & _
               addstring(cSv.GetCell(Row, 0)) & "," & _
               Val(cSv.GetCell(Row, 1)) & _
              "),"
                  
    If i = 100 Then
        i = 0
        sb.Shorten 1
        con.Execute "INSERT INTO #temp(ITEM,ITEM2,QUANT)" & _
                     "VALUES" & _
                     sb.GetAsString
        sb.Clear
    End If
Next

sb.Shorten 1
If sb.length > 0 Then
    con.Execute "INSERT INTO #temp(ITEM,ITEM2,QUANT)" & _
                 "VALUES" & _
                 sb.GetAsString
End If

Me.MousePointer = vbNormal
prog1.Visible = False

Inform " „  Õ„Ì· «·„·›"

myload True
Finally:
clearTemp con
prog1.Visible = False
Me.MousePointer = vbNormal
Exit Function
myerror:
MsgBox Err.Description
GoTo Finally
End Function

