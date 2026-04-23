VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form arrivedatafrm 
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15660
   BeginProperty Font 
      Name            =   "Arabic Transparent"
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
   ScaleHeight     =   11055
   ScaleWidth      =   15660
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   510
      ScaleWidth      =   15660
      TabIndex        =   76
      Top             =   10545
      Width           =   15660
      Begin VB.CheckBox chkMachine 
         Appearance      =   0  'Flat
         Caption         =   "«ŸÂ«— Õ÷Ê— «·„«ﬂÌ‰…"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   135
         TabIndex        =   77
         Top             =   90
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin Threed.SSCommand cmdRemoveAbsLate 
         Height          =   330
         Left            =   2520
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   90
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   582
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
         Caption         =   "Õ–› Ã“«¡«  «· √ŒÌ—"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFixAbsLate 
         Height          =   330
         Left            =   5265
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   90
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   582
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
         Caption         =   "«÷«›… Ã“«¡«  «· √ŒÌ—"
         ButtonStyle     =   3
      End
   End
   Begin VB.PictureBox pctEmp 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   8625
      Left            =   12495
      RightToLeft     =   -1  'True
      ScaleHeight     =   8625
      ScaleWidth      =   3165
      TabIndex        =   72
      Top             =   1770
      Width           =   3165
      Begin VSFlex7Ctl.VSFlexGrid Grid2 
         Height          =   7680
         Left            =   45
         TabIndex        =   73
         Top             =   900
         Width           =   3060
         _cx             =   5397
         _cy             =   13547
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
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin Threed.SSCommand cmdEmp 
         Height          =   375
         Left            =   90
         TabIndex        =   74
         Top             =   0
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmdBranch 
         Height          =   420
         Left            =   90
         TabIndex        =   75
         Top             =   405
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   741
         _Version        =   196610
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ﬂ· «·›—Ê⁄"
         ButtonStyle     =   4
      End
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   " €ÌÌ—  ·ﬁ«∆∆"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   -270
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   1140
      Left            =   -2835
      TabIndex        =   1
      Top             =   -855
      Visible         =   0   'False
      Width           =   3075
      Begin VB.CheckBox chkPenality 
         Appearance      =   0  'Flat
         Caption         =   "»Â«  √ŒÌ—« "
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
         TabIndex        =   4
         Top             =   270
         Width           =   1545
      End
      Begin VB.CheckBox chkMissing 
         Appearance      =   0  'Flat
         Caption         =   "€Ì— „ﬂ „·… «·»Ì«‰« "
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
         TabIndex        =   3
         Top             =   675
         Width           =   1905
      End
      Begin VB.CheckBox chkNoData 
         Appearance      =   0  'Flat
         Caption         =   "€Ì— «·„”Ã·Ì‰"
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
         Left            =   1575
         TabIndex        =   2
         Top             =   270
         Width           =   1410
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Command1"
      Height          =   600
      Left            =   -1260
      TabIndex        =   0
      Top             =   -450
      Visible         =   0   'False
      Width           =   2445
   End
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   10395
      Top             =   -270
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   12510
      Top             =   -270
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   135
      Top             =   -90
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   90
      Top             =   -180
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
   Begin MSAdodcLib.Adodc DATA11 
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   1035
      Top             =   -225
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   6
      Top             =   10395
      Visible         =   0   'False
      Width           =   15660
      _ExtentX        =   27623
      _ExtentY        =   265
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   1770
      ScaleWidth      =   15660
      TabIndex        =   7
      Top             =   0
      Width           =   15660
      Begin Threed.SSCommand cmdConnect 
         Height          =   375
         Left            =   90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
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
         Caption         =   "Connect Server"
         ButtonStyle     =   3
      End
      Begin VB.CommandButton cmdFixData 
         Caption         =   "÷»ÿ Õ÷‹‹Ê— Ê«‰’—«›"
         Height          =   600
         Left            =   15885
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   45
         Width           =   1320
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   17280
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   -45
         Width           =   3030
         Begin VB.CommandButton CmdDelInv 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1710
            MaskColor       =   &H00FFFFFF&
            Picture         =   "visa_search.frx":0000
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.CommandButton cmdExit 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   45
            Picture         =   "visa_search.frx":289A
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   135
            Width           =   1635
         End
      End
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   17325
         TabIndex        =   60
         Top             =   630
         Width           =   2985
         Begin VB.TextBox xdate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1125
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   585
            Width           =   1275
         End
         Begin VB.TextBox xDate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1125
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   225
            Width           =   1275
         End
         Begin VB.CommandButton cmdGo 
            Height          =   330
            Left            =   180
            Picture         =   "visa_search.frx":4D06
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "⁄—÷"
            Top             =   225
            Width           =   915
         End
         Begin VB.CheckBox xCurrent 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "«·Õ«·Ì"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   675
            Width           =   870
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   270
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   225
            Width           =   210
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Õ Ï"
            Height          =   270
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   630
            Width           =   300
         End
         Begin VB.Label xDay1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   -360
            Visible         =   0   'False
            Width           =   780
         End
      End
      Begin VB.Frame Frame_Abs 
         Caption         =   "€Ì«»"
         Height          =   1455
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   270
         Width           =   2535
         Begin VB.CommandButton cmdAbsent 
            Caption         =   " ÕÊÌ·"
            Height          =   375
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   990
            Width           =   2355
         End
         Begin VB.CheckBox xAbsent 
            Appearance      =   0  'Flat
            Caption         =   "÷»ÿ «Ì«„ «·€Ì«»"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   135
            TabIndex        =   57
            Top             =   270
            Width           =   1545
         End
         Begin MSDataListLib.DataCombo xType 
            Height          =   330
            Left            =   90
            TabIndex        =   59
            Top             =   630
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
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   15885
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   585
         Width           =   1365
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
            Picture         =   "visa_search.frx":71F8
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Õ›Ÿ"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.CommandButton CmdUndo 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   45
            MaskColor       =   &H00FFFFFF&
            Picture         =   "visa_search.frx":955B
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   675
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.Frame Frame_comp 
         Caption         =   "÷»ÿ «·»Ì«‰«  €Ì— «·ﬂ«„·…"
         Height          =   1230
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   495
         Width           =   2040
         Begin VB.CommandButton Command3 
            BackColor       =   &H0080FF80&
            Caption         =   "÷»ÿ"
            Height          =   510
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   630
            Width           =   1860
         End
         Begin VB.CheckBox xNo_data 
            Appearance      =   0  'Flat
            Caption         =   "Õ÷Ê— »€Ì— «‰’—«›"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   90
            TabIndex        =   51
            Top             =   315
            Width           =   1905
         End
      End
      Begin VB.Frame Frame_late 
         Caption         =   " «ŒÌ—«  „‰"
         Height          =   1140
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   585
         Width           =   1320
         Begin VB.CommandButton cmdFixLate 
            Caption         =   "÷»ÿ ﬂ· «·«Ì«„"
            Height          =   375
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   720
            Width           =   1140
         End
         Begin VB.TextBox xLate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   390
            Left            =   405
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   270
            Width           =   825
         End
         Begin VB.CheckBox chkLate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   315
            Width           =   240
         End
      End
      Begin VB.Frame Frame_h 
         Height          =   1725
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   0
         Width           =   7980
         Begin VB.CommandButton cmdAbsShow 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   2970
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   945
            Width           =   330
         End
         Begin VB.CommandButton cmdAbsShow 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   225
            Width           =   330
         End
         Begin VB.CommandButton cmdAbsShow 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   585
            Width           =   330
         End
         Begin VB.CommandButton cmdAbsShow 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   945
            Width           =   330
         End
         Begin VB.CommandButton cmdAbsShow 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   1305
            Width           =   330
         End
         Begin VB.CommandButton cmd_emp 
            Caption         =   "»Ì«‰«  «·„ÊŸ›"
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
            Left            =   4365
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   225
            Width           =   1230
         End
         Begin VB.Label xBreak 
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
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   1305
            Width           =   555
         End
         Begin VB.Label xDays_off2 
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
            Left            =   405
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   225
            Width           =   645
         End
         Begin VB.Label xDays_temp2 
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
            Left            =   405
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   585
            Width           =   645
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "„—÷Ì"
            Height          =   270
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   945
            Width           =   495
         End
         Begin VB.Label xdays_Remedy 
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
            Left            =   4005
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   945
            Width           =   600
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "—«Õ…"
            Height          =   270
            Left            =   7155
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   1350
            Width           =   360
         End
         Begin VB.Label xDays_absent 
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
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   945
            Width           =   645
         End
         Begin VB.Label xBreak2 
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
            Left            =   5850
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   1305
            Width           =   600
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ÌÊ„ —«Õ…"
            Height          =   270
            Left            =   7110
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   990
            Width           =   660
         End
         Begin VB.Label xDay_Break 
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
            Left            =   5220
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   945
            Width           =   1815
         End
         Begin VB.Label xCode 
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
            Left            =   5625
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   225
            Width           =   1410
         End
         Begin VB.Label xDesca 
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
            Left            =   3285
            TabIndex        =   34
            Top             =   585
            Width           =   3750
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "«·«”„"
            Height          =   270
            Left            =   7155
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   585
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "«·—ﬁ„"
            Height          =   270
            Left            =   7140
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   270
            Width           =   360
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "€Ì«»"
            Height          =   270
            Left            =   1845
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   990
            Width           =   390
         End
         Begin VB.Label xBreak_all 
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
            Left            =   5220
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   1305
            Width           =   600
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "√Ã«“… ⁄«—÷…"
            Height          =   270
            Left            =   1815
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   675
            Width           =   1005
         End
         Begin VB.Label xDays_Temp 
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
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   585
            Width           =   645
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "√Ã«“… «⁄ Ì«œÌ"
            Height          =   270
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   270
            Width           =   1050
         End
         Begin VB.Label xDays_off 
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
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   225
            Width           =   645
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "«Ã«“… »Œ’„"
            Height          =   270
            Left            =   1845
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1350
            Width           =   900
         End
         Begin VB.Label xDays_Discount 
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
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1305
            Width           =   645
         End
         Begin VB.Label xDays_absent2 
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
            Left            =   405
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   945
            Width           =   645
         End
         Begin VB.Label xDays_Discount2 
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
            Left            =   405
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1305
            Width           =   645
         End
         Begin VB.Label xdays_Remedy2 
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
            Left            =   3330
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   945
            Width           =   645
         End
         Begin VB.Label xdays_abs_late 
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
            Left            =   3330
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   1305
            Width           =   645
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "€Ì‹‹‹«»  «ŒÌ‹‹‹—"
            Height          =   270
            Left            =   4065
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   1350
            Width           =   1110
         End
      End
      Begin VB.Frame Frame3 
         Height          =   645
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1035
         Width           =   1725
         Begin VB.CheckBox chkExc 
            Appearance      =   0  'Flat
            Caption         =   "«Ì«„ «–‰"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   225
            TabIndex        =   10
            Top             =   225
            Width           =   1275
         End
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Align           =   3  'Align Left
      Height          =   8625
      Left            =   0
      TabIndex        =   71
      Top             =   1770
      Width           =   16140
      _cx             =   28469
      _cy             =   15214
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
      BackColorAlternate=   8180472
      GridColor       =   12632256
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   55
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
Attribute VB_Name = "arrivedatafrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sEmp_Code As String
Dim cList As String, cList2 As String, bChange As Boolean
Dim bError As Boolean
Dim cList3 As String, cList4 As String
Dim cList5 As String
Dim aShift As Variant
Dim formMode, emptable As New ADODB.Recordset, shiftTable As New ADODB.Recordset
Dim bResize As Boolean
Dim oSearch As New Search, osearchBranch As Search, oSearchEmp As Search
Dim bRamadan As Boolean
Dim con As New ADODB.Connection
Dim nSort As Integer
Dim bNoReplace As Boolean
Dim cFilesave As String
Dim bNoLoadGrd As Boolean
Dim nMinWidth As Long, nMinHeight As Long
Const LoadMode = 0, DefineMode = 1
Private Sub checkNoData_KeyUp(KeyCode As Integer, Shift As Integer)
Fixgrd
End Sub
Private Sub chkAllEmp_Click()
myLoadGrd2
End Sub
Private Sub chkExc_Click()
myLoadGrd2 True
End Sub
Private Sub chkLate_Click()
xNo_data.Value = 0
myLoadGrd2 True
End Sub

Private Sub chkName_Click()
myLoadGrd2
End Sub

Private Sub cmd_emp_Click()
employfrm.sCode = Xcode.Caption
employfrm.Show
End Sub
Private Sub cmdAbsent_Click()
On Error GoTo myerror
If Not xType.MatchedWithList Then Exit Sub
If MsgBox(" ÕÊÌ· ﬂ· «·€Ì«» «·Ì " & xType.text, vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
Dim i As Long, nBreak As Long, nTaken As Long, nDays_Left As Long, nNeeded As Long
If Not xType.MatchedWithList Then Exit Sub

If Val(xType.BoundText) = 3 Or Val(xType.BoundText) = 4 Or Val(xType.BoundText) = 9 Then
    nDays_Left = Days_left(xType.BoundText, Xcode.Caption, Val(xBreak.Caption), con)
    For i = 1 To grid1.Rows - 1
        If Val(grid1.TextMatrix(i, 8)) = "2" Then
            nNeeded = nNeeded + 1
        End If
    Next
    If nNeeded > nDays_Left Then
        Exit Sub
    End If
ElseIf Val(xType.BoundText) = 10 Then
    For i = 1 To grid1.Rows - 1
        If Val(grid1.TextMatrix(i, 8)) = "2" Then
            If Val(grid1.TextMatrix(i, 24)) = 0 Then
'                MsgBox "ÌÊ„ " & myFormat(grid1.TextMatrix(i, 0)) & " ·Ì” ⁄ÿ·… —”„Ì…"
                Exit Sub
            End If
        End If
    Next
End If

For i = 1 To grid1.Rows - 1
    If Val(grid1.TextMatrix(i, 8)) = "2" Then
        grid1.TextMatrix(i, 8) = xType.BoundText
        grid1.TextMatrix(i, 9) = xType.BoundText
        CheckRow i, 8
'        grid1.TextMatrix(i, 29) = xType.BoundText
'        grid1.TextMatrix(i, 30) = xType.BoundText
        myReplacegrd i
    End If
Next
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdAbsShow_Click(Index As Integer)
If GRID2.Row < 1 Then Exit Sub
Dim aLocal As Variant
If Index = 0 Then
    aLocal = AddFlag(aLocal, "Header1", "«Ì«„ €Ì«» «·„ÊŸ› : " & XDESCA.Caption)
    aLocal = AddFlag(aLocal, "TYPE", "2")
ElseIf Index = 1 Then
    aLocal = AddFlag(aLocal, "Header1", "«Ì«„ «Ã«“… «⁄ Ì«œÌ ··„ÊŸ› : " & XDESCA.Caption)
    aLocal = AddFlag(aLocal, "TYPE", "4")
ElseIf Index = 2 Then
    aLocal = AddFlag(aLocal, "Header1", "«Ì«„ «Ã«“… ⁄«—÷… ··„ÊŸ› : " & XDESCA.Caption)
    aLocal = AddFlag(aLocal, "TYPE", "9")
ElseIf Index = 3 Then
    aLocal = AddFlag(aLocal, "Header1", "«Ì«„ «Ã«“… „—÷Ï ··„ÊŸ› : " & XDESCA.Caption)
    aLocal = AddFlag(aLocal, "TYPE", "5")
ElseIf Index = 4 Then
    aLocal = AddFlag(aLocal, "Header1", "«Ì«„ «Ã«“… »Œ’„ ··„ÊŸ› : " & XDESCA.Caption)
    aLocal = AddFlag(aLocal, "TYPE", "15")
End If

aLocal = AddFlag(aLocal, "CODE", Xcode.Caption)
aLocal = AddFlag(aLocal, "WHERE", "YEAR = " & sYear_period)
aLocal = AddFlag(aLocal, "DATE2", xDate2.text)
aLocal = AddFlag(aLocal, "Header2", "Œ·«· ⁄«„ : " & sYear_period)
show_holidaysfrm.aData = aLocal
show_holidaysfrm.Show 1
End Sub

Private Sub cmdBranch_Click()
Set osearchBranch = New Search
BranchLookup Me, osearchBranch, , , , cmdBranch.Tag <> ""
End Sub

Private Sub cmdConnect_Click()
ConnectServer con, Me
End Sub

Private Sub cmdDelinv_Click()
If Not ValidNum(Xcode.Caption) Then
    MsgBox "ﬂÊœ «·„ÊŸ› €Ì— „”Ã·"
    Exit Sub
End If

If MsgBox("Õ–› »Ì«‰«  «·Õ÷Ê— Ê«·«‰’—› ··„ÊŸ› „‰– »œ«Ì… «·„œ…", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
On Error GoTo myerror
If DelArrive Then
    Inform " „ Õ–› «·»Ì«‰« "
    myLoadGrd
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function DelArrive() As Boolean
con.BeginTrans
On Error GoTo myerror
con.Execute "update hr.checktime set posted = 0 from hr.checktime where checktime.[date] >= " & DateSq(xDate1.text) & " and userid = " & Xcode.Caption, nRecord
con.Execute "Delete From hr.FILE2_50 where date >= " & DateSq(xDate1.text) & " and code =  " & Xcode.Caption, nRecord
con.CommitTrans
prog1.Visible = False
DelArrive = True
Exit Function
myerror:
prog1.Visible = False
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Sub cmdEmp_Click()
On Error GoTo myerror
Set oSearchEmp = New Search
Dim cWhere As String
cWhere = "FILE1_20.MANUAL = 0"
cWhere = cWhere & " AND (FILE1_20.DATE_BEGIN IS NULL OR FILE1_20.DATE_BEGIN <= " & DateSq(xDate2.text) & ")"
cWhere = cWhere & " AND (FILE1_20.DATE_END IS NULL OR FILE1_20.DATE_END >= " & DateSq(xDate1.text) & ")"
    
If FilterEmp Then
    Dim i As Long
    
    Dim sbJson As New ChilkatStringBuilder
    For i = 1 To GRID2.Rows - 1
        sbJson.Append ","
        sbJson.Append GRID2.TextMatrix(i, 0)
    Next
    sbJson.RemoveCharsAt 0, 1
    
    If sbJson.Length > 0 Then
        cWhere = cWhere & " and FILE1_20.code in (" & sbJson.GetAsString() & ")"
    End If
End If



If cmdBranch.Tag <> "" Then
    cWhere = cWhere & " AND " & "FILE1_20.MANAGE = " & cmdBranch.Tag
End If
EmpLookup Me, oSearchEmp, cWhere, , True, , cmdEmp.Tag <> ""
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdInform_Click()
CardLookup
End Sub
Private Sub cmdFixAbsLate_Click()
Me.MousePointer = vbHourglass
If RemoveAbsLate(cmdEmp.Tag, cmdBranch.Tag) Then
    fixArrive sDate_Period, sDate_Period_End, con, prog1, cmdEmp.Tag, cmdBranch.Tag
    AddAbsLate cmdEmp.Tag, cmdBranch.Tag
    fixArrive sDate_Period, sDate_Period_End, con, prog1, cmdEmp.Tag, cmdBranch.Tag, True
    Inform " „  «÷«›… «Ì«„ «· «ŒÌ—«  »‰Ã«Õ"
    myLoadGrd
End If
Me.MousePointer = vbNormal
End Sub

Private Sub cmdFixData_Click()
On Error GoTo myerror
Dim bFixHolidays As Boolean
bFixHolidays = MsgBox("÷»ÿ «·«Ã«“« ", vbOKCancel + vbDefaultButton2) = vbOK
Me.MousePointer = vbHourglass
FIXDATA xDate1.text, xDate2.text, con, prog1, , bFixHolidays, , Me
Me.MousePointer = 0
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdFixLate_Click()
On Error GoTo myerror
If Val(xLate.text) < 1 Or chkLate.Value = 0 Then Exit Sub
If MsgBox("÷»ÿ ﬂ· «Ì«„  «· √ŒÌ— " & xType.text, vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
Fix_late
myLoadGrd2
myLoadGrd
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmdRemoveAbsLate_Click()
Me.MousePointer = vbHourglass
If RemoveAbsLate(cmdEmp.Tag, cmdBranch.Tag) Then
    fixArrive sDate_Period, sDate_Period_End, con, prog1, cmdEmp.Tag, cmdBranch.Tag
    Inform " „  «÷«›… «Ì«„ «· «ŒÌ—«  »‰Ã«Õ"
    myLoadGrd
End If
Me.MousePointer = vbNormal
End Sub

Private Sub cmdSave_Click()
On Error GoTo myerror
If Not myvalid Then Exit Sub
If myReplacegrd Then Inform " „ «·Õ›Ÿ »‰Ã«Õ"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdGo_Click()
myLoadGrd2 True
End Sub
Private Sub Command3_Click()
On Error GoTo myerror
If xNo_data.Value = 0 Then Exit Sub
If MsgBox("÷»ÿ ﬂ· «·«Ì«„ «·€Ì— „ﬂ „·…", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
Fix_Incomplete
myLoadGrd2 True
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Activate()
If bError Then Unload Me
'grid1.Width = Me.Width - 4000
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Not (ActiveControl.Name = grid1.Name And (grid1.Col = 8 Or grid1.Col = 9 Or grid1.Col = 32 Or grid1.Col = 44))) Then
        KeyAscii = 0
    End If
End If
End Sub
Private Sub Form_Load()
On Error GoTo myerror
xLate.text = Myvalue(Val(RetSetting("late", TempSave(Me))))
xDate1.text = sDate_Period
xDate2.text = IIf(xCurrent.Value = 1, sDate_Period_End, myFormat(Date))

bRamadan = False

openCon con

Set data4.Recordset = myCmd("Select * from hr.AR_TYPE WHERE STOP = 0 Order by [ORDER]", con)
Set xType.RowSource = data4
xType.ListField = "Desca"
xType.BoundColumn = "Code"


cList = StrList2("select code,Desca from hr.AR_TYPE WHERE STOP = 0 ORDER BY [ORDER]", con)
cList2 = StrList2("select code,Desca from hr.AR_TYPE WHERE IS_HALF = 1 ORDER BY [ORDER] ", con)
cList3 = StrList2("select code,desca from hr.time_code  where STOP = 0 ORDER BY [CODE]", con)
cList4 = StrList2("SELECT CODE,DESCA FROM hr.MACHINES WHERE STOPED = 0", con)


Fixgrd

Set grid1.DataSource = data10
Set GRID2.DataSource = DATA11


If sEmp_Code = "" Then LoadText Me

If cmdBranch.Tag <> "" Then
    sField = myField("select desca from hr.MANAGE where code = " & cmdBranch.Tag, con) & ""
    If sField <> "" Then
        cmdBranch.Caption = sField
    Else
        cmdBranch.Tag = ""
    End If
End If

If sEmp_Code <> "" Then
    cmdEmp.Tag = sEmp_Code
    cmdEmp.Caption = myField("select desca from hr.FILE1_20 where code =" & sEmp_Code, con)
End If

myLoadGrd2
Exit Sub
myerror:
bError = True
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Resize()
If Not bResize Then
    nMinWidth = grid1.Width
    nMinHeight = GRID2.Height
    bResize = True
End If
grid1.Width = IIf(Me.ScaleWidth - grid1.Left - pctEmp.Width - 50 < nMinWidth, nMinWidth, Me.ScaleWidth - grid1.Left - pctEmp.Width - 50)
'Grid2.Height = IIf(pctEmp.Height - 900 < nMinHeight, nMinHeight, pctEmp.Height - 900)
GRID2.Height = pctEmp.Height - 900
FixWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If sEmp_Code = "" Then
    SaveText Me, , Array(xLate.Name, cmdBranch.Name)
End If
Err.Clear
Set arrivedatafrm = Nothing
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
If Row = 1 Then Exit Sub
If Col = 6 And IsTime(grid1.TextMatrix(Row, 6)) And Trim(grid1.TextMatrix(Row, 44)) = "" Then
    aRet = Ret_Shift(grid1.TextMatrix(Row, 6), myFormat(grid1.TextMatrix(Row, 0)), con)
    If Not IsEmpty(aRet) Then
        grid1.TextMatrix(Row, 44) = retFlag(aRet, "code")
        grid1.TextMatrix(Row, 2) = retFlag(aRet, "time1")
        grid1.TextMatrix(Row, 3) = retFlag(aRet, "time2")
    End If
ElseIf Col = 44 Then
    ChangeShift Row
End If

CheckRow Row, Col

RowDgt Row

myReplacegrd Row
If (Col = 8 Or Col = 9) Then
    myLoadHolidays
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub ChangeShift(Row As Long)
Dim aPrm As Variant

On Error GoTo myerror
If grid1.TextMatrix(Row, 44) = "" Then
    aPrm = AddFlag(Empty, "DATE", myFormat_sp(grid1.TextMatrix(Row, 0)))
    aPrm = AddFlag(aPrm, "CODE", Val(grid1.TextMatrix(Row, grid1.Cols - 2)))

    Dim loctable As ADODB.Recordset
    Set loctable = myCmd("hr.sp_def_shift", con, adStoredProc, aPrm)
    If Not loctable.EOF Then
        grid1.TextMatrix(Row, 44) = loctable!shift_code & ""
        grid1.TextMatrix(Row, 2) = loctable!time_begin & ""
        grid1.TextMatrix(Row, 3) = loctable!time_End & ""
    End If
    Exit Sub
End If

aPrm = AddFlag(Empty, "DATE", myFormat_sp(grid1.TextMatrix(Row, 0)))
aPrm = AddFlag(aPrm, "CODE", Val(grid1.TextMatrix(Row, grid1.Cols - 2)))
aPrm = AddFlag(aPrm, "SHIFT", Val(grid1.TextMatrix(Row, 44)))


Dim cmd As ADODB.Command
Set cmd = mycmdEx("hr.sp_ChangeShift", con, aPrm, 300)

If Not IsNull(cmd.Parameters("@Error_msg")) Then
    MsgBox cmd.Parameters("@Error_msg")
    Exit Sub
ElseIf cmd.Parameters("@Count") > 0 Then
    Inform " „  ⁄œÌ· «·‘›  »‰Ã«Õ"
End If

grid1.TextMatrix(Row, 2) = cmd.Parameters("@TIME1") & ""
grid1.TextMatrix(Row, 3) = cmd.Parameters("@TIME2") & ""

Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_AfterSort(ByVal Col As Long, Order As Integer)
nSort = Col
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If grid1.TextMatrix(Row, grid1.Cols - 1) <> "" Then
    Set times_excfrm.myForm = Me
    times_excfrm.sid = grid1.TextMatrix(Row, grid1.Cols - 1)
    times_excfrm.Show 1
End If
End Sub
Private Sub grid1_DblClick()
With grid1
If .Col = 2 And .TextMatrix(.Row, 2) <> "" And (.TextMatrix(.Row, 6) = "") Then
    .TextMatrix(.Row, 6) = .TextMatrix(.Row, 2)
    Grid1_AfterEdit grid1.Row, 6
ElseIf .Col = 3 And .TextMatrix(.Row, 7) = "" Then
    .TextMatrix(.Row, 7) = .TextMatrix(.Row, 3)
    Grid1_AfterEdit grid1.Row, 7
End If
End With
End Sub
Private Sub Grid1_EnterCell()
With grid1
If .Row = 1 Then
    .Editable = flexEDNone
ElseIf .Col = 6 Or .Col = 7 Or .Col = 8 Or .Col = 9 Or .Col = 25 Or .Col = 30 Or .Col = 31 Or .Col = 32 Or .Col = 42 Or .Col = 43 Or .Col = 44 Or .Col = 45 Or .Col = 46 Or .Col = 47 Or (.Col = 48 And .ValueMatrix(.Row, 50) = 0) Or .Col = .Cols - 3 Then
    .Editable = flexEDKbdMouse
Else
    .Editable = flexEDNone
End If
.Cell(flexcpBackColor, 0, 0, , .Cols - 1) = &H0&
.Cell(flexcpBackColor, 0, .Col, 0, .Col) = vbRed
'Me.Caption = grid1.Col
End With
End Sub

Private Sub Grid1_GotFocus()
Grid1_EnterCell
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1


Dim sType As String, sType_org As String, nDays As Long
If Col = 2 Or Col = 3 Or Col = 6 Or Col = 7 Or Col = 26 Or Col = 27 Or Col = 30 Or Col = 31 Then
    .EditText = RetTime(.EditText)
    If Not IsTime(.EditText) Then .EditText = ""
End If
If (Col = 8 Or Col = 32) And Trim(grid1.EditText) = "" Then
    .EditSelText = ""
ElseIf Col = 8 Or Col = 32 Then
    If Col = 8 Then
        sType = ret_Type_desca(.EditText, "code")
        sType_org = .TextMatrix(Row, Col)
    Else
        sType = ret_Type_desca2(.EditText, "code")
        sType_org = .TextMatrix(Row, Col)
    End If
    If Val(sType) = 10 Then
        If Not Is_Holiday(myFormat(.TextMatrix(Row, 0)), con) Then
'           MsgBox "«·ÌÊ„ ·Ì” " & .EditText
            Cancel = True
            Exit Sub
        End If
    End If
    If ret_Type(sType, "No_data") Then
        If Col = 8 Then
            .TextMatrix(Row, 6) = ""
            .TextMatrix(Row, 7) = ""
            .TextMatrix(Row, 30) = ""
            .TextMatrix(Row, 31) = ""
            .TextMatrix(Row, 32) = ""
            .TextMatrix(Row, 39) = ""
            .TextMatrix(Row, 40) = ""
            .TextMatrix(Row, 41) = ""
            .TextMatrix(Row, 42) = ""
            .TextMatrix(Row, 43) = 0
        Else
'            .TextMatrix(Row, 30) = ""
'            .TextMatrix(Row, 31) = ""
'            .TextMatrix(Row, 32) = ""
        End If
        .TextMatrix(Row, 48) = ""
    ElseIf ret_Type(sType_org, "No_data") Then
        If Col = 8 Then
            .TextMatrix(Row, 6) = .TextMatrix(Row, 4)
            .TextMatrix(Row, 7) = .TextMatrix(Row, 5)
            .TextMatrix(Row, 30) = .TextMatrix(Row, 28)
            .TextMatrix(Row, 31) = .TextMatrix(Row, 29)
        Else
            .TextMatrix(Row, 30) = .TextMatrix(Row, 28)
            .TextMatrix(Row, 31) = .TextMatrix(Row, 29)
        End If
    End If
                
    '    If (Val(sType)) = 3 Then
    '        nDays = Val(xBreak2.Caption) - Val(xBreak.Caption)
    '        If nDays < 1 Then
    '    '        MsgBox "·« ÌÊÃœ —’Ìœ " & .EditText & " ··„ÊŸ›"
    '            'Cancel = True
    '        Else
    '            Inform "»«ﬁÌ " & nDays - 1
    '        End If
    '    End If
End If




If Col <> 8 Then sType = .TextMatrix(Row, 8)
'If Col <> 47 Then
'    If ret_Type(sType, "commission") And grid1.TextMatrix(Row, 47) = "" Then
'        grid1.TextMatrix(Row, 47) = grid1.TextMatrix(Row, 49)
'    ElseIf (Not ret_Type(sType, "commission")) And grid1.TextMatrix(Row, 47) <> "" Then
'        grid1.TextMatrix(Row, 47) = ""
'    End If
'End If
'If Col <> 48 Then
'    If (Not ret_Type(sType, "No_Data")) And grid1.TextMatrix(Row, 48) = "" Then
'        grid1.TextMatrix(Row, 48) = grid1.TextMatrix(Row, 49)
'    ElseIf (ret_Type(sType, "No_Data")) And grid1.TextMatrix(Row, 48) <> "" Then
'        grid1.TextMatrix(Row, 48) = ""
'    End If
'End If

End With
End Sub
Private Sub myload2()
If IsDate(xDate.text) Then
    myLoadGrd
Else
    grid1.Rows = 1
    myDefineHolidays
    xRowCount.Caption = ""
End If
End Sub
Private Sub Handlecontrols(nMode)
CmdDelInv.Enabled = nMode = LoadMode And bedit
'cmdFirst.Enabled = (nMode = LoadMode)
'cmdLast.Enabled = (nMode = LoadMode)
'cmdNext.Enabled = (nMode = LoadMode)
'cmdPrevious.Enabled = (nMode = LoadMode)
'xDoc_No.Enabled = (nMode = DefineMode)
xDate.Enabled = (nMode = DefineMode)
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 46 And grid1.Row <> 0 Then
'    If MsgBox("«·€«¡ »Ì«‰«  «·Õ÷Ê— Ê«·«‰’—«› ··„ÊŸ› ?", vbOKCancel + vbDefaultButton2) = vbOK Then
'        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
'            con.BeginTrans
'            con.Execute "DELETE FROM hr.FILE2_50 WHERE ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
'            con.CommitTrans
'            MyLoadgrd
'        End If
'    End If
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 112 Then
    If (grid1.TextMatrix(grid1.Row, 6) = "" Or grid1.TextMatrix(grid1.Row, 7) = "") And (grid1.TextMatrix(grid1.Row, 6) <> "" Or grid1.TextMatrix(grid1.Row, 7) <> "") Then
        If grid1.TextMatrix(grid1.Row, 7) = "" Then
            grid1.TextMatrix(grid1.Row, 7) = grid1.TextMatrix(grid1.Row, 6)
            grid1.TextMatrix(grid1.Row, 6) = ""
        Else
            grid1.TextMatrix(grid1.Row, 6) = grid1.TextMatrix(grid1.Row, 7)
            grid1.TextMatrix(grid1.Row, 7) = ""
        End If
        Grid1_AfterEdit grid1.Row, grid1.Col
    End If
ElseIf KeyCode = 46 Then
    If grid1.Col = 47 Then
        grid1.TextMatrix(grid1.Row, grid1.Col) = ""
    End If
End If
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(0, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  Format([DATE],'dd-mm-yyyy') FROM hr.SHIFTDAY"

Generalarray(2) = "Order by [date]"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "[date]"
listarray(0, 1) = "##Date##"

GrdArray(0, 0) = "Date"
GrdArray(0, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
Search.Caption = "«” ⁄·«„"
Search.Show 1
End Sub
Private Sub Fixgrd()
Dim nWeek As Integer
With grid1
.TextMatrix(0, 0) = "«· «—ÌŒ"
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 2) = "Õ÷Ê— —”„Ì"
.TextMatrix(0, 3) = "«‰’—«› —”„Ì"
.TextMatrix(0, 4) = "Õ÷Ê— „«ﬂÌ‰…"
.TextMatrix(0, 5) = "«‰’—«› „«ﬂÌ‰…"
.TextMatrix(0, 6) = "Õ÷Ê— „”Ã·"
.TextMatrix(0, 7) = "«‰’—«› „”Ã·"

.TextMatrix(0, 8) = "«· „«„"
.TextMatrix(0, 9) = "‰Ê⁄ «·«‰’—«›"

.TextMatrix(0, 10) = "Õ÷Ê— »⁄œ"
.TextMatrix(0, 11) = "«‰’—«› »⁄œ"
.TextMatrix(0, 12) = "Ã“«¡"
.TextMatrix(0, 13) = "«÷«›Ì"
.TextMatrix(0, 14) = "”«⁄«  —”„Ì…"
.TextMatrix(0, 15) = "”«⁄«  ›⁄·Ì…"

.TextMatrix(0, 16) = "Õ÷Ê— »⁄œ"
.TextMatrix(0, 17) = "«‰’—› »⁄œ"
.TextMatrix(0, 18) = "Ã“«¡"
.TextMatrix(0, 19) = "«÷«›Ì"
.TextMatrix(0, 20) = "”«⁄«  —”„Ì…"
.TextMatrix(0, 21) = "”«⁄«  ›⁄·Ì…"
.TextMatrix(0, 22) = "ÌÊ„ «·—«Õ…"
.TextMatrix(0, 25) = "„·ÕÊŸ…"

.TextMatrix(0, 26) = "„Ì⁄«œ «‰’—«› —«Õ…"
.TextMatrix(0, 27) = "„Ì⁄«œ —ÃÊ⁄ —«Õ…"
.TextMatrix(0, 28) = "«‰’—«› „«ﬂÌ‰…"
.TextMatrix(0, 29) = "—ÃÊ⁄ „«ﬂÌ‰…"
.TextMatrix(0, 30) = "«‰’—«›"
.TextMatrix(0, 31) = "—ÃÊ⁄"
.TextMatrix(0, 32) = " „«„ «·—«Õ…"
.TextMatrix(0, 33) = "«‰’—«› „»ﬂ—"
.TextMatrix(0, 34) = "—ÃÊ⁄ „ √Œ—"
.TextMatrix(0, 35) = "Ã“¡«"
.TextMatrix(0, 36) = "«‰’—«› „»ﬂ—"
.TextMatrix(0, 37) = "—ÃÊ⁄ „ √Œ—"
.TextMatrix(0, 38) = "Ã“«¡"
.TextMatrix(0, 42) = "«–‰"
.TextMatrix(0, 43) = "«·€«¡ «–‰"
.TextMatrix(0, 44) = "«·‘Ì› "

.TextMatrix(0, 45) = "‰”»… Õ÷Ê—"
.TextMatrix(0, 46) = "‰”»… «‰’—«›"

.TextMatrix(0, 47) = "›—⁄ ⁄„Ê·…"
.TextMatrix(0, 48) = "›—⁄ «·»’„…"
.TextMatrix(0, 48) = "›—⁄ «·»’„…"

.TextMatrix(0, .Cols - 3) = "«·€«¡ «·⁄„Ê·…"

.ColHidden(0) = False
.ColHidden(26) = True
.ColHidden(27) = True
.ColHidden(28) = True
.ColHidden(29) = True
.ColHidden(33) = True
.ColHidden(34) = True
.ColHidden(35) = True
.ColHidden(39) = True

.ColHidden(40) = True
.ColHidden(43) = True

.ColHidden(45) = True
.ColHidden(46) = True
.ColHidden(49) = True
.ColHidden(50) = True
.ColHidden(51) = True

.ColComboList(42) = "..."
.ColComboList(44) = cList3
.ColComboList(47) = cList4
.ColComboList(48) = cList4

.ColHidden(9) = True

.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True

.ColDataType(0) = flexDTDate

.ColHidden(4) = chkMachine.Value = 0
.ColHidden(5) = chkMachine.Value = 0

.RowHeight(0) = 700
.WordWrap = True
Dim i As Long

.ColHidden(14) = True
.ColComboList(8) = cList
.ColComboList(9) = cList
.ColComboList(32) = cList2

.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter

For i = 16 To .Cols - 13
    .ColHidden(i) = True
Next

.ColAlignment(26) = flexAlignRightCenter
.ColAlignment(32) = flexAlignRightCenter

.ColHidden(25) = False
.ColHidden(26) = Not bRamadan
.ColHidden(27) = Not bRamadan
.ColHidden(28) = Not bRamadan
.ColHidden(29) = Not bRamadan
.ColHidden(30) = Not bRamadan
.ColHidden(31) = Not bRamadan
.ColHidden(32) = Not bRamadan
.ColHidden(33) = Not bRamadan
.ColHidden(34) = Not bRamadan
.ColHidden(35) = Not bRamadan
.ColHidden(36) = Not bRamadan
.ColHidden(37) = Not bRamadan
.ColHidden(38) = Not bRamadan
.ColHidden(42) = False

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

If grid1.Rows > 1 Then
    .Cell(flexcpBackColor, 1, 0, .Rows - 1, 5) = &H8000000F
    .Cell(flexcpBackColor, 1, 8 + 2, .Rows - 1, .Cols - 1) = &H8000000F
    .Cell(flexcpBackColor, 1, 26, .Rows - 1, 28) = &H8000000F
End If


For i = 1 To .Rows - 1
    RowDgt (i)
Next

For i = 2 To .Rows - 1
    nDay = Weekday(DateValue(.TextMatrix(i, 0)), vbSaturday)
    If nDay = 1 Then nWeek = nWeek + 1
    If nWeek = 1 Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbYellow
    If nWeek = 3 Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbYellow
    If nWeek = 5 Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbYellow
Next i

If grid1.Rows > 1 Then
    .Cell(flexcpFontSize, 1, 44, .Rows - 1, 44) = 10
End If

FixWidth
'Me.Caption = grid1.Cols
End With
End Sub
Private Function validRow(Row As Long) As Boolean
'If Not (ret_Type(grid1.TextMatrix(Row, 8), "SAME") Or ret_Type(grid1.TextMatrix(Row, 9), "SAME") And grid1.TextMatrix(Row, 8) = grid1.TextMatrix(Row, 9)) Then
'    If (Not (IsTime(grid1.TextMatrix(Row, 4)) Or IsTime(grid1.TextMatrix(Row, 5)) Or IsTime(grid1.TextMatrix(Row, 6)) Or IsTime(grid1.TextMatrix(Row, 7)))) Then Exit Function
'End If
validRow = True
End Function
Private Function FixWidth()
Dim i As Long

With grid1
For i = 0 To .Cols - 1
    .ColWidth(i) = 667
Next

    .ColWidth(0) = 1300
    .ColWidth(1) = 750
    .ColWidth(8) = 1100
    .ColWidth(9) = 1100
    .ColWidth(10) = 890
    .ColWidth(11) = 890
    .ColWidth(19) = 700
    .ColWidth(25) = 900
    .ColWidth(32) = 1100
    .ColWidth(33) = 800
    .ColWidth(43) = 600
    .ColWidth(44) = 1600
    .ColWidth(47) = 1200
    .ColWidth(48) = 1200
End With


Dim nRate As Double
For i = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(i) Then
        nwidth = nwidth + grid1.ColWidth(i)
    End If
Next

nRate = (grid1.Width - 420) / (nwidth)

For i = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(i) Then
        grid1.ColWidth(i) = grid1.ColWidth(i) * nRate
    End If
Next
End Function
Private Function Half_Row(Row As Long) As Boolean
With grid1
If ((IsTime(.TextMatrix(Row, 6)) And IsTime(.TextMatrix(Row, 7))) Or (IsTime(.TextMatrix(Row, 30)) And IsTime(.TextMatrix(Row, 31)))) Then Exit Function
If Trim(.TextMatrix(Row, 6)) = "" And Trim(.TextMatrix(Row, 7)) = "" And Trim(.TextMatrix(Row, 30)) = "" And Trim(.TextMatrix(Row, 31)) = "" Then Exit Function
Half_Row = True
End With
End Function
Private Function MissingRow(Row As Long) As Boolean
With grid1
If .TextMatrix(Row, 26) = "" Or .TextMatrix(Row, 27) = "" Then
    If (Trim(.TextMatrix(Row, 6)) = "" Or Trim(.TextMatrix(Row, 7)) = "") And Not (Trim(.TextMatrix(Row, 6)) = "" And Trim(.TextMatrix(Row, 7)) = "") Then
        MissingRow = True
    End If
Else
    If ((Trim(.TextMatrix(Row, 6)) = "" Or Trim(.TextMatrix(Row, 30)) = "") And Not (Trim(.TextMatrix(Row, 6)) = "" And Trim(.TextMatrix(Row, 30)) = "")) Or _
       ((Trim(.TextMatrix(Row, 31)) = "" Or Trim(.TextMatrix(Row, 7)) = "") And Not (Trim(.TextMatrix(Row, 31)) = "" And Trim(.TextMatrix(Row, 7)) = "")) Then
        If (Trim(.TextMatrix(Row, 6)) <> "" And Trim(.TextMatrix(Row, 7)) <> "") And Trim(.TextMatrix(Row, 30)) = "" And Trim(.TextMatrix(Row, 31)) = "" Then
        Else
           MissingRow = True
        End If
    End If
End If
End With
End Function
Private Function PenalityRow(Row As Long) As Boolean
With grid1
If Val(.TextMatrix(Row, 18)) = 0 And Val(.TextMatrix(Row, 28)) = 0 Then Exit Function
End With
PenalityRow = True
End Function
Private Function RetText(ByVal pString) As String
If Trim(pString) = "" Then Exit Function
pString = Trim(pString)
aRet = Split(pString, " ")
For i = 0 To UBound(aRet)
    If Trim(aRet(i)) <> "" Then RetText = RetText & turnFound(RetText, " ") & UCase(Mid(aRet(i), 1, 1))
Next
End Function
Private Sub ColorComp(Row As Long)
With grid1
If Trim(.TextMatrix(Row, 6)) = "" And Trim(.TextMatrix(Row, 7)) = "" Then
'    .Cell(flexcpBackColor, Row, 0, Row, 3) = &H80000005
'    .Cell(flexcpBackColor, row, 6, row, 9) = &H80000005
ElseIf IsTime(.TextMatrix(Row, 6)) And IsTime(.TextMatrix(Row, 7)) Then
'    .Cell(flexcpBackColor, Row, 0, Row, 3) = &HC0FFFF
 '   .Cell(flexcpBackColor, row, 6, row, 9) = &HC0FFFF
Else
'    .Cell(flexcpBackColor, Row, 0, Row, 3) = &HC0C0FF
  '  .Cell(flexcpBackColor, row, 6, row, 9) = &HC0C0FF
End If
End With
End Sub
Private Sub GRID2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If NewRow > 0 Then
    Xcode.Caption = GRID2.TextMatrix(NewRow, 0)
End If
End Sub
Private Sub grid2_EnterCell()
If GRID2.Row > 0 Then Xcode.Caption = GRID2.TextMatrix(GRID2.Row, 0)
End Sub
Private Sub GRID2_GotFocus()
If GRID2.Row > 0 Then Xcode.Caption = GRID2.TextMatrix(GRID2.Row, 0)
End Sub

Private Sub xAbsent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
xNo_data.Value = 0
myLoadGrd2 True
End Sub
Private Sub xBranch_Change()
myLoadGrd2
End Sub
Private Sub xcode_Change()
'On Error GoTo myerror
Me.MousePointer = vbHourglass
myLoadGrd
myLoadHolidays
Me.MousePointer = vbNormal
Exit Sub
myerror:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xCurrent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo myerror
Dim sDate As String, nRow As Long
xDate1.text = sDate_Period
xDate2.text = IIf(xCurrent.Value = 1, sDate_Period_End, myFormat(Date))
myLoadGrd2 True
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xDate1_DblClick()
Set datefrm.oDate = xDate1
datefrm.Show 1
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xdate2_DblClick()
Set datefrm.oDate = xDate2
datefrm.Show 1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDays_absent_Change()
cmdAbsShow(0).Enabled = Val(xDays_absent.Caption) <> 0
End Sub

Private Sub xDays_off_Change()
cmdAbsShow(1).Enabled = Val(xDays_off.Caption) <> 0
End Sub

Private Sub xdays_Remedy_Change()
cmdAbsShow(3).Enabled = Val(xdays_Remedy.Caption) <> 0
End Sub

Private Sub xDays_Temp_Change()
cmdAbsShow(2).Enabled = Val(xDays_temp.Caption) <> 0
End Sub

Private Sub xDesca_Change()
'cmdGo.Enabled = True
End Sub

Private Sub xJob_Click(Area As Integer)
'cmdGo.Enabled = True
End Sub

Private Sub xCode_GotFocus()
myGotFocus Xcode
End Sub

Private Sub xCode_LostFocus()
myLostFocus Xcode
End Sub
Private Sub xMachine_Click()
cmdGo.Enabled = True
End Sub

Private Sub xManage_Change()
Dim nRow As Long
If xManage.MatchedWithList Or Trim(xManage.BoundText) = "" Then
    myLoadGrd2
End If
End Sub

Private Sub xManage_Click(Area As Integer)
cmdGo.Enabled = True
End Sub

Private Sub xSection_Click(Area As Integer)
cmdGo.Enabled = True
End Sub
Private Function FillDataArray(cFile) As Boolean
Dim TextLine As String, aLocal, i As Long
Dim aInsert(4, 1)
On Error GoTo myerror
Open cFile For Input As #1    ' Open file.
'contemp.Execute "Delete * from machine"
con.BeginTrans
Do While Not EOF(1)   ' Loop until end of file.
     i = i + 1
    Line Input #1, TextLine   ' Read line into variable.
    If Trim(TextLine) = "" Then Exit Function
    aLocal = Split(TextLine, vbTab)
    If UBound(aLocal) <> 5 Then
        Close #1   ' Close file.
         Exit Function
     End If
           
    aInsert(0, 0) = "Emp"
    aInsert(0, 1) = addvalue(aLocal(0))
    
    sDate = Trim(aLocal(1))
    sDay = Trim(Mid(sDate, 9, 2))
    smonth = Trim(Mid(sDate, 6, 2))
    sYear = Trim(Mid(sDate, 1, 4))
    sDate = sDay & "-" & smonth & "-" & sYear
     
    aInsert(1, 0) = "[Date]"
    aInsert(1, 1) = addDate(sDate)
    
    aInsert(2, 0) = "[TIME]"
    aInsert(2, 1) = addTime(Trim(Mid(TextLine, 18, 5)))
    
    aInsert(3, 0) = "[TYPE]"
    aInsert(3, 1) = addstring(Trim(aLocal(3)))
    
    aInsert(4, 0) = "ROW"
    aInsert(4, 1) = i
    On Error Resume Next
    con.Execute CreateInsert(aInsert, "hr.MACHINE")
    If Err.Number = -2147467259 Then
        Err.Clear
    ElseIf Err.Number <> 0 Then
        GoTo myerror
    End If
    On Error GoTo myerror
Loop
con.CommitTrans
Close #1   ' Close file.
FillDataArray = True
Exit Function
myerror:
MsgBox Err.Number & vbCrLf & Err.Description
con.RollbackTrans
Err.Clear
Close #1   ' Close file.
End Function
Private Sub GetLogFile()
On Error GoTo myerror
Dim fs As New FileSystemObject, cDrive As String, cFileName As String
cDrive = Trim(LastDrive(True))
cDrive = cDrive & turnFound(cDrive, ":")
If cDrive <> "" Then
    cFileName = cDrive & "\1_attlog.dat"
    If fs.FileExists(cFileName) Then
        If MsgBox("Found New MACHINE Output File! update Data ! are you sure ?", vbOKCancel + vbDefaultButton2) = vbOK Then
            If fs.FileExists(App.Path & "\machine\1_attlog.dat") Then
                If DateValue(fs.GetFile(App.Path & "\machine\1_attlog.dat").DateLastModified) >= DateValue(fs.GetFile(cFileName).DateLastModified) Then
                    If MsgBox("File with modification Date before or with same with the file tried to be replaced !" & vbCrLf & _
                               "Continue replace any How", vbOKCancel + vbDefaultButton2) <> vbOK Then
                        Exit Sub
                    End If
                End If
            End If
            MyCreateFolder App.Path & "\machine"
            fs.CopyFile cFileName, App.Path & "\machine\1_attlog.dat"
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function myReplacegrd(Optional Row As Long = -1) As Boolean
Dim sbJson As New ChilkatStringBuilder
Me.MousePointer = vbHourglass
Dim aInsert As Variant, i As Long, i2 As Long, nId As Long
With grid1
    prog1.Visible = True
    prog1.Value = 0
        
    For i = IIf(Row = -1, 2, Row) To IIf(Row = -1, grid1.Rows - 1, Row)
        prog1.Value = Round(i / (grid1.Rows - 1), 2) * 100
        aInsert = AddFlag(Empty, "DATE", addDate(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "CODE", addvalue(grid1.TextMatrix(i, grid1.Cols - 2)))
        aInsert = AddFlag(aInsert, "EMP_TIME1", addstring(.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "EMP_TIME2", addstring(.TextMatrix(i, 3)))
        
        aInsert = AddFlag(aInsert, "TIME1", addstring(.TextMatrix(i, 6)))
        aInsert = AddFlag(aInsert, "TIME2", addstring(.TextMatrix(i, 7)))
        aInsert = AddFlag(aInsert, "TYPE1", addstring(.TextMatrix(i, 8)))
        aInsert = AddFlag(aInsert, "TYPE2", addstring(.TextMatrix(i, 9)))
        
        aInsert = AddFlag(aInsert, "AMOUNT1", Val(.TextMatrix(i, 16)))
        aInsert = AddFlag(aInsert, "AMOUNT2", Val(.TextMatrix(i, 17)))
                    
        aInsert = AddFlag(aInsert, "PENALITY1", Val(.TextMatrix(i, 18)))
        aInsert = AddFlag(aInsert, "OVERTIME1", Val(.TextMatrix(i, 19)))
        aInsert = AddFlag(aInsert, "WORK_TIME", Val(.TextMatrix(i, 21)))
        aInsert = AddFlag(aInsert, "HOURS", Val(.TextMatrix(i, 20)))
        aInsert = AddFlag(aInsert, "[BREAK]", Val(.TextMatrix(i, 22)))
        aInsert = AddFlag(aInsert, "IS_BREAK", Val(.TextMatrix(i, 23)))
        aInsert = AddFlag(aInsert, "IS_HOLIDAY", Val(.TextMatrix(i, 24)))
        aInsert = AddFlag(aInsert, "[MONTH]", addvalue(sMonth_Period))
        
        aInsert = AddFlag(aInsert, "[YEAR]", addvalue(sYear_period))
        aInsert = AddFlag(aInsert, "NOTE", addstring(.TextMatrix(i, 25)))
        aInsert = AddFlag(aInsert, "TIME1_EXC", addstring(.TextMatrix(i, 39)))
        aInsert = AddFlag(aInsert, "TIME2_EXC", addstring(.TextMatrix(i, 40)))
        
        aInsert = AddFlag(aInsert, "MINUTES_EXC", Val(.TextMatrix(i, 41)))
        aInsert = AddFlag(aInsert, "MINUTES_EXC_CANCEL", Val(.TextMatrix(i, 43)))
        aInsert = AddFlag(aInsert, "SHIFT", addvalue(.TextMatrix(i, 44)))
        
        aInsert = AddFlag(aInsert, "RATE_LATE_ARRIVE", Val(.TextMatrix(i, 45)))
        aInsert = AddFlag(aInsert, "RATE_LATE_LEAVE", Val(.TextMatrix(i, 46)))
        
        aInsert = AddFlag(aInsert, "BRANCH_COM", addvalue(.TextMatrix(i, 47)))
        aInsert = AddFlag(aInsert, "CANCEL_COM", IIf(.ValueMatrix(i, .Cols - 3) = 0, "0", "1"))
        
        aInsert = AddFlag(aInsert, "MACHINE", "CASE WHEN FINGER_PRINT = 0 THEN  " & addvalue(.TextMatrix(i, 47)) & " ELSE MACHINE END")
        aInsert = AddFlag(aInsert, "PENALITY_R", .ValueMatrix(i, 51))
        
        If Not isRmd(.TextMatrix(i, 0)) Then
            aInsert = AddFlag(aInsert, "AMOUNT1_BREAK", 0)
            aInsert = AddFlag(aInsert, "AMOUNT2_BREAK", 0)
            aInsert = AddFlag(aInsert, "PENALITY_BREAK", 0)
            aInsert = AddFlag(aInsert, "TYPE1_BREAK", "NULL")
            aInsert = AddFlag(aInsert, "TYPE2_BREAK", "NULL")
        Else
            
            aInsert = AddFlag(aInsert, "TIME1_BREAK", addstring(.TextMatrix(i, 30)))
            aInsert = AddFlag(aInsert, "TIME2_BREAK", addstring(.TextMatrix(i, 31)))
            
            aInsert = AddFlag(aInsert, "AMOUNT1_BREAK", Val(.TextMatrix(i, 33)))
            aInsert = AddFlag(aInsert, "AMOUNT2_BREAK", Val(.TextMatrix(i, 34)))
            aInsert = AddFlag(aInsert, "PENALITY_BREAK", Val(.TextMatrix(i, 35)))
            
            aInsert = AddFlag(aInsert, "TYPE1_BREAK", addvalue(.TextMatrix(i, 32)))
            aInsert = AddFlag(aInsert, "TYPE2_BREAK", addvalue(.TextMatrix(i, 32)))
        End If
        sbJson.Append addUpdate(aInsert, "hr.FILE2_50", "ID = " & .TextMatrix(i, .Cols - 1)) & ";"
    Next
    prog1.Visible = False
End With

If sbJson.Length > 0 Then
    con.BeginTrans
    On Error GoTo myerror
    con.Execute sbJson.GetAsString()
    con.CommitTrans
End If

Me.MousePointer = vbNormal
myReplacegrd = True
Exit Function
myerror:
Me.MousePointer = vbNormal
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplace()
If Not MsgBox("Õ›Ÿ «·»Ì«‰«  Â· «‰  „Ê«›ﬁ", vbOKCancel) = vbOK Then Exit Sub
If Not myvalid Then Exit Sub
If myReplacegrd Then Inform " „ «·Õ›Ÿ »‰Ã«Õ"
myLoadGrd
End Sub
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = cmdBranch.Name Then
    cmdBranch.Tag = osearchBranch.grid1.TextMatrix(osearchBranch.grid1.Row, 0)
    cmdBranch.Caption = osearchBranch.grid1.TextMatrix(osearchBranch.grid1.Row, 1)
    osearchBranch.Hide
    myLoadGrd2
ElseIf ActiveControl.Name = cmdEmp.Name Then
    cmdEmp.Tag = oSearchEmp.grid1.TextMatrix(oSearchEmp.grid1.Row, 0)
    cmdEmp.Caption = oSearchEmp.grid1.TextMatrix(oSearchEmp.grid1.Row, 1)
    oSearchEmp.Hide
    myLoadGrd2
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub xCode_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then EmpLookupAll Me, oSearchEmp
End Sub
Private Function myvalid() As Boolean
If grid1.Rows < 2 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ··Õ›Ÿ"
    Exit Function
End If
myvalid = True
End Function
Public Sub myLoadGrd()
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "code", TurnValue(Xcode.Caption))
aPrm = AddFlag(aPrm, "Date1", myFormat_sp(xDate1.text))
aPrm = AddFlag(aPrm, "Date2", myFormat_sp(xDate2.text))
aPrm = AddFlag(aPrm, "late", IIf(Val(xLate.text) > 0 And chkLate.Value = 1, Val(xLate.text), Null))
aPrm = AddFlag(aPrm, "Exc", chkExc.Value)
aPrm = AddFlag(aPrm, "Absent", xAbsent.Value)
aPrm = AddFlag(aPrm, "No_Data", xNo_data.Value)
Set data10.Recordset = myCmd("hr.sp_Arrive", con, adStoredProc, aPrm)
With grid1
If grid1.Rows > 1 And grid1.Row < 1 Then grid1.Row = 1

Fixgrd

CalcTotals
'Me.Caption = grid1.Cols
End With
End Sub
Private Sub RowDgt(Row As Long)
With grid1
Dim aRet As Variant, cNote As String
.TextMatrix(Row, 1) = arbDay(.TextMatrix(Row, 0))
aRet = retAr(.TextMatrix(Row, 2), .TextMatrix(Row, 3), .TextMatrix(Row, 6), .TextMatrix(Row, 7), .TextMatrix(Row, 8), Val(.TextMatrix(Row, 45)), Val(.TextMatrix(Row, 46)), Val(.TextMatrix(Row, 44)))
.TextMatrix(Row, 16) = retFlag(aRet, "late_arrive")
.TextMatrix(Row, 17) = retFlag(aRet, "late_leave")
.TextMatrix(Row, 18) = Val(retFlag(aRet, "Penality")) + Val(grid1.TextMatrix(i, 35))
.TextMatrix(Row, 19) = retFlag(aRet, "Over")
.TextMatrix(Row, 20) = retFlag(aRet, "hours")
.TextMatrix(Row, 21) = retFlag(aRet, "work_time")
.TextMatrix(Row, 51) = retFlag(aRet, "penality_r")

With grid1
For i = 16 To 21
    If i = 20 Then
        .TextMatrix(Row, i - 6) = MinToTimeDgt(Val(.TextMatrix(Row, i)) * 60)
    ElseIf i = 17 Then
        .TextMatrix(Row, 11) = MinToTimeDgt(Val(.TextMatrix(Row, i)))
    ElseIf i = 16 Then
        .TextMatrix(Row, 10) = MinToTimeDgt(Val(.TextMatrix(Row, i)))
    ElseIf i = 19 Then
        .TextMatrix(Row, 13) = MinToTimeDgt(Val(.TextMatrix(Row, 19)))
    Else
        .TextMatrix(Row, i - 6) = MinToTimeDgt(.TextMatrix(Row, i))
    End If
Next
End With


.TextMatrix(Row, 45) = Myvalue(.TextMatrix(Row, 45))
.TextMatrix(Row, 46) = Myvalue(.TextMatrix(Row, 46))

cNote = .TextMatrix(Row, 25)
If Weekday(myFormat(grid1.TextMatrix(Row, 0)), vbSaturday) = Val(.TextMatrix(Row, 22)) Then
    If .TextMatrix(Row, 8) = "" Then .TextMatrix(Row, 8) = "3"
    
    .TextMatrix(Row, 23) = "1"
Else
    .TextMatrix(Row, 23) = "0"
End If

If .ValueMatrix(Row, 24) <> 0 Then
    If Val(.TextMatrix(Row, 8)) <> 10 And cNote = "" Then
        .TextMatrix(Row, 25) = .TextMatrix(Row, 25) & turn(.TextMatrix(Row, 25), " - ") & "«Ã«“…"
    End If
    .TextMatrix(Row, 24) = "1"
Else
    .TextMatrix(Row, 24) = "0"
End If

.TextMatrix(Row, 42) = MinToTimeDgt(.TextMatrix(Row, 41))
fixRmd Row
If (Not validRow(Row)) And chkNoData.Value = 0 Then
     .Cell(flexcpBackColor, Row, 6, Row, 25) = &H80000005
ElseIf MissingRow(Row) And chkMissing.Value = 0 Then
    .Cell(flexcpBackColor, Row, 6, Row, 25) = &HC0C0FF
ElseIf grid1.TextMatrix(Row, 8) = "2" Then
     .Cell(flexcpBackColor, Row, 6, Row, 25) = vbCyan
     .Cell(flexcpForeColor, Row, 6, Row, .Cols - 1) = vbBlack
Else
    .Cell(flexcpBackColor, Row, 6, Row, 25) = &H80000005
    .Cell(flexcpForeColor, Row, 6, Row, .Cols - 1) = &H0&
End If
If PenalityRow(Row) And chkPenality.Value = 0 Then
    If Val(.TextMatrix(Row, 16)) > 0 Then
        .Cell(flexcpForeColor, Row, 6, Row, 6) = vbRed
        .Cell(flexcpForeColor, Row, 10, Row, 10) = vbRed
        .Cell(flexcpForeColor, Row, 12, Row, 12) = vbRed
    End If
    If Val(.TextMatrix(Row, 17)) < 0 Then
        .Cell(flexcpForeColor, Row, 7, Row, 7) = vbRed
        .Cell(flexcpForeColor, Row, 11, Row, 11) = vbRed
        .Cell(flexcpForeColor, Row, 12, Row, 12) = vbRed
    End If
    If .TextMatrix(Row, 31) <> "" Then
        .Cell(flexcpForeColor, Row, 26, Row, 26) = vbRed
        .Cell(flexcpForeColor, Row, 28, Row, 28) = vbRed
    End If
    If .TextMatrix(Row, 32) <> "" Then
        .Cell(flexcpForeColor, Row, 27, Row, 32) = vbRed
        .Cell(flexcpForeColor, Row, 28, Row, 32) = vbRed
    End If
End If
End With
End Sub
Private Sub FixDgt(Row As Long)
End Sub
Private Function CheckRow(Row As Long, Col As Long)
With grid1
If (IsTime(.TextMatrix(Row, 6)) Or IsTime(.TextMatrix(Row, 7)) Or IsTime(.TextMatrix(Row, 30)) Or IsTime(.TextMatrix(Row, 31))) And (ret_Type(.TextMatrix(Row, 8), "No_Data") Or Trim(.TextMatrix(Row, 8)) = "") Then
    .TextMatrix(Row, 8) = "1"
ElseIf (Not (IsTime(.TextMatrix(Row, 6)) Or IsTime(.TextMatrix(Row, 7)) Or IsTime(.TextMatrix(Row, 30)) Or IsTime(.TextMatrix(Row, 31)))) Then
    If Not ret_Type(.TextMatrix(Row, 8), "No_Data") Then
        .TextMatrix(Row, 8) = ""
    End If
End If

If Col = 6 Or Col = 7 Or Col = 8 Or Col = .Cols - 3 Then
    fixCom .TextMatrix(Row, 8), Row, Col
End If

If isRmd(grid1.TextMatrix(Row, 0)) Then
    If (IsTime(.TextMatrix(Row, 6)) And IsTime(.TextMatrix(Row, 7))) Or (IsTime(.TextMatrix(Row, 30)) And IsTime(.TextMatrix(Row, 31))) Then
        If Val(.TextMatrix(Row, 32)) = 13 Then .TextMatrix(Row, 32) = ""
    ElseIf Half_Row(Row) And (Not ret_Type(.TextMatrix(Row, 32), "no_Data")) Then
        .TextMatrix(Row, 32) = 13
    ElseIf ret_Type(.TextMatrix(Row, 8), "no_data") Then
        .TextMatrix(Row, 32) = ""
    End If
 Else
    .TextMatrix(Row, 30) = ""
    .TextMatrix(Row, 31) = ""
    .TextMatrix(Row, 32) = ""
    .TextMatrix(Row, 33) = ""
    .TextMatrix(Row, 34) = ""
    .TextMatrix(Row, 35) = ""
    .TextMatrix(Row, 36) = ""
    .TextMatrix(Row, 37) = ""
    .TextMatrix(Row, 38) = ""
End If
End With
End Function
Private Function fixCom(pType As String, Row As Long, Col As Long)
If grid1.ValueMatrix(Row, grid1.Cols - 3) <> 0 Then
    grid1.TextMatrix(Row, 47) = ""
ElseIf Col <> 47 Then
    If pType = "" Then
        grid1.TextMatrix(Row, 47) = ""
    ElseIf ret_Type(pType, "commission", False) And grid1.TextMatrix(Row, 47) = "" Then
        grid1.TextMatrix(Row, 47) = grid1.TextMatrix(Row, 49)
    ElseIf (Not ret_Type(pType, "commission", False)) And grid1.TextMatrix(Row, 47) <> "" Then
        grid1.TextMatrix(Row, 47) = ""
    End If
End If
If Col <> 48 Then
    If pType = "" Then
        grid1.TextMatrix(Row, 48) = ""
    ElseIf (Not ret_Type(pType, "No_Data", False)) And grid1.TextMatrix(Row, 48) = "" Then
        grid1.TextMatrix(Row, 48) = grid1.TextMatrix(Row, 49)
    ElseIf (ret_Type(pType, "No_Data")) And grid1.TextMatrix(Row, 48) <> "" Then
        grid1.TextMatrix(Row, 48) = ""
    End If
End If
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < 7 Then
    grid1.Select Row, IIf(Col < 6, 6, Col + 1)
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 6, 6)
    grid1.ShowCell Row + 1, 1
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Col <> 8 And Col <> 9 And Col <> 32 And Col <> 44 Then CellPos KeyCode, Row, Col
End Sub
Private Sub ValidCheck()
Dim i As Long
If chkNoData.Value = 1 Then
    For i = grid1.Rows - 1 To 1 Step -1
        If validRow(i) Then
            grid1.RemoveItem i
        End If
    Next
End If
If chkMissing.Value = 1 Then
    For i = grid1.Rows - 1 To 1 Step -1
        If Not MissingRow(i) Then
            grid1.RemoveItem i
        End If
    Next
End If
If chkPenality.Value = 1 Then
    For i = grid1.Rows - 1 To 1 Step -1
        If Not PenalityRow(i) Then
            grid1.RemoveItem i
        End If
    Next
End If
End Sub
Private Sub myLoadGrd2(Optional bRefresh As Boolean = False)
Dim sCode As String
Dim sProc As String
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "code", TurnValue(cmdEmp.Tag))
aPrm = AddFlag(aPrm, "manage", TurnValue(cmdBranch.Tag))
aPrm = AddFlag(aPrm, "Date1", myFormat_sp(xDate1.text))
aPrm = AddFlag(aPrm, "Date2", myFormat_sp(xDate2.text))
aPrm = AddFlag(aPrm, "late", IIf(Val(xLate.text) > 0 And chkLate.Value = 1, Val(xLate.text), Null))
aPrm = AddFlag(aPrm, "Exc", chkExc.Value)
aPrm = AddFlag(aPrm, "Absent", xAbsent.Value)
aPrm = AddFlag(aPrm, "No_Data", xNo_data.Value)
'aPrm = AddFlag(aPrm, "Name_Order", chkName.Value)

If GRID2.Row > 0 Then
    sCode = GRID2.TextMatrix(GRID2.Row, 0)
End If

On Error GoTo myerror
If FilterEmp Then
    Set DATA11.Recordset = myCmd("hr.sp_Arrive_emp2", con, adStoredProc, aPrm)
Else
    Set DATA11.Recordset = myCmd("hr.sp_Arrive_emp", con, adStoredProc, aPrm)
End If

With GRID2
    .ColHidden(0) = .Cols > 2
    .ColWidth(0) = 800
    .ColWidth(1) = IIf(.Cols > 2, .Width - 1050, .Width - 1200)
    .ColAlignment(1) = flexAlignRightCenter
    .TextMatrix(0, 0) = "«·ﬂÊœ"
    .TextMatrix(0, 1) = "«·«”„"
    If .Cols > 2 Then
        .ColWidth(2) = 700
        .ColAlignment(2) = flexAlignRightCenter
        .TextMatrix(0, 2) = "«·«Ì«„"
    End If
    
    If .Rows = 1 Then
        Xcode.Caption = ""
    Else
        Dim nRow As Long
        nRow = .FindRow(sCode, , 0)
        
        If nRow > 0 Then
            .Select nRow, 1
            .ShowCell nRow, 1
        ElseIf .Rows > 1 Then
            .Select 1, 1
            .ShowCell .Row, 1
        End If
        If Xcode.Caption <> .TextMatrix(.Row, 0) Then
            Xcode.Caption = .TextMatrix(.Row, 0)
        Else
            If bRefresh Then myLoadGrd
        End If
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function FilterEmp() As Boolean
FilterEmp = xAbsent.Value = 1 Or xNo_data.Value = 1 Or chkExc.Value = 1 Or (Val(xLate.text) > 0 And chkLate.Value = 1)
End Function
Private Sub myLoadHolidays()
myDefineHolidays
If Xcode.Caption = "" Then Exit Sub

Dim aPrm As Variant
aPrm = AddFlag(aPrm, "code", TurnValue(Xcode.Caption))
aPrm = AddFlag(aPrm, "year", sYear_period)
aPrm = AddFlag(aPrm, "Date1", myFormat_sp(xDate1.text))
aPrm = AddFlag(aPrm, "Date2", myFormat_sp(xDate2.text))

Dim loctable As ADODB.Recordset
Set loctable = myCmd("hr.sp_holidays_emp", con, adStoredProc, aPrm)
If Not loctable.EOF Then
    XDESCA.Caption = loctable!DESCA & ""
    
    xDays_absent.Caption = Myvalue(loctable!DAYS_ABSENT)
    xDays_absent2.Caption = Myvalue(loctable!days_absent2)
    
    xDays_Discount.Caption = Myvalue(loctable!days_Discount)
    xDays_Discount2.Caption = Myvalue(loctable!DAYS_DISCOUNT2)
    
    xDays_off.Caption = Myvalue(loctable!DAYS_OFF)
    xDays_off2.Caption = Myvalue(loctable!DAYS_OFF2)
    
    xDays_temp.Caption = Myvalue(loctable!DAYS_TEMP)
    xDays_temp2.Caption = Myvalue(loctable!DAYS_TEMP2)
    
    xdays_Remedy.Caption = Myvalue(loctable!DAYS_REMEDY)
    xdays_Remedy2.Caption = Myvalue(loctable!days_Remedy2)
    xdays_abs_late.Caption = Myvalue(loctable!days_Abs_late + loctable!days_Abs_late2)
    xBreak.Caption = Myvalue(loctable!Break)
    xBreak_all.Caption = Myvalue(loctable!Break_all)
    xBreak2.Caption = Myvalue(loctable!Break_all - loctable!Break)
    
    xDay_Break.Tag = Myvalue(loctable!day_break)
    If Not IsNull(loctable!day_break) Then
        xDay_Break.Caption = NameOfDay(loctable!day_break)
    End If
End If
End Sub
Private Sub xLate_Change()
chkLate.Value = 0
End Sub
Private Sub xNo_data_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
xAbsent.Value = 0
myLoadGrd2 True
End Sub
Private Function ValidBreak() As Long
nBreak = CalcBreak(Val(xDay_Break.Tag))
For i = 1 To grid1.Rows - 1
    If grid1.TextMatrix(i, 8) = "3" Then
        nTaken = nTaken + 1
    End If
Next
End Function
Private Sub myDefineHolidays()
XDESCA.Caption = ""

xDays_absent.Caption = ""
xDays_absent2.Caption = ""

xDays_Discount.Caption = ""
xDays_Discount2.Caption = ""

xDays_off.Caption = ""
xDays_off2.Caption = ""

xDays_temp.Caption = ""
xDays_temp2.Caption = ""

xdays_Remedy.Caption = ""
xdays_Remedy2.Caption = ""
xdays_abs_late.Caption = ""

xBreak.Caption = ""
xBreak_all.Caption = ""
xBreak2.Caption = ""

xDay_Break.Tag = ""
xDay_Break.Caption = ""
End Sub
Private Sub Fix_late()
Dim i As Long
With grid1
For i = 1 To grid1.Rows - 1
    If .TextMatrix(i, 2) <> "" And .TextMatrix(i, 6) <> "" Then .TextMatrix(i, 6) = .TextMatrix(i, 2)
    If .TextMatrix(i, 3) <> "" And .TextMatrix(i, 7) <> "" Then .TextMatrix(i, 7) = .TextMatrix(i, 3)
    If .TextMatrix(i, 33) <> "" And .TextMatrix(i, 26) <> "" Then .TextMatrix(i, 26) = .TextMatrix(i, 33)
    If .TextMatrix(i, 34) <> "" And .TextMatrix(i, 27) <> "" Then .TextMatrix(i, 27) = .TextMatrix(i, 34)
    Grid1_AfterEdit i, 7
Next
End With
End Sub
Private Sub fixRmd(Row As Long)
Dim aRet As Variant
With grid1

aRet = retAr_m(.TextMatrix(Row, 26), .TextMatrix(Row, 27), .TextMatrix(Row, 30), .TextMatrix(Row, 31), .TextMatrix(Row, 32))
.TextMatrix(Row, 33) = retFlag(aRet, "early_leave")
.TextMatrix(Row, 34) = retFlag(aRet, "late_arrive")
.TextMatrix(Row, 35) = retFlag(aRet, "penality")

.TextMatrix(Row, 36) = MinToTimeDgt(Val(retFlag(aRet, "early_leave")))
.TextMatrix(Row, 37) = MinToTimeDgt(Val(retFlag(aRet, "late_arrive")))
.TextMatrix(Row, 38) = MinToTimeDgt(retFlag(aRet, "penality"))
End With
End Sub
Private Sub Fix_Incomplete()
Dim i As Long
With grid1
For i = 1 To grid1.Rows - 1
    If Trim(.TextMatrix(i, 26)) = "" Or Trim(.TextMatrix(i, 27)) = "" Then
        If Trim(.TextMatrix(i, 6)) = "" And Trim(.TextMatrix(i, 7)) <> "" Then
            .TextMatrix(i, 6) = .TextMatrix(i, 2)
            Grid1_AfterEdit i, 6
        End If
        If Trim(.TextMatrix(i, 7)) = "" And Trim(.TextMatrix(i, 6)) <> "" Then
            .TextMatrix(i, 7) = .TextMatrix(i, 3)
            Grid1_AfterEdit i, 7
        End If
    Else
        If Trim(.TextMatrix(i, 6)) = "" And Trim(.TextMatrix(i, 30)) <> "" Then
            .TextMatrix(i, 6) = .TextMatrix(i, 2)
            Grid1_AfterEdit i, 6
        End If
        
        If Trim(.TextMatrix(i, 30)) = "" And Trim(.TextMatrix(i, 6)) <> "" Then
            .TextMatrix(i, 30) = .TextMatrix(i, 26)
            Grid1_AfterEdit i, 30
        End If
        
        If Trim(.TextMatrix(i, 7)) = "" And Trim(.TextMatrix(i, 31)) <> "" Then
            .TextMatrix(i, 7) = .TextMatrix(i, 3)
            Grid1_AfterEdit i, 7
        End If
        
        If Trim(.TextMatrix(i, 31)) = "" And Trim(.TextMatrix(i, 7)) <> "" Then
            .TextMatrix(i, 31) = .TextMatrix(i, 27)
            Grid1_AfterEdit i, 31
        End If
    End If
Next
End With
End Sub
Private Sub fillgrd2()
With GRID2
.TextMatrix(0, 0) = "«·ÌÊ„"
.TextMatrix(0, 1) = "«· «—ÌŒ"
nDays = DateDiff("d", xDate1.text, xDate2.text)
.Rows = 1
For i = nDays To 0 Step -1
    sDate = myFormat(DateAdd("d", i, xDate1.text))
    .AddItem ""
    .TextMatrix(.Rows - 1, 1) = sDate
    .TextMatrix(.Rows - 1, 0) = arbDay(.TextMatrix(.Rows - 1, 1))
Next
.ColAlignment(0) = flexAlignRightCenter
.ColAlignment(1) = flexAlignRightCenter
If GRID2.Rows > 1 And GRID2.Row < 1 Then
    GRID2.Select 1, 0
End If
End With
End Sub
Private Sub CalcTotals()
With grid1
grid1.SubtotalPosition = flexSTAbove
grid1.Subtotal flexSTSum, -1, 16, "00", vbYellow, vbRed, , "≈Ã„«·Ì"
grid1.Subtotal flexSTSum, -1, 17, "00", vbYellow, vbRed, , "≈Ã„«·Ì"
grid1.Subtotal flexSTSum, -1, 18, "00", vbYellow, vbRed, , "≈Ã„«·Ì"
grid1.Subtotal flexSTSum, -1, 19, "00", vbYellow, vbRed, , "≈Ã„«·Ì"
grid1.Subtotal flexSTSum, -1, 21, "00", vbYellow, vbRed, , "≈Ã„«·Ì"
If grid1.Rows > 1 Then
    .TextMatrix(1, 10) = MinToTimeDgt(Val(.TextMatrix(1, 16)))
    .TextMatrix(1, 11) = MinToTimeDgt(Val(.TextMatrix(1, 17)))
    .TextMatrix(1, 12) = MinToTimeDgt(Val(.TextMatrix(1, 18)))
    .TextMatrix(1, 13) = MinToTimeDgt(Val(.TextMatrix(1, 19)))
    .TextMatrix(1, 15) = MinToTimeDgt(Val(.TextMatrix(1, 21)))
    .TextMatrix(1, 16) = MinToTimeDgt(Val(.TextMatrix(1, 21)))
End If
End With
End Sub
Private Function AddAbsLate(Optional pCode As String = "", Optional pManage As String = "") As Boolean
Dim cmd As New ADODB.Command
aInsert = AddFlag(Empty, "DATE1", myFormat_sp(sDate_Period))
aInsert = AddFlag(aInsert, "DATE2", myFormat_sp(sDate_Period_End))
aInsert = AddFlag(aInsert, "CODE", TurnValue(pCode))
aInsert = AddFlag(aInsert, "MANAGE", TurnValue(pManage))
Set cmd = mycmdEx("hr.sp_add_abs_late", con, aInsert)
If Not IsNull(cmd.Parameters("@Error_Msg").Value) Then
    MsgBox cmd.Parameters("@Error_Msg").Value
    Exit Function
End If
AddAbsLate = True
End Function
Private Function RemoveAbsLate(Optional pCode As String = "", Optional pManage As String = "") As Boolean
Dim cmd As New ADODB.Command
aInsert = AddFlag(Empty, "DATE1", myFormat_sp(sDate_Period))
aInsert = AddFlag(aInsert, "DATE2", myFormat_sp(sDate_Period_End))
aInsert = AddFlag(aInsert, "CODE", TurnValue(pCode))
aInsert = AddFlag(aInsert, "MANAGE", TurnValue(pManage))
Set cmd = mycmdEx("hr.sp_remove_abs_late", con, aInsert)
If Not IsNull(cmd.Parameters("@Error_Msg").Value) Then
    MsgBox cmd.Parameters("@Error_Msg").Value
    Exit Function
End If
RemoveAbsLate = True
End Function

