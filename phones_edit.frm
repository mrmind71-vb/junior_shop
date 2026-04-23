VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form phoneEditfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«” ⁄·«„ ›Ê« Ì— "
   ClientHeight    =   9420
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15150
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   15150
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ckhPhone 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " ·Ì›Ê‰«  €Ì— ’«·Õ…"
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
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   225
      Width           =   1995
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   630
      Width           =   5820
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   13
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
         Picture         =   "phones_edit.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdDateFix 
         Height          =   555
         Left            =   1305
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   979
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
         Caption         =   "÷»ÿ «· Ê—«ÌŒ"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "phones_edit.frx":2323
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   4500
         TabIndex        =   15
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
         Picture         =   "phones_edit.frx":44A6
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   3195
         TabIndex        =   11
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
         Picture         =   "phones_edit.frx":7359
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   9015
      Begin VB.TextBox XPHONE 
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   3435
      End
      Begin VB.TextBox XDATE2 
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
         Left            =   90
         MaxLength       =   15
         TabIndex        =   4
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox XDATE1 
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
         Left            =   90
         MaxLength       =   15
         TabIndex        =   3
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox xDesca 
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   3435
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   330
         Left            =   3960
         TabIndex        =   0
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·› —… «·Ì"
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
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   585
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰   ·Ì›Ê‰"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·› —… „‰"
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
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰  ⁄„Ì·"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   1125
      End
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   225
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   90
      Top             =   -225
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
      Height          =   7260
      Left            =   180
      TabIndex        =   16
      Top             =   1395
      Width           =   14820
      _cx             =   26141
      _cy             =   12806
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
      Cols            =   4
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   18
      Top             =   8955
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   24
      Top             =   8760
      Visible         =   0   'False
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "phoneEditfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CMD_EXEL_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub cmdDateFix_Click()
Me.MousePointer = vbHourglass
Dim cString  As String
Dim cm As New ADODB.command
cString = "Update SUBCUST " & _
           " Set SUBCUST.F_DATE = t.MaxDate" & _
           " From SUBCUST" & _
           "  INNER JOIN (" & _
           " SELECT PHONE, MAX(DATE) AS MaxDate" & _
           " From QFILE6_20H" & _
           " GROUP BY PHONE" & _
           " ) T ON SUBCUST.PHONE = T.PHONE"
Set cm = cmd(cString, con, adText)
cm.Execute
Me.MousePointer = vbNormal
MsgBox " „  ⁄œÌ· «· «—ÌŒ »‰Ã«Õ"
End Sub

Private Sub cmdExcel_Click()
Me.MousePointer = 11
If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 19)
    aRow = AddFlag(aRow, aSub)
End If
ToFileExelNew grid1, , , aRow, , 0.9, , , , , , Me, Array(Me.Caption)
Me.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
Me.MousePointer = vbHourglass
myLoadGrd
Me.MousePointer = vbNormal
End Sub
Private Sub Form_Load()
    openCon con
    
    Set data1.Recordset = cmd("Select Code,DescA From STORE_BR", con).Execute
    Set XSTORE.RowSource = data1
    XSTORE.ListField = "Desca"
    XSTORE.BoundColumn = "Code"
    If Not XSTORE.MatchedWithList Then
        XSTORE.BoundText = cBranchStore
    End If
    XSTORE.Enabled = False
    
    XDATE1.text = myFormat_p(Date)
    XDATE2.text = myFormat_p(Date)
    
    Set grid1.DataSource = DATA2
    
    fixGrd
End Sub
Private Sub myLoadGrd()
    If IsDate(XDATE1.text) Then
        aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(XDATE1.text))
    End If
    
    If IsDate(XDATE2.text) Then
        aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(XDATE2.text))
    End If
    
    If XSTORE.MatchedWithList Then
        aPrm = AddFlag(aPrm, "STORE", XSTORE.BoundText)
    End If
        
    If Trim(xDesca.text) <> "" Then
        aPrm = AddFlag(aPrm, "DESCA", Trim(xDesca.text))
    End If
        
    If Trim(XPHONE.text) <> "" Then
        aPrm = AddFlag(aPrm, "PHONE", Trim(XPHONE.text))
    End If
                        
    Set DATA2.Recordset = cmd("[dbo].[SP_CUST_PHONE_SHOW]", con, adStoredProc, aPrm).Execute
    fixGrd
    If ckhPhone.Value = 1 Then
        CheckPhones
    End If
    If grid1.Rows = 1 Then
        panel1(0).Caption = "·«  ÊÃœ ”Ã·« "
    Else
        panel1(0).Caption = "⁄œœ «·”Ã·«  : " & grid1.Rows - 1
    End If
End Sub
Sub fixGrd()
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .RowHeight(0) = 600
    .WordWrap = True
    
    .TextMatrix(0, 0) = "«·⁄„Ì·"
    .TextMatrix(0, 1) = "«· ·Ì›Ê‰"
    .TextMatrix(0, 2) = " «—ÌŒ «Œ—  ⁄«„·"
    
    .ColWidth(0) = 6000
    .ColWidth(1) = 3000
    .ColWidth(2) = 2000
    .ColWidth(3) = 1500
    .ColFormat(2) = "yyyy/m/d"
    .ColHidden(.Cols - 1) = True
    '.ColComboList(2) = StrList("SELECT CODE,DESCA FROM BRANCH", con)
    '.ColDataType(3) = flexDTDate
    'XCOUNT.Caption = "⁄œœ «·⁄„·«¡ : " & grid1.Rows - 1
    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub Form_Resize()
Dim nHeight As Long
nHeight = Me.Height - grid1.Top - 1200
If nHeight > 3000 Then
    grid1.Height = nHeight
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1

If Col = 1 Then
    If IsValidMobile(.TextMatrix(Row, Col)) Then
        Dim sPhone  As String
        sPhone = myField("SELECT TOP 1 PHONE FROM SUBCUST WHERE ID = " & .TextMatrix(Row, .Cols - 1), con) & ""
        If sPhone <> "" And Trim(sPhone) <> Trim(grid1.TextMatrix(grid1.Row, Col)) Then
            con.BeginTrans
            con.Execute "UPDATE SUBCUST" & _
                        " SET PHONE = " & addstring(.TextMatrix(Row, Col)) & _
                        " WHERE ID = " & .TextMatrix(Row, .Cols - 1)
            
            con.Execute "UPDATE FILE6_20H " & _
                        " SET PHONE = " & addstring(.TextMatrix(Row, Col)) & _
                        " WHERE FILE6_20H.PHONE = " & MyParn(sPhone)
            con.CommitTrans
            Inform "  „ «· ⁄œÌ· "
        End If
    End If
ElseIf Col = 0 And Trim(.TextMatrix(Row, Col)) <> "" Then
    If UpdatePhones(.TextMatrix(Row, 1), .TextMatrix(Row, Col), "", con) Then
        Inform " „ «· ⁄œÌ·"
    End If
End If
End With
End Sub
Private Sub grid1_EnterCell()
With grid1
'    If .Col = 0 Or .Col = 1 And bopt2 = True Then
    If .Col = 0 Or .Col = 1 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'If grid1.Col = 12 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If MsgBox("Õ–› «· ·Ì›Ê‰ ", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
    Dim sDoc_No As String
    sDoc_No = myField("SELECT TOP 1 DOC_NO FROM QFILE6_20H WHERE PHONE = " & MyParn(grid1.TextMatrix(grid1.Row, 1)), con) & ""
    If sDoc_No <> "" Then
        MsgBox "«· ·Ì›Ê‰ „ÊÃÊœ ›Ï «·›« Ê—… —ﬁ„ " & sDoc_No
        Exit Sub
    End If
    On Error GoTo myError
    
    con.Execute "DELETE FROM SUBCUST WHERE PHONE = " & MyParn(grid1.TextMatrix(grid1.Row, 1))
    MsgBox " „ Õ–› «·”Ã·"
    grid1.RemoveItem grid1.Row
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(grid1.EditText) < 3 Then
        Cancel = True
    End If
ElseIf Col = 1 Then
    If Not IsValidMobile(grid1.EditText) Then
        Cancel = True
        Exit Sub
    End If
End If
End Sub
Private Sub XPHONE_GotFocus()
myGotFocus XPHONE
End Sub
Private Sub XPHONE_LostFocus()
myLostFocus XPHONE
End Sub
Private Sub XDATE2_GotFocus()
myGotFocus XDATE2
End Sub
Private Sub XDATE2_LostFocus()
myLostFocus XDATE2
myValidDate XDATE2
End Sub
Private Sub XDATE1_GotFocus()
myGotFocus XDATE1
End Sub
Private Sub XDATE1_LostFocus()
myLostFocus XDATE1
myValidDate XDATE1
End Sub
Private Sub xDesca_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus XSTORE
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus XSTORE
If Not XSTORE.MatchedWithList Then XSTORE.BoundText = ""
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 12 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col < 2 Then
    grid1.Col = Col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 0)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub CheckPhones()
Dim I As Long
Dim Row As Long, nRows As Long
grid1.Redraw = flexRDNone
nRows = grid1.Rows - 1
prog1.Visible = True
For Row = grid1.Rows - 1 To 1 Step -1
   I = I + 1
   prog1.Value = Round(I / nRows, 2) * 100
   If IsValidMobile(grid1.TextMatrix(Row, 1)) Then
        grid1.RemoveItem Row
   End If
Next
prog1.Visible = False
grid1.Redraw = flexRDBuffered
End Sub
