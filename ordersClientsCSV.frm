VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form OrdersClientsCsv 
   Caption         =   "”Õ» ”œ«œ CSV"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   405
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   45
      Width           =   2265
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ﬂ„Ì«  „”·„… ›ﬁÿ"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   315
         Value           =   1  'Checked
         Width           =   1680
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   13470
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·„·› <> «·ÿ·»Ì…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   9
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   315
         Width           =   1770
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„”·„ = «·ÿ·»Ì…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   8
         Left            =   2005
         RightToLeft     =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   315
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„”·„ <> «·ÿ·»Ì…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   3605
         RightToLeft     =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   315
         Width           =   1725
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   12600
         RightToLeft     =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   315
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "»Ì«‰«  ’«·Õ…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   11085
         RightToLeft     =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   315
         Width           =   1365
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "€Ì— ’«·Õ…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   9710
         RightToLeft     =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   315
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "€Ì— „ÊÃÊœ… ›Ì «·ÿ·»Ì…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   7480
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   315
         Width           =   2130
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„”·„ «ﬂ»— „‰ «·ÿ·»Ì…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   5430
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   315
         Width           =   1950
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   16200
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   4065
      Begin Threed.SSCommand CMD_SEND 
         Height          =   555
         Left            =   1125
         TabIndex        =   3
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         Picture         =   "ordersClientsCSV.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "ordersClientsCSV.frx":2925
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         Picture         =   "ordersClientsCSV.frx":5179
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdWeb 
         Height          =   555
         Left            =   2205
         TabIndex        =   5
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   979
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
         Caption         =   "Excel «” —œ«œ „‰"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8790
      Left            =   -45
      TabIndex        =   0
      Top             =   855
      Width           =   20310
      _ExtentX        =   35825
      _ExtentY        =   15505
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "«Œÿ«¡ „·› «·«” ·«„"
      TabPicture(0)   =   "ordersClientsCSV.frx":749C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "«·«’‰«› «·„” ·„…"
      TabPicture(1)   =   "ordersClientsCSV.frx":74B8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Grid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   8970
         Left            =   -74910
         TabIndex        =   6
         Top             =   360
         Width           =   20130
         _cx             =   35507
         _cy             =   15822
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
         BackColorFixed  =   192
         ForeColorFixed  =   16777215
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
         Cols            =   11
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
      Begin VSFlex7Ctl.VSFlexGrid Grid1 
         Height          =   8340
         Left            =   90
         TabIndex        =   7
         Top             =   360
         Width           =   20130
         _cx             =   35507
         _cy             =   14711
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
         BackColorFixed  =   12648384
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
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
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            Height          =   870
            Left            =   -3330
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   2475
            Visible         =   0   'False
            Width           =   2940
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "ﬂ„Ì«  „” ·„…"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1425
               RightToLeft     =   -1  'True
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   450
               Width           =   1455
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "€Ì— „”·„…"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   450
               Width           =   1140
            End
         End
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   10860
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   4635
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
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
End
Attribute VB_Name = "OrdersClientsCsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public myForm As Form
Dim sPO_NO As String
Dim sDate As String
Dim sCode As String
Private Sub Check1_Click()
FilterGrd grid1
FilterGrd grid2
End Sub

Private Sub CMD_SEND_Click()
If Not myValid Then
    Exit Sub
End If
Dim sDoc_no_add As String, nCount As Long
nCount = Val(myreplace(sDoc_no_add) & "")
If nCount >= 0 Then
    'MsgBox " „  ”ÃÌ· " & nCount & " ”Ã·"
    myForm.sDoc_no_add = sDoc_no_add
    Unload Me
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdWeb_Click()
cmdWeb.Enabled = False
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
grid1.Rows = 1
grid2.Rows = 1
sDate = ""
sCode = ""
sPO_NO = ""

getCSV

FilterGrd grid1
FilterGrd grid2
cmdWeb.Enabled = True
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = True
SSTab1.Tab = 1
End Sub
Private Sub Form_Load()
openCon con
Set grid1.DataSource = data1
Fixgrd
Fixgrd2
End Sub
Private Sub Fixgrd()
With grid1
.Cols = 11
.TextMatrix(0, 0) = "#"
.TextMatrix(0, 1) = "ASIN"
.TextMatrix(0, 2) = "Order No"
.TextMatrix(0, 3) = "Description"
.TextMatrix(0, 4) = "Quant"
.TextMatrix(0, 5) = "Quant Po"
.TextMatrix(0, 6) = "Quant Order"
.TextMatrix(0, 7) = "Price"
.TextMatrix(0, 8) = "item"
.TextMatrix(0, 9) = "id"
.TextMatrix(0, .Cols - 1) = "Error"
.RowHeight(0) = 600
.ColWidth(0) = 600
.ColWidth(1) = 1600
.ColWidth(2) = 1600
.ColWidth(3) = 8000
.ColWidth(4) = 1300
.ColWidth(5) = 1300
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColWidth(.Cols - 1) = 3000
.ColHidden(.Cols - 3) = True
.ColHidden(.Cols - 2) = True

Dim col As Long
For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
    If .TextMatrix(i, .Cols - 2) = "" Then
        .TextMatrix(i, .Cols - 1) = "€Ì— „ÊÃÊœ ›Ï «·ÿ·»Ì…"
    ElseIf .ValueMatrix(i, 6) <> .ValueMatrix(i, 5) Then
        .TextMatrix(i, .Cols - 1) = "ﬂ„Ì… «·ÿ·»Ì… <> ﬂ„Ì… «·„·›"
    ElseIf .ValueMatrix(i, 4) > .ValueMatrix(i, 5) Then
        .TextMatrix(i, .Cols - 1) = "«·ﬂ„Ì… «·„”·„… «ﬂ»— „‰ «·ÿ·»Ì…"
    End If
    
    If .TextMatrix(i, .Cols - 1) <> "" Then
        If .ValueMatrix(i, 4) > 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H8080FF
        Else
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0C0FF
        End If
        grid2.AddItem ""
        For col = 0 To .Cols - 1
            grid2.TextMatrix(grid2.Rows - 1, col) = .TextMatrix(i, col)
        Next
    End If
Next

If .Rows > 1 Then
    .Cell(flexcpFontSize, 1, .Cols - 1, .Rows - 1, .Cols - 1) = 11
End If
End With
End Sub
Private Sub Fixgrd2()
With grid2
.TextMatrix(0, 0) = "#"
.TextMatrix(0, 1) = "ASIN"
.TextMatrix(0, 2) = "Order No"
.TextMatrix(0, 3) = "Description"
.TextMatrix(0, 4) = "Quant"
.TextMatrix(0, 5) = "Quant Po"
.TextMatrix(0, 6) = "Quant Order"
.TextMatrix(0, 7) = "Price"
.TextMatrix(0, 8) = "item"

.TextMatrix(0, 9) = "ID"
.TextMatrix(0, 10) = "Error"


.RowHeight(0) = 600
.ColWidth(0) = 600
.ColWidth(1) = 1600
.ColWidth(2) = 1600
.ColWidth(3) = 8000
.ColWidth(4) = 1300
.ColWidth(5) = 1300
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColWidth(.Cols - 1) = 3000

.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True
If .Rows > 1 Then
    .Cell(flexcpFontSize, 1, .Cols - 1, .Rows - 1, .Cols - 1) = 11
End If
End With
End Sub
Private Sub Form_Resize()
SSTab1.Height = IIf(Me.Height - SSTab1.Top - 1000 < 1000, 1000, Me.Height - SSTab1.Top - 1000)
grid1.Height = SSTab1.Height - 500
grid2.Height = SSTab1.Height - 500
End Sub

Private Sub grid1_EnterCell()
'If Grid1.col = Grid1.Cols - 1 And Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1) = "«·ﬂ„Ì… «·„”·„… «ﬂ»— „‰ «·ÿ·»Ì…" Then
'    Grid1.Editable = flexEDKbdMouse
'Else
'    Grid1.Editable = True
'End If
End Sub

Private Sub Option1_Click(Index As Integer)
FilterGrd grid1
FilterGrd grid2
End Sub
Private Sub FilterGrd(pGrid)
Dim i As Long
With pGrid
For i = 1 To .Rows - 1
    If Option1(0).Value Then
        .RowHidden(i) = False
    ElseIf Option1(1).Value Then
        .RowHidden(i) = .ValueMatrix(i, 4) = 0
    ElseIf Option1(2).Value Then
        .RowHidden(i) = .ValueMatrix(i, 4) > 0
    ElseIf Option1(3).Value Then
        .RowHidden(i) = .TextMatrix(i, .Cols - 1) <> ""
    ElseIf Option1(4).Value Then
        .RowHidden(i) = .TextMatrix(i, .Cols - 1) = ""
    ElseIf Option1(5).Value Then
        .RowHidden(i) = .TextMatrix(i, .Cols - 2) <> ""
    ElseIf Option1(6).Value Then
        .RowHidden(i) = .ValueMatrix(i, 4) <= .ValueMatrix(i, 5)
    ElseIf Option1(7).Value Then
        .RowHidden(i) = (.ValueMatrix(i, 4) = .ValueMatrix(i, 5)) Or .ValueMatrix(i, 4) = 0 Or .TextMatrix(i, .Cols - 2) = ""
    ElseIf Option1(8).Value Then
        .RowHidden(i) = (.ValueMatrix(i, 4) <> .ValueMatrix(i, 5))
    ElseIf Option1(9).Value Then
        .RowHidden(i) = (.ValueMatrix(i, 5) = .ValueMatrix(i, 6)) Or .TextMatrix(i, .Cols - 2) = ""
    End If
    
    If Check1.Value = 1 Then
        If (Not .RowHidden(i)) And .ValueMatrix(i, 4) = 0 Then
            .RowHidden(i) = True
        End If
    End If
Next
End With
End Sub

Private Function myValid() As Boolean
Dim nCount As Long
If grid1.Rows < 2 Then
    MsgBox "·« ÌÊÃœ »Ì«‰«  ·Õ›ŸÂ«"
    Exit Function
End If

'For i = 1 To grid1.Rows - 1
'    If grid1.ValueMatrix(i, 4) <> 0 Then
'        If grid1.TextMatrix(i, grid1.Cols - 1) <> "" Then
'            MsgBox grid1.TextMatrix(i, grid1.Cols - 1)
'            If grid1.TextMatrix(i, grid1.Cols - 1) <> "«·ﬂ„Ì… «·„”·„… «ﬂ»— „‰ «·ÿ·»Ì…" Then
'                'Exit Function
'            End If
'        End If
'    End If
'Next
myValid = True
End Function
Private Function getCSV() As String
Dim cFileName As String
Common1.InitDir = App.Path & "\CSV\AMAZON"
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
Dim Doc_No_ As String, date_ As String, desca2_ As String
Dim asin_ As String, desca_ As String, quant_ As String, price_ As String

Doc_No_ = "Order Id"
asin_ = "ASIN"
date_ = "Unloaded Date"
quant_ = "Unloaded Quantity"
QUANT_PO = "PO Quantity"


prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption

sPO_NO = cSv.GetCellByName(0, Doc_No_)
If sPO_NO = "" Then
   MsgBox "—ﬁ„ «·„” ‰œ €Ì— „”Ã·"
   Exit Function
End If
        
        
sDate = cSv.GetCellByName(0, date_)
If Not IsDate(sDate) Then
    sDate = myFormat(Date)
End If
        
Dim ordTable As New ADODB.Recordset
Set ordTable = cmd("select * from file6_51H WHERE DOC_NO = " & MyParn(cSv.GetCellByName(0, Doc_No_)), con).Execute
If ordTable.EOF Then
   MsgBox "„” ‰œ «·ÿ·»Ì… €Ì— „ÊÃÊœ"
   Exit Function
End If

If IsNull(ordTable!code) Then
   MsgBox "„” ‰œ «·ÿ·»Ì… »œÊ‰ ⁄„Ì·"
   Exit Function
End If

sCode = ordTable!code
Set ordTable = Nothing

If myField("select * from file6_52H WHERE PO_NO = " & MyParn(sPO_NO), con) & "" <> "" Then
    If MsgBox("„” ‰œ «·ÿ·»Ì… „ÊÃÊœ ›Ï «–Ê‰ «· ”·Ì„", vbDefaultButton2 + vbOKCancel) <> vbOK Then
        Exit Function
    End If
End If

Dim Tb As New ChilkatStringBuilder
Dim sb As New ChilkatStringBuilder

For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    
    If Tb.length <> 0 Then
        Tb.AppendLine " UNION ALL ", 1
    End If
    
    Tb.AppendLine "SELECT " & MyParn(cSv.GetCellByName(i, Doc_No_)) & " AS DOC_NO,", 1
    Tb.AppendLine MyParn(cSv.GetCellByName(i, asin_)) & " AS ASIN,", 1
    Tb.AppendLine Val(cSv.GetCellByName(i, quant_)) & " AS QUANT,", 1
    Tb.AppendLine Val(cSv.GetCellByName(i, QUANT_PO)) & " AS QUANT_PO,", 1
    Tb.AppendLine i + 1 & " AS record_number", 1
Next

sb.AppendLine "WITH tb AS (", 1
sb.AppendLine Tb.GetAsString, 1
sb.AppendLine ")", 1
sb.AppendLine "SELECT tb.ASIN," & _
               "FILE6_51.DOC_NO," & _
               "FILE6_51.DESCA," & _
               "tb.QUANT," & _
               "tb.QUANT_PO," & _
               "FILE6_51.QUANT," & _
               "FILE6_51.PRICE," & _
               "FILE6_51.ITEM," & _
               "FILE6_51.ID" & _
               " FROM TB " & _
               " LEFT JOIN FILE6_51 ON TB.ASIN = FILE6_51.ASIN AND TB.DOC_NO = FILE6_51.DOC_NO " & _
               " ORDER BY tb.Record_Number ", 1

Set data1.Recordset = cmd(sb.GetAsString, con).Execute
Fixgrd
Fixgrd2

prog1.Visible = False
Me.Caption = sCaption
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function myreplace(ByRef pDoc_no As String) As String
Dim cString As New ChilkatStringBuilder
Dim sDoc_no As String, sCaption As String
Dim i As Long
sDoc_no = myField("(SELECT dbo.fn_retzero((SELECT COALESCE(MAX(DOC_NO),0) FROM FILE6_52H),6,1))", con)
cString.Append "INSERT INTO FILE6_52H"
cString.Append "("
cString.Append "DOC_NO,"
cString.Append "PO_NO,"
cString.Append "CODE,"
cString.Append "[DATE]"
cString.Append ")"

cString.Append "VALUES"
cString.Append "("
cString.Append addstring(sDoc_no) & ","
cString.Append addstring(sPO_NO) & ","
cString.Append addstring(sCode) & ","
cString.Append addDate(sDate)
cString.Append ");"

On Error GoTo myerror
con.BeginTrans
con.Execute cString.GetAsString
sCaption = Me.Caption
For i = 1 To grid1.Rows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & grid1.Rows - 1
    prog1.Value = Round(i / (grid1.Rows - 1), 2) * 100
    If grid1.ValueMatrix(i, 4) <> 0 Then
        cString.Clear
        cString.Append "INSERT INTO FILE6_52"
        cString.Append "("
        cString.Append "DOC_NO,"
        cString.Append "ITEM,"
        cString.Append "[ASIN],"
        cString.Append "[DESCA],"
        cString.Append "Quant,"
        cString.Append "Quant_PO,"
        cString.Append "PRICE,"
        cString.Append "ID_PO"
        cString.Append ")"
        
        cString.Append "VALUES"
        cString.Append "("
        cString.Append MyParn(sDoc_no) & ","
        cString.Append addvalue(grid1.TextMatrix(i, grid1.Cols - 3)) & ","
        cString.Append addstring(grid1.TextMatrix(i, 1)) & ","
        cString.Append addstring(grid1.TextMatrix(i, 3)) & ","
        cString.Append grid1.ValueMatrix(i, 4) & ","
        cString.Append grid1.ValueMatrix(i, 5) & ","
        cString.Append grid1.ValueMatrix(i, 7) & ","
        cString.Append grid1.ValueMatrix(i, grid1.Cols - 2)
        cString.Append ")"
        con.Execute cString.GetAsString
        myreplace = myRepalce + 1
    End If
Next
con.CommitTrans
prog1.Visible = False
Me.Caption = sCaption
pDoc_no = sDoc_no
Exit Function
myerror:
myreplace = -1
MsgBox Err.Description
Err.Clear
If TransCount(con) > 0 Then
    con.RollbackTrans
End If
End Function

