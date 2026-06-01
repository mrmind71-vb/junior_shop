VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form onlineCollectCSVnewfrm 
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   7020
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   4560
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "»Ì«‰«  €Ì— ’«·Õ…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   1725
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "»Ì«‰«  ’«·Õ…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   1410
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   780
      Left            =   11610
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   3930
      Begin Threed.SSCommand CMD_SEND 
         Height          =   600
         Left            =   2655
         TabIndex        =   6
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1058
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
         Picture         =   "collectCSVnew.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "collectCSVnew.frx":2925
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   600
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1058
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
         Picture         =   "collectCSVnew.frx":5179
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   600
         Left            =   1305
         TabIndex        =   12
         Top             =   135
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1058
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
         Picture         =   "collectCSVnew.frx":749C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   15570
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   4695
      Begin Threed.SSCommand cmdWeb 
         Height          =   600
         Left            =   2340
         TabIndex        =   2
         Top             =   135
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   1058
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "”Õ» ”œ«œ „‰ Œ·«· «·„Êﬁ⁄"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdShipping 
         Height          =   600
         Left            =   45
         TabIndex        =   3
         Top             =   135
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   1058
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "ﬂ‘› Õ”«» ‘—ﬂ… «·‘Õ‰"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9780
      Left            =   0
      TabIndex        =   0
      Top             =   855
      Width           =   20310
      _ExtentX        =   35825
      _ExtentY        =   17251
      _Version        =   393216
      Tabs            =   2
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
      TabCaption(0)   =   "»Ê«·’ «” —Ã«⁄ «Ê  »œÌ· »œÊ‰ ›« Ê—…"
      TabPicture(0)   =   "collectCSVnew.frx":98B0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GRID2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "”œ«œ ‘—ﬂ… «·‘Õ‰"
      TabPicture(1)   =   "collectCSVnew.frx":98CC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid1"
      Tab(1).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Height          =   9330
         Left            =   45
         TabIndex        =   13
         Top             =   360
         Width           =   20175
         _cx             =   35586
         _cy             =   16457
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
         BackColorFixed  =   -2147483633
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
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   9330
         Left            =   -74955
         TabIndex        =   14
         Top             =   360
         Width           =   20175
         _cx             =   35586
         _cy             =   16457
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
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   10860
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   6300
      Top             =   7000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "onlineCollectCSVnewfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public bIgPaid As Boolean
Public bIgMinus As Boolean
Public sDoc_no As String
Public sBank As String
Public sShip As String
Private Sub CMD_SEND_Click()
If Not myValid Then
    Exit Sub
End If

Dim nCount As Long
nCount = myreplace
If nCount >= 0 Then
    MsgBox " „  ”ÃÌ· " & nCount & " ”Ã·"
    Unload Me
End If
End Sub

Private Sub cmdExcel_Click()
Me.MousePointer = 11
Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aRow = AddFlag(aRow, aSub)

For i = 1 To grid1.Rows - 2
    If grid1.TextMatrix(i, 2) = "" Then
        aSub = AddFlag(Empty, "row", i)
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "word_wrap", False)
        aSub = AddFlag(aSub, "back_color", 19)
        aRow = AddFlag(aRow, aSub)
    End If
Next
If SSTab1.Tab = 1 Then
    ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, SSTab1.TabCaption(2))
ElseIf SSTab1.Tab = 0 Then
    ToFileExelNew GRID2, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, SSTab1.TabCaption(1))
End If
Me.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdShipping_Click()
cmdShipping.Enabled = False
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False

grid1.Rows = 1
GRID2.Rows = 1

getDataShip

SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = True

cmdShipping.Enabled = True
If GRID2.Rows > 1 Then
    SSTab1.Tab = 0
End If

Option1(0).Value = True
End Sub

Private Sub cmdWeb_Click()
cmdWeb.Enabled = False
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
grid1.Rows = 1
GRID2.Rows = 1
getDataWeb
cmdWeb.Enabled = True
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = True
Option1(0).Value = True
Fixgrd
End Sub
Private Function getDataWeb()
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


If cSv.NumRows < 1 Then Exit Function

If sBank = "" Then
    MsgBox "ﬂÊœ »‰ﬂ €Ì— „ÊÃÊœ"
    Exit Function
End If

Dim loctable As New ADODB.Recordset
Set loctable = mycmd("SELECT * FROM FILE5_10 WHERE CODE = " & MyParn(sBank), con)
If loctable.EOF Then
    MsgBox "»Ì«‰ «·»‰ﬂ €Ì— ’«·Õ"
    Exit Function
End If

Dim id_header As String, type_header As String, gross_header As String
Dim fees_header As String, vat_header As String, net_header As String


id_header = loctable!id_header & ""
type_header = loctable!type_header & ""
gross_header = loctable!gross_header & ""
fees_header = loctable!fees_header & ""
vat_header = loctable!vat_header & ""
net_header = loctable!net_header & ""

If id_header = "" Then
    MsgBox "⁄‰Ê«‰ ⁄„Êœ ﬂÊœ «· ⁄—Ì› €Ì— „”Ã· ›Ï »Ì«‰«  «·»‰ﬂ"
    Exit Function
End If

If type_header = "" Then
    MsgBox "⁄‰Ê«‰ ⁄„Êœ ‰Ê⁄ «·”œ«œ €Ì— „”Ã· ›Ï »Ì«‰«  «·»‰ﬂ"
    Exit Function
End If

If gross_header = "" Then
    MsgBox "⁄‰Ê«‰ ⁄„Êœ «·≈Ã„«·Ì €Ì— „”Ã· ›Ï »Ì«‰«  «·»‰ﬂ"
    Exit Function
End If

If fees_header = "" Then
    MsgBox "⁄‰Ê«‰ ⁄„Êœ «·—”Ê„ €Ì— „”Ã· ›Ï »Ì«‰«  «·»‰ﬂ"
    Exit Function
End If

If net_header = "" Then
    MsgBox "⁄‰Ê«‰ ⁄„Êœ «·—”Ê„ €Ì— „”Ã· ›Ï »Ì«‰«  «·»‰ﬂ"
    Exit Function
End If


Dim Tb As New ChilkatStringBuilder
Dim cString As New ChilkatStringBuilder
prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption

Dim sMsg As String, bDone As Boolean
Dim sid As String
Dim nGross As Double

For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    
    sid = cSv.GetCellByName(i, id_header)
    If sid <> "" Then
        grid1.AddItem ""
        sType = cSv.GetCellByName(i, type_header)
        nGross = Val(cSv.GetCellByName(i, gross_header))
        nFees = Val(cSv.GetCellByName(i, fees_header)) + Val(cSv.GetCellByName(i, vat_header))
        nNet = Val(cSv.GetCellByName(i, net_header))
    
        If LCase(Mid(sType, 1, 1)) = "r" Then
            nGross = nGross * -1
            nNet = nNet * -1
        End If
            
        grid1.TextMatrix(grid1.Rows - 1, 0) = sid
        sMsg = ""
        
        bDone = GrdDesc(sid, grid1.Rows - 1, sMsg)
        If sMsg <> "" Then
            grid1.TextMatrix(grid1.Rows - 1, 2) = sMsg
        End If
        grid1.TextMatrix(grid1.Rows - 1, 10) = nGross
        grid1.TextMatrix(grid1.Rows - 1, 11) = nFees
        grid1.TextMatrix(grid1.Rows - 1, 12) = nNet
        
        grid1.TextMatrix(grid1.Rows - 1, grid1.Cols - 1) = IIf(bDone, -1, 0)
        If Not bDone Then
            grid1.Cell(flexcpBackColor, grid1.Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = &HC0C0FF
        Else
            grid1.Cell(flexcpBackColor, grid1.Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = &HFFFFFF
        End If
    End If
Next
prog1.Visible = False
Me.Caption = sCaption
End Function
Private Sub Form_Load()
openCon con
Fixgrd
Fixgrd2
End Sub
Private Sub Fixgrd()
With grid1
.TextMatrix(0, 0) = "Payment id"
.TextMatrix(0, 1) = "—ﬁ„ «·ÿ·»Ì…"
.TextMatrix(0, 2) = "«·«”„"
.TextMatrix(0, 3) = "«· ·Ì›Ê‰"
.TextMatrix(0, 4) = "⁄œœ «·›Ê« Ì—"
.TextMatrix(0, 5) = "≈Ã„«·Ì «·›Ê« Ì—"
.TextMatrix(0, 6) = "≈Ã„«·Ì „— Ã⁄"
.TextMatrix(0, 7) = "’«›Ì «·›Ê« Ì—"
.TextMatrix(0, 8) = "≈Ã„«·Ì «·„”œœ"
.TextMatrix(0, 9) = "«·»«ﬁÌ"
.TextMatrix(0, 10) = "„”œœ"
.TextMatrix(0, 11) = "⁄„Ê·…"
.TextMatrix(0, 12) = "«·’«›Ì"
.TextMatrix(0, 13) = "OK"

.RowHeight(0) = 600
.ColWidth(0) = 3000
.ColWidth(1) = 1200
.ColWidth(2) = 2500
.ColWidth(3) = 2000
.ColWidth(4) = 800
.ColWidth(5) = 1200
.ColWidth(6) = 1200
.ColWidth(7) = 1200
.ColWidth(8) = 1200
.ColWidth(9) = 1200
.ColWidth(10) = 1200
.ColWidth(11) = 1200
.ColWidth(12) = 1200
.ColWidth(13) = 700

.ColDataType(.Cols - 1) = flexDTBoolean

For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
If grid1.Rows > 1 Then
    .Cell(flexcpFontSize, 1, 0, .Rows - 1, 0) = 9
End If
End With
End Sub
Private Sub Fixgrd2()
With GRID2
.TextMatrix(0, 0) = "»Ê·Ì’… «·‘Õ‰"
.TextMatrix(0, 1) = "—ﬁ„ «·ÿ·»Ì…"
.TextMatrix(0, 2) = "«·«”„"
.TextMatrix(0, 3) = "«· ·Ì›Ê‰"
.TextMatrix(0, 4) = "⁄œœ «·›Ê« Ì—"
.TextMatrix(0, 5) = "≈Ã„«·Ì «·›Ê« Ì—"
.TextMatrix(0, 6) = "≈Ã„«·Ì „— Ã⁄"
.TextMatrix(0, 7) = "’«›Ì «·›Ê« Ì—"
.TextMatrix(0, 8) = "≈Ã„«·Ì «·„”œœ"
.TextMatrix(0, 9) = "«·»«ﬁÌ"
.TextMatrix(0, 10) = "COD"
.TextMatrix(0, 11) = "ﬁÌ„… «·‘Õ‰"
.TextMatrix(0, 12) = "’«›Ì"
.TextMatrix(0, 13) = "OK"

.RowHeight(0) = 600
.ColWidth(0) = 1500
.ColWidth(1) = 2000
.ColWidth(2) = 3000
.ColWidth(3) = 2000
.ColWidth(4) = 800
.ColWidth(5) = 1200
.ColWidth(6) = 1200
.ColWidth(7) = 1200
.ColWidth(8) = 1200
.ColWidth(9) = 1200
.ColWidth(10) = 1200
.ColWidth(11) = 1200
.ColWidth(12) = 1200
.ColWidth(13) = 700

.ColDataType(.Cols - 1) = flexDTBoolean

For i = 1 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Private Function GrdDesc(sPayment_id As String, Row As Long, ByRef pMsg As String) As Boolean
Dim sOrder_No As String
sOrder_No = rsFunc("[dbo].[fn_order_collect]", , addstring(sPayment_id), "default") & ""
If sOrder_No = "" Then
    pMsg = "—ﬁ„ " & "Payment id" & " €Ì— ’ÕÌÕ"
    Exit Function
End If


Dim loctable As New ADODB.Recordset
Set loctable = myRs("SELECT FILE6_90H.*," & _
                    "vw_online_order_total.total," & _
                    "vw_online_order_total.total_plus ," & _
                    "vw_online_order_total.total_minus ," & _
                    "vw_online_order_total.count_of," & _
                    "vw_online_order_pay.value as total_pay" & _
                    " FROM FILE6_90H " & _
                    " LEFT JOIN vw_online_order_total ON FILE6_90H.DOC_NO = vw_online_order_total.ORDER_NO" & _
                    " LEFT JOIN vw_online_order_pay ON FILE6_90H.DOC_NO = vw_online_order_pay.ORDER_NO" & _
                    " WHERE DOC_NO = " & MyParn(sOrder_No), con)

If loctable.EOF Then
    MsgBox "—ﬁ„ ÿ·»»… €Ì— ’ÕÌÕ"
    Exit Function
End If

grid1.TextMatrix(Row, 0) = loctable!PAYMENT_ID & ""
grid1.TextMatrix(Row, 1) = loctable!doc_no & ""
grid1.TextMatrix(Row, 2) = loctable!Name & ""
grid1.TextMatrix(Row, 3) = loctable!phone & ""
grid1.TextMatrix(Row, 4) = loctable!count_of & ""
grid1.TextMatrix(Row, 5) = loctable!total_plus & ""
grid1.TextMatrix(Row, 6) = loctable!total_minus & ""
grid1.TextMatrix(Row, 7) = loctable!total & ""
grid1.TextMatrix(Row, 8) = loctable!total_pay & ""
grid1.TextMatrix(Row, 9) = mRound(loctable!total & "") - mRound(loctable!total_pay)
GrdDesc = True
End Function
Private Sub Form_Resize()
SSTab1.Height = IIf(Me.Height - SSTab1.Top - 1000 < 1000, 1000, Me.Height - SSTab1.Top - 1000)
grid1.Height = SSTab1.Height - 500
GRID2.Height = SSTab1.Height - 500
End Sub

Private Sub GRID1_Click()
'If grid1.col = grid1.Cols - 1 And grid1.ValueMatrix(grid1.Row, grid1.col) = -1 Then
'    grid1.Editable = flexEDKbdMouse
'Else
'    grid1.Editable = flexEDNone
'End If
End Sub

Private Sub Option1_Click(Index As Integer)
Dim i As Long
For i = 1 To grid1.Rows - 1
    If Option1(0).Value Then
        grid1.RowHidden(i) = False
    ElseIf Option1(1).Value Then
        grid1.RowHidden(i) = grid1.ValueMatrix(i, grid1.Cols - 1) = 0
    ElseIf Option1(2).Value Then
        grid1.RowHidden(i) = grid1.ValueMatrix(i, grid1.Cols - 1) <> 0
    End If
Next

For i = 1 To GRID2.Rows - 1
    If Option1(0).Value Then
        GRID2.RowHidden(i) = False
    ElseIf Option1(1).Value Then
        GRID2.RowHidden(i) = GRID2.ValueMatrix(i, GRID2.Cols - 1) = 0
    ElseIf Option1(2).Value Then
        GRID2.RowHidden(i) = GRID2.ValueMatrix(i, GRID2.Cols - 1) <> 0
    End If
Next
End Sub
Private Function myreplace() As Long
Dim i As Long
Dim aInsert As Variant

On Error GoTo myerror
Dim sCaption As String

sCaption = Me.Caption
prog1.Visible = True
With grid1
con.BeginTrans
For i = 1 To .Rows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & grid1.Rows - 1
    prog1.Value = Round(i / (grid1.Rows - 1), 2) * 100
    If .ValueMatrix(i, .Cols - 1) <> 0 Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(sDoc_no))
        aInsert = AddFlag(aInsert, "PAYMENT_ID", addstring(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "ORDER_NO", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "[TOTAL]", grid1.ValueMatrix(i, 10))
        aInsert = AddFlag(aInsert, "[COMMISSION]", grid1.ValueMatrix(i, 11))
        aInsert = AddFlag(aInsert, "[TYPE]", "1")
        con.Execute addInsert(aInsert, "FILE9_10"), nAffect
        myreplace = myreplace + nAffect
    End If
Next
End With


Me.Caption = Caption
prog1.Visible = True

With GRID2
For i = 1 To .Rows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & GRID2.Rows - 1
    prog1.Value = Round(i / (GRID2.Rows - 1), 2) * 100
    If .ValueMatrix(i, .Cols - 1) <> 0 Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(sDoc_no))
        aInsert = AddFlag(aInsert, "SHIP_NO", addstring(.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "ORDER_NO", addstring(.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "[COD]", .ValueMatrix(i, 10))
        aInsert = AddFlag(aInsert, "[FEE]", .ValueMatrix(i, 11))
        aInsert = AddFlag(aInsert, "[TYPE]", "2")
        con.Execute addInsert(aInsert, "FILE9_10"), nAffect
        myreplace = myreplace + nAffect
    End If
Next
End With
con.CommitTrans
Finally:
Me.Caption = sCaption
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
myreplace = -1
GoTo Finally
End Function
Private Function myValid() As Boolean
Dim nCount As Long
If grid1.Rows + GRID2.Rows <= grid1.FixedRows + GRID2.FixedRows Then
    MsgBox "·« ÌÊÃœ »Ì«‰«  ·Õ›ŸÂ«"
    'Exit Function
End If

For i = 1 To grid1.Rows - 1
    If grid1.ValueMatrix(i, grid1.Cols - 1) = 0 Then
        MsgBox "Œÿ√ ›Ì «·»Ì«‰« "
        'Exit Function
    End If
Next

For i = 1 To GRID2.Rows - 1
    If GRID2.ValueMatrix(i, GRID2.Cols - 1) = 0 Then
        MsgBox "Œÿ√ ›Ì «·»Ì«‰« "
       'Exit Function
    End If
Next
myValid = True
End Function
Private Function grdDesc2(sShip_no As String, Row As Long, ByRef pMsg As String) As Boolean
Dim sOrder_No As String

sOrder_No = rsFunc("[dbo].[fn_order_collect]", con, "default", addstring(sShip_no)) & ""
If sOrder_No = "" Then
    'MsgBox "—ﬁ„ »Ê·Ì’… «·‘Õ‰ €Ì— ’ÕÌÕ"
    pMsg = "—ﬁ„ »Ê·Ì’… «·‘Õ‰ €Ì— ’ÕÌÕ"
    Exit Function
End If

Dim loctable As New ADODB.Recordset
Set loctable = myRs("SELECT FILE6_90H.*," & _
                    "vw_online_order_total.total," & _
                    "vw_online_order_total.total_plus ," & _
                    "vw_online_order_total.total_minus ," & _
                    "vw_online_order_total.count_of," & _
                    "vw_online_order_pay.value as total_pay" & _
                    " FROM FILE6_90H " & _
                    " LEFT JOIN vw_online_order_total ON FILE6_90H.DOC_NO = vw_online_order_total.ORDER_NO" & _
                    " LEFT JOIN vw_online_order_pay ON FILE6_90H.DOC_NO = vw_online_order_pay.ORDER_NO" & _
                    " WHERE DOC_NO = " & MyParn(sOrder_No), con)
GRID2.TextMatrix(Row, 1) = loctable!doc_no & ""
GRID2.TextMatrix(Row, 2) = loctable!Name
GRID2.TextMatrix(Row, 3) = loctable!phone & ""
GRID2.TextMatrix(Row, 4) = loctable!count_of & ""
GRID2.TextMatrix(Row, 5) = loctable!total_plus & ""
GRID2.TextMatrix(Row, 6) = loctable!total_minus & ""
GRID2.TextMatrix(Row, 7) = loctable!total & ""
GRID2.TextMatrix(Row, 8) = loctable!total_pay & ""
GRID2.TextMatrix(Row, 9) = mRound(loctable!total & "") - mRound(loctable!total_pay)
grdDesc2 = True
End Function
Private Function getDataShip()
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

Dim loctable As New ADODB.Recordset
Set loctable = mycmd("SELECT * FROM SHIP WHERE CODE = " & MyParn(sShip), con)
If loctable.EOF Then
    MsgBox "»Ì«‰ ‘—ﬂ… «·‘Õ‰ €Ì— ’«·Õ"
    Exit Function
End If

Dim id_header As String, type_header As String, gross_header As String
Dim fees_header As String, vat_header As String, net_header As String


order_header = loctable!order_header & ""
cod_HEADER = loctable!cod_HEADER & ""
fees_header = loctable!fees_header & ""
vat_header = loctable!vat_header & ""
net_header = loctable!net_header & ""

If order_header = "" Then
    MsgBox "⁄‰Ê«‰ ⁄„Êœ —ﬁ„ «·»Ê·Ì’… €Ì— „”Ã· ›Ï »Ì«‰«  ‘—ﬂ… «·‘Õ‰"
    Exit Function
End If

If cod_HEADER = "" Then
    MsgBox "⁄‰Ê«‰ ⁄„Êœ COD €Ì— „”Ã· ›Ï »Ì«‰«  ‘—ﬂ… «·‘Õ‰"
    Exit Function
End If

If fees_header = "" Then
    MsgBox "⁄‰Ê«‰ ⁄„Êœ «·—”Ê„ €Ì— „”Ã· ›Ï »Ì«‰«  ‘—ﬂ… «·‘Õ‰"
    Exit Function
End If

If net_header = "" Then
    MsgBox "⁄‰Ê«‰ ⁄„Êœ «·—”Ê„ €Ì— „”Ã· ›Ï »Ì«‰«  ‘—ﬂ… «·‘Õ‰"
    Exit Function
End If



If cSv.NumRows < 1 Then Exit Function

Dim Tb As New ChilkatStringBuilder
Dim cString As New ChilkatStringBuilder


prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption

Dim sMsg As String

Dim ship_no As String
Dim bDone As Boolean
Dim cod As Double
Dim fee As Double
Dim vat As Double
Dim Value As Double

For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    
    ship_no = cSv.GetCellByName(i, order_header)
    If ship_no <> "" Then
        cod = Val(cSv.GetCellByName(i, cod_HEADER))
        fees = Val(cSv.GetCellByName(i, fees_header))
        vat = Val(cSv.GetCellByName(i, vat_header))
        net = Val(cSv.GetCellByName(i, net_header))
            
            GRID2.AddItem ""
            GRID2.TextMatrix(GRID2.Rows - 1, 0) = ship_no
            sMsg = ""
            bDone = grdDesc2(ship_no, GRID2.Rows - 1, sMsg)
            If sMsg <> "" Then
                GRID2.TextMatrix(GRID2.Rows - 1, 2) = sMsg
            End If
            GRID2.TextMatrix(GRID2.Rows - 1, 10) = cod
            GRID2.TextMatrix(GRID2.Rows - 1, 11) = fees + vat
            GRID2.TextMatrix(GRID2.Rows - 1, 12) = net
            GRID2.TextMatrix(GRID2.Rows - 1, 13) = IIf(bDone, -1, 0)
            If Not bDone Then
                GRID2.Cell(flexcpBackColor, GRID2.Rows - 1, 0, GRID2.Rows - 1, GRID2.Cols - 1) = &HC0C0FF
            Else
                GRID2.Cell(flexcpBackColor, GRID2.Rows - 1, 0, GRID2.Rows - 1, GRID2.Cols - 1) = &HFFFFFF
            End If
    End If
Next
prog1.Visible = False
Me.Caption = sCaption
End Function
Private Sub Option2_Click(Index As Integer)
Dim i As Long
For i = 1 To GRID2.Rows - 1
    If Option2(0).Value Then
        GRID2.RowHidden(i) = False
    ElseIf Option2(1).Value Then
        GRID2.RowHidden(i) = GRID2.ValueMatrix(i, 8) = 0
    ElseIf Option2(2).Value Then
        GRID2.RowHidden(i) = GRID2.ValueMatrix(i, 8) <> 0
    End If
Next
End Sub
