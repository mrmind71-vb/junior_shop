VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form orders_soldfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "„— Ã⁄«  «’‰«›"
   ClientHeight    =   10200
   ClientLeft      =   165
   ClientTop       =   510
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
   ScaleHeight     =   10200
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   8370
      Width           =   3930
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
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
         Picture         =   "order_sold2.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand CMD_SEND 
         Height          =   555
         Left            =   1620
         TabIndex        =   4
         Top             =   180
         Width           =   2265
         _ExtentX        =   3995
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
         Picture         =   "order_sold2.frx":2323
         Caption         =   " ”ÃÌ· «’‰«› «·ÿ·»Ì…"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "order_sold2.frx":498A
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4110
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   20220
      _cx             =   35666
      _cy             =   7250
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
      Cols            =   15
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   4110
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4230
      Width           =   20220
      _cx             =   35666
      _cy             =   7250
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
      ForeColorSel    =   0
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   14
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   -765
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
      Left            =   0
      Top             =   -765
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   18135
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   8325
      Width           =   2220
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   555
         Left            =   45
         TabIndex        =   6
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
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
         Picture         =   "order_sold2.frx":6C80
         Caption         =   "«Œ Ì«— ﬂ· «·›« Ê—…"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "order_sold2.frx":9156
      End
   End
End
Attribute VB_Name = "orders_soldfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
Public sDoc_no_order As String

Public myForm As salesTax
Dim bStopCell As Boolean
Dim con As New ADODB.Connection
Dim conServer As New ADODB.Connection
Private Sub CMD_SEND_Click()
If Not MYVALID Then Exit Sub
Dim sNewDoc As String
sNewDoc = myreplace
If sNewDoc = "" Then Exit Sub
MsgBox " „ «· —ÃÌ· «·Ì „” ‰œ „»Ì⁄«  —ﬁ„ " & sNewDoc
myForm.myProcOrder sNewDoc
Unload Me
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSelectAll_Click()
GRID1.Rows = 1
CalcTotals
addRow
End Sub

Private Sub Form_Load()
openCon con

Set GRID1.DataSource = data1
Set GRID2.DataSource = DATA2

myloadgrd2
fixGrd
CalcTotals

CellPos 13 + 1, 0, GRID1.Cols - 1
End Sub
Private Sub myloadgrd2()
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DOC_NO", sDoc_no_order)
aPrm = AddFlag(aPrm, "DOC_NO_IG", TurnValue(sDoc_no))
Set DATA2.Recordset = mycmd("[dbo].[sp_ORDERS_SOLD]", con, adStoredProc, aPrm)
Fixgrd2
If GRID2.Rows > 1 Then
    GRID2.Cell(flexcpFontSize, 1, 4 + 1, GRID2.Rows - 1, 4 + 1) = 10
End If
For i = 1 To GRID2.Rows - 1
    GRID2.TextMatrix(i, 0) = i
Next
End Sub
Sub Fixgrd2()
With GRID2
    .RowHeight(0) = 600

    .TextMatrix(0, 0) = "„"
    .TextMatrix(0, 0 + 1) = "ﬂÊœ"
    .TextMatrix(0, 1 + 1) = "«·„Ê”„"
    .TextMatrix(0, 2 + 1) = "«·„’‰⁄"
    .TextMatrix(0, 3 + 1) = "„ÊœÌ· «·„’‰⁄"
    
    .TextMatrix(0, 4 + 1) = "«·’‰›"
    .TextMatrix(0, 5 + 1) = "„ﬁ«”"
    .TextMatrix(0, 6 + 1) = "«··Ê‰"
    .TextMatrix(0, 7 + 1) = "”⁄— «·„” Â·ﬂ"
    .TextMatrix(0, 8 + 1) = "«·ﬂ„Ì…"
    .TextMatrix(0, 9 + 1) = "«·”⁄—"
    .TextMatrix(0, 10 + 1) = "«·≈Ã„«·Ì"
    .TextMatrix(0, 11 + 1) = "«·»«∆⁄"
    
    .TextMatrix(0, 12 + 1) = "„»«⁄"
    .TextMatrix(0, 13 + 1) = "«·»«ﬁÌ"
    .TextMatrix(0, 14 + 1) = "«· ﬂ·›…"
    
    .ColWidth(0) = 600
    .ColWidth(0 + 1) = 1000
    .ColWidth(1 + 1) = 1000
    .ColWidth(2 + 1) = 1500
    .ColWidth(3 + 1) = 1400
    .ColWidth(4 + 1) = 6000
    .ColWidth(5 + 1) = 1000
    .ColWidth(6 + 1) = 1000
    .ColWidth(7 + 1) = 1000
    .ColWidth(8 + 1) = 1000
    .ColWidth(9 + 1) = 1000
    .ColWidth(10 + 1) = 1000
    .ColWidth(11 + 1) = 1000
    .ColWidth(12 + 1) = 1000
    .ColWidth(13 + 1) = 1000
    .ColWidth(14 + 1) = 1000

    .ColHidden(11 + 1) = True
    .ColHidden(.Cols - 4) = True
    .ColHidden(.Cols - 3) = True
    .ColHidden(.Cols - 2) = True
    .ColHidden(.Cols - 1) = True

    .ColWidth(GRID2.Cols - 1) = 1000
    
    Dim i As Long
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColAlignment(4 + 1) = flexAlignLeftCenter
End With
End Sub
Private Sub CalcTotals()
GRID1.Subtotal flexSTClear
GRID2.Subtotal flexSTClear

With GRID2
For i = 1 To .Rows - 1
    nFound = GRID1.FindRow(.TextMatrix(i, 0 + 1), , 0 + 1)
    If nFound <> -1 Then
         .TextMatrix(i, 12 + 1) = .ValueMatrix(i, .Cols - 1) + GRID1.ValueMatrix(nFound, 8 + 1)
         .TextMatrix(i, 13 + 1) = .ValueMatrix(i, 8 + 1) - .ValueMatrix(i, 12 + 1)
    Else
         .TextMatrix(i, 12 + 1) = .ValueMatrix(i, .Cols - 1)
         .TextMatrix(i, 13 + 1) = .ValueMatrix(i, 8 + 1) - .ValueMatrix(i, 12 + 1)
    End If
    
    If .ValueMatrix(i, 13 + 1) = 0 Then
        .Cell(flexcpBackColor, i, 0 + 1, i, .Cols - 1) = &HC0C0FF
    ElseIf .ValueMatrix(i, 13 + 1) > 0 And .ValueMatrix(i, 13 + 1) < .ValueMatrix(i, 8 + 1) Then
        .Cell(flexcpBackColor, i, 0 + 1, i, .Cols - 1) = &H80FFFF
    ElseIf .ValueMatrix(i, 12 + 1) = .ValueMatrix(i, .Cols - 1) Then
        .Cell(flexcpBackColor, i, 0 + 1, i, .Cols - 1) = vbWhite
    End If
Next


.SubtotalPosition = flexSTBelow
.Subtotal flexSTSum, -1, 8 + 1, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
.Subtotal flexSTSum, -1, 10 + 1, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
.Subtotal flexSTSum, -1, 11 + 1, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
.Subtotal flexSTSum, -1, 12 + 1, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
.Subtotal flexSTSum, -1, 13 + 1, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
End With

For i = 1 To GRID1.Rows - 1
    GRID1.TextMatrix(i, 0) = i
Next
If GRID1.Rows > 1 Then
    GRID1.Cell(flexcpFontSize, 1, 4 + 1, GRID1.Rows - 1, 4 + 1) = 10
End If

With GRID1
    .Subtotal flexSTClear
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 8 + 1, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, 10 + 1, "#.##", &HE0E0E0, , True, "«·≈Ã„«·Ì"
End With
End Sub
Sub fixGrd()
With GRID1
    .RowHeight(0) = 600
    '.Cols = 13
    .TextMatrix(0, 0) = "„"
    .TextMatrix(0, 0 + 1) = "ﬂÊœ"
    .TextMatrix(0, 1 + 1) = "«·„Ê”„"
    .TextMatrix(0, 2 + 1) = "«·„’‰⁄"
    .TextMatrix(0, 3 + 1) = "„ÊœÌ· «·„’‰⁄"
    
    .TextMatrix(0, 4 + 1) = "«·’‰›"
    .TextMatrix(0, 5 + 1) = "„ﬁ«”"
    .TextMatrix(0, 6 + 1) = "«··Ê‰"
    .TextMatrix(0, 7 + 1) = "”⁄— «·„” Â·ﬂ"
    .TextMatrix(0, 8 + 1) = "«·ﬂ„Ì…"
    .TextMatrix(0, 9 + 1) = "«·”⁄—"
    .TextMatrix(0, 10 + 1) = "«·≈Ã„«·Ì"
    
    .TextMatrix(0, 11 + 1) = "«· ﬂ·›…"
    .TextMatrix(0, 12 + 1) = "DESCA2"
    .TextMatrix(0, 13 + 1) = "SKU"
        
    .ColWidth(0) = 600
    .ColWidth(0 + 1) = 1000
    .ColWidth(1 + 1) = 1000
    .ColWidth(2 + 1) = 1500
    .ColWidth(3 + 1) = 1400
    .ColWidth(4 + 1) = 7800
    .ColWidth(5 + 1) = 1000
    .ColWidth(6 + 1) = 1000
    .ColWidth(7 + 1) = 1200
    .ColWidth(8 + 1) = 750
    .ColWidth(9 + 1) = 1200
    .ColWidth(10 + 1) = 1400
    .ColWidth(11 + 1) = 1200
    
    .ColHidden(.Cols - 3) = True
    .ColHidden(.Cols - 2) = True
    .ColHidden(.Cols - 1) = True
        
    Dim i As Long
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColAlignment(4 + 1) = flexAlignLeftCenter
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set orders_soldfrm = Nothing
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
CalcTotals
End Sub
Private Sub grid1_EnterCell()
If GRID1.col = 8 + 1 Then
    GRID1.Editable = flexEDKbdMouse
Else
    GRID1.Editable = flexEDNone
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If Val(GRID1.EditText) <= 0 Then
    Cancel = True
    Exit Sub
End If

Dim nFound As Long
nFound = GRID2.FindRow(GRID1.TextMatrix(Row, 0 + 1), , 0)
If GRID2.ValueMatrix(nFound, 8 + 1) - GRID2.ValueMatrix(nFound, GRID1.Cols - 1) < Val(GRID1.EditText) Then
    Cancel = True
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf col < 8 + 1 Then
    GRID1.col = 8 + 1
ElseIf Row < GRID1.Rows - 1 Then
    GRID1.Select Row + 1, NextEmpty(GRID1, Row + 1, 8 + 1, 8 + 1)
    GRID1.ShowCell GRID1.Row, 8 + 1
Else
    GRID1.Select Row, col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 4 Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, GRID1.Row, GRID1.col
ElseIf KeyCode = 46 And GRID1.Row <> GRID1.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
        myRemove GRID1.Row
        CalcTotals
        bStopCell = True
    End If
End If
End Sub
Private Sub myRemove(Row As Long)
GRID1.RemoveItem Row
End Sub
Private Function MYVALID() As Boolean
Dim i As Long
Dim nFound As Long

If GRID1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«› „”Ã·…"
    Exit Function
End If

With GRID2
For i = 1 To .Rows - 2
    nFound = GRID1.FindRow(.TextMatrix(i, 0 + 1), , 0)
    If nFound = -1 Then
        If .ValueMatrix(i, 13 + 1) > 0 Then
            If MsgBox("«·’‰› " & .TextMatrix(i, 4 + 1) & " ·„ Ì”Ã· ›Ï „·› «· ”·Ì„", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
        End If
    ElseIf .TextMatrix(i, 13 + 1) < 0 Then
        MsgBox "«·’‰› " & .TextMatrix(i, 4 + 1) & " «ﬂ»— „‰ „·› «· ”·Ì„ » " & Abs(.ValueMatrix(i, 13 + 1))
        Exit Function
    ElseIf .TextMatrix(i, 13 + 1) > 0 Then
        If MsgBox("«·’‰› " & .TextMatrix(i, 4 + 1) & "  „ »ﬁÌ »ﬂ„Ì… " & .ValueMatrix(i, 13 + 1), vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
    End If
Next

If Not IsDate(myForm.xDate.text) Then
    MsgBox " «—ÌŒ «·›« Ê—… €Ì— „”Ã·"
    Exit Function
End If

If Not myForm.XSTORE.MatchedWithList Then
    MsgBox "«·„Œ“‰ €Ì— „”Ã·"
    Exit Function
End If

If Not myForm.XMAN.MatchedWithList Then
    MsgBox "«·»«∆⁄ €Ì— „”Ã·"
    Exit Function
End If

MYVALID = True
End With
End Function
Private Sub Grid2_dblClick()
With GRID2
If GRID2.Row = 0 Or GRID2.Row = GRID2.Rows - 1 Then Exit Sub
addRow GRID2.Row
End With
End Sub
Private Sub addRow(Optional pRow As Long = -1)
Dim i As Long
Dim Row As Long

'grid1.Subtotal flexSTClear
For Row = IIf(pRow = -1, 1, pRow) To IIf(pRow = -1, GRID2.Rows - 2, pRow)
    If GRID2.ValueMatrix(Row, 13 + 1) > 0 Then
        GRID1.AddItem ""
        For i = 0 + 1 To 10 + 1
            GRID1.TextMatrix(GRID1.Rows - 1, i) = GRID2.TextMatrix(Row, i)
        Next
        GRID1.TextMatrix(GRID1.Rows - 1, 11 + 1) = GRID2.TextMatrix(Row, 14 + 1)
        GRID1.TextMatrix(GRID1.Rows - 1, 12 + 1) = GRID2.TextMatrix(Row, 15 + 1)
        GRID1.TextMatrix(GRID1.Rows - 1, 13 + 1) = GRID2.TextMatrix(Row, 16 + 1)
    End If
Next

CalcTotals
End Sub
Private Function myreplace() As String
    Dim sDoc_New As String
    Dim cString As New ChilkatStringBuilder
    Dim aInsert As Variant
    
    Dim orderTable As New ADODB.Recordset
    cString.Append "Select * from FILE6_52H where DOC_NO = " & MyParn(sDoc_no_order)
    Set orderTable = cmd(cString.GetAsString, con).Execute
    
    If orderTable.EOF Then
        MsgBox "·« ÌÊÃœ „” ‰œ «” ·«„"
        Exit Function
    End If
    
        
    If IsNull(orderTable!code) Then
        MsgBox "„”” ‰œ «· ”·Ì„ »œÊ‰ ⁄„Ì·"
        Exit Function
    End If
    
  
    cString.Clear
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(myForm.xDoc_No.text))
    aInsert = AddFlag(aInsert, "CODE", addstring(orderTable!code))
    aInsert = AddFlag(aInsert, "[Date]", addDate(myForm.xDate.text))
    aInsert = AddFlag(aInsert, "STORE", addstring(myForm.XSTORE.BoundText))
    aInsert = AddFlag(aInsert, "CASH", "0")
    aInsert = AddFlag(aInsert, "ORDER_NO", addstring(orderTable!doc_no))
    aInsert = AddFlag(aInsert, "PO_NO", addstring(orderTable!PO_NO))
    aInsert = AddFlag(aInsert, "USERNAME", addstring(myForm.xusername.Caption))
    aInsert = AddFlag(aInsert, "TIME1", addDate("GETDATE()"))
    aInsert = AddFlag(aInsert, "USER_IP", addstring(cIpName))
    aInsert = AddFlag(aInsert, "MAN", addstring(myForm.XMAN.BoundText))
    aInsert = AddFlag(aInsert, "ISINVOICE", "1")
    aInsert = AddFlag(aInsert, "ISRET", "0")
    
    con.BeginTrans
    con.Execute addInsert(aInsert, "FILE6_20H")
            
    cString.Clear
    Dim nAdd As Long
    With GRID1
    For i = 1 To .Rows - 2
        If nAdd >= 1000 Then
            cString.Shorten 1
            con.Execute cString.GetAsString
            cString.Clear
            nAdd = 0
        End If
        
        If .ValueMatrix(i, 8) > 0 Then
            nAdd = nAdd + 1
            If cString.Length = 0 Then
                cString.Append "insert into FILE6_20(" & _
                    "DOC_NO," & _
                    "ITEM," & _
                    "QUANT," & _
                    "PRICE," & _
                    "[ROW]," & _
                    "[MAN]," & _
                    "[USER_IP]," & _
                    "[PRICE_C]," & _
                    "[COST]," & _
                    "[DESCA2]" & _
                    ")" & _
                    "VALUES"
            End If
            cString.Append "("
            cString.Append addstring(myForm.xDoc_No.text) & ","
            cString.Append addstring(.TextMatrix(i, 0 + 1)) & ","
            cString.Append .ValueMatrix(i, 8 + 1) & ","
            cString.Append .ValueMatrix(i, 9 + 1) & ","
            cString.Append i & ","
            cString.Append addstring(myForm.XMAN.BoundText) & ","
            cString.Append addstring(GetComputerNamecIpName) & ","
            cString.Append .ValueMatrix(i, 7 + 1) & ","
            cString.Append .ValueMatrix(i, 11 + 1) & ","
            cString.Append addstring(.TextMatrix(i, 12 + 1)) & ")"
            cString.Append ","
        End If
    Next
    End With
    
    If cString.Length > 0 Then
        cString.Shorten 1
        con.Execute cString.GetAsString()
        UpdateInvTotal sDoc_New, con
    End If
    
    con.CommitTrans
    
    myreplace = myForm.xDoc_No.text
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
