VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form online_stage_invoice_trans 
   BackColor       =   &H00FFFFFF&
   Caption         =   " ÕÊÌ· ÿ·»Ì«  «Ê‰ ·«Ì‰ «·Ì „»Ì⁄« "
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
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   5535
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   270
      Width           =   4110
      Begin Threed.SSCommand cmdSend 
         Height          =   600
         Left            =   1485
         TabIndex        =   15
         Top             =   180
         Width           =   2580
         _ExtentX        =   4551
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
         Caption         =   " ÕÊÌ· «·Ì ›« Ê—… „»Ì⁄« "
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_stage_inovice_trans.frx":0000
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   600
         Left            =   45
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
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
         Picture         =   "online_stage_inovice_trans.frx":22F6
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   9675
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   45
      Width           =   3255
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   585
         Width           =   1680
      End
      Begin VB.TextBox xShip_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   225
         Width           =   1680
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«· «—ÌŒ"
         Height          =   330
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   630
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»Ê·Ì’… «·‘Õ‰"
         Height          =   330
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   270
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   7305
      Begin VB.Label lbldoc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·„” ‰œ «·›—⁄Ì"
         Height          =   330
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label xdoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   2085
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰Ê⁄ «·„” ‰œ «·›—⁄"
         Height          =   330
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label xtype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   2085
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·ÿ·»Ì…"
         Height          =   330
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   825
      End
      Begin VB.Label xorder_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   2130
      End
      Begin VB.Label xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   2130
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«· «—ÌŒ"
         Height          =   330
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   735
      ScaleWidth      =   20370
      TabIndex        =   3
      Top             =   9465
      Width           =   20370
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4110
      Left            =   4320
      TabIndex        =   0
      Top             =   1170
      Width           =   15945
      _cx             =   28125
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
      Cols            =   6
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   4110
      Left            =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5310
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
      Cols            =   13
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
   Begin VSFlex7Ctl.VSFlexGrid grdTotal 
      Height          =   4110
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1170
      Width           =   4245
      _cx             =   7488
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
Attribute VB_Name = "online_stage_invoice_trans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sOrder_No As String
Public sDoc_no As String
Public myForm As Form
Dim bStopCell As Boolean
Dim con As New ADODB.Connection
Dim conServer As New ADODB.Connection

Private Sub CMD_SEND_Click()
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSendInvoice_Click()
If Not myValid Then Exit Sub
Dim sNewDoc As String
If xtype.Tag = "1" Or xtype.Tag = 3 Then
    sNewDoc = myReplaceOnline
End If
If sNewDoc = "" Then Exit Sub
MsgBox " „ «· —ÃÌ· «·Ì „” ‰œ „»Ì⁄«  —ﬁ„ " & sNewDoc
myForm.myload
Unload Me
End Sub

Private Sub Form_Load()

Fixgrd
fixGrdTotal
myAddItem

myload
myLoadGrd
myLoadDiffer
Fixgrd2

CellPos 13, 0, grid1.Cols - 1
End Sub
Private Sub myLoadGrd()
Dim cString As String
cString = "SELECT v.ITEM," & _
          "CAST(NULL AS VARCHAR(10))," & _
          "FACT.DESCA," & _
          "FILE1_10.DESCA," & _
          "FILE1_10.COLOR," & _
          "FILE1_10.SCAL," & _
          "SUM(v.QUANT), " & _
          "MAX(v.PRICE)," & _
          "CAST(NULL AS VARCHAR(10))," & _
          "file1_10.price," & _
          "file1_10.costitem, " & _
          " 0 as QUANT_SELECTED," & _
          " SUM(v.QUANT) as differ " & _
          " From vw_online_items v" & _
          " INNER JOIN FILE1_10 ON v.ITEM = FILE1_10.ITEM " & _
          " INNER JOIN FACT ON FILE1_10.code = FACT.CODE " & _
          " WHERE v.ORDER_NO = " & MyParn(sOrder_No) & _
          " AND v.DOC_NO = " & Val(sDoc_no)

cString = cString & _
           " GROUP BY v.ITEM," & _
            "FACT.DESCA," & _
            "FILE1_10.DESCA," & _
            "FILE1_10.COLOR," & _
            "FILE1_10.SCAL," & _
            "FILE1_10.PRICE," & _
            "FILE1_10.COSTITEM"

On Error GoTo myerror

Set grid2.DataSource = myRs(cString)
Finally:
Set loctable = Nothing
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Resume Finally
End Sub
Sub Fixgrd2()
With grid2
    .RowHeight(0) = 600
    '.Cols = 13
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "SKU"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "«·’‰›"
    .TextMatrix(0, 4) = "«··Ê‰"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«·ﬂ„Ì…"
    .TextMatrix(0, 7) = "price"
    .TextMatrix(0, 8) = "item name"
    .TextMatrix(0, 9) = "Price_item"
    .TextMatrix(0, 10) = "cost"
    .TextMatrix(0, .Cols - 2) = "ﬂ„Ì… „Œ «—…"
    .TextMatrix(0, .Cols - 1) = "«·›—ﬁ"
    
    .ColHidden(1) = True
    .ColHidden(7) = True
    .ColHidden(8) = True
    .ColHidden(9) = True
    .ColHidden(10) = True
    
    
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 2000
    .ColWidth(2) = 1500
    .ColWidth(3) = 5000
    .ColWidth(4) = 3000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    
    .ColWidth(7) = 2000
    .ColWidth(grid2.Cols - 1) = 1000
    
    
    Dim i As Long
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    
    .Subtotal flexSTClear
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTCount, -1, 5, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, 6, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, .Cols - 2, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, .Cols - 1, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
End With
End Sub
Sub Fixgrd()
With grid1
    .RowHeight(0) = 600
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "«·’‰›"
    .TextMatrix(0, 3) = "«··Ê‰"
    .TextMatrix(0, 4) = "„ﬁ«”"
    .TextMatrix(0, 5) = "«·ﬂ„Ì…"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 1500
    .ColWidth(2) = 5000
    .ColWidth(3) = 3000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    
    Dim i As Long
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub myLoadDiffer()
Dim i As Long
With grid2

.Subtotal flexSTClear
If .Rows < 2 Then Exit Sub

For i = 1 To .Rows - 1
    nFound = grdTotal.FindRow(.TextMatrix(i, 0), , 0)
    If nFound = -1 Then
         .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbWhite
         .TextMatrix(i, grid2.Cols - 2) = 0
         .TextMatrix(i, grid2.Cols - 1) = .ValueMatrix(i, 6)
    Else
        .TextMatrix(i, grid2.Cols - 2) = grdTotal.ValueMatrix(nFound, 1)
        .TextMatrix(i, grid2.Cols - 1) = .ValueMatrix(i, 6) - .ValueMatrix(i, grid2.Cols - 2)
        If .ValueMatrix(i, grid2.Cols - 1) = 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0FFC0
        ElseIf .ValueMatrix(i, grid2.Cols - 1) > 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H80FFFF
        Else
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H8080FF
        End If
    End If
Next
End With

With grid1
For i = 1 To .Rows - 2
    nFound = grid2.FindRow(.TextMatrix(i, 0), , 0)
    If nFound = -1 Then
        .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H8080FF
        '.TextMatrix(i, 5) = "e1"
    Else
        If grid2.ValueMatrix(nFound, grid2.Cols - 1) = 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0FFC0
        ElseIf grid2.ValueMatrix(nFound, grid2.Cols - 1) > 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H80FFFF
             '.TextMatrix(i, 8) = "w"
        Else
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H8080FF
            '.TextMatrix(i, 8) = "e2"
        End If
    End If
Next
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set OnlineCheck = Nothing
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
If Not validRow(Row, col) Then
    Exit Sub
End If

If Row = grid1.Rows - 1 Then
    myAddItem
    grid1.Select Row, grid1.Cols - 1
End If

myLoadGrdTotal
Fixgrd
End Sub
Private Sub grid1_EnterCell()
If grid1.col = 0 Or grid1.col = 5 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Function validRow(Row As Long, Optional col As Long = -1) As Boolean
With grid1
    If Not isNum(.TextMatrix(Row, 0)) Then Exit Function
    If .ValueMatrix(Row, 5) <= 0 Then Exit Function
End With
validRow = True
End Function
Private Function GrdDesc(ByRef sitem As String, Row As Long) As Boolean
With grid1
    .TextMatrix(Row, 1) = ""
    .TextMatrix(Row, 2) = ""
    .TextMatrix(Row, 3) = ""
    .TextMatrix(Row, 4) = ""
    
    If Trim(sitem) = "" Then Exit Function
               
    Dim loctable As New ADODB.Recordset
    Dim cString As String
    cString = "SELECT FILE1_10.ITEM," & _
          "FACT.DESCA AS FACT," & _
          "FILE1_10.desca," & _
          "FILE1_10.COLOR," & _
          "FILE1_10.SCAL " & _
          "FROM FILE1_10 " & _
          "INNER JOIN FACT " & _
          "ON FILE1_10.code = FACT.CODE " & _
          "WHERE FILE1_10.ITEM = " & sitem
    
    Set loctable = myRs(cString)
    If loctable.EOF Then Exit Function
    
    .TextMatrix(Row, 1) = loctable!Fact & ""
    .TextMatrix(Row, 2) = loctable!DESCA & ""
    .TextMatrix(Row, 3) = loctable!color & ""
    .TextMatrix(Row, 4) = loctable!SCAL & ""
    If .TextMatrix(Row, 0) = "" Then .TextMatrix(Row, 5) = 1
    GrdDesc = True
End With
End Function
Private Sub myAddItem()
With grid1
.AddItem ""
End With
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 0 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
        Exit Sub
    End If
    
    grid1.EditText = rsItemCode(grid1.EditText)
    If Not isNum(grid1.EditText) Then
        Cancel = True
        Exit Sub
    ElseIf Not GrdDesc(grid1.EditText, Row) Then
        Cancel = True
        Exit Sub
    End If
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf col < 5 Then
    grid1.col = 5
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 5)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, col
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
    'If grid1.Col = 4 Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
        myRemove grid1.Row
        myLoadGrdTotal
        Fixgrd
        bStopCell = True
    End If
End If
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
'CalcTotals
End Sub
Private Sub myLoadGrdTotal()
Dim cString As New ChilkatStringBuilder

Dim i As Long
If grid1.Rows < 3 Then
    grdTotal.Rows = 1
    Exit Sub
End If

cString.Append "WITH tb as ( "
For i = 1 To grid1.Rows - 2
    If i > 1 Then
        cString.Append " UNION ALL "
    End If
    
    cString.Append "SELECT " & _
                grid1.TextMatrix(i, 0) & " AS ITEM " & "," & _
                grid1.TextMatrix(i, 5) & " AS QUANT " & "," & _
                i & " AS ITEM_ORDER"
Next
cString.Append ")"
cString.Append "SELECT ITEM,SUM(QUANT) FROM tb GROUP BY ITEM ORDER BY MIN(ITEM_ORDER)"

Set grdTotal.DataSource = myRs(cString.GetAsString())

fixGrdTotal

myLoadDiffer
Fixgrd2

End Sub
Sub fixGrdTotal()
With grdTotal
    .RowHeight(0) = 600
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·ﬂ„Ì…"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 1000
    
    
    Dim i As Long
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next

    .Subtotal flexSTClear
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTCount, -1, 0, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, 1, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"

End With
End Sub
Private Sub myload()
Dim loctable As New ADODB.Recordset
Dim strSql As String
strSql = "SELECT v.ORDER_NO," & _
          " NULLIF(v.DOC_NO,0) AS DOC_NO," & _
          " v.STAGE," & _
          " v.DATE," & _
          " v.[TYPE]," & _
          " v.type_desca," & _
          " STAGES_CODES.DESCA AS STAGE_DESCA," & _
          " FILE6_25.DESCA AS MAKE_DESCA" & _
          " FROM vw_online_order v " & _
          " INNER JOIN STAGES_CODES ON v.STAGE = STAGES_CODES.CODE" & _
          " LEFT JOIN FILE6_25 ON v.MAN = FILE6_25.CODE" & _
          " WHERE v.ORDER_NO = " & MyParn(sOrder_No) & _
          " AND v.doc_no = " & Val(sDoc_no)

On Error GoTo myerror

Set loctable = myRs(strSql)
If Not loctable.EOF Then
    xorder_no.Caption = loctable!ORDER_NO & ""
    xdoc_no.Caption = loctable!DOC_NO & ""
    xDate.Caption = myFormat_p(loctable!Date)
    xtype.Caption = loctable!TYPE_dESCA
    xtype.Tag = loctable!Type
    xDate.Caption = myFormat_p(loctable!Date)
    xdate1.text = myFormat_p(rsDateBranch(sBranchOnline))
    
    If loctable!Type = 0 Or loctable!Type = 1 Then
        cmdSend.Caption = " ÕÊÌ· «·Ì „»Ì⁄« "
    ElseIf loctable!Type = 2 Then
        cmdSend.Caption = " ÕÊÌ· «·Ì „— Ã⁄"
    ElseIf loctable!Type = 3 Then
        cmdSend.Caption = "«—”«· ··«” »œ«·"
    ElseIf loctable!Type = 4 Then
        cmdSend.Caption = "«” »œ«· «·„— Ã⁄"
    End If
    
    lbldoc.Visible = loctable!Type <> 0
    xdoc_no.Visible = loctable!Type <> 0
End If
Finally:
Set loctable = Nothing
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Resume Finally
End Sub
Private Function myReplaceOnline() As String
    Dim sDoc_New As String
    Dim sDate As String
    Dim cAddress As String
    
    Dim cString As New ChilkatStringBuilder
                                  
    cString.Clear
    cString.Append " SELECT V.ORDER_NO," & _
              " V.DOC_NO," & _
              " V.DATE," & _
              " V.SHIP_NO," & _
              " V.SALES_RET," & _
              " V.MAN," & _
              " V.TYPE," & _
              " V.DISCOUNT," & _
              " FILE6_90H.NAME," & _
              " FILE6_90H.PHONE," & _
              " FILE6_90H.Shipping_City," & _
              " FILE6_90H.CITY," & _
              " FILE6_90H.STREET, " & _
              " FILE6_90H.SHIPPING, " & _
              " FILE6_90H.Payment_Method," & _
              " FILE6_90H.PAYMENT_ID" & _
              " FROM FILE6_90H " & _
              " INNER JOIN vw_online_doc ON v.ORDER_NO =  FILE6_90H.DOC_NO " & _
              " WHERE v.ORDER_NO = " & MyParn(sOrder_No) & _
              " DOC_NO = " & sDoc_no
    
           
    Dim loctable As New ADODB.Recordset
    Set loctable = myRs(cString.GetAsString, conServer)
    
    If loctable.EOF And loctable.BOF Then
        MsgBox "«·„” ‰œ €Ì— „”Ã·"
        Exit Function
    End If
        
            
    aInsert = AddFlag(Empty, "CODE", addstring("0000"))
    aInsert = AddFlag(aInsert, "[Date]", addDate(sDate))
    aInsert = AddFlag(aInsert, "STORE", addstring(sStoreOnline))
    aInsert = AddFlag(aInsert, "DISCOUNT", Val(loctable!discount))
    aInsert = AddFlag(aInsert, "BOX", addstring(cManBox))
    aInsert = AddFlag(aInsert, "userName", addstring(cUserName))
    aInsert = AddFlag(aInsert, "branch", addstring(cBranch))
    aInsert = AddFlag(aInsert, "MAN", addstring(sMan))
    aInsert = AddFlag(aInsert, "PHONE", addstring(loctable!phone))
    aInsert = AddFlag(aInsert, "PAYMENT_ID", addstring(loctable!PAYMENT_ID))
    
    If Val(loctable!discount & "") <> 0 Then
        aInsert = AddFlag(aInsert, "username_disc", addstring("«Ê‰·«Ì‰"))
        aInsert = AddFlag(aInsert, "NOTE_DISC", addstring(loctable!discount_code))
    End If
    
    aInsert = AddFlag(aInsert, "ONLINE", "2")
    aInsert = AddFlag(aInsert, "USER_IP", addstring(GetComputerName))
    
    cAddress = loctable!Shipping_City & ""
    If loctable!Street & "" <> "" Then cAddress = cAddress & Tr(cAddress, " ") & loctable!Street
    
    aInsert = AddFlag(aInsert, "ADDRESS", addstring(cAddress))
    aInsert = AddFlag(aInsert, "[NAME]", addstring(loctable!Name))
    aInsert = AddFlag(aInsert, "[CHARGE2]", addstring(loctable!Shipping))
    aInsert = AddFlag(aInsert, "[online_doc]", addstring(sOrder_No))
    
    If xtype.Tag <> "0" Then
        aInsert = AddFlag(aInsert, "[doc_no_online]", addvalue(sDoc_no))
    End If
        
    aInsert = AddFlag(aInsert, "[city]", addstring(loctable!city))
    aInsert = AddFlag(aInsert, "[Shipping_City]", addstring(loctable!Shipping_City))
    aInsert = AddFlag(aInsert, "[street]", addstring(loctable!Street))
    aInsert = AddFlag(aInsert, "[Payment_Method]", addstring(loctable!Payment_Method))
    
    aInsert = AddFlag(aInsert, "[ship_no]", addstring(xShip_No.text))
    aInsert = AddFlag(aInsert, "[ship]", addstring(xship.text))

    sDoc_New = NewflagDoc(sDate, cManBox, con)
    aInsert = AddFlag(aInsert, "[doc_no]", addstring(sDoc_New))
    
    con.BeginTrans
    con.Execute addInsert(aInsert, "FILE6_20H")
    
    cString.Clear
    cString.Append "insert into FILE6_20(" & _
                    "DOC_NO," & _
                    "ITEM," & _
                    "QUANT," & _
                    "PRICE," & _
                    "[ROW]," & _
                    "[MAN]," & _
                    "[USER_IP]," & _
                    "[PRICE_C]," & _
                    "[PRICE_C2]," & _
                    "[DESCA2]," & _
                    "[COST]," & _
                    "[SKU]" & _
                    ")"
    With grid2
    
    cString.Append " VALUES "
    For i = 1 To .Rows - 2
        If .ValueMatrix(i, .Cols - 2) > 0 And .ValueMatrix(i, .Cols - 1) >= 0 Then
            cString.Append "("
            cString.Append addstring(sDoc_New) & ","
            cString.Append addstring(.TextMatrix(i, 0)) & ","
            cString.Append addstring(.TextMatrix(i, .Cols - 2)) & ","
            cString.Append .ValueMatrix(i, 7) & ","
            cString.Append i & ","
            cString.Append addstring(sMan) & ","
            cString.Append addstring(GetComputerNamecIpName) & ","
            cString.Append .ValueMatrix(i, 9) & ","
            cString.Append .ValueMatrix(i, 7) & ","
            cString.Append addstring(.TextMatrix(i, 8)) & ","
            cString.Append addstring(.TextMatrix(i, 10)) & ","
            cString.Append addstring(.TextMatrix(i, 1)) & ")"
            
            cString.Append ","
        End If
    Next
    cString.Shorten 1
    End With
    
    con.Execute cString.GetAsString()
    UpdateInvTotal sDoc_New, con
    
    con.Execute " UPDATE FILE6_90H SET SALES_DOC = " & addstring(sDoc_New) & ", SALES_DATE = " & addDate(sDate) & " WHERE DOC_NO = " & MyParn(sDoc_no)
    con.CommitTrans
    
    myReplaceOnline = sDoc_New
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function myValid() As Boolean
Dim i As Long
Dim nFound As Long

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«› „”Ã·…"
    Exit Function
End If

With grid1
For i = 1 To .Rows - 2
    nFound = grid2.FindRow(.TextMatrix(i, 0), , 0)
    If nFound = -1 Then
        MsgBox "«·’‰› " & .TextMatrix(i, 0) & " €Ì— „ÊÃÊœ ›Ï «·ÿ·»Ì…"
        Exit Function
    End If
Next
End With

With grid2
For i = 1 To .Rows - 2
    nFound = grid1.FindRow(.TextMatrix(i, 0), , 0)
    If nFound = -1 Then
        If MsgBox("«·’‰› " & .TextMatrix(i, 0) & " ·„ Ì”Ã· ›Ï «·›« Ê—…", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
    ElseIf .TextMatrix(i, grid2.Cols - 1) < 0 Then
        MsgBox "«·’‰› " & .TextMatrix(i, 0) & " «ﬂ»— „‰ «·›« Ê—… » " & Abs(.ValueMatrix(i, grid2.Cols - 1))
        Exit Function
    ElseIf .TextMatrix(i, grid2.Cols - 1) > 0 Then
        If MsgBox("«·’‰› " & .TextMatrix(i, 0) & " «ﬁ· „‰ «·›« Ê—… » " & Abs(.ValueMatrix(i, grid2.Cols - 1)), vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
    End If
Next
myValid = True
End With
End Function
Private Sub grid2_EnterCell()
grid2.Editable = flexEDKbdMouse
End Sub
