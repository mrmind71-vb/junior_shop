VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdOrdersDelivery_popfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   " ﬁ—Ì— ÿ·»Ì«  - «” ·«„ - „»Ì⁄«  - —’Ìœ - ‰”»…  Ê—Ìœ  «·√’‰«›  ›’Ì·Ï"
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   90
      Width           =   4155
      Begin VB.Label xPO_NO 
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
         Height          =   375
         Left            =   1035
         TabIndex        =   9
         Top             =   225
         Width           =   2895
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   8
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   90
      Width           =   4785
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   3600
         TabIndex        =   3
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
         Picture         =   "grdOrdersDelivery_pop.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   2475
         TabIndex        =   2
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
         Picture         =   "grdOrdersDelivery_pop.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdOrdersDelivery_pop.frx":4699
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   1260
         TabIndex        =   1
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "grdOrdersDelivery_pop.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSql 
         Height          =   555
         Left            =   45
         TabIndex        =   0
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "grdOrdersDelivery_pop.frx":8D47
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7395
      Left            =   90
      TabIndex        =   4
      Top             =   855
      Width           =   20175
      _cx             =   35586
      _cy             =   13044
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   10095
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   3330
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
End
Attribute VB_Name = "grdOrdersDelivery_popfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sPO_NO As String
Dim con As New ADODB.Connection
Dim aHeader()
Private Sub cmdExcel_Click()
Me.MousePointer = 11

Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aRow = AddFlag(aRow, aSub)


If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 40)
    aRow = AddFlag(aRow, aSub)
End If

ToFileExelNew grid1, , , aRow, , 1.2, , , , , , Me, Array(Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), retHeader(aHeader, 4, 5))

Me.MousePointer = 0
fixGrd
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
Me.MousePointer = vbHourglass
myload
Me.MousePointer = vbNormal
End Sub

Private Sub cmdPdf_Click()
myPrint "1", True
End Sub
Private Sub cmdPrint_Click()
myPrint
End Sub

Private Sub cmdSql_Click()
myload True
End Sub
Private Sub Form_Resize()
grid1.Height = IIf(Me.Height - grid1.Top - 1000 < 3000, 3000, Me.Height - grid1.Top - 1000)
End Sub
Private Sub Form_Load()
openCon con
XPO_NO.Caption = sPO_NO
Set grid1.DataSource = DATA11
fixGrd
myload
End Sub
Private Sub myload(Optional bString As Boolean = False)
With grid1

Dim cString As New ChilkatStringBuilder
ReDim aHeader(0)
cString.Append "SELECT FILE1_10.ITEM," & _
          "FACT.DESCA," & _
          "FILE1_10.MODELFACT0," & _
          "COALESCE(FILE1_10.DESCA2,FILE1_10.DESCA) AS DESCA," & _
          "FILE1_10.COLOR," & _
          "FILE1_10.SCAL," & _
          "FILE1_10.COST," & _
          "FILE1_10.costitem," & _
          "FILE1_10.PRICE," & _
          "FILE1_10.PRICE2," & _
          "FILE1_10.BARCODE," & _
          "FILE1_10.ENA," & _
          "FILE4_10.DESCA," & _
          "SUM(vw_orders_delivery.QUANT_ORDERED) AS QUANT," & _
          "SUM(vw_orders_delivery.QUANT_CONFIRMED) AS QUANT_CONFIRMED," & _
          "COALESCE(vw.QUANT_DONE,0)," & _
          "SUM(vw_orders_delivery.QUANT_RECEIVED) AS RECEVD," & _
          "SUM(vw_orders_delivery.QUANT_SOLD) AS SALES," & _
          "COUNT(DISTINCT vw_orders_delivery.PO_NO) AS COUNT1," & _
          "COUNT(DISTINCT vw_orders_delivery.PO_NO_RCV) AS COUNT2," & _
          "COUNT(DISTINCT vw_orders_delivery.PO_NO_SALES) AS COUNT3"
          
cString.Append " FROM FILE1_10" & _
                " INNER JOIN FACT " & _
                " ON FILE1_10.FACT = FACT.CODE" & _
                " INNER JOIN FILE4_10" & _
                " ON FILE4_10.CODE = FILE1_10.CODE" & _
                " INNER JOIN vw_orders_delivery" & _
                " INNER JOIN FILE6_51H ON vw_orders_delivery.PO_NO = FILE6_51H.DOC_NO" & _
                " ON FILE1_10.ITEM = vw_orders_delivery.ITEM" & _
                " LEFT JOIN vw_purchase_orders_clients AS vw ON vw_orders_delivery.PO_NO = vw.ORDER_NO" & _
                " AND vw.ITEM = vw_orders_delivery.ITEM "

Dim cWhere As New ChilkatStringBuilder
If SSP_NO <> "" <> "" Then
    cWhere.Append "PO_NO = " & MyParn(sPO_NO)
    cWhere.Append " AND "
    aHeader(0) = "—ﬁ„ «·ÿ·»Ì… : " & sPO_NO
End If


If cWhere.length <> 0 Then
    cWhere.Shorten 5
    cString.Append " WHERE " & cWhere.GetAsString()
End If

cString.Append " GROUP BY FILE1_10.model," & _
               " FACT.DESCA," & _
               " FILE1_10.DESCA," & _
               " FILE1_10.DESCA2," & _
               " FILE1_10.MOSM," & _
               " FILE1_10.MODELFACT0," & _
               " FILE1_10.COLOR," & _
               " FILE1_10.SCAL," & _
               " FILE1_10.COST," & _
               " FILE1_10.costitem," & _
               " FILE1_10.PRICE," & _
               " FILE1_10.PRICE2," & _
               " FILE1_10.BARCODE," & _
               " FILE1_10.ENA," & _
               " FILE4_10.DESCA," & _
               " FILE1_10.ITEM," & _
               " FILE1_10.MOSM," & _
               " FILE1_10.C_SCAL," & _
               " COALESCE(vw.QUANT_DONE,0)"

cString.Append " ORDER BY FILE1_10.model," & _
               " FACT.DESCA," & _
               " FILE1_10.MOSM," & _
               " FILE1_10.MODELFACT0," & _
               " FILE1_10.COLOR," & _
               " FILE1_10.C_SCAL"

If bString Then
    Clipboard.Clear
    Clipboard.SetText cString.GetAsString
    Exit Sub
End If

Set DATA11.Recordset = cmd(cString.GetAsString, con).Execute
End With
fixGrd
End Sub
Sub fixGrd()
With grid1
    .RowHeight(0) = 600
    .WordWrap = True
    
    .TextMatrix(0, 0) = "«·’‰›"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "SKU"
    .TextMatrix(0, 3) = "«·«”„"
    .TextMatrix(0, 4) = "«··Ê‰"
    .TextMatrix(0, 5) = "«·„ﬁ«”"
    
    .TextMatrix(0, 6) = "«·”⁄—"
    .TextMatrix(0, 7) = "Net Purchese PRICE"
    .TextMatrix(0, 8) = "Amazon cost"
    .TextMatrix(0, 9) = "List Price"
    
    .TextMatrix(0, 10) = "BARCODE"
    .TextMatrix(0, 11) = "ASIN"
    
    .TextMatrix(0, 12) = "«·„Ê—œ"
    .TextMatrix(0, 13) = "ﬂ„Ì… ÿ·»Ì…"
    .TextMatrix(0, 14) = "ﬂ„Ì… „ƒﬂœ…"
    .TextMatrix(0, 14 + 1) = "ﬂ„Ì… „‰›–…"
    .TextMatrix(0, 15 + 1) = "ﬂ„»… „”·„…"
    .TextMatrix(0, 16 + 1) = "ﬂ„»… „»«⁄…"
    
    .TextMatrix(0, 17 + 1) = "⁄œœ «·ÿ·»Ì« "
    .TextMatrix(0, 18 + 1) = "ÿ·»Ì«  „” ·„…"
    .TextMatrix(0, 19 + 1) = "ÿ·»Ì«  „»«⁄…"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 1300
    .ColWidth(3) = 3200
    .ColWidth(4) = 1700
    .ColWidth(5) = 800
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    .ColWidth(9) = 1300
    
    .ColWidth(10) = 1500
    .ColWidth(11) = 1700
    
    .ColWidth(12) = 1400
    
    .ColWidth(13) = 1000
    .ColWidth(14) = 1000
    .ColWidth(14 + 1) = 950
    .ColWidth(15 + 1) = 950
    .ColWidth(16 + 1) = 950
    .ColWidth(17 + 1) = 950
    .ColWidth(18 + 1) = 950
    .ColWidth(19 + 1) = 950
    
    .ColHidden(6) = True
    .ColHidden(7) = True
    .ColHidden(8) = True
    .ColHidden(9) = True
    .ColHidden(12) = True

    For i = 13 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next

    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 11
    

    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 13, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 14, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 14 + 1, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 15 + 1, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 16 + 1, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdOrdersDeliveryfrm = Nothing
End Sub
Private Function MYVALID() As Boolean
MYVALID = True
End Function
Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
If grid1.Rows < 3 Then Exit Sub
Dim aRow As Variant, aSub As Variant

aSub = AddFlag(Empty, "row", grid1.Rows - 1)
aSub = AddFlag(aSub, "col", 1)
aSub = AddFlag(aSub, "cols", 2)
aSub = AddFlag(aSub, "text", "≈Ã„«·Ì")
aRow = AddFlag(aRow, aSub)

Set printGrdNew.myForm = Me
printGrdNew.DOPRINT Me.grid1, 0.75, 0, "ÃÊ‰ÌÊ—", Me.Caption, retHeader(aHeader, 0, 2), , True, True, 8, , aRow, Array(1)

If Not bIgPreview Then
    printGrdNew.Show 1
Else
    Unload printGrdNew
End If
End Sub

