VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ShowOrderfrm 
   Caption         =   "«ŸÂ«— «’‰«› ÿ·»Ì…"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   15735
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   4200
      Top             =   5100
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Height          =   4425
      Left            =   90
      TabIndex        =   0
      Top             =   720
      Width           =   15585
      _cx             =   27490
      _cy             =   7805
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5130
      Width           =   3660
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1215
         Picture         =   "showOrder.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2415
         Picture         =   "showOrder.frx":27EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "showOrder.frx":4C15
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   15540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   15540
   End
End
Attribute VB_Name = "ShowOrderfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
Dim aHeader(2) As String
Dim aRow As Variant, aXl As Variant
Private Sub cmdExel_Click()
ToFileExelNew grid1, , , aXl, , 1
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdPrint_Click()
'Dim aRow(0) As Variant
'aRow(0) = AddFlag(Empty, "row", 1)
'aRow(0) = AddFlag(aRow(0), "col", 0)
'aRow(0) = AddFlag(aRow(0), "cols", IIf(retFlag(aData, "ITEM") = "", 5, 4))
printGrdNew.doprint grid1, 0.75, -3, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 4), retHeader(aHeader, 6, 4), False, False, 10, , aRow
printGrdNew.Show 1
End Sub
Private Sub Form_Load()
myload
End Sub
Sub Fixgrd()
With grid1
        
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "SKU"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "«·’‰›"
    .TextMatrix(0, 4) = "«··Ê‰"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«·ﬂ„Ì…"
    .TextMatrix(0, 7) = "«·”⁄—"
    .TextMatrix(0, 8) = "«·≈Ã„«·Ì"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1800
    .ColWidth(2) = 1500
    .ColWidth(3) = 3500
    .ColWidth(4) = 2500
    .ColWidth(5) = 800
    .ColWidth(6) = 800
    .ColWidth(7) = 1000
    .ColWidth(8) = 1400
    
    .ColHidden(.Cols - 2) = True
    .ColHidden(.Cols - 1) = True
        
                
    .ExplorerBar = flexExSort
    
    .SubtotalPosition = flexSTBelow
    
    
    .Subtotal flexSTSum, -1, 6, "#", &HC0FFC0, , True, "≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, 8, "#", &HC0FFC0, , True, "≈Ã„«·Ì"
    
    
    If grid1.Rows > 1 Then
        aSub = AddFlag(Empty, "row", grid1.Rows - 1)
        aSub = AddFlag(aSub, "col", 0)
        aSub = AddFlag(aSub, "cols", 6)
        aRow = AddFlag(aRow, aSub)
    
        aSub = AddFlag(Empty, "row", grid1.Rows - 1)
        aSub = AddFlag(aSub, "col", 0)
        aSub = AddFlag(aSub, "cols", 6)
        aSub = AddFlag(aSub, "text", "«·≈Ã„«·Ì")
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "back_color", 19)
        aXl = AddFlag(aXl, aSub)
        
        
        If grid1.ValueMatrix(1, 9) <> 0 Then
            grid1.AddItem ""
            grid1.AddItem ""
            
            aSub = AddFlag(Empty, "row", grid1.Rows - 2)
            aSub = AddFlag(aSub, "col", 0)
            aSub = AddFlag(aSub, "cols", 8)
            aRow = AddFlag(aRow, aSub)
            
            aSub = AddFlag(Empty, "row", grid1.Rows - 2)
            aSub = AddFlag(aSub, "col", 0)
            aSub = AddFlag(aSub, "cols", 8)
            aSub = AddFlag(aSub, "bold", True)
            aSub = AddFlag(aSub, "back_color", 19)
            aXl = AddFlag(aXl, aSub)
                
            aSub = AddFlag(Empty, "row", grid1.Rows - 1)
            aSub = AddFlag(aSub, "col", 0)
            aSub = AddFlag(aSub, "cols", 8)
            aRow = AddFlag(aRow, aSub)
            
            aSub = AddFlag(Empty, "row", grid1.Rows - 1)
            aSub = AddFlag(aSub, "col", 0)
            aSub = AddFlag(aSub, "cols", 8)
            aSub = AddFlag(aSub, "bold", True)
            aSub = AddFlag(aSub, "back_color", 19)
            aXl = AddFlag(aXl, aSub)
                                               
            
            For i = 0 To 7
                If i <= 5 Then grid1.TextMatrix(grid1.Rows - 3, i) = "≈Ã„«·Ì"
                grid1.TextMatrix(grid1.Rows - 2, i) = "«·Œ’„"
                grid1.TextMatrix(grid1.Rows - 1, i) = "«·’«›Ì"
            Next
            
            grid1.TextMatrix(grid1.Rows - 2, 8) = grid1.ValueMatrix(1, 9)
            grid1.TextMatrix(grid1.Rows - 1, 8) = Round(grid1.ValueMatrix(grid1.Rows - 3, 8) - grid1.ValueMatrix(grid1.Rows - 2, 8), 2)
            
            grid1.MergeCells = flexMergeRestrictRows
            
            grid1.MergeRow(grid1.Rows - 3) = True
            grid1.MergeRow(grid1.Rows - 2) = True
            grid1.MergeRow(grid1.Rows - 1) = True
            
            .Cell(flexcpBackColor, .Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = &HC0FFC0
            .Cell(flexcpBackColor, .Rows - 2, 0, grid1.Rows - 1, grid1.Cols - 1) = &HC0FFC0
            
            
        End If
    End If
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .Select .Rows - 1, 0
    .ShowCell .Rows - 1, 0
End With
End Sub
Private Sub myload()
Dim cString As String
cString = " SELECT FILE6_90.ITEM," & _
          " FILE6_90.SKU," & _
          " FACT.DESCA," & _
          " COALESCE(FILE1_10.desca,FILE6_90.ITEM_NAME)," & _
          " FILE1_10.COLOR," & _
          " FILE1_10.SCAL," & _
          " FILE6_90.QUANT," & _
          " FILE6_90.PRICE," & _
          " FILE6_90.TOTAL," & _
          " FILE6_90H.DISCOUNT," & _
          " FILE6_90H.DATE " & _
          " FROM  FILE6_90 " & _
          " INNER JOIN FILE6_90H ON FILE6_90.DOC_NO = FILE6_90H.DOC_NO " & _
          " LEFT  JOIN FILE1_10 ON FILE6_90.ITEM = FILE1_10.ITEM" & _
          " LEFT  JOIN  FACT ON FILE1_10.code = FACT.CODE" & _
          " WHERE FILE6_90H.DOC_NO =  " & MyParn(sDoc_no)
cString = cString & " ORDER BY DATE"
Set grid1.DataSource = myRs(cString)
If grid1.Rows > 1 Then
    Label1.Caption = "»Ì«‰«  ÿ·»Ì… —ﬁ„ : " & sDoc_no
    Label2.Caption = "» «—ÌŒ : " & myFormat_p(grid1.TextMatrix(1, 10))
    
    aHeader(0) = Label1.Caption
    aHeader(1) = Label2.Caption
End If
Fixgrd
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set ShowOrderfrm = Nothing
End Sub


