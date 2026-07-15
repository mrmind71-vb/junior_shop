VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form ShowOrdersInvoicesfrm 
   Caption         =   "≈ŸÂ«— ›Ê« Ì— «·ÿ·»Ì« "
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   13245
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4155
      Left            =   90
      TabIndex        =   0
      Top             =   495
      Width           =   13110
      _cx             =   23125
      _cy             =   7329
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
      Cols            =   8
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
      Top             =   4635
      Width           =   3660
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1215
         Picture         =   "showOrdersInvoices.frx":0000
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
         Picture         =   "showOrdersInvoices.frx":27EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "showOrdersInvoices.frx":4C15
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Label xOrder_No 
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
      Left            =   8820
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   135
      Width           =   3165
   End
   Begin VB.Label Label1 
      Caption         =   "—ﬁ„ «·ÿ·»Ì… :"
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
      Left            =   12060
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   135
      Width           =   1050
   End
End
Attribute VB_Name = "ShowOrdersInvoicesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sOrder_No As String
Public sWhere As String
Dim aHeader(2) As String
Dim aRow As Variant, aXl As Variant
Dim dbm As New DBManage
Private Sub cmdExel_Click()
ToFileExelNew grid1, , , aXl, , 1
End Sub
Private Sub cmdExit_Click()
Set dbm = Nothing
Unload Me
End Sub
Private Sub cmdPrint_Click()
'Dim aRow(0) As Variant
'aRow(0) = AddFlag(Empty, "row", 1)
'aRow(0) = AddFlag(aRow(0), "col", 0)
'aRow(0) = AddFlag(aRow(0), "cols", IIf(retFlag(aData, "ITEM") = "", 5, 4))
printGrdNew.doprint grid1, 0.85, -3, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 4), retHeader(aHeader, 6, 4), False, False, 10
printGrdNew.Show 1
End Sub
Private Sub Form_Load()
xOrder_No.Caption = sOrder_No
myload
End Sub
Sub fixGrd()
With grid1
        
    .RowHeight(0) = 800
    .WordWrap = True
  
    .ColWidth(0) = 1800
    .ColWidth(1) = 1800
    .ColWidth(2) = 1500
    .ColWidth(3) = 1000
    .ColWidth(4) = 1100
    .ColWidth(5) = 1800
    .ColWidth(6) = 1300
    .ColWidth(7) = 2500
                            
    .ExplorerBar = flexExSort
    
    .SubtotalPosition = flexSTBelow
    .ColDataType(2) = flexDTDate
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    
    .Subtotal flexSTSum, -1, 3, "#", &HC0FFC0, , True, "≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, 4, "#", &HC0FFC0, , True, "≈Ã„«·Ì"
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .Select .Rows - 1, 0
    .ShowCell .Rows - 1, 0
End With
End Sub
Private Sub myload()
Dim strSql As String
    strSql = "SELECT " & _
            " FILE6_20H.[DOC_NO] AS [—ﬁ„ «·›« Ê—…]" & _
            ",ONLINE_TYPES.DESCA AS [‰Ê⁄ «·›« Ê—…]" & _
            ",FORMAT(FILE6_20H.[DATE],'yyyy/M/d') AS [ «—ÌŒ «·ÿ·»Ì…]" & _
            ",FILE6_20H.[TOTAL_QUANT] AS [«·ﬂ„Ì…]" & _
            ",FILE6_20H.[TOTAL_ITEM] - FILE6_20H.[DISCOUNT] AS [«·ﬁÌ„…]" & _
            ",COALESCE(FILE6_20H.SHIP_NO,FILE6_90S.SHIP_NO) AS [»Ê·Ì’… «·‘Õ‰]" & _
            ",FORMAT(COALESCE(FILE6_90SH.[DATE],FILE6_20H.DATE_SHIP),'yyy/M/d') AS [ «—ÌŒ «·‘Õ‰]" & _
            ",FILE6_25.DESCA AS [«·„‰œÊ»]" & _
            " FROM FILE6_20H" & _
            " INNER JOIN ONLINE_TYPES ON FILE6_20H.INV_TYPE_ONLINE = ONLINE_TYPES.CODE" & _
            " LEFT JOIN FILE6_90S ON FILE6_20H.DOC_NO = FILE6_90S.INV_NO" & _
            " LEFT JOIN FILE6_90SH ON FILE6_90S.DOC_NO = FILE6_90SH.DOC_NO " & _
            " LEFT JOIN FILE6_25 ON FILE6_20H.MAN = FILE6_25.CODE "
strSql = strSql & " WHERE FILE6_20H.ONLINE_DOC = " & MyParn(sOrder_No)
If sWhere <> "" Then
    strSql = strSql & " AND " & sWhere
End If
strSql = strSql & " ORDER BY FILE6_20H.DATE,FILE6_20H.DOC_NO"
Set grid1.DataSource = dbm.myRs(strSql)
fixGrd
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set dbm = Nothing
Set ShowOrdersInvoicesfrm = Nothing
End Sub


