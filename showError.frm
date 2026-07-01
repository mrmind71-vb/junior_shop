VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form ShowError 
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6990
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   3885
      _cx             =   6853
      _cy             =   12330
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
      Rows            =   1
      Cols            =   2
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
      Editable        =   2
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
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   7065
      Width           =   3885
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   1260
         Picture         =   "showError.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   2520
         Picture         =   "showError.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "зяж"
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "showError.frx":4C15
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1185
      End
   End
End
Attribute VB_Name = "ShowError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sCaption As String
Private Sub cmdExel_Click()
ToFileExelNew grid1, , , , , , , , , , , , Array(Me.Caption)
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdPrint_Click()
printGrdNew.doprint grid1, 0.75, -3, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 4), retHeader(aHeader, 6, 4), False, False, 10, , aRow
printGrdNew.Show 1
End Sub
Private Sub Form_Load()
Me.Caption = sCaption
Fixgrd
'myload
End Sub
Sub Fixgrd()
With grid1
        
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = "#"
    .TextMatrix(0, 1) = "Item"
    
    .ColWidth(0) = 500
    .ColWidth(1) = 3000
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set ShowError = Nothing
End Sub


