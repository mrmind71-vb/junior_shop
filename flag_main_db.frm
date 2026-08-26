VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form flag_dbfrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   7710
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   600
      Left            =   135
      MaskColor       =   &H00FFFFFF&
      Picture         =   "flag_main_db.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   4905
      UseMaskColor    =   -1  'True
      Width           =   1590
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4740
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   7440
      _cx             =   13123
      _cy             =   8361
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
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   1755
      TabIndex        =   2
      Top             =   4815
      Width           =   5820
      Begin VB.TextBox xDesca 
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
         Height          =   405
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   3660
      End
      Begin VB.Label Label1 
         Caption         =   "«·»Ì«‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   1860
      End
   End
End
Attribute VB_Name = "flag_dbfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Public sCaption As String, sTable As String, sFieldCode As String, sFieldDesca As String, sFieldName1 As String, sFieldName2 As String, nZero As Long
Public sFilter As String
Private Sub myload()
Dim cString As String, cWhere As String
cString = "SELECT " & sFieldCode & "," & sFieldDesca & " FROM " & sTable
If Trim(xDesca.text) <> "" Then
    cWhere = MyParnAnd(xDesca.text, sFieldDesca)
End If
If sFilter <> "" Then cWhere = cWhere & Tr(cWhere) & sFilter
If cWhere <> "" Then cString = cString & " where " & cWhere
Dim db As New clsDb
Set grid1.DataSource = db.myRs(cString)
Set db = Nothing
myAddItem
fixGrd
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set flag_dbfrm = Nothing
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
With grid1
If Not validRow(Row) Then Exit Sub
If Row = grid1.Rows - 1 Then myAddItem

Dim aInsert As Variant
Dim db As New clsDb
aInsert = AddFlag(aInsert, sFieldDesca, addstring(.TextMatrix(Row, 1)))
If grid1.TextMatrix(Row, 0) = "" Then
    .TextMatrix(Row, 0) = IIf(nZero = -1, db.NewFlag(sTable, sFieldCode), RetZero(db.NewFlag(sTable, sFieldCode), nZero))
    aInsert = AddFlag(aInsert, sFieldCode, IIf(nZero = -1, addvalue(.TextMatrix(Row, 0)), addstring(.TextMatrix(Row, 0))))
   If db.Execute(addInsert(aInsert, sTable)) = -1 Then myload
Else
   If db.Execute(addUpdate(aInsert, sTable, "CODE = " & IIf(nZero = -1, .TextMatrix(Row, 0), MyParn(.TextMatrix(Row, 0))))) = -1 Then myload
End If
End With
cleanUp:
Set db = Nothing
End Sub
Private Sub grid1_EnterCell()
If bEdit And grid1.col = 1 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim db As New clsDb
If KeyCode = 46 And bEdit Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("«·€«¡ «·”Ã· ?? Â· «‰  „ √ﬂœ", vbYesNo) = vbYes Then
                If db.Execute("Delete  from " & sTable & "  where code = " & IIf(nZero = -1, grid1.TextMatrix(grid1.Row, 0), MyParn(grid1.TextMatrix(grid1.Row, 0)))) <> -1 Then
                grid1.RemoveItem grid1.Row
            End If
        End If
    End If
End If
Set db = Nothing
End Sub
Private Sub Form_Load()
bEdit = True
Me.Caption = sCaption
If sFieldCode = "" Then sFieldCode = "CODE"
If sFieldDesca = "" Then sFieldDesca = "desca"
If sFieldName1 = "" Then sFieldName1 = "«·ﬂÊœ"
If sFieldName2 = "" Then sFieldName2 = "«·»Ì«‰"
If nZero = 0 Then nZero = -1
Label1.Caption = sFieldName2 & turn(sFieldName2, " :")

myload

CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, col
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox sFieldName1 & "€Ì— „ÊÃÊœ"
        Cancel = True
    End If
End If
End Sub

Private Sub xDesca_Change()
myload
End Sub

Private Function MYVALID(nRow) As Boolean
If nZero = -1 Then
    If Not IsNumeric(grid1.TextMatrix(nRow, 0)) Then Exit Function
Else
    If Trim(grid1.TextMatrix(nRow, 0)) = "" Then Exit Function
End If
If Trim(grid1.TextMatrix(nRow, 1)) = "" Then Exit Function
MYVALID = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, 0) = "" Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
    End If
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 And grid1.TextMatrix(.Row, 0) = "" Then
    .RemoveItem .Row
End If
End With
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Trim(.TextMatrix(nRow, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub fixGrd()
With grid1
.ColWidth(0) = 1000
.ColWidth(1) = 6000
.TextMatrix(0, 0) = sFieldName1
.TextMatrix(0, 1) = sFieldName2
For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpBackColor, 1, 0, .Rows - 1, 0) = &H8000000F
End With
End Sub
Private Sub myAddItem()
grid1.AddItem ""
grid1.Cell(flexcpBackColor, 1, 0, grid1.Rows - 1, 0) = &H8000000F
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If col < grid1.Cols - 1 Then
    grid1.Select Row, col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.ShowCell Row + 1, 0
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 1)
Else
    grid1.Select Row, col
End If
End Sub
