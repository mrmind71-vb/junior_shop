VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form online_orders_cash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   11175
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4740
      Left            =   270
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _cx             =   19076
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
   Begin Threed.SSCommand cmdExit 
      Height          =   600
      Left            =   270
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4770
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
      Picture         =   "online_orders_cash.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -450
      Top             =   720
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   -495
      Top             =   1080
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   -1665
      Top             =   1215
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
End
Attribute VB_Name = "online_orders_cash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Public sOrder_no As String
Dim cList As String
Dim dbm As New DBManage
Dim con As New ADODB.Connection
Private Sub myload(Optional bSql As Boolean = False)
Dim cString As String
cString = "SELECT [SHIP] AS [‘—ﬂ… «·‘Õ‰], " & _
           " SHIP_NO AS [»Ê·Ì’… «·‘Õ‰]," & _
           " FORMAT(DATE,'yyyy/M/d') AS [«· «—ÌŒ]," & _
           "[VALUE] AS [«·ﬁÌ„…], " & _
           "[DESCA] AS [«·»Ì«‰], " & _
           "ID " & _
           "FROM FILE6_90C  " & _
           "WHERE ORDER_NO = " & MyParn(sOrder_no)

Set grid1.DataSource = dbm.myRs(cString)
myAddItem
fixGrd
End Sub

Private Sub cmdExcel_Click()

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set online_orders_cash = Nothing
Set dbm = Nothing
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
With grid1

If Not validRow(Row) Then Exit Sub

If grid1.Row = grid1.Rows - 2 Then
    grid1.TextMatrix(grid1.Rows - 1, 2) = grid1.TextMatrix(grid1.Rows - 2, 2)
End If

If Row = .Rows - 1 Then
   myAddItem
End If

Dim aInsert As Variant
aInsert = AddFlag(Empty, "[SHIP]", addstring(.TextMatrix(Row, 0)))
aInsert = AddFlag(aInsert, "SHIP_NO", addstring(.TextMatrix(Row, 1)))
aInsert = AddFlag(aInsert, "[DATE]", addDate(.TextMatrix(Row, 2)))
aInsert = AddFlag(aInsert, "[VALUE]", .ValueMatrix(Row, 3))
aInsert = AddFlag(aInsert, "[DESCA]", addstring(.TextMatrix(Row, 4)))
If .TextMatrix(Row, grid1.Cols - 1) = "" Then
    aInsert = AddFlag(aInsert, "ORDER_NO", addstring(sOrder_no))
    dbm.Execute (addInsert(aInsert, "FILE6_90C"))
    myload
Else
    If Not dbm.Execute(addUpdate(aInsert, "FILE6_90C", "ID = " & .TextMatrix(Row, .Cols - 1))) Then myload
End If
End With
End Sub
Private Sub grid1_EnterCell()
If Not bEdit Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Form_Load()
cList = dbm.strList("SELECT * FROM SHIP WHERE STOPED = 0")

myload

CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "‘—ﬂ… «·‘Õ‰ „ÿ·Ê»…"
        Cancel = True
    End If
ElseIf col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "»Ê·Ì÷… «·‘Õ‰ „ÿ·Ê»…"
        Cancel = True
    End If
ElseIf col = 2 Then
    If Not IsDate(grid1.EditText) Then
        MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
        Cancel = True
    Else
        grid1.EditText = myFormat_p(grid1.EditText)
    End If
ElseIf col = 3 Then
    If Val(grid1.EditText) <= 0 Then
        Cancel = True
    End If
End If
End Sub
Private Sub fixGrd()
With grid1
.ColComboList(0) = cList
.ColWidth(0) = 1500
.ColWidth(1) = 3500
.ColWidth(2) = 1300
.ColWidth(3) = 1000
.ColWidth(4) = 3000
.ColHidden(.Cols - 1) = True


For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Function validRow(Row) As Boolean
With grid1
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If .TextMatrix(Row, 1) = "" Then Exit Function
If Not IsDate(.TextMatrix(Row, 2)) Then Exit Function
If .ValueMatrix(Row, 3) <= 0 Then Exit Function
End With
validRow = True
End Function
Private Sub myAddItem()
With grid1
    .AddItem ""
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If col < grid1.Cols - 3 Then
    grid1.Select Row, col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 1)
    grid1.ShowCell Row + 1, 0
Else
    grid1.Select Row, col
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
If KeyCode = 46 And bEdit Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("Õ–›?? Â· √‰  „Ê«›ﬁ", vbYesNo + vbDefaultButton2 + vbCritical, "Õ–›") = vbYes Then
            If IsNumeric(grid1.TextMatrix(grid1.Row, 0)) Then
                If Not dbm.Execute("Delete From FILE6_90C WHERE id =  " & MyParn(grid1.TextMatrix(grid1.Row, grid1.Cols - 1))) Then Exit Sub
            End If
            grid1.RemoveItem grid1.Row
        End If
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
myload
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If grid1.col = 0 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If col = 0 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
