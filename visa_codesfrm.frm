VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form visa_codes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«ﬂÊ«œ «·›Ì“"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13935
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   13935
   Begin VB.Frame Frame2 
      Height          =   1140
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   6120
      Width           =   4650
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1410
      End
      Begin Threed.SSCommand cmdGoChange 
         Height          =   330
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
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
         Caption         =   " ⁄œÌ· «·‰”»"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label Label3 
         Caption         =   " ⁄œÌ· „‰  «—ÌŒ "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   1260
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   -90
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6045
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   13785
      _cx             =   24315
      _cy             =   10663
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
      Cols            =   4
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   4770
      TabIndex        =   7
      Top             =   6120
      Width           =   9105
      Begin VB.CheckBox chkStoped 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«ŸÂ«— «·„ Êﬁ›…"
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
         Height          =   285
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   630
         Width           =   1545
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "visa_codesfrm.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1500
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   585
         Width           =   2490
      End
      Begin MSDataListLib.DataCombo xBank 
         Height          =   315
         Left            =   3780
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "»‰ﬂ «·›Ì“«"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8010
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   855
      End
      Begin VB.Label label2 
         Caption         =   "‰Ê⁄ «·›Ì“«"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8055
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   0
      Top             =   0
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
Attribute VB_Name = "visa_codes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Dim con As New adodb.Connection
Dim CLIST As String
Private Sub myload()
Dim cString As String
Dim cWhere As String
cString = "SELECT [CODE] AS [«·ﬂÊœ]" & _
        ",[DESCA] AS [«·»Ì«‰]" & _
        ",[BANK] AS [«·Ì‰ﬂ]" & _
        ",[RATE] AS [‰”»… «·»‰ﬂ]" & _
        ",[RATE_INSTALL] AS [‰”»… «· ﬁ”Ìÿ]" & _
        ",[WITH_POINTS] AS [«·‰ﬁ«ÿ]" & _
        ",[BANK_NAME] AS [«”„ «·»‰ﬂ]" & _
        ",[STOPED] AS [„ Êﬁ›]" & _
        " FROM [visa_codes]"

If chkStoped.Value = 0 Then
    cWhere = cWhere & Tr(cWhere) & "VISA_CODES.STOPED = 0"
End If

If Trim(xdesca.text) <> "" Then
    cWhere = cWhere & Tr(cWhere) & "desca like " & MyParn(xdesca.text & "%")
End If

If xBank.MatchedWithList Then
    cWhere = cWhere & Tr(cWhere) & "bank = " & MyParn(xBank.BoundText)
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

cString = cString & " order by CODE"
Set data1.Recordset = cmd(cString, con).Execute
myAddItem
Fixgrd
End Sub

Private Sub chkStoped_Click()
myload
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdGoChange_Click()
If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— „”Ã·"
    Exit Sub
End If

'If DateDiff("D", myFormat(xDate.text), Date) > 10 Then
'    MsgBox "⁄œœ «·«Ì«„ «ﬂ»— „‰ 10 «Ì«„"
'    Exit Sub
'End If

If MsgBox(" ⁄œÌ· «·‰”»", vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub

Dim cString As String
Dim nRecords As Long
cString = "UPDATE FILE6_20_VISA SET " & _
          " FILE6_20_VISA.RATE_VISA = VISA_CODES.RATE" & _
          " FROM FILE6_20_VISA " & _
          " INNER JOIN VISA_CODES ON FILE6_20_VISA.VISA_CODE = VISA_CODES.CODE" & _
          " INNER JOIN FILE6_20H ON FILE6_20_VISA.DOC_NO = FILE6_20H.DOC_NO" & _
          " WHERE FILE6_20H.DATE >= " & DateSq(xDate.text)

con.BeginTrans
On Error GoTo myError
con.Execute cString, nRecords
con.CommitTrans

MsgBox " „  ⁄œÌ· " & nRecords & " ”Ã·"
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
If TransCount(con) > 0 Then
    con.RollbackTrans
End If
myload
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Set visa_cods = Nothing
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
With grid1
Dim sCode As String
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
   myAddItem
End If

Dim aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(.TextMatrix(Row, 1)))
aInsert = AddFlag(aInsert, "BANK", addstring(.TextMatrix(Row, 2)))
aInsert = AddFlag(aInsert, "RATE", .ValueMatrix(Row, 3))
aInsert = AddFlag(aInsert, "RATE_INSTALL", .ValueMatrix(Row, 4))
aInsert = AddFlag(aInsert, "WITH_POINTS", .ValueMatrix(Row, 5))
aInsert = AddFlag(aInsert, "BANK_NAME", addstring(.TextMatrix(Row, 6)))
aInsert = AddFlag(aInsert, "STOPED", IIf(.ValueMatrix(Row, 7) = 0, "0", "1"))
On Error GoTo myError
If .TextMatrix(Row, 0) = "" Then
    aInsert = AddFlag(aInsert, "CODE", myField("select coalesce(max(code),0) + 1 from VISA_CODES", con, , , 1))
    con.Execute addInsert(aInsert, "VISA_CODES")
    myload
    grid1.ShowCell grid1.Rows - 1, 0
Else
    con.Execute addUpdate(aInsert, "VISA_CODES", "code = " & .TextMatrix(Row, 0))
End If
End With
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
myload
End Sub
Private Sub grid1_EnterCell()
If (grid1.col = 0) Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo myError
If KeyCode = 46 Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("Õ–›?? Â· √‰  „Ê«›ﬁ", vbYesNo + vbCritical, "Õ–›") = vbYes Then
            If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
                con.BeginTrans
                con.Execute "Delete From VISA_CODES where code =  " & MyParn(grid1.TextMatrix(grid1.Row, 0))
                con.CommitTrans
            End If
            grid1.RemoveItem grid1.Row
        End If
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
myload
End Sub
Private Sub Form_Load()
openCon con

CLIST = StrList(listBank, con)

Set DATA2.Recordset = cmd(listBank, con).Execute
Set xBank.RowSource = DATA2
xBank.ListField = "Desca"
xBank.BoundColumn = "Code"

Set grid1.DataSource = data1
With grid1
myload
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End With
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "»Ì«‰ «·›Ì“« „ÿ·Ê»"
        Cancel = True
    End If
ElseIf col = 2 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "»‰ﬂ «·›Ì“« „ÿ·Ê»"
        Cancel = True
    End If
End If
End Sub
Private Sub Fixgrd()
With grid1
.RowHeight(0) = 600
.ColWidth(0) = 700
.ColWidth(1) = 3300
.ColWidth(2) = 3300
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 2000
.ColWidth(7) = 800
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ColComboList(2) = CLIST
.Cell(flexcpBackColor, 1, 0, .Rows - 1, 0) = &H8000000F
End With
End Sub
Private Sub xbank_Change()
If xBank.MatchedWithList Then
    myload
ElseIf xBank.BoundText = "" Then
    myload
End If
End Sub

Private Sub xDesca_Change()
myload
End Sub
Private Function validRow(Row As Long) As Boolean
With grid1
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 2)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub myAddItem()
With grid1
    .AddItem ""
    grid1.Cell(flexcpBackColor, 1, 0, grid1.Rows - 1, 0) = &H8000000F
    grid1.ShowCell grid1.Rows - 1, 0
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If col < grid1.Cols - 1 Then
    grid1.Select Row, col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.ShowCell Row + 1, 1
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 1)
Else
    grid1.Select Row, col
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If col = 2 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
     CellPos KeyCode, Row, col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid1.col = 2 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Function listBank() As String
listBank = "SELECT CODE,DESCA " & _
          " FROM FILE5_10 " & _
          " WHERE FILE5_10.IS_VISA = 1"

'Dim sTb As String
'sTb = "SELECT BANK FROM VISA_CODES"
'If chkStoped.Value = 0 Then
'    sTb = sTb & " WHERE STOPED = 0"
'End If
'
'listBank = listBank & sTb & ")"

cString = cString & " ORDER BY FILE5_10.DESCA"
End Function
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xdesca_GotFocus()
myGotFocus xdesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xdesca
End Sub
Private Sub xbank_GotFocus()
myGotFocus xBank
End Sub
Private Sub xbank_LostFocus()
myLostFocus xBank
If Not xBank.MatchedWithList Then
    xBank.BoundText = ""
    myload
End If
End Sub
