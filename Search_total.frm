VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Search_total 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈” ⁄·«„"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   13500
   Tag             =   "Factory"
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   13440
      TabIndex        =   4
      Top             =   8670
      Width           =   13500
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   7335
         ScaleHeight     =   465
         ScaleWidth      =   4380
         TabIndex        =   6
         Top             =   0
         Width           =   4380
         Begin Threed.SSCommand cmdFilter 
            Height          =   285
            Left            =   135
            TabIndex        =   10
            Top             =   90
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   196610
            ForeColor       =   4210752
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Filter"
            ButtonStyle     =   3
         End
         Begin VB.CheckBox xBegin 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Begin"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   285
            Left            =   3420
            TabIndex        =   9
            Top             =   90
            Width           =   915
         End
         Begin VB.CheckBox xEnter 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Enter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H002F2F2F&
            Height          =   285
            Left            =   2520
            TabIndex        =   8
            Top             =   90
            Width           =   825
         End
         Begin Threed.SSCommand cmdPrint 
            Height          =   285
            Left            =   1620
            TabIndex        =   11
            Top             =   90
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            _Version        =   196610
            ForeColor       =   4210752
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Print"
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExcel 
            Height          =   285
            Left            =   855
            TabIndex        =   12
            Top             =   90
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   196610
            ForeColor       =   32768
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Excel"
            ButtonStyle     =   3
         End
      End
      Begin VB.Label lblCount 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   45
         Width           =   2760
      End
   End
   Begin MSAdodcLib.Adodc Ado1 
      Height          =   330
      Left            =   300
      Top             =   1125
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Bindings        =   "Search_total.frx":0000
      Height          =   7845
      Left            =   90
      TabIndex        =   7
      Top             =   90
      Width           =   13335
      _cx             =   23521
      _cy             =   13838
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   300
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
      TabBehavior     =   0
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
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   45
      TabIndex        =   2
      Top             =   7965
      Width           =   5115
      Begin VB.TextBox txtlookup 
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
         Height          =   390
         Index           =   0
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Visible         =   0   'False
         Width           =   3165
      End
      Begin MSDataListLib.DataCombo cmbLookup 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
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
         Index           =   0
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Visible         =   0   'False
         Width           =   660
      End
   End
End
Attribute VB_Name = "Search_total"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sid As String, aFormat As Variant, bTotal As Boolean, bPrint As Boolean
Public aRowExcel As Long
Public sCaption As String
Public aFilter As Variant
Public sFlag As String, bEnter As Boolean
Dim Generalarray, listarray, GrdArray
Dim cString As String, bAtChange As Boolean
Dim aValue As Variant
Dim con As New ADODB.Connection
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdGo_Click()
fillgrd
fixGrd
If grid1.Rows > 1 Then
    grid1.SetFocus
End If
End Sub
Private Sub cmbLookup_Change(Index As Integer)
If cmbLookup(Index).MatchedWithList Or Trim(cmbLookup(Index).BoundText) = "" Then
    fillgrd
    If listarray(Index - 1, 5) = "LAST_CLICKED" Then
        addSetting "SEARCH_LIST" & Index, cmbLookup(Index).BoundText, TempSave(Generalarray(0))
    End If
End If
End Sub
Private Sub cmbLookup_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
End If
'If KeyCode = 46 Then cmbLookup(Index).BoundText = ""
End Sub
Private Sub cmbLookup_LostFocus(Index As Integer)
On Error Resume Next
If Index = UBound(listarray) + 1 Then
    grid1.SetFocus
    If grid1.Rows > 1 Then grid1.Row = 11
End If
Err.Clear
Exit Sub
End Sub

Private Sub cmbLookup_Validate(Index As Integer, Cancel As Boolean)
If Not cmbLookup(Index).MatchedWithList Then
    cmbLookup(Index).BoundText = ""
    fillgrd
End If
End Sub
Private Sub cmdExcel_Click()
Me.MousePointer = 11
If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 19)
    aRow = AddFlag(aRow, aSub)
End If
ToFileExelNew grid1, , , aRow, , 1, , , , , , Me, Array(Me.Caption)
Me.MousePointer = 0
End Sub

Private Sub cmdFilter_Click()
If grid1.Rows > 1001 Then
    MsgBox "⁄œœ «·”ÿÊ— «ﬂ»— „‰ 1000"
    Exit Sub
End If
Dim cString As String, nCol As Long, sType As String, cFilter As String
nCol = Val(retFlag(aFilter, "col"))
sType = IIf(retFlag(aField, "type") = "", "s", retFlag(aField, "type"))
For i = 1 To grid1.Rows - 1
    cFilter = cFilter & IIf(cFilter = "", "", ",") & IIf(sType = "s", MyParn(grid1.TextMatrix(i, nCol)), grid1.TextMatrix(i, nCol))
Next
If cFilter <> "" Then
    cFilter = retFlag(aFilter, "field") & " IN(" & cFilter & ")"
End If
Generalarray(0).myproc2 cFilter
End Sub
Private Sub cmdPrint_Click()
Generalarray(0).myPrint sFlag
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
'    If UBound(Generalarray) < 5 Then
        Unload Me
'    Else
'        If Generalarray(5) Then Me.Hide
'        If Not Generalarray(5) Then Unload Me
'    End If
End If
End Sub
Private Sub Form_Activate()
On Error Resume Next
'If txtlookup(1).Visible Then txtlookup(1).SetFocus
For i = 0 To UBound(listarray)
    If listarray(i, 3) <> "" And listarray(i, 2) = "" Then
        txtlookup(i + 1).text = listarray(i, 3)
    End If
Next
Err.Clear
'If Generalarray(4) Then fillgrd
End Sub
Private Sub Form_Load()
'FileName = App.Path & "\SKINS\winaqua.skn"
'If FileName <> "" Then
'    Skin1.LoadSkin FileName ' Loads another skin into Skin component
'    Skin1.ApplySkin Me.hWnd ' Applies the skin to this window and its child controls
'End If
Me.Caption = sCaption

openCon con

Set grid1.DataSource = Ado1
'StatusBar1.Panels(1).Width = 2500
cmdPrint.Visible = bPrint
grid1.ExplorerBar = flexExSort

'Ado1.CommandType = adCmdText

Generalarray = searchArray(0)
listarray = searchArray(1)
GrdArray = searchArray(2)

Frame2.Width = Generalarray(3)
grid1.Cols = UBound(GrdArray) + 1
LoadControls
SetValue

xEnter.Visible = Not bEnter
xEnter.Value = IIf(retFlag(aValue, "enter") Or bEnter, 1, 0)
xBegin.Value = IIf(retFlag(aValue, "begin"), 1, 0)

If UBound(Generalarray) = 3 Then
    fillgrd
Else
    If UBound(Generalarray) >= 4 Then
        If Not Generalarray(4) Then fillgrd Else fixGrd
    End If
End If

cmdFilter.Visible = retFlag(aFilter, "filter")
Handlecontrols
Picture2.Left = Me.Width - Picture2.Width - 100
End Sub

Private Sub Form_Resize()
'Picture2.Width = 4155
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set Search_total = Nothing
End Sub

Private Sub grid1_DblClick()
If grid1.Row > IIf(bTotal, 1, 0) Then
    If sFlag = "" Then Generalarray(0).myProc Else Generalarray(0).myProc sFlag
End If
End Sub
Private Sub grid1_GotFocus()
For i = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(i) Then Exit For
Next
If grid1.Rows > 1 Then grid1.Select 1, i
End Sub
Sub fillgrd()
On Error GoTo myerror
cString = Generalarray(1)
For i = 0 To UBound(listarray)
   If listarray(i, 4) = "" Then
        If txtlookup(i + 1).text <> "" Then
            cCond = Replace(listarray(i, 1), "%cFilter%", FixString(txtlookup(i + 1).text))
            cCond = FixMulti(cCond, txtlookup(i + 1).text)
            cCond = FixValue(cCond, txtlookup(i + 1).text)
            cCond = FixZero(cCond, txtlookup(i + 1).text)
            If UBound(listarray, 2) <= 4 Then cCond = FixDate(cCond, txtlookup(i + 1).text, "=") Else cCond = FixDate(cCond, txtlookup(i + 1).text, IIf(IsEmpty(listarray(i, 5)), "=", listarray(i, 5)))
            cCond = Replace(cCond, "cFilter", txtlookup(i + 1).text)
            cString = cString & Space(1) & turn(cString) & Space(1) & cCond
        End If
    Else
       If cmbLookup(i + 1).BoundText <> "" Then
            cCond = Replace(listarray(i, 1), "cFilter", cmbLookup(i + 1).BoundText)
            cString = cString & Space(1) & turn(cString) & Space(1) & cCond
        End If
    End If
Next
cString = cString & Space(1) & Generalarray(2)
Set Ado1.Recordset = cmd(cString, con).Execute
fixGrd
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Handlecontrols()
End Sub
Private Sub fixGrd()
For i = 0 To grid1.Cols - 1
    grid1.TextMatrix(0, i) = GrdArray(i, 0)
    grid1.ColWidth(i) = GrdArray(i, 1)
    If Not retFlag(aFormat, "left") Then grid1.ColAlignment(i) = flexAlignRightCenter
    nwidth = nwidth + grid1.ColWidth(i)
    If UBound(GrdArray, 2) = 2 Then
        If GrdArray(i, 2) = "d" Then grid1.ColDataType(i) = flexDTDate
    End If
    If GrdArray(i, 3) Then
        bTotal = True
        grid1.SubtotalPosition = flexSTAbove
        grid1.Subtotal flexSTSum, -1, i, GrdArray(i, 4), &HC0FFC0, vbBlack, True, "≈Ã„«·Ï"
    End If
Next
grid1.Width = nwidth + 400
Me.Width = grid1.Width + 400
If retFlag(aFormat, "left") = True Then
    grid1.RightToLeft = False
    grid1.FontSize = 10
End If
'Label2.Caption = IIf(Grid1.Rows = 1, "·«  ÊÃœ ”Ã·« ", "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & Grid1.Rows - 1)
lblCount.Caption = IIf(grid1.Rows < 3, "·«  ÊÃœ ”Ã·« ", "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.Rows - 2)
End Sub
Private Sub LoadControls()
nVSpace = 420
nFrame = Frame2.Height
For i = 0 To UBound(listarray)
    nRow = nRow + 1
    Frame2.Height = nFrame + (nVSpace * (nRow - 1))
    If listarray(i, 4) = "" Then
        Load txtlookup(nRow)
        txtlookup(nRow).Visible = True
        txtlookup(nRow).Top = txtlookup(0).Top + (nVSpace * (nRow - 1))
        txtlookup(nRow).TabIndex = i
    Else
        Load cmbLookup(nRow)
        cmbLookup(nRow).Visible = True
        cmbLookup(nRow).Top = cmbLookup(0).Top + (nVSpace * (nRow - 1))
        Load data2(nRow)
        Set data2(nRow).Recordset = cmd(listarray(i, 2) & "", con).Execute
        Set cmbLookup(nRow).RowSource = data2(nRow)
        cmbLookup(nRow).BoundColumn = listarray(i, 3)
        cmbLookup(nRow).ListField = listarray(i, 4)
        cmbLookup(nRow).TabIndex = i
        If listarray(i, 5) = "LAST_CLICKED" Then
            cmbLookup(nRow).BoundText = RetSetting("SEARCH_LIST" & nRow, TempSave(Generalarray(0)))
        Else
            cmbLookup(nRow).BoundText = listarray(i, 5)
        End If
        If Not cmbLookup(nRow).MatchedWithList Then cmbLookup(nRow).BoundText = ""
    End If
    
    Load Label1(nRow)
    Label1(nRow).Top = Label1(0).Top + (nVSpace * (nRow - 1))
    Label1(nRow).Caption = listarray(i, 0) & " :"
    lblWidth = IIf(lblWidth < Label1(nRow).Width, Label1(nRow).Width, lblWidth)
Next
If nRow >= 2 Then
    Me.Height = Me.Height + (nVSpace * (nRow - 1))
End If
For i = 1 To Label1.Count - 1
    If listarray(i - 1, 4) = "" Then
        txtlookup(i).Width = Frame2.Width - (lblWidth + 400)
        Label1(i).Left = txtlookup(i).Left + 100 + txtlookup(i).Width
        If listarray(i - 1, 2) <> "" Then txtlookup(i).text = listarray(i - 1, 2)
    Else
        cmbLookup(i).Width = Frame2.Width - (lblWidth + 400)
        Label1(i).Left = cmbLookup(i).Left + 100 + cmbLookup(i).Width
    End If
'    Load Label1(i)
    Label1(i).Caption = ArbString(Label1(i).Caption)
    Label1(i).Left = Label1(i).Left
    Label1(i).Top = Label1(i).Top
    Label1(i).Visible = True
Next
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    grid1_DblClick
End If
End Sub

Private Sub txtlookup_Change(Index As Integer)
If xEnter.Value = 0 Then fillgrd
End Sub
Private Sub txtlookup_GotFocus(Index As Integer)
txtlookup(Index).SelStart = 0
txtlookup(Index).SelLength = Len(txtlookup(Index).text)
End Sub
Private Function FixString(cString)
aString = Split(Trim(cString), " ")
For i = 0 To UBound(aString)
    If Trim(aString(i)) <> "" Then FixString = FixString & " " & Trim(aString(i))
Next
FixString = "%" & Replace(Trim(FixString), " ", "%") & "%"
End Function

Private Sub txtlookup_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 And xEnter.Value = 1 Then fillgrd
End Sub

Private Sub txtlookup_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    csource = Ado1.RecordSource
    If bEnter Then fillgrd
    If grid1.Rows = 2 And csource = Ado1.RecordSource Then
        KeyCode = 0
        grid1.Row = 1
        grid1_DblClick
    ElseIf grid1.Rows > 2 And Not bEnter Then
        grid1.SetFocus
    End If
ElseIf grid1.Rows > 1 And (KeyCode = 40 Or KeyCode = 38) Then
    grid1.SetFocus
End If
End Sub

Private Sub txtlookup_LostFocus(Index As Integer)
'On Error Resume Next
'If Index = UBound(listarray) + 1 Then
'    Grid1.SetFocus
'    Grid1.Row = 1
'End If
'Err.Clear
End Sub
Private Function FixMulti(ByVal cString, cSearch) As String
Dim nFound As Long, nFound2 As Long, aString As Variant, cField As String
FixMulti = cString
For i = 1 To Len(FixMulti)
    If xBegin.Value = 0 Then
        nFound = InStr(1, FixMulti, "%%")
        If nFound = 0 Then Exit Function
        nFound2 = InStr(nFound + 3, FixMulti, "%%")
        cField = Mid(FixMulti, nFound + 2, nFound2 - (nFound + 2))
        aString = Split(Trim(cSearch), " ")
        cString2 = ""
        For i2 = 0 To UBound(aString)
            If Trim(aString(i2)) <> "" Then cString2 = cString2 & IIf(cString2 = "", "", " and ") & cField & " Like " & "'%" & aString(i2) & "%'"
        Next
        FixMulti = Replace(FixMulti, "%%" & cField & "%%", "(" & cString2 & ")")
    Else
        nFound = InStr(1, FixMulti, "%%")
        If nFound = 0 Then Exit Function
        nFound2 = InStr(nFound + 3, FixMulti, "%%")
        cField = Mid(FixMulti, nFound + 2, nFound2 - (nFound + 2))
        FixMulti = Replace(FixMulti, "%%" & cField & "%%", "(" & cField & " Like " & MyParn(cSearch & "%") & ")")
    End If
Next
End Function
Private Function FixValue(ByVal cString, cSearch) As String
Dim cSign As String, nAfter As Integer
FixValue = cString
For i = 1 To Len(FixValue)
    nFound = InStr(1, FixValue, "**")
    If nFound = 0 Then Exit Function
    nFound2 = InStr(nFound + 3, FixValue, "**")
    nAfter = nFound2 + 1
    
    If nAfter < Len(cString) Then
        If Mid(cString, nAfter + 1, 1) = ">" Or Mid(cString, nAfter + 1, 1) = "<" Then cSign = Mid(cString, nAfter + 1, 1)
        If nAfter + 1 < Len(cString) Then If Mid(cString, nAfter + 2, 1) = "=" Then cSign = cSign & "="
    End If
    
    cField = Mid(FixValue, nFound + 2, nFound2 - (nFound + 2))

    If IsNumeric(cSearch) Then
        cReplace = cField & Space(1) & IIf(cSign = "", "=", cSign) & Space(1) & Val(cSearch)
    Else
         cReplace = "1 = 2"
    End If
    FixValue = Replace(FixValue, "**" & cField & "**" & cSign, "(" & cReplace & ")")
Next
End Function
Private Function FixZero(ByVal cString, cSearch) As String
Dim nAfter As Integer, nZero As Integer
FixZero = cString
For i = 1 To Len(FixZero)
    nFound = InStr(1, FixZero, "@@")

    If nFound = 0 Then Exit Function
    nFound2 = InStr(nFound + 3, FixZero, "@@")
    nAfter = nFound2 + 1

    If nAfter < Len(cString) Then
        nZero = Val(Mid(cString, nAfter + 1, 2))
    End If

    cField = Mid(FixZero, nFound + 2, nFound2 - (nFound + 2))

    cReplace = cField & Space(1) & " = " & Space(1) & MyParn(RetZero(cSearch, nZero))
    FixZero = Replace(FixZero, "@@" & cField & "@@" & nZero, "(" & cReplace & ")")
Next
End Function
Private Function FixDate(ByVal cString, cSearch, pSign) As String
FixDate = cString
For i = 1 To Len(FixDate)
    cString2 = ""
    nFound = InStr(1, FixDate, "##")
    If nFound = 0 Then Exit Function
    nFound2 = InStr(nFound + 3, FixDate, "##")
    cField = Mid(FixDate, nFound + 2, nFound2 - (nFound + 2))
    
    If IsDate(cSearch) Then
        If DateValue(cSearch) >= DateValue("01-01-1753") Then
            cString2 = cField & " " & pSign & " " & DateSq(Format(cSearch, "YYYY-MM-DD"))
        Else
            cString2 = "(1 = 3)"
        End If
    Else
        cString2 = "(1 = 3)"
    End If
    FixDate = Replace(FixDate, "##" & cField & "##", cString2)
Next
End Function
Private Function SetValue() As Variant
Dim cField As String
cField = "search_enter:" & Generalarray(0).Name
cField = cField & turn(sid, ":") & sid
aValue = AddFlag(Empty, "enter", RetSetting(cField, TempSave(Generalarray(0), sid)) = "TRUE")

cField = "search_begin:" & Generalarray(0).Name
cField = cField & turn(sid, ":") & sid
aValue = AddFlag(aValue, "begin", RetSetting(cField, TempSave(Generalarray(0), sid)) = "TRUE")
End Function
Private Sub xEnter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cField = "search_enter:" & Generalarray(0).Name
cField = cField & turn(sid, ":") & sid
addSetting cField, IIf(xEnter.Value = 1, "TRUE", "FALSE"), TempSave(Generalarray(0))
End Sub
Private Sub xBegin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cField = "search_begin:" & Generalarray(0).Name
cField = cField & turn(sid, ":") & sid
addSetting cField, IIf(xBegin.Value = 1, "TRUE", "FALSE"), TempSave(Generalarray(0))
End Sub

